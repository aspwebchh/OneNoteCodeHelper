<#
.SYNOPSIS
    自动识别语言的回归测试。

.DESCRIPTION
    对 Tools\detect-samples 下的每个样本跑 LanguageRegistry.Detect。子目录名就是期望结果，
    none 目录里放的是应当识别不出（返回 null）的内容。认错或认不出时列出各语言得分，调权重就看它。
    最后跑一个性能用例：大段空行夹着代码，识别必须在 1 秒内完成。

    不碰 OneNote，只通过反射调用构建出来的 DLL，普通权限即可运行。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Tools\detect-test.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Tools\detect-test.ps1 -Configuration Debug -ShowScores
#>
[CmdletBinding()]
param(
    [ValidateSet('Release', 'Debug')]
    [string]$Configuration = 'Release',

    # 直接指定要测的 DLL，优先于 -Configuration
    [string]$DllPath,

    # 通过的样本也列出各语言得分
    [switch]$ShowScores
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
if (-not $DllPath) {
    $DllPath = Join-Path $root "bin\$Configuration\net48\OneNoteCodeHelper.dll"
}

if (-not (Test-Path $DllPath)) {
    Write-Error "找不到 $DllPath，请先构建。"
    exit 1
}

# 按字节加载，不锁 bin 下的 DLL，免得影响下一次构建
$assembly = [System.Reflection.Assembly]::Load([System.IO.File]::ReadAllBytes((Resolve-Path $DllPath)))
$flags = [System.Reflection.BindingFlags]'NonPublic, Public, Static'
$detect = $assembly.GetType('OneNoteCodeHelper.Highlighting.LanguageRegistry', $true).GetMethod('Detect', $flags)
$describe = $assembly.GetType('OneNoteCodeHelper.Services.RenderDiagnostics', $true).GetMethod('DescribeDetectionScores', $flags)

function Get-DetectedId([string]$code) {
    $language = $detect.Invoke($null, @($code))
    if ($language) { $language.Id } else { $null }
}

function Get-Scores([string]$code) {
    if ($describe) { $describe.Invoke($null, @($code)) } else { '' }
}

$samplesDir = Join-Path $PSScriptRoot 'detect-samples'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$total = 0
$failed = 0
$corpus = New-Object System.Text.StringBuilder

foreach ($dir in Get-ChildItem $samplesDir -Directory | Sort-Object Name) {
    $expected = if ($dir.Name -eq 'none') { $null } else { $dir.Name }

    foreach ($file in Get-ChildItem $dir.FullName -File -Filter *.txt | Sort-Object Name) {
        $total++
        $code = [System.IO.File]::ReadAllText($file.FullName, $utf8)
        [void]$corpus.Append($code).Append("`n" * 400)

        $actual = Get-DetectedId $code
        $name = '{0}/{1}' -f $dir.Name, $file.BaseName

        if ($actual -eq $expected) {
            if ($ShowScores) {
                Write-Host ('  ok    {0,-32} {1}' -f $name, (Get-Scores $code))
            }
            continue
        }

        $failed++
        $shown = if ($actual) { $actual } else { '(无法确定)' }
        Write-Host ('  FAIL  {0,-32} -> {1,-12} {2}' -f $name, $shown, (Get-Scores $code)) -ForegroundColor Red
    }
}

# 性能：LikelihoodPatterns 里记过的病态输入是成片的空行。把全部样本拼起来、每段后面跟几百个空行，
# 各种语言的特征都会大量命中，是最坏情况。先跑一次预热，免得把正则的构造时间算进去。
$stress = $corpus.ToString()
while ($stress.Length -lt 64 * 1024) {
    $stress += $corpus.ToString()
}

[void](Get-DetectedId $stress)
$watch = [System.Diagnostics.Stopwatch]::StartNew()
[void](Get-DetectedId $stress)
$watch.Stop()

$total++
$elapsed = $watch.ElapsedMilliseconds
if ($elapsed -gt 1000) {
    $failed++
    Write-Host ("  FAIL  性能：{0:N0} 字符的混合输入识别用了 {1} ms，超过 1000 ms" -f $stress.Length, $elapsed) -ForegroundColor Red
}
else {
    Write-Host ("  ok    性能：{0:N0} 字符的混合输入识别用了 {1} ms" -f $stress.Length, $elapsed)
}

Write-Host ''
if ($failed -eq 0) {
    Write-Host "全部通过：$total / $total" -ForegroundColor Green
    exit 0
}

Write-Host ("通过 {0} / {1}，失败 {2}" -f ($total - $failed), $total, $failed) -ForegroundColor Red
exit 1
