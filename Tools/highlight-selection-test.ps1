<#
.SYNOPSIS
    「高亮选中」的回归测试：从页面 XML 里认出选中的代码、拼成源码、换成代码框。

.DESCRIPTION
    不碰 OneNote，只通过反射调用构建出来的 DLL 里 RenderDiagnostics.HighlightSelection，普通权限即可运行。
    重点是缩进：在 OneNote 里敲代码时行首按 Tab 会把段落挂到上一段下面，选区里的段落父节点各不相同，
    这种选区要能高亮，缩进换成制表符，没选中的下级段落不能跟着删掉。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Tools\highlight-selection-test.ps1
#>
[CmdletBinding()]
param(
    [ValidateSet('Release', 'Debug')]
    [string]$Configuration = 'Release',

    # 直接指定要测的 DLL，优先于 -Configuration
    [string]$DllPath
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
$diagnostics = $assembly.GetType('OneNoteCodeHelper.Services.RenderDiagnostics', $true)

function Invoke-Diag([string]$name, [object[]]$arguments) {
    try {
        $diagnostics.GetMethod($name, $flags).Invoke($null, $arguments)
    }
    catch [System.Reflection.TargetInvocationException] {
        throw $_.Exception.InnerException
    }
}

$script:total = 0
$script:failed = 0

function Assert-Equal([string]$name, $actual, $expected) {
    $script:total++
    if ($actual -ceq $expected) {
        Write-Host "  ok    $name"
        return
    }

    $script:failed++
    Write-Host "  FAIL  $name" -ForegroundColor Red
    Write-Host "        期望：$expected" -ForegroundColor Red
    Write-Host "        实际：$actual" -ForegroundColor Red
}

# 拼一个页面，每个参数是一个文本框（one:Outline）里的段落。-Title 是标题里的段落
function New-Page([string[]]$outlines, [string]$Title) {
    $head = if ($Title) { "<one:Title><one:OE><one:T selected=`"all`"><![CDATA[$Title]]></one:T></one:OE></one:Title>" } else { '' }
    $i = 0
    $body = ($outlines | ForEach-Object {
        $i++
        "<one:Outline objectID=`"{O$i}`"><one:OEChildren>$_</one:OEChildren></one:Outline>"
    }) -join ''
    "<one:Page xmlns:one=`"http://schemas.microsoft.com/office/onenote/2013/onenote`">$head$body</one:Page>"
}

# 一个段落。-Sel 表示拖选时选中了（OneNote 只把 one:T 标成 all），-Children 是缩进的下级段落
function New-Line([string]$text, [switch]$Sel, [string]$Children) {
    $mark = if ($Sel) { ' selected="all"' } else { '' }
    $sub = if ($Children) { "<one:OEChildren>$Children</one:OEChildren>" } else { '' }
    "<one:OE><one:T$mark><![CDATA[$text]]></one:T>$sub</one:OE>"
}

function New-Table([string[]]$cells) {
    $body = ($cells | ForEach-Object { "<one:Cell><one:OEChildren>$_</one:OEChildren></one:Cell>" }) -join ''
    "<one:OE><one:Table><one:Row>$body</one:Row></one:Table></one:OE>"
}

function Test-Highlight([string]$name, [string]$page, [string]$expected) {
    Assert-Equal $name (Invoke-Diag 'HighlightSelection' @($page)) $expected
}

$A = New-Line 'A'
$B = New-Line 'B'

Write-Host '高亮选中：'

Test-Highlight '平铺的几行，选中中间两行' `
    (New-Page @("$A$(New-Line 's1' -Sel)$(New-Line 's2' -Sel)$B")) 's1|s2 => A|#|B'

$method = New-Line 'void f() {' -Sel -Children (New-Line 'return;' -Sel)
$class = New-Line 'class X {' -Sel -Children "$method$(New-Line '}' -Sel)"
Test-Highlight '行首按 Tab 缩进的代码：每层补一个制表符，整块换掉' `
    (New-Page @("$A$class$(New-Line '}' -Sel)$B")) `
    'class X {|\tvoid f() {|\t\treturn;|\t}|} => A|#|B'

Test-Highlight '选区从缩进里开始：代码框放在外层，删空的 OEChildren 去掉' `
    (New-Page @("$(New-Line 'if (a) {' -Children (New-Line 'c1' -Sel))$(New-Line '}' -Sel)$B")) `
    '\tc1|} => if (a) {|#|B'

Test-Highlight '选区在缩进块中间结束：没选中的下级段落留在代码框后面' `
    (New-Page @("$A$(New-Line 'if (a) {' -Sel -Children "$(New-Line 'c1' -Sel)$(New-Line 'c2')")$B")) `
    'if (a) {|\tc1 => A|#|c2|B'

Test-Highlight '只选了缩进的几行：按最浅的一层对齐，代码框留在缩进里' `
    (New-Page @("$(New-Line 'if (a) {' -Children "$(New-Line 'c1' -Sel -Children (New-Line 'c2' -Sel))$(New-Line 'c3')")")) `
    'c1|\tc2 => if (a) {|>#|>c3'

Test-Highlight '缩进里的空行不补制表符' `
    (New-Page @("$(New-Line 'p' -Sel -Children "$(New-Line '' -Sel)$(New-Line 'c' -Sel)")")) `
    'p||\tc => #'

Test-Highlight '表格单元格里的缩进' `
    (New-Page @("$A$(New-Table @("$(New-Line 'x' -Sel -Children (New-Line 'y' -Sel))"))")) `
    'x|\ty => #'

Test-Highlight '同时选了表格内外' `
    (New-Page @("$(New-Line 's' -Sel)$(New-Table @("$(New-Line 'x' -Sel)"))")) `
    'FAIL: 选中的内容跨越了不同的区块（比如同时选了表格内外）。请只选中同一块里的代码。'

Test-Highlight '同时选了两个文本框' `
    (New-Page @("$(New-Line 's1' -Sel)", "$(New-Line 's2' -Sel)")) `
    'FAIL: 选中的内容跨越了不同的区块（比如同时选了表格内外）。请只选中同一块里的代码。'

Test-Highlight '标题里的文字' (New-Page @($A) -Title 't') 'FAIL: 选中的内容不在一个可编辑的区块里，无法替换。'

Test-Highlight '没选中' (New-Page @("$A$B")) 'FAIL: 没有检测到选中的文本。请先在页面上选中要高亮的代码，再点「高亮选中」。'

Write-Host ''
if ($script:failed -eq 0) {
    Write-Host "全部通过：$($script:total) / $($script:total)" -ForegroundColor Green
    exit 0
}

Write-Host ("通过 {0} / {1}，失败 {2}" -f ($script:total - $script:failed), $script:total, $script:failed) -ForegroundColor Red
exit 1
