# 使用模拟进程验证卸载清理逻辑；不会操作真实进程或注册表。
$ErrorActionPreference = 'Stop'

foreach ($scriptPath in @(
    (Join-Path $PSScriptRoot '..\install.ps1'),
    (Join-Path $PSScriptRoot 'addin-surrogate.ps1'),
    $PSCommandPath
)) {
    $bytes = [IO.File]::ReadAllBytes($scriptPath)
    if ($bytes.Length -lt 3 -or $bytes[0] -ne 239 -or $bytes[1] -ne 187 -or $bytes[2] -ne 191) {
        throw "$scriptPath 缺少 UTF-8 BOM。"
    }

    $tokens = $null
    $parseErrors = $null
    [System.Management.Automation.Language.Parser]::ParseFile(
        $scriptPath, [ref] $tokens, [ref] $parseErrors) | Out-Null
    if ($parseErrors.Count -gt 0) {
        throw "$scriptPath 在 PowerShell 5.1 中解析失败：$($parseErrors[0].Message)"
    }
}

. (Join-Path $PSScriptRoot 'addin-surrogate.ps1')

$targetArgument = '/Processid:{441360A0-59D3-4969-9F91-7AF166C512BE}'

function New-FakeProcess {
    param([int] $Id, [int] $SessionId, [string] $Argument, [string] $Name = 'dllhost.exe')
    [pscustomobject]@{
        ProcessId   = $Id
        SessionId   = $SessionId
        Name        = $Name
        CommandLine = "C:\Windows\System32\DllHost.exe $Argument"
    }
}

function Reset-FakeProcesses {
    param([object[]] $Processes)
    $script:FakeProcesses = @($Processes)
    $script:QueryCount = 0
    $script:SleepCount = 0
    $script:StoppedIds = @()
    $script:ExitOnSleep = 0
    $script:ExitId = 0
    $script:QueryFails = $false
    $script:StopFails = $false
    $script:KeepAfterStop = $false
    $script:ReplaceAtQuery = 0
}

function Get-CimInstance {
    [CmdletBinding()]
    param([string] $ClassName, [string] $Filter)
    if ($ClassName -ne 'Win32_Process' -or $Filter -ne "Name = 'dllhost.exe'") {
        throw '查询范围不符合预期。'
    }
    $script:QueryCount++
    if ($script:QueryFails) { throw '模拟查询失败' }
    if ($script:ReplaceAtQuery -eq $script:QueryCount) {
        $script:FakeProcesses = @(New-FakeProcess -Id 101 -SessionId 1 -Argument '/Processid:{OTHER}')
    }
    $script:FakeProcesses
}

function Start-Sleep {
    param([int] $Milliseconds)
    if ($Milliseconds -ne 500) { throw '轮询间隔不符合预期。' }
    $script:SleepCount++
    if ($script:ExitOnSleep -eq $script:SleepCount) {
        $script:FakeProcesses = @($script:FakeProcesses |
            Where-Object { $_.ProcessId -ne $script:ExitId })
    }
}

function Stop-Process {
    [CmdletBinding()]
    param([int] $Id, [switch] $Force)
    if (-not $Force) { throw '结束代理进程时未指定 -Force。' }
    $script:StoppedIds += $Id
    if ($script:StopFails) { throw '模拟拒绝访问' }
    if (-not $script:KeepAfterStop) {
        $script:FakeProcesses = @($script:FakeProcesses | Where-Object { $_.ProcessId -ne $Id })
    }
}

function Assert-Equal {
    param($Expected, $Actual, [string] $Case)
    if ($Expected -ne $Actual) {
        throw "$Case：预期 $Expected，实际 $Actual。"
    }
}

function Assert-Throws {
    param([scriptblock] $Action, [string] $Pattern, [string] $Case)
    $message = $null
    try { & $Action } catch { $message = $_.Exception.Message }
    if (-not $message -or $message -notmatch $Pattern) {
        throw "$Case：预期错误包含 $Pattern，实际为 $message。"
    }
}

# 其他 AppID、其他会话、名称不符和参数后有额外字符都不能被结束。
Reset-FakeProcesses @(
    (New-FakeProcess -Id 102 -SessionId 1 -Argument '/Processid:{OTHER}'),
    (New-FakeProcess -Id 103 -SessionId 2 -Argument $targetArgument),
    (New-FakeProcess -Id 104 -SessionId 1 -Argument $targetArgument -Name 'other.exe'),
    (New-FakeProcess -Id 105 -SessionId 1 -Argument ($targetArgument + '-extra'))
)
Stop-AddInSurrogates -SessionId 1
Assert-Equal 0 $script:StoppedIds.Count '非目标进程'
Assert-Equal 0 $script:SleepCount '无目标进程'

# OneNote 即使已经退出，只要代理进程自行退出，就不应强制结束。
Reset-FakeProcesses @((New-FakeProcess -Id 101 -SessionId 1 -Argument $targetArgument))
$script:ExitOnSleep = 2
$script:ExitId = 101
Stop-AddInSurrogates -SessionId 1
Assert-Equal 0 $script:StoppedIds.Count '自行退出'
Assert-Equal 2 $script:SleepCount '自行退出等待'

Reset-FakeProcesses @((New-FakeProcess -Id 101 -SessionId 1 -Argument $targetArgument))
Stop-AddInSurrogates -SessionId 1
Assert-Equal 1 $script:StoppedIds.Count '残留代理进程'
Assert-Equal 101 $script:StoppedIds[0] '残留代理进程 PID'
Assert-Equal 10 $script:SleepCount '强制结束前等待'

# 等待期间 PID 被其他代理进程复用时，结束前的重新查询必须阻止误杀。
Reset-FakeProcesses @((New-FakeProcess -Id 101 -SessionId 1 -Argument $targetArgument))
$script:ReplaceAtQuery = 12
Stop-AddInSurrogates -SessionId 1
Assert-Equal 0 $script:StoppedIds.Count 'PID 复用'

Reset-FakeProcesses @((New-FakeProcess -Id 101 -SessionId 1 -Argument $targetArgument))
$script:StopFails = $true
Assert-Throws { Stop-AddInSurrogates -SessionId 1 } '模拟拒绝访问' '结束失败'
Assert-Equal 1 $script:StoppedIds.Count '结束失败尝试次数'

Reset-FakeProcesses @((New-FakeProcess -Id 101 -SessionId 1 -Argument $targetArgument))
$script:KeepAfterStop = $true
Assert-Throws { Stop-AddInSurrogates -SessionId 1 } '仍在运行' '结束后仍残留'

Reset-FakeProcesses @()
$script:QueryFails = $true
Assert-Throws { Stop-AddInSurrogates -SessionId 1 } '模拟查询失败' '查询失败'
Assert-Equal 0 $script:StoppedIds.Count '查询失败'

$unknown = New-FakeProcess -Id 106 -SessionId 1 -Argument $targetArgument
$unknown.CommandLine = $null
Reset-FakeProcesses @($unknown)
Assert-Throws { Stop-AddInSurrogates -SessionId 1 } '无法读取.*命令行' '命令行不可读'
Assert-Equal 0 $script:StoppedIds.Count '命令行不可读'

Write-Host '插件代理进程清理测试通过。' -ForegroundColor Green
