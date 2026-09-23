<#
.SYNOPSIS
    一键安装 / 卸载 OneNote 代码高亮外接程序。

.DESCRIPTION
    把「关闭 OneNote -> 构建 -> 写注册表 -> 重新启动 OneNote」串成一步。
    COM 类注册需要管理员权限（会自动提权）；OneNote 的外接程序清单项写在 HKCU。

    为什么必须关 OneNote：运行中的加载项会让代理进程锁住 DLL，
    而且 OneNote 只在启动时读一次外接程序注册表。

.PARAMETER Configuration
    构建并注册哪个配置，默认 Release。

.PARAMETER SkipBuild
    跳过构建，直接注册已有的输出。

.PARAMETER Uninstall
    卸载而不是安装。

.PARAMETER Force
    OneNote 没能正常退出时强制结束进程。默认只礼貌地请求关闭，
    失败就停下来让你自己处理，避免丢掉没保存的内容。

.PARAMETER NoRestart
    完成后不自动重新启动 OneNote。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File install.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File install.ps1 -Configuration Debug

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall
#>
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Release',

    [switch] $SkipBuild,
    [switch] $Uninstall,
    [switch] $Force,
    [switch] $NoRestart
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------
# 注册 COM 类必须写 HKLM（mscoree 的 DllGetClassObject 不读 HKCU 下的 CLSID，
# 写在那儿激活会以 0x80070002 失败），所以这一步需要管理员。
# 没提权就带着原来的参数，以管理员身份重新启动自己。
# 新窗口加 -NoExit 留着不关，否则 UAC 起的窗口做完就消失，你看不到结果。
# ---------------------------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host ''
    Write-Host '  需要管理员权限：COM 类必须注册到 HKLM。' -ForegroundColor Yellow
    Write-Host '  正在以管理员身份重新启动，请在 UAC 提示里确认…'

    $scriptArg = '"' + $PSCommandPath + '"'
    $argList = @('-NoExit', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptArg)
    foreach ($kv in $PSBoundParameters.GetEnumerator()) {
        if ($kv.Value -is [switch]) {
            if ($kv.Value.IsPresent) { $argList += "-$($kv.Key)" }
        }
        else {
            $argList += @("-$($kv.Key)", '"' + $kv.Value + '"')
        }
    }

    try {
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $argList
    }
    catch {
        Write-Error '提权被取消或失败。请自己开一个管理员身份的 PowerShell 再运行本脚本。'
        exit 1
    }

    exit
}

function Write-Step {
    param([string] $Text)
    Write-Host ''
    Write-Host "==> $Text" -ForegroundColor Cyan
}

function Write-Ok {
    param([string] $Text)
    Write-Host "    $Text" -ForegroundColor Green
}

# ---------------------------------------------------------------
# 关闭 OneNote，返回它原来的 exe 路径（供之后重启用），没在运行就返回 $null
# ---------------------------------------------------------------
function Stop-OneNote {
    $running = @(Get-Process ONENOTE -ErrorAction SilentlyContinue)
    if ($running.Count -eq 0) {
        Write-Ok 'OneNote 没有在运行。'
        return $null
    }

    $exePath = ($running | Where-Object { $_.Path } | Select-Object -First 1).Path
    Write-Host '    正在请求 OneNote 退出…'
    foreach ($p in $running) {
        try { [void] $p.CloseMainWindow() } catch { }
    }

    # 最多等 20 秒
    for ($i = 0; $i -lt 40; $i++) {
        Start-Sleep -Milliseconds 500
        if (-not (Get-Process ONENOTE -ErrorAction SilentlyContinue)) { break }
    }

    $stubborn = @(Get-Process ONENOTE -ErrorAction SilentlyContinue)
    if ($stubborn.Count -gt 0) {
        if (-not $Force) {
            Write-Host ''
            Write-Warning 'OneNote 没有退出，可能有未保存的内容或正等待你确认某个对话框。'
            Write-Warning '请手动关闭 OneNote 后重新运行本脚本；确定可以强制结束就加 -Force。'
            exit 1
        }

        Write-Host '    -Force：强制结束 OneNote 进程。'
        $stubborn | Stop-Process -Force
        Start-Sleep -Seconds 1
    }

    Write-Ok 'OneNote 已退出。'
    return $exePath
}

# ---------------------------------------------------------------
# 找出 ONENOTE.EXE 的位置。用于 OneNote 本来就没在运行、
# 但装完之后仍然该把它打开的情况（比如上一次安装中途失败后重跑）。
# ---------------------------------------------------------------
function Get-OneNotePath {
    $appPaths = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\ONENOTE.EXE'
    if (Test-Path $appPaths) {
        $registered = (Get-ItemProperty $appPaths).'(default)'
        if ($registered -and (Test-Path $registered)) {
            return $registered
        }
    }

    $candidates = @()
    if ($env:ProgramFiles) {
        $candidates += (Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\ONENOTE.EXE')
    }
    $pf86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
    if ($pf86) {
        $candidates += (Join-Path $pf86 'Microsoft Office\root\Office16\ONENOTE.EXE')
    }

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    return $null
}

function Start-OneNote {
    param(
        [string] $ExePath,
        # 装完之后即使 OneNote 本来没开着也把它打开；卸载时不这么做。
        [switch] $EvenIfWasClosed
    )

    if ($NoRestart) {
        Write-Host '    -NoRestart：不自动启动 OneNote。'
        return $false
    }

    if (-not $ExePath -and $EvenIfWasClosed) {
        $ExePath = Get-OneNotePath
    }

    if (-not $ExePath) {
        return $false
    }

    Write-Step '启动 OneNote'
    Start-Process -FilePath $ExePath
    Write-Ok '已启动。'
    return $true
}

# ---------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------
Push-Location $PSScriptRoot
try {
    $title = if ($Uninstall) { '卸载 OneNote 代码高亮' } else { '安装 OneNote 代码高亮' }
    Write-Host ''
    Write-Host "  $title" -ForegroundColor White
    Write-Host '  ----------------------------------------'

    $outputDll = Join-Path $PSScriptRoot "bin\$Configuration\net48\OneNoteCodeHelper.dll"

    # --- 1. 先关闭宿主，释放代理进程对 DLL 的占用 ---
    Write-Step '关闭 OneNote'
    $oneNotePath = Stop-OneNote

    # --- 2. 构建 ---
    if (-not $Uninstall -and -not $SkipBuild) {
        Write-Step "构建（$Configuration）"

        $dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
        if (-not $dotnet) {
            Write-Error '找不到 dotnet 命令。请先安装 .NET SDK，或用 -SkipBuild 跳过构建直接注册。'
            exit 1
        }

        # 构建前确认输出文件没被运行中的代理进程占用。
        if (Test-Path $outputDll) {
            try {
                $stream = [IO.File]::Open($outputDll, 'Open', 'ReadWrite', 'None')
                $stream.Close()
            }
            catch {
                Write-Host ''
                Write-Warning "输出文件被其他进程占用，构建一定会失败："
                Write-Warning "  $outputDll"
                Write-Warning '最常见的原因：OneNote 或插件的 dllhost.exe 代理进程还在运行。关掉 OneNote 后再试。'
                Write-Warning '也可以加 -SkipBuild 跳过构建，直接用已有的输出注册。'
                exit 1
            }
        }

        & dotnet build -c $Configuration --nologo -v quiet
        if ($LASTEXITCODE -ne 0) {
            Write-Error "构建失败（退出码 $LASTEXITCODE）。"
            exit 1
        }

        Write-Ok '构建成功。'
    }

    # --- 3. 注册 / 注销 ---
    if ($Uninstall) {
        Write-Step '删除注册表项'
        & (Join-Path $PSScriptRoot 'Tools\unregister.ps1')
    }
    else {
        Write-Step '写入 COM 注册（HKLM）和 OneNote 加载项配置（HKCU）'
        & (Join-Path $PSScriptRoot 'Tools\register.ps1') -Configuration $Configuration
    }

    if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) {
        Write-Error '注册步骤失败，详见上面的输出。'
        exit 1
    }

    # --- 4. 启动 OneNote ---
    $logPath = Join-Path $env:LOCALAPPDATA 'OneNoteCodeHelper\log.txt'
    $logLinesBefore = if (Test-Path $logPath) { @(Get-Content $logPath).Count } else { 0 }
    $started = Start-OneNote -ExePath $oneNotePath -EvenIfWasClosed:(-not $Uninstall)

    if (-not $Uninstall -and $started) {
        Write-Step '确认 OneNote 已加载插件'
        $addInKey = 'HKCU:\Software\Microsoft\Office\OneNote\AddIns\OneNoteCodeHelper.AddIn'
        $loaded = $false
        $behavior = $null
        $maxTries = 60
        $recoveryWarned = $false
        for ($i = 0; $i -lt $maxTries; $i++) {
            Start-Sleep -Milliseconds 500
            $behavior = (Get-ItemProperty -Path $addInKey -ErrorAction SilentlyContinue).LoadBehavior
            if ($behavior -eq 2) { break }

            $oneNote = Get-Process ONENOTE -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $recoveryWarned -and $oneNote.MainWindowTitle -like '*您希望如何启动 OneNote*') {
                Write-Warning 'OneNote 正在等待恢复启动选择。请在 OneNote 窗口点击“正常启动”；安装程序会继续等待。'
                $recoveryWarned = $true
                $maxTries = 240
            }

            if (Test-Path $logPath) {
                $newLines = @(Get-Content $logPath | Select-Object -Skip $logLinesBefore)
                $connected = @($newLines | Where-Object { $_ -match '已连接到 OneNote' }).Count -gt 0
                $ribbonLoaded = @($newLines | Where-Object { $_ -match '功能区已加载' }).Count -gt 0
                $loaded = $connected -and $ribbonLoaded
            }
            if ($loaded) { break }
        }

        if (-not $loaded) {
            Write-Error "OneNote 没有成功加载插件（LoadBehavior=$behavior）。请查看 $logPath 和 OneNote 的 COM 加载项对话框。"
            exit 1
        }
        Write-Ok '已确认 OneNote 调用了插件并加载功能区。'
    }

    # --- 5. 收尾提示 ---
    Write-Host ''
    if ($Uninstall) {
        Write-Host '  卸载完成。' -ForegroundColor Green
        Write-Host '  设置与日志没有删除，需要的话自己清理：'
        Write-Host "    $env:APPDATA\OneNoteCodeHelper\settings.xml"
        Write-Host "    $env:LOCALAPPDATA\OneNoteCodeHelper\log.txt"
    }
    else {
        Write-Host '  安装完成。' -ForegroundColor Green
        Write-Host '  在 OneNote 的「开始」选项卡末尾找「代码高亮」组。'
        Write-Host ''
        Write-Host '  没看到按钮的话按顺序查：'
        Write-Host "    1. 日志 $env:LOCALAPPDATA\OneNoteCodeHelper\log.txt"
        Write-Host '       连「OnConnection 开始」都没有 = 没被加载；有异常堆栈 = 代码问题'
        Write-Host '    2. OneNote 的「文件 / 选项 / 加载项」里本项的状态'
        Write-Host '    3. 注册表 HKCU\Software\Microsoft\Office\16.0\OneNote\Resiliency\DisabledItems'
    }
    Write-Host ''
}
finally {
    Pop-Location
}
