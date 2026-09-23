<#
.SYNOPSIS
    把 OneNote 代码高亮外接程序注册到本机。

.DESCRIPTION
    需要管理员权限。

    为什么必须要管理员：mscoree.dll 的 DllGetClassObject 只读 HKCR 的机器部分
    （也就是 HKLM\SOFTWARE\Classes），不读每用户部分。COM 类写在 HKCU 里它看不见，
    激活会以 0x80070002 失败，OneNote 随即把 LoadBehavior 降成 2 —— 表现就是
    「装好了但功能区里什么都没有」。本机实测：同一个 CLSID 写 HKCU 失败、写 HKLM 成功。
    RegAsm 只写 HKLM 也是同一个原因。

    所以分两处写：
      COM 类（CLSID / ProgID）      -> HKLM，需要管理员
      OneNote 外接程序清单与容错项  -> HKCU，每用户

.PARAMETER Configuration
    要注册哪个构建配置的输出，默认 Release。

.PARAMETER DllPath
    直接指定 DLL 路径，给上后忽略 Configuration。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Tools\register.ps1
    powershell -ExecutionPolicy Bypass -File Tools\register.ps1 -Configuration Debug
#>
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Release',
    [string] $DllPath
)

$ErrorActionPreference = 'Stop'

$ClassId  = '{441360A0-59D3-4969-9F91-7AF166C512BE}'
$ProgId   = 'OneNoteCodeHelper.AddIn'
$AsmName  = 'OneNoteCodeHelper'
$Friendly = 'OneNote 代码高亮'
$Descr    = '代码语法高亮（Java、Lua、PowerShell、Bat、Bash、XML、HTML、CSS）'

# ---------- 1. 必须是管理员 ----------
$isAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Error '需要管理员权限：COM 类必须注册到 HKLM，mscoree 不读 HKCU 下的 CLSID。请用管理员身份重新运行（或直接用根目录的 install.ps1，它会帮你提权）。'
    exit 1
}

# ---------- 2. 找到 DLL，读出真实标识 ----------
if (-not $DllPath) {
    $projectRoot = Split-Path -Parent $PSScriptRoot
    $DllPath = Join-Path $projectRoot "bin\$Configuration\net48\$AsmName.dll"
}

if (-not (Test-Path $DllPath)) {
    Write-Error "找不到 $DllPath。请先构建项目：dotnet build -c $Configuration"
    exit 1
}

$DllPath = (Resolve-Path $DllPath).Path
# 用 String.Replace 而不是 -replace：后者第一个参数是正则，单个反斜杠是非法模式。
$codeBase = 'file:///' + $DllPath.Replace('\', '/')

# 从 DLL 里读真实的程序集标识，不要硬编码：版本或密钥一变，硬编码的值就会让 CLR
# 找不到程序集，而且失败时只报一个含糊的 FileNotFoundException。
$asmName = [Reflection.AssemblyName]::GetAssemblyName($DllPath)
$AsmVer  = $asmName.Version.ToString()

# 强名称是 CodeBase 生效的前提：Fusion 只对强名称程序集认「指向应用程序目录之外的 codeBase」。
$token = $asmName.GetPublicKeyToken()
if (-not $token -or $token.Length -eq 0) {
    Write-Error '程序集没有强名称签名，注册了也加载不起来。请确认 csproj 里有 SignAssembly / AssemblyOriginatorKeyFile 后重新构建。'
    exit 1
}

# ---------- 3. OneNote 最好先关掉 ----------
# 运行中的 OneNote 会锁住 DLL（之后重新构建会失败），而且它只在启动时读一次
# 外接程序注册表，不重启等于白注册。
if (Get-Process ONENOTE -ErrorAction SilentlyContinue) {
    Write-Warning 'OneNote 正在运行。注册后必须完全退出并重新启动 OneNote 才会生效。'
}

# ---------- 4. 写注册表 ----------
function Set-Key {
    param([string] $Path, [hashtable] $Values, [string] $DefaultValue)

    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }

    if ($PSBoundParameters.ContainsKey('DefaultValue')) {
        Set-ItemProperty -Path $Path -Name '(default)' -Value $DefaultValue
    }

    if ($Values) {
        foreach ($name in $Values.Keys) {
            $value = $Values[$name]
            $type = if ($value -is [int]) { 'DWord' } else { 'String' }
            Set-ItemProperty -Path $Path -Name $name -Value $value -Type $type
        }
    }
}

# 4a. 清掉早期版本脚本写在 HKCU 下的 COM 类注册。
#     HKCU\Software\Classes 在 HKCR 合并视图里优先级高于 HKLM，留着会盖住正确的那份。
foreach ($stale in @("HKCU:\Software\Classes\CLSID\$ClassId", "HKCU:\Software\Classes\$ProgId")) {
    if (Test-Path $stale) {
        Remove-Item $stale -Recurse -Force
        Write-Host "  已清理旧的每用户注册：$stale"
    }
}

# 4b. COM 类注册 -> HKLM。OneNote 以 CLSCTX_LOCAL_SERVER 激活加载项，
#     Windows 的系统 DLL 代理进程按 AppID/DllSurrogate 加载 mscoree.dll。
$clsidRoot = "HKLM:\SOFTWARE\Classes\CLSID\$ClassId"

Set-Key -Path $clsidRoot -DefaultValue $ProgId -Values @{ 'AppID' = $ClassId }
Set-Key -Path "HKLM:\SOFTWARE\Classes\AppID\$ClassId" -Values @{ 'DllSurrogate' = '' }

$inproc = @{
    'ThreadingModel' = 'Both'
    'Class'          = $ProgId
    'Assembly'       = $asmName.FullName
    'RuntimeVersion' = 'v4.0.30319'
    'CodeBase'       = $codeBase
}
Set-Key -Path "$clsidRoot\InprocServer32" -Values $inproc -DefaultValue 'mscoree.dll'

# 同样的内容再写一份到版本号子键，这是 RegAsm 的标准做法。
Set-Key -Path "$clsidRoot\InprocServer32\$AsmVer" -Values $inproc

Set-Key -Path "$clsidRoot\ProgId" -DefaultValue $ProgId
Set-Key -Path "$clsidRoot\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}"

Set-Key -Path "HKLM:\SOFTWARE\Classes\$ProgId" -DefaultValue $Friendly
Set-Key -Path "HKLM:\SOFTWARE\Classes\$ProgId\CLSID" -DefaultValue $ClassId

# 4c. OneNote 的外接程序清单 -> HKCU。LoadBehavior=3 表示启动时加载。
Set-Key -Path "HKCU:\Software\Microsoft\Office\OneNote\AddIns\$ProgId" -Values @{
    'FriendlyName' = $Friendly
    'Description'  = $Descr
    'LoadBehavior' = 3
    'CommandLineSafe' = 1
}

# 4d. 开发期保险：只要抛一次异常，Office 就会把插件丢进「已禁用项目」再也不加载。
#     列进 DoNotDisableAddinList 可以避免一次失败就永久失效。
Set-Key -Path 'HKCU:\Software\Microsoft\Office\16.0\OneNote\Resiliency\DoNotDisableAddinList' -Values @{
    $ProgId = 1
}

# 4e. DisabledItems 可能包含其他加载项，不能整项清空。
$disabled = 'HKCU:\Software\Microsoft\Office\16.0\OneNote\Resiliency\DisabledItems'
if (Test-Path $disabled) {
    $names = (Get-Item $disabled).GetValueNames()
    if ($names.Count -gt 0) {
        Write-Warning 'OneNote 的“已禁用项目”里有记录。若本插件被列出，请在 OneNote 选项中单独启用。'
    }
}

# ---------- 5. 自检：OneNote 使用的本地服务器激活能否成功 ----------
# 普通 New-Object -ComObject 会允许进程内激活，无法发现缺少 DllSurrogate 的故障。
$probe = Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -Wait -PassThru `
    -ArgumentList @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
                    ('"' + (Join-Path $PSScriptRoot 'probe-com.ps1') + '"'))
$activated = ($probe.ExitCode -eq 0)

# ---------- 6. 报告 ----------
Write-Host ''
if ($activated) {
    Write-Host '注册完成，COM 自检通过。' -ForegroundColor Green
} else {
    Write-Host '注册表已写入，但 COM 自检没通过，OneNote 无法加载插件。' -ForegroundColor Red
}
Write-Host "  程序集 : $DllPath"
Write-Host "  标识   : $($asmName.FullName)"
Write-Host "  CLSID  : $ClassId  (HKLM)"
Write-Host "  ProgId : $ProgId"
Write-Host ''
Write-Host '接下来：完全退出 OneNote（任务管理器里确认 ONENOTE.EXE 已结束），再重新打开。'
Write-Host '「开始」选项卡末尾应出现「代码高亮」组。'
Write-Host "出问题先看日志：$env:LOCALAPPDATA\OneNoteCodeHelper\log.txt"

if (-not $activated) {
    exit 1
}
