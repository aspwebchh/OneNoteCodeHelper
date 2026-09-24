<#
.SYNOPSIS
    卸载 OneNote 代码高亮外接程序。

.DESCRIPTION
    需要管理员权限：COM 类注册在 HKLM 下（原因见 register.ps1 的说明）。
    OneNote 的外接程序清单项在 HKCU，一并删掉。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Tools\unregister.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$ClassId = '{441360A0-59D3-4969-9F91-7AF166C512BE}'
$ProgId  = 'OneNoteCodeHelper.AddIn'

$isAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Error '需要管理员权限：COM 类注册在 HKLM 下。请用管理员身份重新运行（或用根目录的 uninstall.ps1，它会帮你提权）。'
    exit 1
}

if (Get-Process ONENOTE -ErrorAction SilentlyContinue) {
    Write-Warning 'OneNote 正在运行。卸载后需要重启 OneNote，按钮才会消失。'
}

$paths = @(
    # COM 类（HKLM）
    "HKLM:\SOFTWARE\Classes\CLSID\$ClassId",
    "HKLM:\SOFTWARE\Classes\AppID\$ClassId",
    "HKLM:\SOFTWARE\Classes\$ProgId",
    # 早期版本脚本写在 HKCU 下的 COM 类，顺手清掉
    "HKCU:\Software\Classes\CLSID\$ClassId",
    "HKCU:\Software\Classes\$ProgId",
    # OneNote 外接程序清单（HKCU）
    "HKCU:\Software\Microsoft\Office\OneNote\AddIns\$ProgId"
)

foreach ($path in $paths) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Recurse -Force
        Write-Host "已删除 $path"
    }
}

$resiliency = 'HKCU:\Software\Microsoft\Office\16.0\OneNote\Resiliency\DoNotDisableAddinList'
if (Test-Path $resiliency) {
    Remove-ItemProperty -Path $resiliency -Name $ProgId -ErrorAction SilentlyContinue
    Write-Host "已从 DoNotDisableAddinList 移除 $ProgId"
}

Write-Host ''
Write-Host '卸载完成。重启 OneNote 后「代码高亮」组会消失。' -ForegroundColor Green
Write-Host '设置与日志没有删除，分别在：'
Write-Host "  $env:APPDATA\OneNoteCodeHelper\settings.xml"
Write-Host "  $env:LOCALAPPDATA\OneNoteCodeHelper\log.txt"
