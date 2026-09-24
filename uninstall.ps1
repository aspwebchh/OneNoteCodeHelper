<#
.SYNOPSIS
    卸载 OneNote 代码高亮外接程序。

.DESCRIPTION
    关闭 OneNote，注销 COM 类和 OneNote 加载项，并在需要时重新启动 OneNote。
    复用 install.ps1 的卸载流程；脚本会自动请求管理员权限。

.PARAMETER Force
    OneNote 没能正常退出时强制结束进程。

.PARAMETER NoRestart
    卸载完成后不自动重新启动 OneNote。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File uninstall.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File uninstall.ps1 -Force -NoRestart
#>
[CmdletBinding()]
param(
    [switch] $Force,
    [switch] $NoRestart
)

$ErrorActionPreference = 'Stop'
$global:LASTEXITCODE = 0

& (Join-Path $PSScriptRoot 'install.ps1') -Uninstall -Force:$Force -NoRestart:$NoRestart
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
