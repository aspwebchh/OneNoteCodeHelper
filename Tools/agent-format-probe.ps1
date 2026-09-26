<#
.SYNOPSIS
    显式创建独立 OneNote 测试分区，验证字体、局部样式、对齐和段间距往返。
.DESCRIPTION
    先运行 agent-test.ps1 构建测试程序。此探针不属于普通回归；仅编辑自己创建的测试页。
    测试页和 XML 留在 OutputDirectory，便于在 OneNote 内查看。不会调用 AI 接口。
#>
[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$OutputDirectory, [ValidateSet('Release','Debug')][string]$Configuration = 'Release')
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$program = Join-Path $repoRoot "Tests\bin\$Configuration\net48\OneNoteCodeHelper.AgentTests.exe"
if (-not (Test-Path -LiteralPath $program)) { throw '请先运行 Tools\agent-test.ps1。' }
& $program --probe ([IO.Path]::GetFullPath($OutputDirectory))
exit $LASTEXITCODE
