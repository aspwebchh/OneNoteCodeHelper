<#
.SYNOPSIS
    Agent 离线回归：格式保留、工具调用、流式协议、冲突和撤销；不调用真实 AI 或 OneNote。
#>
[CmdletBinding()]
param([string]$DllPath, [ValidateSet('Release','Debug')][string]$Configuration = 'Release')
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $DllPath) { $DllPath = Join-Path $repoRoot "bin\$Configuration\net48\OneNoteCodeHelper.dll" }
$DllPath = (Resolve-Path -LiteralPath $DllPath).Path
$project = Join-Path $repoRoot 'Tests\OneNoteCodeHelper.AgentTests.csproj'
& dotnet build $project -c $Configuration "-p:AgentDllPath=$DllPath" --nologo
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& (Join-Path $repoRoot "Tests\bin\$Configuration\net48\OneNoteCodeHelper.AgentTests.exe")
exit $LASTEXITCODE
