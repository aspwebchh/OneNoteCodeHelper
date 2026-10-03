<#
.SYNOPSIS
    MCP 离线回归：模拟页面、隔离命名管道和真实 stdio 子进程；不调用真实 AI 或 OneNote。
#>
#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$McpExe,
    [ValidateSet('Release', 'Debug')][string]$Configuration = 'Release'
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $McpExe) { $McpExe = Join-Path $repoRoot "Mcp\Server\bin\$Configuration\net10.0-windows\OneNoteCodeHelper.Mcp.exe" }
$McpExe = (Resolve-Path -LiteralPath $McpExe).Path
$dll = Join-Path $repoRoot "bin\$Configuration\net48\OneNoteCodeHelper.dll"
$project = Join-Path $repoRoot 'Tests\OneNoteCodeHelper.AgentTests.csproj'
& dotnet build $project -c $Configuration "-p:AgentDllPath=$dll" --nologo
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& (Join-Path $repoRoot "Tests\bin\$Configuration\net48\OneNoteCodeHelper.AgentTests.exe") --mcp-only $McpExe
exit $LASTEXITCODE
