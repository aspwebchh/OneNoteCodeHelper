<#
.SYNOPSIS
    发布 Windows x64 自包含 MCP 文件夹；不安装、不注册、不操作 OneNote。
#>
#requires -Version 5.1
[CmdletBinding()]
param([ValidateSet('Release', 'Debug')][string]$Configuration = 'Release')
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repoRoot 'Mcp\Server\OneNoteCodeHelper.Mcp.csproj'
& dotnet publish $project -c $Configuration -r win-x64 --self-contained true -p:PublishTrimmed=false -p:PublishSingleFile=false --nologo
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
$executable = Join-Path $repoRoot "Mcp\Server\bin\$Configuration\net10.0-windows\win-x64\publish\OneNoteCodeHelper.Mcp.exe"
if (-not (Test-Path -LiteralPath $executable)) { throw '发布完成但没有找到 MCP EXE。' }
Write-Host "MCP 自包含文件夹已发布：$executable"
