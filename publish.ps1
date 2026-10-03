<#
.SYNOPSIS
    一键发布插件和 Windows x64 自包含 MCP，生成发布目录、ZIP 及 SHA256 校验文件。
.DESCRIPTION
    只构建和打包，不执行安装、COM 注册或 OneNote 操作。不覆盖以前的发布目录。
    发布包保留原安装脚本所需的 bin 和 Mcp 目录结构；目标机安装时使用 install.ps1 -SkipBuild。
.PARAMETER Configuration
    构建配置，默认 Release。
.PARAMETER OutputDirectory
    发布包的父目录，默认仓库的 bin\publish。相对路径按调用脚本时的工作目录解释。
    仓库内的输出必须位于 bin 下，避免自包含运行时被 net48 的文件扫描误收。
.PARAMETER SkipBuild
    使用已有插件构建和 MCP 自包含发布输出，不执行 dotnet 命令。
.PARAMETER NoZip
    只生成发布目录及目录内校验文件，不压缩 ZIP。
.PARAMETER PluginDirectory
    配合 SkipBuild 使用已验证的独立插件输出目录，避免正在运行的插件占用默认 DLL。
.PARAMETER McpDirectory
    配合 SkipBuild 使用已验证的 Windows x64 自包含 MCP 输出目录。
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File publish.ps1
.EXAMPLE
    pwsh -File publish.ps1 -OutputDirectory D:\Releases
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File publish.ps1 -SkipBuild -NoZip
#>
#requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('Release', 'Debug')][string]$Configuration = 'Release',
    [string]$OutputDirectory,
    [string]$PluginDirectory,
    [string]$McpDirectory,
    [switch]$SkipBuild,
    [switch]$NoZip
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$repoRoot = $PSScriptRoot

function Assert-File {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "缺少发布文件：$Path" }
}

function Copy-ReleaseFile {
    param([string]$Source, [string]$Destination)
    Assert-File $Source
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Destination))
    Copy-Item -LiteralPath $Source -Destination $Destination
}

function Assert-McpOutput {
    param([string]$Directory)
    foreach ($name in @('OneNoteCodeHelper.Mcp.exe', 'OneNoteCodeHelper.Mcp.dll',
        'OneNoteCodeHelper.Mcp.deps.json', 'OneNoteCodeHelper.Mcp.runtimeconfig.json', 'coreclr.dll', 'hostfxr.dll', 'hostpolicy.dll')) {
        Assert-File (Join-Path $Directory $name)
    }
    $config = Get-Content -LiteralPath (Join-Path $Directory 'OneNoteCodeHelper.Mcp.runtimeconfig.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $config.runtimeOptions.PSObject.Properties['includedFrameworks']) { throw 'MCP 输出没有自包含运行时，请先运行 Tools\publish-mcp.ps1。' }
    $deps = Get-Content -LiteralPath (Join-Path $Directory 'OneNoteCodeHelper.Mcp.deps.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($deps.runtimeTarget.name -notmatch '/win-x64$') { throw 'MCP 输出不是 Windows x64 发布。' }
}

try {
    if (($PluginDirectory -or $McpDirectory) -and -not $SkipBuild) { throw '自定义构建目录只允许与 -SkipBuild 一起使用。' }
    if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repoRoot 'bin\publish' }
    $outputRoot = [IO.Path]::GetFullPath($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory))
    $repoBoundary = $repoRoot.TrimEnd([char[]]'\/') + [IO.Path]::DirectorySeparatorChar
    $binBoundary = (Join-Path $repoRoot 'bin').TrimEnd([char[]]'\/') + [IO.Path]::DirectorySeparatorChar
    $outputBoundary = $outputRoot.TrimEnd([char[]]'\/') + [IO.Path]::DirectorySeparatorChar
    if ($outputBoundary.StartsWith($repoBoundary, [StringComparison]::OrdinalIgnoreCase) -and
        -not $outputBoundary.StartsWith($binBoundary, [StringComparison]::OrdinalIgnoreCase)) {
        throw '仓库内的发布输出必须位于 bin 目录下；也可以指定仓库外的目录。'
    }

    if (-not $SkipBuild) {
        [void](Get-Command dotnet -CommandType Application -ErrorAction Stop)
        Write-Host "构建解决方案（$Configuration）…"
        Push-Location -LiteralPath $repoRoot
        try {
            & dotnet build (Join-Path $repoRoot 'OneNoteCodeHelper.sln') -c $Configuration --nologo
            if ($LASTEXITCODE -ne 0) { throw "解决方案构建失败，退出码 $LASTEXITCODE。" }
        }
        finally { Pop-Location }
    }

    $pluginSource = Join-Path $repoRoot "bin\$Configuration\net48"
    if ($PluginDirectory) { $pluginSource = (Resolve-Path -LiteralPath $PluginDirectory).Path }
    $pluginFiles = @('OneNoteCodeHelper.dll', 'Extensibility.dll', 'Microsoft.Office.Interop.OneNote.dll')
    foreach ($name in $pluginFiles) { Assert-File (Join-Path $pluginSource $name) }
    $identity = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $pluginSource 'OneNoteCodeHelper.dll'))
    if ($identity.GetPublicKeyToken().Length -eq 0) { throw '插件 DLL 没有强名称签名，不能发布。' }

    $stamp = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmssfff')
    $packageName = "OneNoteCodeHelper-$($identity.Version)-win-x64-$stamp-$([Guid]::NewGuid().ToString('N').Substring(0, 8))"
    $packagePath = [IO.Path]::GetFullPath((Join-Path $outputRoot $packageName))
    if (-not $packagePath.StartsWith($outputBoundary, [StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $packagePath)) {
        throw '发布目录无效或已存在，停止打包。'
    }
    [void][IO.Directory]::CreateDirectory($packagePath)
    $pluginTarget = Join-Path $packagePath "bin\$Configuration\net48"
    foreach ($name in $pluginFiles) { Copy-ReleaseFile (Join-Path $pluginSource $name) (Join-Path $pluginTarget $name) }
    if (Test-Path -LiteralPath (Join-Path $pluginSource 'OneNoteCodeHelper.pdb')) {
        Copy-ReleaseFile (Join-Path $pluginSource 'OneNoteCodeHelper.pdb') (Join-Path $pluginTarget 'OneNoteCodeHelper.pdb')
    }

    $mcpRelative = "Mcp\Server\bin\$Configuration\net10.0-windows\win-x64\publish"
    $mcpTarget = Join-Path $packagePath $mcpRelative
    if ($SkipBuild) {
        $mcpSource = Join-Path $repoRoot $mcpRelative
        if ($McpDirectory) { $mcpSource = (Resolve-Path -LiteralPath $McpDirectory).Path }
        Assert-McpOutput $mcpSource
        [void][IO.Directory]::CreateDirectory($mcpTarget)
        foreach ($file in Get-ChildItem -LiteralPath $mcpSource -Recurse -File) {
            # 只复制运行时、程序集和发布元数据，不携带个人设置或日志。
            if ($file.Extension -notin @('.dll', '.exe', '.pdb') -and $file.Name -notin @('OneNoteCodeHelper.Mcp.deps.json', 'OneNoteCodeHelper.Mcp.runtimeconfig.json')) { continue }
            $relative = $file.FullName.Substring($mcpSource.Length + 1)
            Copy-ReleaseFile $file.FullName (Join-Path $mcpTarget $relative)
        }
    }
    else {
        Write-Host '发布 MCP（Windows x64，自带运行时）…'
        & (Join-Path $repoRoot 'Tools\publish-mcp.ps1') -Configuration $Configuration -OutputDirectory $mcpTarget
        if ($LASTEXITCODE -ne 0) { throw "MCP 发布失败，退出码 $LASTEXITCODE。" }
    }
    Assert-McpOutput $mcpTarget

    foreach ($relative in @('install.ps1', 'uninstall.ps1', 'Tools\register.ps1', 'Tools\unregister.ps1',
        'Tools\addin-surrogate.ps1', 'Tools\probe-com.ps1', 'README.md', 'docs\onenote-mcp.md')) {
        Copy-ReleaseFile (Join-Path $repoRoot $relative) (Join-Path $packagePath $relative)
    }

    $instructions = @"
# OneNoteCodeHelper 发布包

构建配置：$Configuration；插件版本：$($identity.Version)；MCP：Windows x64 自包含。

解压到固定目录后，在发布包根目录执行：

    powershell -ExecutionPolicy Bypass -File install.ps1 -Configuration $Configuration -SkipBuild

安装会请求管理员权限、关闭并重新打开 OneNote。目标机需要 OneNote 桌面版 Office16 和 .NET Framework 4.8，不需要 .NET 10 SDK 或运行时。
安装后不要移动发布目录，因为 COM 注册绑定其中的 DLL 路径；迁移目录后需要重新注册。

打开 OneNote 并加载插件后，执行：

    & '.\$mcpRelative\OneNoteCodeHelper.Mcp.exe' --doctor

Codex 配置示例及调用流程见 docs/onenote-mcp.md；把示例 EXE 路径改成当前解压目录中的实际路径。
卸载使用根目录 uninstall.ps1。SHA256SUMS.txt 列出发布文件的 SHA256 校验值。
本包不包含源码、签名私钥、个人 AI 配置或日志。
"@
    [IO.File]::WriteAllText((Join-Path $packagePath 'INSTALL.md'), $instructions, [Text.UTF8Encoding]::new($true))
    $manifest = [ordered]@{
        format_version = 1
        created_utc = [DateTime]::UtcNow.ToString('o')
        configuration = $Configuration
        plugin_assembly_version = $identity.Version.ToString()
        plugin_path = "bin/$Configuration/net48/OneNoteCodeHelper.dll"
        mcp_path = ($mcpRelative + '\OneNoteCodeHelper.Mcp.exe').Replace('\', '/')
        runtime_identifier = 'win-x64'
        self_contained = $true
    }
    [IO.File]::WriteAllText((Join-Path $packagePath 'release-manifest.json'), ($manifest | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    $prefix = $packagePath + [IO.Path]::DirectorySeparatorChar
    $hashes = @(Get-ChildItem -LiteralPath $packagePath -Recurse -File | Sort-Object FullName | ForEach-Object {
        $relative = $_.FullName.Substring($prefix.Length).Replace('\', '/')
        (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant() + '  ' + $relative
    })
    [IO.File]::WriteAllLines((Join-Path $packagePath 'SHA256SUMS.txt'), [string[]]$hashes, [Text.UTF8Encoding]::new($false))

    if (-not $NoZip) {
        $zipPath = $packagePath + '.zip'
        Write-Host '压缩发布包…'
        Compress-Archive -LiteralPath $packagePath -DestinationPath $zipPath -CompressionLevel Optimal
        $zipHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
        [IO.File]::WriteAllText(($zipPath + '.sha256'), ($zipHash + '  ' + [IO.Path]::GetFileName($zipPath) + [Environment]::NewLine), [Text.UTF8Encoding]::new($false))
        Write-Host "ZIP：$zipPath"
    }
    Write-Host "发布目录：$packagePath"
    Write-Host "安装说明：$(Join-Path $packagePath 'INSTALL.md')"
}
catch {
    Write-Error $_ -ErrorAction Continue
    exit 1
}
