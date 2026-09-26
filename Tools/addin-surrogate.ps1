# 卸载时只清理当前会话中承载本插件的 COM 代理进程。
$script:AddInSurrogateArgumentPattern = '(?:^|\s)' + `
    [regex]::Escape('/Processid:{441360A0-59D3-4969-9F91-7AF166C512BE}') + '(?=\s|$)'

function Get-AddInSurrogates {
    param([int] $SessionId)

    try {
        $processes = @(Get-CimInstance -ClassName Win32_Process -Filter "Name = 'dllhost.exe'" -ErrorAction Stop)
    }
    catch {
        throw "无法查询 dllhost.exe 进程，卸载已停止：$($_.Exception.Message)"
    }

    foreach ($process in $processes) {
        if ($process.Name -ne 'dllhost.exe' -or
            $null -eq $process.SessionId -or $process.SessionId -ne $SessionId) {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($process.CommandLine)) {
            throw "无法读取当前会话中 dllhost.exe PID $($process.ProcessId) 的命令行，卸载已停止。"
        }

        if ($process.CommandLine -match $script:AddInSurrogateArgumentPattern) {
            $process
        }
    }
}

function Stop-AddInSurrogates {
    param([int] $SessionId = -1)

    if ($SessionId -lt 0) {
        $SessionId = [Diagnostics.Process]::GetCurrentProcess().SessionId
    }

    $remaining = @(Get-AddInSurrogates -SessionId $SessionId)
    if ($remaining.Count -eq 0) {
        Write-Host '    插件代理进程没有在运行。'
        return
    }

    Write-Host '    等待插件代理进程自行退出（最多 5 秒）…'
    for ($i = 0; $i -lt 10; $i++) {
        Start-Sleep -Milliseconds 500
        $remaining = @(Get-AddInSurrogates -SessionId $SessionId)
        if ($remaining.Count -eq 0) {
            Write-Host '    插件代理进程已退出。'
            return
        }
    }

    Write-Host '    插件代理进程仍在运行，正在结束本插件专属的 dllhost.exe…'
    foreach ($candidate in $remaining) {
        $surrogateId = [int] $candidate.ProcessId
        # 等待期间进程可能已退出、PID 可能被复用；结束前重新核对命令行和会话。
        $current = @(Get-AddInSurrogates -SessionId $SessionId |
            Where-Object { $_.ProcessId -eq $surrogateId })
        if ($current.Count -eq 0) { continue }

        try {
            Stop-Process -Id $surrogateId -Force -ErrorAction Stop
        }
        catch {
            $stopError = $_
            $current = @(Get-AddInSurrogates -SessionId $SessionId |
                Where-Object { $_.ProcessId -eq $surrogateId })
            if ($current.Count -gt 0) {
                throw "无法结束插件代理进程 PID $surrogateId，卸载已停止：$($stopError.Exception.Message)"
            }
        }
    }

    for ($i = 0; $i -lt 10; $i++) {
        $remaining = @(Get-AddInSurrogates -SessionId $SessionId)
        if ($remaining.Count -eq 0) {
            Write-Host '    插件代理进程已退出。'
            return
        }
        Start-Sleep -Milliseconds 500
    }

    $remaining = @(Get-AddInSurrogates -SessionId $SessionId)
    if ($remaining.Count -gt 0) {
        $ids = ($remaining | ForEach-Object { $_.ProcessId }) -join ', '
        throw "插件代理进程仍在运行（PID: $ids），卸载已停止。"
    }
    Write-Host '    插件代理进程已退出。'
}
