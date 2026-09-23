$ErrorActionPreference = 'Stop'

try {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class ComActivationProbe
{
    [DllImport("ole32.dll", PreserveSig = true)]
    public static extern int CoCreateInstance(
        ref Guid clsid, IntPtr outer, uint context, ref Guid iid, out IntPtr instance);
}
'@

    $clsid = [Guid] '441360A0-59D3-4969-9F91-7AF166C512BE'
    $iid = [Guid] 'B65AD801-ABAF-11D0-BB8B-00A0C90F2744'
    $instance = [IntPtr]::Zero
    $hr = [ComActivationProbe]::CoCreateInstance(
        [ref] $clsid, [IntPtr]::Zero, 4, [ref] $iid, [ref] $instance)

    if ($hr -ne 0 -or $instance -eq [IntPtr]::Zero) {
        throw ('OneNote 的本地 COM 激活失败：0x{0:X8}' -f ($hr -band 0xffffffff))
    }

    [void] [Runtime.InteropServices.Marshal]::Release($instance)
    exit 0
}
catch {
    Write-Error $_
    exit 1
}
