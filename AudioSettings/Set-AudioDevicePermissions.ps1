<#
.SYNOPSIS
    Takes ownership of MMDevices Audio registry keys and grants SYSTEM full control.
.NOTES
    Author:  andrew.webber@ultima.com
    Version: 1.0.0
    Run at system startup (elevated/SYSTEM) before Set-AudioDeviceSettings.ps1 fires.
    Only needs to run once — ownership persists until AudioEndpointBuilder recreates the keys.
    Log:     C:\_source\logs\AudioRegistryPermissions_yyyyMMdd.log  (daily rolling)
#>

# ============================================================
# CONFIGURATION
# ============================================================

$renderPath  = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render"
$capturePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture"
$logDir      = "C:\_source\logs"
$logFile     = Join-Path $logDir "AudioRegistryPermissions_$(Get-Date -Format 'yyyyMMdd').log"

# ============================================================
# P/INVOKE — PRIVILEGE ESCALATION
#
# The MMDevices registry keys are owned by NT SERVICE\AudioEndpointBuilder.
# Even NT AUTHORITY\SYSTEM cannot write them without first taking ownership.
# PowerShell has no built-in way to enable SE_PRIVILEGE_ENABLED on a token,
# so we use P/Invoke to call the Win32 AdjustTokenPrivileges API directly.
#
# Privileges required:
#   SeTakeOwnershipPrivilege — allows taking ownership of objects we don't own
#   SeRestorePrivilege       — allows writing to registry keys we own after taking ownership
# ============================================================

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class TokenPrivilege {
    [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
    internal static extern bool AdjustTokenPrivileges(IntPtr htok, bool disall, ref TokPriv1Luid newst, int len, IntPtr prev, IntPtr relen);

    [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
    internal static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr phtok);

    [DllImport("advapi32.dll", SetLastError = true)]
    internal static extern bool LookupPrivilegeValue(string host, string name, ref long pluid);

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    internal struct TokPriv1Luid {
        public int Count;
        public long Luid;
        public int Attr;
    }

    internal const int SE_PRIVILEGE_ENABLED = 0x00000002;
    internal const int TOKEN_QUERY          = 0x00000008;
    internal const int TOKEN_ADJUST_PRIVS   = 0x00000020;

    public static bool Enable(long processHandle, string privilege) {
        TokPriv1Luid tp;
        IntPtr hproc = new IntPtr(processHandle);
        IntPtr htok  = IntPtr.Zero;
        OpenProcessToken(hproc, TOKEN_ADJUST_PRIVS | TOKEN_QUERY, ref htok);
        tp.Count = 1;
        tp.Luid  = 0;
        tp.Attr  = SE_PRIVILEGE_ENABLED;
        LookupPrivilegeValue(null, privilege, ref tp.Luid);
        return AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
    }
}
"@

# ============================================================
# FUNCTIONS
# ============================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Add-Content -Path $logFile -Value $entry
    Write-Host $entry
}

function Enable-Privileges {
    # Enable the two privileges needed to take ownership of AudioEndpointBuilder-owned keys.
    # Must be done before any registry ownership operations.
    $handle = [System.Diagnostics.Process]::GetCurrentProcess().Handle
    [TokenPrivilege]::Enable($handle, "SeTakeOwnershipPrivilege") | Out-Null
    [TokenPrivilege]::Enable($handle, "SeRestorePrivilege")        | Out-Null
}

function Set-RegistryKeyOwnership {
    param([string]$PSPath)

    # Takes ownership of the given key and grants SYSTEM full control with
    # ContainerInherit + ObjectInherit, so all child device subkeys are covered
    # without needing to repeat the operation per device.
    try {
        $system = [System.Security.Principal.NTAccount]"NT AUTHORITY\SYSTEM"
        $subKey = $PSPath -replace "^HKLM:\\", ""

        # Step 1: Take ownership (requires SeTakeOwnershipPrivilege)
        $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey(
            $subKey,
            [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree,
            [System.Security.AccessControl.RegistryRights]::TakeOwnership
        )
        $acl = $key.GetAccessControl([System.Security.AccessControl.AccessControlSections]::None)
        $acl.SetOwner($system)
        $key.SetAccessControl($acl)
        $key.Close()

        # Step 2: Now we own it, grant SYSTEM full control with inheritance
        $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey(
            $subKey,
            [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree,
            [System.Security.AccessControl.RegistryRights]::ChangePermissions
        )
        $acl  = $key.GetAccessControl()
        $rule = New-Object System.Security.AccessControl.RegistryAccessRule(
            "NT AUTHORITY\SYSTEM",
            "FullControl",
            "ContainerInherit,ObjectInherit",
            "None",
            "Allow"
        )
        $acl.AddAccessRule($rule)
        $key.SetAccessControl($acl)
        $key.Close()

        Write-Log "Ownership set on: $PSPath"
    }
    catch {
        Write-Log "FAIL Could not set ownership on $PSPath -- $($_.Exception.Message)" -Level "WARN"
    }
}

# ============================================================
# MAIN
# ============================================================

# Mutex prevents concurrent execution if WMI fires multiple events per device insertion.
# WaitOne(0) = non-blocking acquire. If another instance holds the mutex, exit immediately.
$mutexName = "Global\AudioDevicePermissions"
$mutex     = New-Object System.Threading.Mutex($false, $mutexName)

if (-not $mutex.WaitOne(0)) {
    exit 0
}

try {
    Write-Log "========================================"
    Write-Log "Audio registry permissions script started"
    Write-Log "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log "========================================"

    Enable-Privileges
    Write-Log "Privileges enabled: SeTakeOwnershipPrivilege, SeRestorePrivilege"

    # Take ownership of the Render and Capture parent keys.
    # ContainerInherit+ObjectInherit means all device subkeys inherit SYSTEM full control.
    Set-RegistryKeyOwnership -PSPath $renderPath
    Set-RegistryKeyOwnership -PSPath $capturePath

    Write-Log "========================================"
    Write-Log "Audio registry permissions script completed"
    Write-Log "========================================"
}
finally {
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}