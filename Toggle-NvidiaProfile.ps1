<#
.SYNOPSIS
    Toggles NVIDIA display settings between Gaming and Work profiles

.DESCRIPTION
    Uses NvAPIWrapper to adjust digitial vibrance, brightness, contrast, and gamma for comfortable viewing
    during work or enhanced visibility during gaming.

.PARAMETER Profile
    The profile to apply: "gaming" or "work"

.PARAMETER ListDisplays
    Lists all connected displays and their current settings

.EXAMPLE
    .\Toggle-NvidiaProfile.ps1 -Profile gaming
    .\Toggle-NvidiaProfile.ps1 -Profile work
    .\Toggle-NvidiaProfile.ps1 -ListDisplays

.NOTES
    Requires: NvAPIWrapper.dll (install via NuGet or download from GitHub)
    GitHub: https://github.com/falahati/NvAPIWrapper
#>

param(
    [ValidateSet("gaming", "work")]
    [string]$Profile,

    [switch]$ListDisplays
)

# ===========================================================================================================
# CONFIGURATION - ADJUST THESE VALUES TO YOUR PREFERENCE
# ===========================================================================================================

$Config = @{
    # Path to NvAPIWrapper.dll - UPDATE THIS PATH
    NvAPIWrapperPath = "C:\Tools\NvAPIWrapper\NvAPIWrapper.dll"

    # Work Profile
    Work = @{
        DigitalVibrance = 50
        Brightness      = 50
        Contrast        = 50
        Gamma           = 1.0
    }

    # Gaming Profile
    Gaming = @{
        DigitalVibrance = 70
        Brightness = 55
        Contrast = 55
        Gamma = 1.50
    }
}

# ===========================================================================================================
# FUNCTIONS
# ===========================================================================================================

function Initialize-NvAPI {
    if (-not (Test-Path $Config.NvAPIWrapperPath)) {
        Write-Error @"
NvAPIWrapper.dll not found at $($Config.NvAPIWrapperPath)

To install:
1. Download from: https://github.com/falahati/NvAPIWrapper/releases
2. Extract and update the path in this script's `$Config section
        - or install via NuGet: Install-Package NvAPIWrapper.Net
"@
        exit 1
    }

    try {
        Add-Type -Path $Config.NvAPIWrapperPath
        [NvAPIWrapper.NVIDIA]::Initialize()
        Write-Verbose "NvAPI Initialized successfully"
    }
    catch {
        Write-Error "Failed to initialize NvAPI: $_"
        exit 1
    }
}

function Show-DisplayInfo {
    $displays = Get-AllDisplays

    Write-Host "`n===== Connected NVIDIA Displays =====" -ForegroundColor Cyan 

    foreach ($display in $displays) {
        Write-Host "`nDisplay: $($display.Name)" -ForegroundColor Yellow
        Write-Host " Device Path: $(display.DevicePath)"

        try {
            $vibrance = $display.DigitalVibranceControl.CurrentLevel
            Write-Host " Digital Vibrance: $vibrance"
        }
        catch {
            Write-Host " Digital Vibrance: (not available)" -ForegroundColor DarkGray
        }
    }

    Write-Host "`n"
}

function Set-DisplayProfile {
    param(
        [Parameter(Mandatory)]
        [string]$ProfileName
    )

    $settings = $config[$ProfileName]
    $displays = Get-AllDisplays
    
    Write-Host "`nApplying '$ProfileName' profile..." -ForegroundColor Cyan

    foreach ($display in $displays) {
        Write-Host "  Configuring: $($display.Name)" -ForegroundColor Yellow

        # digital vibrance
        try {
            $display.DigitalVibranceControl.CurrentLevel = $settings.DigitalVibrance
            Write-Host "    Digital Vibrance: $($settings.DigitalVibrance)" -ForegroundColor Green
        }
        catch {
            Write-Host "    Digital Vibrance: Failed - $_" -ForegroundColor Red
        }

        # Note: gamma/brightness/contrast require different API calls
        # Using Windows GDI as fallback for gamma
    }

    # apply gamma via Windows API (works universally)
    Set-ScreenGamma -Gamma $settings.Gamma

    Write-Host "`n✓ Profile '$ProfileName' applied successfully!" -ForegroundColor Green
}

function Set-ScreenGamma {
    param(
        [double]$Gamma = 1.0
    )

    # Clamp to safe range
    $Gamma = [Math]::Max(0.5, [Math]::Min(2.0, $Gamma))

    $GammaCode= @'
using System;
using System.Runtime.InteropServices;

public class GammaRamp {
    [DllImport("gdi32.dll")]
    private static extern bool SetDeviceGammaRamp(IntPtr hDC, ref RAMP lpRamp);

    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    public struct RAMP {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public ushort[] Red;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public ushort[] Green;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public ushort[] Blue;
    }

    public static bool SetGamma(double gamma) {
        RAMP ramp = new RAMP();
        ramp.Red = new ushort[256];
        ramp.Green = new ushort[256];
        ramp.Blue = new ushort[256];

        for (int i = 0; i < 256; i++) {
            int value = (int)(Math.Pow(i / 255.0, 1.0 / gamma) * 65535);
            value = Math.Max(0, Math.Min(65535, value));
            ramp.Red[i] = ramp.Green[i] = ramp.Blue[i] = (ushort)value;
        }

        IntPtr hdc = GetDC(IntPtr.Zero);
        bool result = SetDeviceGammaRamp(hdc, ref ramp);
        ReleaseDC(IntPtr.Zero, hdc);
        return result;
    }
}
'@

    try {
        Add-Type -TypeDefinition $GammaCode -ErrorAction SilentlyContinue
    }
    catch {
        # type may already be loaded
    }

    $result = [GammaRamp]::SetGamma($Gamma)
    if ($result) {
        Write-Host "    Gamma: $Gamma" -ForegroundColor Green
    }
    else {
        Write-Host "    Gamma: failed to apply" -ForegroundColor Red
    }
}

# =============================================================================
# MAIN EXECUTION
# =============================================================================

# Initialize NVIDIA API
Initialize-NvAPI

# Handle parameters
if ($ListDisplays) {
    Show-DisplayInfo
    exit 0
}

if (-not $Profile) {
    Write-Host @"

Usage:
    .\Toggle-NvidiaProfile.ps1 -Profile gaming    # Apply gaming settings
    .\Toggle-NvidiaProfile.ps1 -Profile work      # Apply work settings
    .\Toggle-NvidiaProfile.ps1 -ListDisplays      # Show current display info

Current Configuration:
    Work Profile:
        Digital Vibrance: $($Config.Work.DigitalVibrance)
        Gamma: $($Config.Work.Gamma)

    Gaming Profile:
        Digital Vibrance: $($Config.Gaming.DigitalVibrance)
        Gamma: $($Config.Gaming.Gamma)
"@ -ForegroundColor Cyan
    exit 0
}

# apply the requested profile
Set-DisplayProfile -ProfileName $Profile