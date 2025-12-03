<#
.SYNOPSIS
    Sets the screen resolution to 2048x1080 using native Windows API.

.DESCRIPTION
    This script was created to address a recurring issue during remote connections to hosts with newly reinstalled operating systems.
    Often, those systems defaulted to a basic resolution (e.g., 1024x768), which made them difficult to work with.
    This script enforces a more practical resolution (2048x1080) that fits well with the author's local display setup,
    ensuring consistency and better usability when managing machines remotely.

.NOTES
    Author: Avraham Makovsky
#>

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class Display
{
    [StructLayout(LayoutKind.Sequential)]
    public struct DEVMODE
    {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmDeviceName;
        public short dmSpecVersion;
        public short dmDriverVersion;
        public short dmSize;
        public short dmDriverExtra;
        public int dmFields;
        public int dmPositionX;
        public int dmPositionY;
        public int dmDisplayOrientation;
        public int dmDisplayFixedOutput;
        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmFormName;
        public short dmLogPixels;
        public short dmBitsPerPel;
        public int dmPelsWidth;
        public int dmPelsHeight;
        public int dmDisplayFlags;
        public int dmDisplayFrequency;
        public int dmICMMethod;
        public int dmICMIntent;
        public int dmMediaType;
        public int dmDitherType;
        public int dmReserved1;
        public int dmReserved2;
        public int dmPanningWidth;
        public int dmPanningHeight;
    }

    [DllImport("user32.dll")]
    public static extern int EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE devMode);

    [DllImport("user32.dll")]
    public static extern int ChangeDisplaySettings(ref DEVMODE devMode, int flags);

    public const int ENUM_CURRENT_SETTINGS = -1;
    public const int CDS_UPDATEREGISTRY = 0;

    public static void SetResolution(int width, int height)
    {
        DEVMODE dm = new DEVMODE();
        dm.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));
        EnumDisplaySettings(null, ENUM_CURRENT_SETTINGS, ref dm);

        dm.dmPelsWidth = width;
        dm.dmPelsHeight = height;

        ChangeDisplaySettings(ref dm, CDS_UPDATEREGISTRY);
    }
}
"@

# Set resolution to 2048x1080
[Display]::SetResolution(2048, 1080)
Write-Host "✅ Screen resolution set to 2048x1080."
