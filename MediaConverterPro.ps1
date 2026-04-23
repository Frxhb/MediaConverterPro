<#
.SYNOPSIS
    Media Converter Pro - A powerful media conversion, editing, and downloading tool built with PowerShell and WPF.
.DESCRIPTION
    This script provides a GUI for ffmpeg, ffprobe, yt-dlp, and upscayl. It allows batch processing,
    queuing, downloading from various sites, and AI-based audio transcription/image upscaling.
#>

# ==============================================================================
# 1. APPLICATION INITIALIZATION & WINDOWS API SETUP
# ==============================================================================

# C# Code to import kernel32 and user32 functions. 
# This hides the background PowerShell console window securely.
$hideConsoleCode = @'
using System;
using System.Runtime.InteropServices;
public class ConsoleHelper {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    public static void HideConsole() {
        IntPtr hWnd = GetConsoleWindow();
        if (hWnd != IntPtr.Zero) {
            ShowWindow(hWnd, 0); // 0 = SW_HIDE
        }
    }
}
'@
if (-not ("ConsoleHelper" -as [type])) {
    Add-Type -TypeDefinition $hideConsoleCode
}
[ConsoleHelper]::HideConsole()

# C# Code to bypass User Interface Privilege Isolation (UIPI) if running as Admin.
# This ensures Drag & Drop functionality works even if the script is elevated.
$code = @'
[DllImport("user32.dll")]
public static extern bool ChangeWindowMessageFilterEx(IntPtr hWnd, uint msg, uint action, IntPtr pAttributes);
'@
if (-not ("WinApi.WinApiDragDrop" -as [type])) {
    $winApi = Add-Type -MemberDefinition $code -Name "WinApiDragDrop" -Namespace WinApi -PassThru
}
else {
    $winApi = [WinApi.WinApiDragDrop]
}
$WM_DROPFILES = 0x233
$WM_COPYDATA = 0x004A
$WM_COPYGLOBALDATA = 0x0049
$MSGFLT_ALLOW = 1

$shortPathCode = @'
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace WinApi {
    public class PathHelper {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern uint GetShortPathName(string lpszLongPath, StringBuilder lpszShortPath, uint cchBuffer);
    }
}
'@
if (-not ("WinApi.PathHelper" -as [type])) {
    Add-Type -TypeDefinition $shortPathCode
}

# Save the WinAPI variables to the script scope so we can apply them safely after the WPF Window renders
$script:winApi = $winApi
$script:WM_DROPFILES = $WM_DROPFILES
$script:WM_COPYDATA = $WM_COPYDATA
$script:WM_COPYGLOBALDATA = $WM_COPYGLOBALDATA
$script:MSGFLT_ALLOW = $MSGFLT_ALLOW

# Determine the directory where this script resides to establish a working path
$ScriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ScriptDir)) { 
    # Fallback if running from ISE or as a compiled executable
    $ScriptDir = [System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
}
if ([string]::IsNullOrWhiteSpace($ScriptDir)) { $ScriptDir = $PWD.Path }

# Define relative paths for configuration and log files
$ConfigDir = Join-Path $ScriptDir "config"
$LogDir = Join-Path $ScriptDir "logs"

# Create directories if they do not exist
foreach ($dir in @($ConfigDir, $LogDir)) {
    if (-not (Test-Path $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
}

# Define file paths for app configuration, presets, queue state, and logs
$PresetFile = Join-Path $ConfigDir "mcp_presets.json"
$ConfigFile = Join-Path $ConfigDir "mcp_config.json"
$QueueFile = Join-Path $ConfigDir "mcp_queue.json"
$CrashLog = Join-Path $LogDir "crash.log"
$ConvertLog = Join-Path $LogDir "convert.log"

# Helper functions for logging application events and errors (Thread/Lock safe via retry loop)
function Write-CrashLog { 
    param([string]$Message)
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $Message`r`n"
    for ($i = 0; $i -lt 3; $i++) { 
        try { 
            $fs = [System.IO.File]::Open($CrashLog, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
            try {
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($entry)
                $fs.Write($bytes, 0, $bytes.Length)
            }
            finally {
                if ($null -ne $fs) { $fs.Dispose() }
            }
            break 
        } 
        catch { if ($i -eq 2) { Write-Warning "CrashLog Write Failed: $_" }; Start-Sleep -Milliseconds 50 } 
    }
}
function Write-ConvertLog { 
    param([string]$Message)
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message`r`n"
    for ($i = 0; $i -lt 3; $i++) { try { [System.IO.File]::AppendAllText($ConvertLog, $entry); break } catch { Start-Sleep -Milliseconds 50 } }
}

# Initialize a thread-safe dictionary to prevent queue overwrite race conditions during parallel tasks
$script:ReservedFilenames = [System.Collections.Concurrent.ConcurrentDictionary[string, byte]]::new([StringComparer]::OrdinalIgnoreCase)

# Function to generate a unique filename to prevent overwriting existing files
function Get-UniqueFileName ([string]$FilePath) {
    $dir = [System.IO.Path]::GetDirectoryName($FilePath)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $ext = [System.IO.Path]::GetExtension($FilePath)
    $newPath = $FilePath
    $counter = 1
    
    while ((Test-Path -LiteralPath $newPath) -or $script:ReservedFilenames.ContainsKey($newPath)) {
        $newPath = Join-Path $dir "$name ($counter)$ext"
        $counter++
    }
    [void]$script:ReservedFilenames.TryAdd($newPath, 0)
    return $newPath
}

# ==============================================================================
# 2. MAIN APPLICATION LOGIC & STATE MANAGEMENT
# ==============================================================================
try {
    # Set console encoding to UTF8 to handle special characters in paths/logs
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { Write-Warning "Failed to set UTF8 console encoding: $_" }
    $ErrorActionPreference = "Continue" 
    
    # Load required .NET assemblies for WPF and Windows Forms
    Add-Type -AssemblyName PresentationFramework, System.Windows.Forms, System.Drawing, Microsoft.VisualBasic

    # Setup System Tray Notification Icon (Events are bound later after UI loads)
    $script:TrayIcon = New-Object System.Windows.Forms.NotifyIcon
    $script:TrayIcon.Icon = [System.Drawing.SystemIcons]::Information
    $script:TrayIcon.Visible = $true
    $script:TrayIcon.Text = "Media Converter Pro"

    # Setup Native Windows 10/11 Toast Notifications with fallback
    function Show-Toast {
        param([string]$Title, [string]$Message)
        try {
            # Safely load the native Windows Runtime APIs
            $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
            $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime]
            
            $xmlString = @"
<toast>
    <visual>
        <binding template="ToastGeneric">
            <text>$Title</text>
            <text>$Message</text>
        </binding>
    </visual>
</toast>
"@
            $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
            $xml.LoadXml($xmlString)
            $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
            
            # Using the PowerShell executable's registered AppID to bypass strict Focus Assist rules
            $appId = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"
            [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show($toast)
        }
        catch {
            # Graceful fallback to the legacy balloon tip if WinRT fails
            if ($script:TrayIcon) {
                $script:TrayIcon.ShowBalloonTip(3000, $Title, $Message, [System.Windows.Forms.ToolTipIcon]::Info)
            }
        }
    }

    # Global state object to track the active job queue, tool paths, and live log status
    $script:State = @{
        ffmpeg                 = "ffmpeg.exe"
        ffprobe                = "ffprobe.exe"
        ytdlp                  = "yt-dlp.exe"
        upscayl                = ""
        python                 = "python.exe"
        whisperFound           = $false
        ffmpegFound            = $false
        ffprobeFound           = $false
        ytdlpFound             = $false
        handbrakecli           = "HandBrakeCLI.exe"
        handbrakeFound         = $false
        jsRuntimeFound         = $false
        upscaylFound           = $false
        BatchQueue             = @()
        CurrentJobIndex        = 0
        p                      = $null # Active background process
        totalDuration          = 0     # For progress bar calculation
        lastOutDir             = ""
        tempLog                = Join-Path $env:TEMP "mcp_live.log"
        tempLogErr             = Join-Path $env:TEMP "mcp_live_err.log"
        lastLogPos             = 0
        SupportedSitesCache    = $null
        CustomFilenames        = @{}
        IsAutoUpdatingFilename = $false
        Regex                  = @{
            FFmpegTime  = [regex]::new("time=\s*(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?", [System.Text.RegularExpressions.RegexOptions]::Compiled)
            FFmpegFrame = [regex]::new("frame=\s*(\d+)", [System.Text.RegularExpressions.RegexOptions]::Compiled)
            HbProg      = [regex]::new("Encoding: task \d+ of \d+, (\d+\.\d+) %", [System.Text.RegularExpressions.RegexOptions]::Compiled)
            HbETA       = [regex]::new("ETA ([\dhms]+)", [System.Text.RegularExpressions.RegexOptions]::Compiled)
            YtDlpProg   = [regex]::new("\[download\]\s+(\d+\.?\d*)%", [System.Text.RegularExpressions.RegexOptions]::Compiled)
            YtDlpETA    = [regex]::new("ETA\s+(\d+:\d+)", [System.Text.RegularExpressions.RegexOptions]::Compiled)
            UpscaylProg = [regex]::new("(\d+\.\d+)%", [System.Text.RegularExpressions.RegexOptions]::Compiled)
            WhisperTime = [regex]::new("-->\s*(?:(\d{2,3}):)?(\d{2}):(\d{2})\.(\d{3})", [System.Text.RegularExpressions.RegexOptions]::Compiled)
            Speed       = [regex]::new("speed=\s*(\d+(?:\.\d+)?)x", [System.Text.RegularExpressions.RegexOptions]::Compiled)
        }
    }

    # Function to locate all required third-party tools on the host system
    function Find-Tools {
        $script:State.ffmpegFound = $false; $script:State.ffprobeFound = $false; $script:State.ytdlpFound = $false; $script:State.jsRuntimeFound = $false; $script:State.upscaylFound = $false; $script:isWinGetVersion = $false

        # 1. Attempt to load fast from config cache
        if (Test-Path $ConfigFile) {
            try {
                $cached = Get-Content $ConfigFile | ConvertFrom-Json
                if ($cached.ffmpeg -and (Test-Path $cached.ffmpeg)) { $script:State.ffmpeg = $cached.ffmpeg; $script:State.ffmpegFound = $true }
                if ($cached.ffprobe -and (Test-Path $cached.ffprobe)) { $script:State.ffprobe = $cached.ffprobe; $script:State.ffprobeFound = $true }
                if ($cached.handbrakecli -and (Test-Path $cached.handbrakecli)) { $script:State.handbrakecli = $cached.handbrakecli; $script:State.handbrakeFound = $true }                 if ($cached.ytdlp -and (Test-Path $cached.ytdlp)) { $script:State.ytdlp = $cached.ytdlp; $script:State.ytdlpFound = $true; $script:isWinGetVersion = [bool]$cached.isWinGetVersion }
                if ($cached.upscayl -and (Test-Path $cached.upscayl)) { 
                    $script:State.upscayl = $cached.upscayl; $script:State.upscaylModels = $cached.upscaylModels; $script:State.upscaylWorkDir = $cached.upscaylWorkDir; $script:State.upscaylFound = $true 
                }
                # Check for Node.js quickly
                if ((Get-Command "node.exe" -ErrorAction SilentlyContinue) -or (Get-Command "deno.exe" -ErrorAction SilentlyContinue)) { $script:State.jsRuntimeFound = $true }
                
                if ($script:State.ffmpegFound -and $script:State.handbrakeFound -and $script:State.ytdlpFound) { return } # Skip the heavy loop if everything essential is found!
            }
            catch {}
        }

        # Ensure environment path is fresh for the loop
        $env:PATH = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")

        # Check for system-installed yt-dlp (specifically WinGet/WindowsApps)
        $sysCheck = Get-Command "yt-dlp.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($sysCheck) {
            if ($sysCheck.Source -match "WinGet" -or $sysCheck.Source -match "WindowsApps") {
                $script:State.ytdlp = $sysCheck.Source
                $script:State.ytdlpFound = $true
                $script:isWinGetVersion = $true
            }
        }

        # Fallback check in standard WinGet link folder
        if (-not $script:isWinGetVersion) {
            $winGetLinkPath = Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Links\yt-dlp.exe"
            if (Test-Path $winGetLinkPath) {
                $script:State.ytdlp = $winGetLinkPath
                $script:State.ytdlpFound = $true
                $script:isWinGetVersion = $true
            }
        }

        # Locate ffmpeg and ffprobe globally
        $sysFfmpeg = Get-Command "ffmpeg.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($sysFfmpeg) { 
            $script:State.ffmpeg = $sysFfmpeg.Source
            $script:State.ffmpegFound = $true 
        }

        $sysFfprobe = Get-Command "ffprobe.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($sysFfprobe) { 
            $script:State.ffprobe = $sysFfprobe.Source
            $script:State.ffprobeFound = $true 
        }
        
        # Check for JavaScript runtime (required by yt-dlp to bypass some protections)
        if ((Get-Command "node.exe" -ErrorAction SilentlyContinue) -or (Get-Command "deno.exe" -ErrorAction SilentlyContinue)) { 
            $script:State.jsRuntimeFound = $true 
        }

        # Look for local executables in the script directory if system ones fail
        if (-not $script:isWinGetVersion) {
            $localY = Join-Path $ScriptDir "yt-dlp.exe"
            if (Test-Path $localY) {
                $script:State.ytdlp = $localY
                $script:State.ytdlpFound = $true
            }
        }
    
        if (-not $script:State.ffmpegFound -and (Test-Path (Join-Path $ScriptDir "ffmpeg.exe"))) {
            $script:State.ffmpeg = Join-Path $ScriptDir "ffmpeg.exe"
            $script:State.ffmpegFound = $true
        }
        if (-not $script:State.ffprobeFound -and (Test-Path (Join-Path $ScriptDir "ffprobe.exe"))) {
            $script:State.ffprobe = Join-Path $ScriptDir "ffprobe.exe"
            $script:State.ffprobeFound = $true
        }
        $sysHb = Get-Command "HandBrakeCLI.exe" -ErrorAction SilentlyContinue
        if ($sysHb) { $script:State.handbrakecli = $sysHb.Source; $script:State.handbrakeFound = $true }
        elseif (-not $script:State.handbrakeFound -and (Test-Path (Join-Path $ScriptDir "HandBrakeCLI.exe"))) {
            $script:State.handbrakecli = Join-Path $ScriptDir "HandBrakeCLI.exe"
            $script:State.handbrakeFound = $true         
        }
        
        # Attempt to locate the Upscayl binary (AI Image Upscaling)
        $script:State.upscaylFound = $false
        $script:State.upscayl = ""
        $script:State.upscaylModels = ""
        $script:State.upscaylWorkDir = ""

        # Common installation paths for Upscayl
        $uPaths = @(
            (Join-Path $env:LOCALAPPDATA "Programs\upscayl\resources\bin\upscayl-bin.exe"),
            (Join-Path $env:LOCALAPPDATA "Programs\Upscayl\resources\bin\upscayl-bin.exe"),
            (Join-Path $env:ProgramFiles "upscayl\resources\bin\upscayl-bin.exe"),
            (Join-Path $env:ProgramFiles "Upscayl\resources\bin\upscayl-bin.exe"),
            (Join-Path $ScriptDir "realesrgan\realesrgan-ncnn-vulkan.exe")
        )

        # Loop through possible Upscayl paths and bind model directories if found
        foreach ($p in $uPaths) {
            if (Test-Path -LiteralPath $p) {
                $script:State.upscaylFound = $true
                $script:State.upscayl = $p
                
                $workDir = Split-Path $p -Parent
                $m1 = Join-Path $workDir "models"
                $m2 = Join-Path (Split-Path $workDir -Parent) "models"
                $m3 = Join-Path (Split-Path (Split-Path $workDir -Parent) -Parent) "models"

                if (Test-Path -LiteralPath $m1) { $script:State.upscaylModels = $m1; $script:State.upscaylWorkDir = $workDir }
                elseif (Test-Path -LiteralPath $m2) { $script:State.upscaylModels = $m2; $script:State.upscaylWorkDir = (Split-Path $workDir -Parent) }
                elseif (Test-Path -LiteralPath $m3) { $script:State.upscaylModels = $m3; $script:State.upscaylWorkDir = (Split-Path (Split-Path $workDir -Parent) -Parent) }
                
                break
            }
        }

        # (Removed automatic yt-dlp cache wipe here to preserve YouTube JS signatures and prevent 429 rate limits)
        <#

        old code:
        if ($script:State.ytdlpFound) {
            [void](Start-Process -FilePath $script:State.ytdlp -ArgumentList "--rm-cache-dir" -WindowStyle Hidden)
        }        

        #>

        # --- NEW CACHE SAVING BLOCK ---
        # Save found paths to the config file to speed up the next launch
        if ($script:State.ffmpegFound -or $script:State.ytdlpFound -or $script:State.upscaylFound) {
            $cacheObj = [PSCustomObject]@{}
            if (Test-Path $ConfigFile) {
                try { $cacheObj = Get-Content $ConfigFile -Raw | ConvertFrom-Json } catch {}
            }
            
            # Append or update tool paths in the config object
            if ($script:State.ffmpegFound) { $cacheObj | Add-Member -MemberType NoteProperty -Name "ffmpeg" -Value $script:State.ffmpeg -Force }
            if ($script:State.ffprobeFound) { $cacheObj | Add-Member -MemberType NoteProperty -Name "ffprobe" -Value $script:State.ffprobe -Force }
            if ($script:State.handbrakeFound) { $cacheObj | Add-Member -MemberType NoteProperty -Name "handbrakecli" -Value $script:State.handbrakecli -Force }             if ($script:State.ytdlpFound) { 
                $cacheObj | Add-Member -MemberType NoteProperty -Name "ytdlp" -Value $script:State.ytdlp -Force 
                $cacheObj | Add-Member -MemberType NoteProperty -Name "isWinGetVersion" -Value $script:isWinGetVersion -Force 
            }
            if ($script:State.upscaylFound) { 
                $cacheObj | Add-Member -MemberType NoteProperty -Name "upscayl" -Value $script:State.upscayl -Force
                $cacheObj | Add-Member -MemberType NoteProperty -Name "upscaylModels" -Value $script:State.upscaylModels -Force
                $cacheObj | Add-Member -MemberType NoteProperty -Name "upscaylWorkDir" -Value $script:State.upscaylWorkDir -Force
            }

            # Save silently only if the cache was actually updated to prevent SSD wear and I/O latency
            try { 
                $newJson = $cacheObj | ConvertTo-Json -Depth 5 
                $oldJson = if (Test-Path $ConfigFile) { Get-Content $ConfigFile -Raw } else { "" }
                if ($newJson -ne $oldJson) {
                    Set-Content $ConfigFile -Value $newJson -Encoding UTF8 -Force 
                }
            } catch {
                Write-CrashLog "Failed to save configuration cache: $($_.Exception.Message)"
            }
        }
        # ------------------------------
    }
    
    # Run the tool finder
    Find-Tools

    # ==============================================================================
    # 3. CONFIGURATION MANAGEMENT
    # ==============================================================================
    
    # Load configuration from JSON or define defaults as a proper PSCustomObject
    $DefaultConfig = [PSCustomObject]@{ Theme = "Light"; WhisperModel = "base"; PlaySound = $true; AlwaysOnTop = $false; ThreadLimit = "Auto"; DefaultOutDir = ""; AutoDelete = $false }
    if (Test-Path $ConfigFile) { try { $Config = Get-Content $ConfigFile -Raw | ConvertFrom-Json } catch { $Config = $DefaultConfig } } else { $Config = $DefaultConfig }
    
    # Ensure ALL required UI properties exist (repairs the file if Find-Tools created it first)
    if ($null -eq $Config.Theme) { $Config | Add-Member -MemberType NoteProperty -Name "Theme" -Value "Light" }
    if ($null -eq $Config.WhisperModel) { $Config | Add-Member -MemberType NoteProperty -Name "WhisperModel" -Value "base" }
    if ($null -eq $Config.PlaySound) { $Config | Add-Member -MemberType NoteProperty -Name "PlaySound" -Value $true }
    if ($null -eq $Config.AlwaysOnTop) { $Config | Add-Member -MemberType NoteProperty -Name "AlwaysOnTop" -Value $false }
    if ($null -eq $Config.ThreadLimit) { $Config | Add-Member -MemberType NoteProperty -Name "ThreadLimit" -Value "Auto" }
    if ($null -eq $Config.DefaultOutDir) { $Config | Add-Member -MemberType NoteProperty -Name "DefaultOutDir" -Value "" }
    if ($null -eq $Config.AutoDelete) { $Config | Add-Member -MemberType NoteProperty -Name "AutoDelete" -Value $false }

    # ==============================================================================
    # 4. WPF UI DEFINITION (XAML)
    # ==============================================================================
    
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Media Converter Pro v1.0" Width="1200" Height="1000" MinWidth="900" MinHeight="800" WindowStartupLocation="CenterScreen" Background="{DynamicResource BgBrush}" AllowDrop="True" FontFamily="Segoe UI">
    <Window.TaskbarItemInfo>
        <TaskbarItemInfo x:Name="TaskbarProgress" ProgressState="None"/>
    </Window.TaskbarItemInfo>
    <Window.Resources>
        <SolidColorBrush x:Key="BgBrush" Color="#0F172A"/>
        <SolidColorBrush x:Key="CardBrush" Color="#1E293B"/>
        <SolidColorBrush x:Key="TextBrush" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="MutedBrush" Color="#94A3B8"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#334155"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#6366F1"/>
        <SolidColorBrush x:Key="InputBgBrush" Color="#0F172A"/>
        <SolidColorBrush x:Key="DragBrush" Color="#1E1B4B"/>

        <Style TargetType="TextBlock"><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/></Style>
        
        <Style TargetType="TabItem">
            <Setter Property="Background" Value="Transparent"/><Setter Property="Foreground" Value="{DynamicResource MutedBrush}"/>
            <Setter Property="Padding" Value="20,12"/><Setter Property="FontSize" Value="15"/><Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderThickness" Value="0"/><Setter Property="Margin" Value="0,0,5,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="Border" Background="{TemplateBinding Background}" CornerRadius="8">
                            <ContentPresenter x:Name="ContentSite" VerticalAlignment="Center" HorizontalAlignment="Center" ContentSource="Header" Margin="15,8"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource BorderBrush}"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource AccentBrush}"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SubTabStyle" TargetType="TabItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource MutedBrush}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="Border" Background="{TemplateBinding Background}" CornerRadius="6" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" Margin="0,0,5,10">
                            <ContentPresenter ContentSource="Header" Margin="15,6" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource BorderBrush}"/>
                                <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <Style TargetType="MenuItem">
            <Setter Property="Background" Value="{DynamicResource CardBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>

        <Style TargetType="ListBoxItem">
            <Setter Property="Background" Value="{DynamicResource InputBgBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
            <Setter Property="Padding" Value="5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border CornerRadius="4" Background="{TemplateBinding Background}" Margin="2">
                            <ContentPresenter Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentBrush}"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="{DynamicResource InputBgBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
            <Setter Property="Padding" Value="8"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border CornerRadius="4" Background="{TemplateBinding Background}" Margin="2">
                            <ContentPresenter Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource BorderBrush}"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentBrush}"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{DynamicResource InputBgBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton x:Name="ToggleButton" Grid.Column="2" Focusable="false"
                                          IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press"
                                          Background="{TemplateBinding Background}"
                                          BorderBrush="{TemplateBinding BorderBrush}"
                                          BorderThickness="{TemplateBinding BorderThickness}">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6">
                                            <Path x:Name="Arrow" Fill="{Binding Foreground, RelativeSource={RelativeSource AncestorType=ComboBox}}" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,10,0" Data="M0,0 L4,4 L8,0 z"/>
                                        </Border>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter x:Name="ContentSite" TextElement.Foreground="{TemplateBinding Foreground}" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" Margin="{TemplateBinding Padding}" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            
                            <Popup x:Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False">
                                <Border Background="{DynamicResource CardBrush}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="6" Margin="0,4,0,0" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="250">
                                    <ScrollViewer Margin="2" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TextBox"><Setter Property="Background" Value="{DynamicResource InputBgBrush}"/><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/><Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/><Setter Property="BorderThickness" Value="1"/><Setter Property="Padding" Value="8"/><Setter Property="VerticalContentAlignment" Value="Center"/><Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="0"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ListBox"><Setter Property="Background" Value="{DynamicResource InputBgBrush}"/><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/><Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/><Setter Property="BorderThickness" Value="1"/><Setter Property="Padding" Value="5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBox">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6">
                            <Grid>
                                <TextBlock x:Name="WatermarkText" Text="&#x1F4C1; Drag &amp; Drop Media Here" Foreground="{DynamicResource MutedBrush}" FontSize="14" FontWeight="SemiBold" HorizontalAlignment="Center" VerticalAlignment="Center" IsHitTestVisible="False" Visibility="Collapsed"/>
                                <ScrollViewer Focusable="false" Padding="{TemplateBinding Padding}">
                                    <ItemsPresenter SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}"/>
                                </ScrollViewer>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="HasItems" Value="False">
                                <Setter TargetName="WatermarkText" Property="Visibility" Value="Visible"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <Style TargetType="CheckBox"><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/><Setter Property="VerticalContentAlignment" Value="Center"/></Style>
        <Style TargetType="Border"><Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/></Style>
        
        <Style TargetType="Button">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="6" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>

        <Grid Grid.Row="0" Margin="0,0,0,20">
            <StackPanel>
                <TextBlock Text="Media Converter Pro" Foreground="{DynamicResource AccentBrush}" FontSize="32" FontWeight="Black"/>
                <TextBlock x:Name="TxtSubtitle" Text="Open Source Media Editor, Converter and Downloader" Foreground="{DynamicResource MutedBrush}" FontSize="14" Margin="0,2,0,0"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top">
                <Button x:Name="BtnUpdate" Content="Update Tools" Width="110" Height="40" FontSize="14" Background="{DynamicResource CardBrush}" Foreground="{DynamicResource TextBrush}" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" Cursor="Hand" Margin="0,0,10,0" ToolTip="Update Dependencies"/>
                <Button x:Name="BtnSettings" Content="Settings" Width="90" Height="40" FontSize="14" Background="{DynamicResource CardBrush}" Foreground="{DynamicResource TextBrush}" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" Cursor="Hand" ToolTip="Settings"/>
            </StackPanel>
        </Grid>
        
        <TabControl x:Name="MainTabs" Grid.Row="1" Background="{DynamicResource CardBrush}" BorderBrush="Transparent" BorderThickness="0" Padding="0">
            
            <TabItem x:Name="TabAudio">
                <TabItem.Header>
                    <TextBlock Text="Audio" ToolTip="Convert, extract, and normalize audio files" />
                </TabItem.Header>
                <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="15">
                    <StackPanel>
                        <Border Background="{DynamicResource CardBrush}" BorderThickness="1" CornerRadius="8" Padding="20" Margin="0,0,0,15">
                            <StackPanel>
                                <TextBlock Text="Queue (Drag and Drop here)" FontWeight="Bold" FontSize="16" Margin="0,0,0,10"/>
                                <Grid Margin="0,0,0,15">
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="130"/><ColumnDefinition Width="90"/></Grid.ColumnDefinitions>
                                    <ListBox x:Name="A_InList" Height="100" AllowDrop="True" SelectionMode="Extended">
                                        <ListBox.ContextMenu>
                                            <ContextMenu x:Name="A_CtxMenu">
                                                <MenuItem x:Name="A_CtxRemove" Header="Remove Selected"/>
                                                <MenuItem x:Name="A_CtxClear" Header="Clear All"/>
                                            </ContextMenu>
                                        </ListBox.ContextMenu>
                                    </ListBox>
                                    <StackPanel Grid.Column="1" Margin="10,0,0,0">
                                        <Button x:Name="A_BtnAdd" Content="+ Files" Height="45" Background="{DynamicResource AccentBrush}" Foreground="White" BorderThickness="0" Cursor="Hand" Margin="0,0,0,10"/>
                                        <Button x:Name="A_BtnClear" Content="Clear" Height="45" Background="#EF4444" Foreground="White" BorderThickness="0" Cursor="Hand"/>
                                    </StackPanel>
                                    <Button x:Name="A_BtnInfo" Grid.Column="2" Content="Info" Margin="10,0,0,0" Background="{DynamicResource MutedBrush}" Foreground="White" BorderThickness="0" Cursor="Hand" Height="45" VerticalAlignment="Top"/>
                                </Grid>
                                <Grid>
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="130"/><ColumnDefinition Width="90"/></Grid.ColumnDefinitions>
                                    <TextBox x:Name="A_OutDir" ToolTip="The folder where processed files will be saved" Text="Select target folder..." IsReadOnly="True" Cursor="Arrow"/>
                                    <Button x:Name="A_BtnOut" Grid.Column="1" Content="Select Folder" Margin="10,0,0,0" Height="40" Background="#4B5563" Foreground="White" BorderThickness="0" Cursor="Hand"/>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <TabControl Background="Transparent" BorderThickness="0">
                            <TabItem Style="{StaticResource SubTabStyle}">
                                <TabItem.Header><TextBlock Text="Encoding Settings"/></TabItem.Header>
                                <Border Background="{DynamicResource BgBrush}" CornerRadius="8" Padding="20">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                        <Grid.RowDefinitions><RowDefinition/><RowDefinition/></Grid.RowDefinitions>
                                        <StackPanel Grid.Row="0" Margin="10"><TextBlock Text="Format" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/><ComboBox x:Name="A_CFormat" SelectedIndex="0"><ComboBoxItem>MP3</ComboBoxItem><ComboBoxItem>M4A</ComboBoxItem><ComboBoxItem>WAV</ComboBoxItem><ComboBoxItem>FLAC</ComboBoxItem><ComboBoxItem>Copy (Extract Original)</ComboBoxItem></ComboBox></StackPanel>
                                        <StackPanel Grid.Row="0" Grid.Column="1" Margin="10"><TextBlock Text="Quality (Bitrate)" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/><ComboBox x:Name="A_CQual" SelectedIndex="2"><ComboBoxItem>128k</ComboBoxItem><ComboBoxItem>192k</ComboBoxItem><ComboBoxItem>320k</ComboBoxItem><ComboBoxItem>Lossless</ComboBoxItem></ComboBox></StackPanel>
                                        <StackPanel Grid.Row="1" Margin="10"><TextBlock Text="Channels" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/><ComboBox x:Name="A_CChan" SelectedIndex="0"><ComboBoxItem>Original</ComboBoxItem><ComboBoxItem>Mono</ComboBoxItem><ComboBoxItem>Stereo</ComboBoxItem></ComboBox></StackPanel>
                                        <StackPanel Grid.Row="1" Grid.Column="1" Margin="10"><TextBlock Text="Metadata" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/><ComboBox x:Name="A_CMeta" SelectedIndex="0"><ComboBoxItem>Keep</ComboBoxItem><ComboBoxItem>Remove</ComboBoxItem></ComboBox></StackPanel>
                                    </Grid>
                                </Border>
                            </TabItem>
                            
                            <TabItem Style="{StaticResource SubTabStyle}">
                                <TabItem.Header><TextBlock Text="Filters &amp; Advanced"/></TabItem.Header>
                                <Border Background="{DynamicResource BgBrush}" CornerRadius="8" Padding="20" Focusable="True">
                                    <StackPanel Focusable="True">
                                        <Grid Margin="0,0,0,15">
                                            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                            
                                            <Border Grid.Column="0" Background="{DynamicResource CardBrush}" CornerRadius="6" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" Padding="15" Margin="0,0,10,0">
                                                <StackPanel>
                                                    <TextBlock Text="Trim Audio" FontSize="14" FontWeight="Bold" Foreground="#EF4444" Margin="0,0,0,10"/>
                                                    <Grid Margin="0,0,0,10">
                                                        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                        <StackPanel Margin="0,0,5,0">
                                                            <TextBlock Text="Start (00:00:00)" FontSize="12" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,2"/>
                                                            <TextBox x:Name="A_TrimStart" Text="00:00:00"/>
                                                        </StackPanel>
                                                        <StackPanel Grid.Column="1" Margin="5,0,0,0">
                                                            <TextBlock Text="End (00:00:00)" FontSize="12" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,2"/>
                                                            <TextBox x:Name="A_TrimEnd" Text="00:00:00"/>
                                                        </StackPanel>
                                                    </Grid>
                                                    <Grid>
                                                        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                        <Slider x:Name="A_SliderTrimStart" Minimum="0" Maximum="0" Value="0" Margin="0,0,5,0" IsEnabled="False" ToolTip="Drag to set Start Time"/>
                                                        <Slider x:Name="A_SliderTrimEnd" Minimum="0" Maximum="0" Value="0" Grid.Column="1" Margin="5,0,0,0" IsEnabled="False" ToolTip="Drag to set End Time"/>
                                                    </Grid>
                                                </StackPanel>
                                            </Border>

                                            <Border Grid.Column="1" Background="{DynamicResource CardBrush}" CornerRadius="6" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" Padding="15" Margin="10,0,0,0">
                                                <StackPanel VerticalAlignment="Center">
                                                    <TextBlock Text="Quick Operations" FontSize="14" FontWeight="Bold" Foreground="{DynamicResource TextBrush}" Margin="0,0,0,10"/>
                                                    <CheckBox x:Name="A_CheckNorm" Content="Normalize Audio (R128)" Margin="0,0,0,12" FontWeight="Bold"/>
                                                    <CheckBox x:Name="A_CheckExtract" Content="Extract Audio from Video (.mp4, .mkv)" FontWeight="Bold" Foreground="{DynamicResource AccentBrush}"/>
                                                </StackPanel>
                                            </Border>
                                        </Grid>

                                        <Border Background="{DynamicResource CardBrush}" CornerRadius="6" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" Padding="15">
                                            <StackPanel>
                                                <CheckBox x:Name="A_CheckCustomParams" Content="Add custom FFmpeg params (Live Preview)" FontWeight="Bold"/>
                                                <StackPanel x:Name="A_CustomParamsPanel" Visibility="Collapsed" Margin="0,15,0,0">
                                                    <TextBlock Text="Final FFmpeg command preview:" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                    <TextBox x:Name="A_ParamsPreview" IsReadOnly="True" TextWrapping="Wrap" Background="#0F172A" Foreground="#10B981" FontFamily="Consolas" FontSize="13" Padding="12" Margin="0,0,0,10" MinHeight="60"/>
                                                    <TextBlock Text="Add extra arguments (e.g. -af volume=2.0):" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                    <TextBox x:Name="A_CustomParams"/>
                                                </StackPanel>
                                            </StackPanel>
                                        </Border>
                                    </StackPanel>
                                </Border>
                            </TabItem>
                        </TabControl>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <TabItem x:Name="TabVideo">
                <TabItem.Header>
                    <TextBlock Text="Video" ToolTip="Convert, compress, and optimize video files" />
                </TabItem.Header>
                <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="15">
                    <StackPanel>
                        <Border Background="{DynamicResource CardBrush}" BorderThickness="1" CornerRadius="8" Padding="20" Margin="0,0,0,15">
                            <StackPanel>
                                <TextBlock Text="Queue (Drag and Drop here)" FontWeight="Bold" FontSize="16" Margin="0,0,0,10"/>
                                <Grid Margin="0,0,0,15">
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="130"/><ColumnDefinition Width="90"/></Grid.ColumnDefinitions>
                                    <ListBox x:Name="V_InList" Height="100" AllowDrop="True" SelectionMode="Extended">
                                        <ListBox.ContextMenu>
                                            <ContextMenu x:Name="V_CtxMenu">
                                                <MenuItem x:Name="V_CtxRemove" Header="Remove Selected"/>
                                                <MenuItem x:Name="V_CtxClear" Header="Clear All"/>
                                            </ContextMenu>
                                        </ListBox.ContextMenu>
                                    </ListBox>
                                    <StackPanel Grid.Column="1" Margin="10,0,0,0">
                                        <Button x:Name="V_BtnAdd" Content="+ Files" Height="45" Background="{DynamicResource AccentBrush}" Foreground="White" BorderThickness="0" Cursor="Hand" Margin="0,0,0,10"/>
                                        <Button x:Name="V_BtnClear" Content="Clear" Height="45" Background="#EF4444" Foreground="White" BorderThickness="0" Cursor="Hand"/>
                                    </StackPanel>
                                    <Button x:Name="V_BtnInfo" Grid.Column="2" Content="Info" Margin="10,0,0,0" Background="{DynamicResource MutedBrush}" Foreground="White" BorderThickness="0" Cursor="Hand" Height="45" VerticalAlignment="Top"/>
                                </Grid>
                                <Grid>
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="130"/><ColumnDefinition Width="90"/></Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                                    <TextBox x:Name="V_OutDir" Grid.Row="0" ToolTip="The folder where processed files will be saved" Text="Select target folder..." IsReadOnly="True" Cursor="Arrow" Margin="0,0,0,10" Height="40"/>
                                    <Button x:Name="V_BtnOut" Grid.Row="0" Grid.Column="1" Content="Select Folder" Margin="10,0,0,10" Height="40" Background="#4B5563" Foreground="White" BorderThickness="0" Cursor="Hand"/>
                                    <TextBox x:Name="V_OutFilename" Grid.Row="1" Grid.Column="0" ToolTip="Edit the final filename for the selected item" Text="" Cursor="IBeam" Height="40"/>
                                    <CheckBox x:Name="V_CheckSmartName" Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2" Content="Auto-Update Name" IsChecked="False" IsEnabled="False" VerticalAlignment="Center" Margin="10,0,0,0" ToolTip="Automatically replace audio/video tags in filename (e.g. dtshd to eac3)" FontWeight="Bold" Foreground="{DynamicResource AccentBrush}"/>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <TabControl Background="Transparent" BorderThickness="0">
                            <TabItem Style="{StaticResource SubTabStyle}">
                                <TabItem.Header><TextBlock Text="Encoding Options"/></TabItem.Header>
                                <Border Background="{DynamicResource BgBrush}" CornerRadius="8" Padding="20">
                                    <StackPanel>
                                        <StackPanel Margin="10,0,10,15">
                                            <TextBlock Text="Profile (Presets)" FontSize="13" Foreground="#8B5CF6" FontWeight="Bold" Margin="0,0,0,5"/>
                                            <Grid>
                                                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="150"/></Grid.ColumnDefinitions>
                                                <ComboBox x:Name="V_Preset" SelectedIndex="0">
                                                    <ComboBoxItem>Set Manually</ComboBoxItem>
                                                    <ComboBoxItem>Standard MP4 (H.264, 1080p, CRF 23)</ComboBoxItem>
                                                    <ComboBoxItem>WhatsApp / Web (720p, CRF 28)</ComboBoxItem>
                                                </ComboBox>
                                                <Button x:Name="V_BtnSavePreset" Grid.Column="1" Content="Save as Preset" Background="#10B981" Foreground="White" BorderThickness="0" Cursor="Hand" Margin="10,0,0,0"/>
                                            </Grid>
                                        </StackPanel>

                                        <Grid>
                                            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                            <Grid.RowDefinitions><RowDefinition/><RowDefinition/><RowDefinition/></Grid.RowDefinitions>
                                            
                                            <StackPanel Grid.Row="0" Margin="10"><TextBlock Text="Container Format" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/><ComboBox x:Name="V_CFormat" SelectedIndex="1"><ComboBoxItem>MP4</ComboBoxItem><ComboBoxItem>MKV</ComboBoxItem><ComboBoxItem>AVI</ComboBoxItem><ComboBoxItem>WEBM</ComboBoxItem><ComboBoxItem>GIF</ComboBoxItem></ComboBox></StackPanel>
                                            <StackPanel Grid.Row="0" Grid.Column="1" Margin="10"><TextBlock Text="Hardware Acceleration" FontSize="13" Foreground="#10B981" FontWeight="Bold" Margin="0,0,0,5"/><ComboBox x:Name="V_CHWAccel" SelectedIndex="0"><ComboBoxItem>CPU (x264/x265/AV1)</ComboBoxItem><ComboBoxItem>NVIDIA GPU (NVENC)</ComboBoxItem><ComboBoxItem>AMD GPU (AMF)</ComboBoxItem><ComboBoxItem>Intel GPU (QSV)</ComboBoxItem></ComboBox></StackPanel>

                                            <Grid Grid.Row="1">
                                                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                <StackPanel Margin="10,0,5,0"><TextBlock Text="Video Codec" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                    <ComboBox x:Name="V_CCodec" SelectedIndex="3" ToolTip="Select 'Copy' to keep original video untouched.">
                                                        <ComboBoxItem>H.264 (AVC)</ComboBoxItem>
                                                        <ComboBoxItem>H.265 (HEVC)</ComboBoxItem>
                                                        <ComboBoxItem>AV1 (Next-Gen)</ComboBoxItem>
                                                        <ComboBoxItem>Copy Video</ComboBoxItem>
                                                    </ComboBox>
                                                </StackPanel>
                                                <StackPanel Grid.Column="1" Margin="5,0,5,0"><TextBlock Text="Audio Codec" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                    <ComboBox x:Name="V_CAudio" SelectedIndex="3" ToolTip="Select 'Copy' to keep original audio untouched.">
                                                        <ComboBoxItem>AAC</ComboBoxItem>
                                                        <ComboBoxItem>AC3 (Dolby)</ComboBoxItem>
                                                        <ComboBoxItem>EAC3 (Dolby+)</ComboBoxItem>
                                                        <ComboBoxItem>Copy Audio</ComboBoxItem>
                                                    </ComboBox>
                                                </StackPanel>
                                                <StackPanel Grid.Column="2" Margin="5,0,10,0"><TextBlock Text="Subtitles" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                    <ComboBox x:Name="V_CSub" SelectedIndex="1" ToolTip="Choose how to handle existing subtitles.">
                                                        <ComboBoxItem>Remove Subs</ComboBoxItem>
                                                        <ComboBoxItem>Copy All Subs</ComboBoxItem>
                                                    </ComboBox>
                                                </StackPanel>
                                            </Grid>
                                        </Grid>
                                    </StackPanel>
                                </Border>
                            </TabItem>

                            <TabItem Style="{StaticResource SubTabStyle}">
                                <TabItem.Header><TextBlock Text="Filters &amp; Editing"/></TabItem.Header>
                                <Border Background="{DynamicResource BgBrush}" CornerRadius="8" Padding="20">
                                    <StackPanel>
                                        <Grid Margin="0,0,0,15">
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="*"/>
                                            </Grid.ColumnDefinitions>

                                            <Border Grid.Column="0" Background="{DynamicResource CardBrush}" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}" Padding="15" Margin="0,0,10,0">
                                                <StackPanel>
                                                    <TextBlock Text="Video Settings" FontWeight="Bold" Foreground="{DynamicResource TextBrush}" Margin="0,0,0,10"/>
                                                    <StackPanel Margin="0,0,0,10"><TextBlock Text="Resolution" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/><ComboBox x:Name="V_CRes" SelectedIndex="0"><ComboBoxItem>Original</ComboBoxItem><ComboBoxItem>1080p (FHD)</ComboBoxItem><ComboBoxItem>720p (HD)</ComboBoxItem></ComboBox></StackPanel>
                                                    <StackPanel Margin="0,0,0,10"><TextBlock Text="Framerate (FPS)" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                        <ComboBox x:Name="V_CFPS" SelectedIndex="0">
                                                            <ComboBoxItem>Original</ComboBoxItem>
                                                            <ComboBoxItem>60 FPS</ComboBoxItem>
                                                            <ComboBoxItem>30 FPS</ComboBoxItem>
                                                            <ComboBoxItem>24 FPS (Cinematic)</ComboBoxItem>
                                                        </ComboBox>
                                                    </StackPanel>
                                                    <StackPanel><TextBlock Text="Video Speed" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/><ComboBox x:Name="V_CSpeed" SelectedIndex="0"><ComboBoxItem>Original (1.0x)</ComboBoxItem><ComboBoxItem>0.5x (Slow Motion)</ComboBoxItem><ComboBoxItem>1.25x (Fast)</ComboBoxItem><ComboBoxItem>1.5x (Faster)</ComboBoxItem><ComboBoxItem>2.0x (Very Fast)</ComboBoxItem></ComboBox></StackPanel>
                                                </StackPanel>
                                            </Border>

                                            <Border Grid.Column="1" Background="{DynamicResource CardBrush}" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}" Padding="15" Margin="10,0,0,0">
                                                <StackPanel>
                                                    <TextBlock Text="Audio &amp; Subtitles" FontWeight="Bold" Foreground="{DynamicResource TextBrush}" Margin="0,0,0,10"/>
                                                    <StackPanel Margin="0,0,0,10"><TextBlock Text="Audio Tracks (Multi-Track)" FontSize="13" Foreground="{DynamicResource AccentBrush}" FontWeight="Bold" Margin="0,0,0,5"/><ComboBox x:Name="V_CAudioTracks" SelectedIndex="1"><ComboBoxItem>Default Only (Track 1)</ComboBoxItem><ComboBoxItem>Keep ALL Tracks &amp; Subs</ComboBoxItem></ComboBox></StackPanel>
                                                    <Grid Margin="0,0,0,10">
                                                        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                        <StackPanel Margin="0,0,5,0"><TextBlock Text="Volume" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/><ComboBox x:Name="V_CVol" SelectedIndex="0"><ComboBoxItem>Original</ComboBoxItem><ComboBoxItem>Boost 150%</ComboBoxItem><ComboBoxItem>Normalize (EBU)</ComboBoxItem></ComboBox></StackPanel>
                                                        <StackPanel Grid.Column="1" Margin="5,0,0,0"><TextBlock Text="Delay (Sec)" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/><TextBox x:Name="V_AudioDelay" Text="0.0"/></StackPanel>
                                                    </Grid>
                                                    <StackPanel><TextBlock Text="Burn Subtitles (.srt)" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBox x:Name="V_SubPath" IsReadOnly="True"/><Button x:Name="V_BtnSub" Grid.Column="1" Content="..." Width="40" Background="#4B5563" Foreground="White" BorderThickness="0" Cursor="Hand" Margin="5,0,0,0"/></Grid>
                                                    </StackPanel>
                                                </StackPanel>
                                            </Border>
                                        </Grid>

                                        <Border Background="{DynamicResource CardBrush}" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}" Padding="15">
                                            <StackPanel>
                                                <TextBlock Text="Timeline &amp; Trimming" FontWeight="Bold" Foreground="#EF4444" Margin="0,0,0,10"/>
                                                <Grid>
                                                    <Grid.ColumnDefinitions>
                                                        <ColumnDefinition Width="200"/>
                                                        <ColumnDefinition Width="*"/>
                                                    </Grid.ColumnDefinitions>
                                                    
                                                    <StackPanel Grid.Column="0" Margin="0,0,15,0" VerticalAlignment="Center">
                                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                            <StackPanel Margin="0,0,5,0"><TextBlock Text="Start" FontSize="12" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,2"/><TextBox x:Name="V_TrimStart" Text="00:00:00"/></StackPanel>
                                                            <StackPanel Grid.Column="1" Margin="5,0,0,0"><TextBlock Text="End" FontSize="12" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,2"/><TextBox x:Name="V_TrimEnd" Text="00:00:00"/></StackPanel>
                                                        </Grid>
                                                        <Button x:Name="V_BtnGenPreview" Content="Generate Visual Timeline" Margin="0,15,0,0" Height="30" Background="#8B5CF6" Foreground="White" BorderThickness="0" Cursor="Hand" FontWeight="Bold"/>
                                                    </StackPanel>

                                                    <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                                        <Grid>
                                                            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                            <Slider x:Name="V_SliderTrimStart" Minimum="0" Maximum="0" Value="0" Margin="0,0,5,0" IsEnabled="False" ToolTip="Drag to set Start Time"/>
                                                            <Slider x:Name="V_SliderTrimEnd" Minimum="0" Maximum="0" Value="0" Grid.Column="1" Margin="5,0,0,0" IsEnabled="False" ToolTip="Drag to set End Time"/>
                                                        </Grid>
                                                        <ScrollViewer x:Name="V_PreviewScroll" Height="120" Visibility="Collapsed" VerticalScrollBarVisibility="Disabled" HorizontalScrollBarVisibility="Auto" Margin="0,15,0,0">
                                                            <StackPanel x:Name="V_PreviewStack" Orientation="Horizontal"/>
                                                        </ScrollViewer>
                                                    </StackPanel>
                                                </Grid>
                                            </StackPanel>
                                        </Border>
                                    </StackPanel>
                                </Border>
                            </TabItem>

<TabItem Style="{StaticResource SubTabStyle}">
                                <TabItem.Header><TextBlock Text="Advanced Quality"/></TabItem.Header>
                                <Border Background="{DynamicResource BgBrush}" CornerRadius="8" Padding="20">
                                    <StackPanel>
                                        
                                        <Border Margin="10,0,10,15" Background="{DynamicResource CardBrush}" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" Padding="15" CornerRadius="6">
                                            <Grid>
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="Auto"/>
                                                    <ColumnDefinition Width="100"/>
                                                    <ColumnDefinition Width="*"/>
                                                </Grid.ColumnDefinitions>
                                                <CheckBox x:Name="V_CheckTargetSize" Content="Compress to exact target size (Overrides CRF): " VerticalAlignment="Center" Foreground="{DynamicResource AccentBrush}" FontWeight="Bold"/>
                                                <TextBox x:Name="V_TargetSizeMB" Grid.Column="1" Text="24.5" Margin="10,0,0,0" ToolTip="Target in MB (e.g. 24.5 for Discord)"/>
                                                <TextBlock Grid.Column="2" Text="MB" VerticalAlignment="Center" Margin="10,0,0,0" Foreground="{DynamicResource MutedBrush}"/>
                                            </Grid>
                                        </Border>

                                        <StackPanel Margin="10,0,10,15">
                                            <TextBlock Text="Quality / CRF (Lower = Better, 23 = Default)" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                            <Grid>
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="*"/>
                                                    <ColumnDefinition Width="40"/>
                                                    <ColumnDefinition Width="Auto"/>
                                                </Grid.ColumnDefinitions>
                                                <Slider x:Name="V_SliderCRF" Minimum="0" Maximum="51" Value="23" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center"/>
                                                <TextBlock x:Name="V_CRFText" Grid.Column="1" Text="23" TextAlignment="Center" VerticalAlignment="Center" FontSize="15" FontWeight="Bold"/>
                                                <TextBlock x:Name="V_CRFDesc" Grid.Column="2" Text="Balanced (Standard)" Margin="10,0,0,0" VerticalAlignment="Center" Foreground="{DynamicResource MutedBrush}"/>
                                            </Grid>
                                        </StackPanel>

                                        <Border Background="{DynamicResource CardBrush}" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}" Padding="15" Margin="10,0,10,15">
                                            <CheckBox x:Name="V_UseHandbrake" Content="Use HandBrake Engine (Best for Audio Sync, VFR, and heavy compression)" FontWeight="Bold" Foreground="#10B981" ToolTip="Uses HandBrakeCLI instead of FFmpeg. Video/Audio Copy features are disabled."/>
                                        </Border>
                                        
                                        <StackPanel Margin="10,5,10,0">
                                            <CheckBox x:Name="V_CheckCustomParams" Content="Add custom FFmpeg params (Live Preview)" FontWeight="Bold"/>
                                            <StackPanel x:Name="V_CustomParamsPanel" Visibility="Collapsed" Margin="0,15,0,0">
                                                <TextBlock Text="Final FFmpeg command preview:" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                <TextBox x:Name="V_ParamsPreview" IsReadOnly="True" TextWrapping="Wrap" Background="#0F172A" Foreground="#10B981" FontFamily="Consolas" FontSize="13" Padding="12" Margin="0,0,0,10" MinHeight="60"/>
                                                <TextBlock Text="Add extra arguments (e.g. -vf hue=s=0 -preset slow):" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                <TextBox x:Name="V_CustomParams"/>
                                            </StackPanel>
                                        </StackPanel>
                                        
                                    </StackPanel>
                                </Border>
                            </TabItem>

                        </TabControl>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <TabItem x:Name="TabImage">
                <TabItem.Header>
                    <TextBlock Text="Images" ToolTip="Convert, scale, and manage metadata for images" />
                </TabItem.Header>
                <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="15">
                    <StackPanel>
                        <Border Background="{DynamicResource CardBrush}" BorderThickness="1" CornerRadius="8" Padding="20" Margin="0,0,0,15">
                            <StackPanel>
                                <TextBlock Text="Queue (Drag and Drop here)" FontWeight="Bold" FontSize="16" Margin="0,0,0,10"/>
                                <Grid Margin="0,0,0,15">
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="130"/><ColumnDefinition Width="90"/></Grid.ColumnDefinitions>
                                    <ListBox x:Name="I_InList" Height="100" AllowDrop="True" SelectionMode="Extended">
                                        <ListBox.ContextMenu>
                                            <ContextMenu x:Name="I_CtxMenu">
                                                <MenuItem x:Name="I_CtxRemove" Header="Remove Selected"/>
                                                <MenuItem x:Name="I_CtxClear" Header="Clear All"/>
                                            </ContextMenu>
                                        </ListBox.ContextMenu>
                                    </ListBox>
                                    <StackPanel Grid.Column="1" Margin="10,0,0,0">
                                        <Button x:Name="I_BtnAdd" Content="+ Files" Height="45" Background="{DynamicResource AccentBrush}" Foreground="White" BorderThickness="0" Cursor="Hand" Margin="0,0,0,10"/>
                                        <Button x:Name="I_BtnClear" Content="Clear" Height="45" Background="#EF4444" Foreground="White" BorderThickness="0" Cursor="Hand"/>
                                    </StackPanel>
                                    <Button x:Name="I_BtnInfo" Grid.Column="2" Content="Info" Margin="10,0,0,0" Background="{DynamicResource MutedBrush}" Foreground="White" BorderThickness="0" Cursor="Hand" Height="45" VerticalAlignment="Top"/>
                                </Grid>
                                <Grid>
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="130"/><ColumnDefinition Width="90"/></Grid.ColumnDefinitions>
                                    <TextBox x:Name="I_OutDir" ToolTip="The folder where processed files will be saved" Text="Select target folder..." IsReadOnly="True" Cursor="Arrow"/>
                                    <Button x:Name="I_BtnOut" Grid.Column="1" Content="Select Folder" Margin="10,0,0,0" Height="40" Background="#4B5563" Foreground="White" BorderThickness="0" Cursor="Hand"/>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <TabControl Background="Transparent" BorderThickness="0">
                            <TabItem Style="{StaticResource SubTabStyle}">
                                <TabItem.Header><TextBlock Text="Conversion"/></TabItem.Header>
                                <Border Background="{DynamicResource BgBrush}" CornerRadius="8" Padding="20">
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>

                                        <Border Grid.Column="0" Background="{DynamicResource CardBrush}" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}" Padding="15" Margin="0,0,10,0">
                                            <StackPanel>
                                                <TextBlock Text="Encoding Settings" FontWeight="Bold" Foreground="{DynamicResource TextBrush}" Margin="0,0,0,10"/>
                                                <StackPanel Margin="0,0,0,10">
                                                    <TextBlock Text="Format" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                    <ComboBox x:Name="I_CFormat" SelectedIndex="0">
                                                        <ComboBoxItem>JPG</ComboBoxItem>
                                                        <ComboBoxItem>PNG</ComboBoxItem>
                                                        <ComboBoxItem>WEBP</ComboBoxItem>
                                                        <ComboBoxItem>ICO (Windows Icon)</ComboBoxItem>
                                                        <ComboBoxItem>HEIC</ComboBoxItem>
                                                        <ComboBoxItem>BMP</ComboBoxItem>
                                                    </ComboBox>
                                                </StackPanel>
                                                <StackPanel>
                                                    <TextBlock Text="Quality (JPG/WEBP only)" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                    <ComboBox x:Name="I_CQual" SelectedIndex="0">
                                                        <ComboBoxItem>High (Default)</ComboBoxItem>
                                                        <ComboBoxItem>Medium</ComboBoxItem>
                                                        <ComboBoxItem>Low</ComboBoxItem>
                                                    </ComboBox>
                                                </StackPanel>
                                            </StackPanel>
                                        </Border>

                                        <Border Grid.Column="1" Background="{DynamicResource CardBrush}" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}" Padding="15" Margin="10,0,0,0">
                                            <StackPanel>
                                                <TextBlock Text="Adjustments" FontWeight="Bold" Foreground="{DynamicResource TextBrush}" Margin="0,0,0,10"/>
                                                <StackPanel Margin="0,0,0,15">
                                                    <TextBlock Text="Scaling" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                    <ComboBox x:Name="I_CRes" SelectedIndex="0">
                                                        <ComboBoxItem>Original</ComboBoxItem>
                                                        <ComboBoxItem>Max Width 1920px</ComboBoxItem>
                                                        <ComboBoxItem>Max Width 1280px</ComboBoxItem>
                                                        <ComboBoxItem>Max Width 800px</ComboBoxItem>
                                                    </ComboBox>
                                                </StackPanel>
                                                <CheckBox x:Name="I_CheckMeta" Content="Completely remove EXIF &amp; Metadata" Foreground="#EF4444" FontWeight="Bold" VerticalAlignment="Center" Margin="0,5,0,0"/>
                                            </StackPanel>
                                        </Border>
                                    </Grid>
                                </Border>
                            </TabItem>
                        </TabControl>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <TabItem x:Name="TabMuxing">
                <TabItem.Header>
                    <TextBlock Text="Muxing" ToolTip="Merge separate video and audio files instantly without re-encoding" />
                </TabItem.Header>
                <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="15">
                    <StackPanel>
                        <Border Background="{DynamicResource CardBrush}" BorderThickness="1" CornerRadius="8" Padding="20">
                            <StackPanel>
                                <TextBlock Text="Merge Video &amp; Audio without re-encoding" Foreground="{DynamicResource AccentBrush}" FontSize="18" FontWeight="Bold" Margin="0,0,0,20"/>
                                
                                <TextBlock Text="1. Select video without audio (Drop file here)" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,8"/>
                                <Grid Margin="0,0,0,20"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="120"/></Grid.ColumnDefinitions><TextBox x:Name="M_InVideo" IsReadOnly="True"/><Button x:Name="M_BtnVid" Grid.Column="1" Content="Browse" Margin="10,0,0,0" Height="40" Background="{DynamicResource AccentBrush}" Foreground="White" BorderThickness="0" Cursor="Hand"/></Grid>

                                <TextBlock Text="2. Select audio file (Drop file here)" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,8"/>
                                <Grid Margin="0,0,0,20"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="120"/></Grid.ColumnDefinitions><TextBox x:Name="M_InAudio" IsReadOnly="True"/><Button x:Name="M_BtnAud" Grid.Column="1" Content="Browse" Margin="10,0,0,0" Height="40" Background="{DynamicResource AccentBrush}" Foreground="White" BorderThickness="0" Cursor="Hand"/></Grid>

                                <TextBlock Text="3. Save target as..." FontSize="14" FontWeight="SemiBold" Margin="0,0,0,8"/>
                                <Grid Margin="0,0,0,10"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="120"/></Grid.ColumnDefinitions><TextBox x:Name="M_OutFile" IsReadOnly="True" /><Button x:Name="M_BtnOut" Grid.Column="1" Content="Select" Margin="10,0,0,0" Height="40" Background="#4B5563" Foreground="White" BorderThickness="0" Cursor="Hand"/></Grid>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <TabItem x:Name="TabDownload">
                <TabItem.Header>
                    <TextBlock Text="Download" ToolTip="Download videos and audio from YouTube and supported sites" />
                </TabItem.Header>
                <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="15">
                    <StackPanel>
                        <Border Background="{DynamicResource CardBrush}" BorderThickness="1" CornerRadius="8" Padding="20" Margin="0,0,0,15">
                            <StackPanel>
                                <TextBlock Text="Web Video / Audio Link (YouTube, SoundCloud, Arte, etc.)" FontWeight="Bold" FontSize="16" Margin="0,0,0,5"/>
                                <TabControl x:Name="Y_InputTabs" Background="Transparent" BorderThickness="0" Margin="0,0,0,10" Padding="0">
                                    <TabItem Style="{StaticResource SubTabStyle}">
                                        <TabItem.Header><TextBlock Text="Single Link"/></TabItem.Header>
                                        <Grid Margin="0,10,0,5">
                                            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="130"/></Grid.ColumnDefinitions>
                                            <TextBox x:Name="Y_Link" Text="https://" Height="45"/>
                                            <Button x:Name="Y_BtnPreview" Grid.Column="1" Content="Fetch Video Info" Margin="10,0,0,0" Height="45" Background="#8B5CF6" Foreground="White" BorderThickness="0" Cursor="Hand"/>
                                        </Grid>
                                    </TabItem>
                                    <TabItem Style="{StaticResource SubTabStyle}">
                                        <TabItem.Header><TextBlock Text="Batch File (.txt)"/></TabItem.Header>
                                        <Grid Margin="0,10,0,5">
                                            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="130"/></Grid.ColumnDefinitions>
                                            <TextBox x:Name="Y_BatchFile" ToolTip="Select a .txt file containing one URL per line" IsReadOnly="True" Text="Select batch text file..." Height="45" Cursor="Arrow"/>
                                            <Button x:Name="Y_BtnBatchBrowse" Grid.Column="1" Content="Browse .txt" Margin="10,0,0,0" Height="45" Background="#10B981" Foreground="White" BorderThickness="0" Cursor="Hand"/>
                                        </Grid>
                                    </TabItem>
                                </TabControl>
                                
                                <Grid>
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="130"/></Grid.ColumnDefinitions>
                                    <TextBox x:Name="Y_OutDir" ToolTip="The folder where downloaded files will be saved" Text="Select target folder..." IsReadOnly="True" Cursor="Arrow"/>
                                    <Button x:Name="Y_BtnOut" Grid.Column="1" Content="Select Folder" Margin="10,0,0,0" Height="40" Background="#4B5563" Foreground="White" BorderThickness="0" Cursor="Hand"/>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <TabControl Background="Transparent" BorderThickness="0">
                            <TabItem Style="{StaticResource SubTabStyle}">
                                <TabItem.Header><TextBlock Text="Format &amp; Quality"/></TabItem.Header>
                                <Border Background="{DynamicResource BgBrush}" CornerRadius="8" Padding="20">
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>

                                        <Border Grid.Column="0" Background="{DynamicResource CardBrush}" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}" Padding="15" Margin="0,0,10,0">
                                            <StackPanel>
                                                <TextBlock Text="Media Request" FontWeight="Bold" Foreground="{DynamicResource TextBrush}" Margin="0,0,0,10"/>
                                                <StackPanel Margin="0,0,0,15">
                                                    <TextBlock Text="Download Type" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                    <ComboBox x:Name="Y_Type" SelectedIndex="0">
                                                        <ComboBoxItem>Video</ComboBoxItem>
                                                        <ComboBoxItem>Audio Only</ComboBoxItem>
                                                    </ComboBox>
                                                </StackPanel>
                                                <StackPanel>
                                                    <TextBlock Text="Quality / Resolution" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                    <ComboBox x:Name="Y_Res" SelectedIndex="0">
                                                        <ComboBoxItem>Best Possible</ComboBoxItem>
                                                        <ComboBoxItem>Max 1080p</ComboBoxItem>
                                                        <ComboBoxItem>Max 720p</ComboBoxItem>
                                                        <ComboBoxItem>Audio Only (Highest Quality)</ComboBoxItem>
                                                    </ComboBox>
                                                </StackPanel>
                                            </StackPanel>
                                        </Border>

                                        <Border Grid.Column="1" Background="{DynamicResource CardBrush}" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}" Padding="15" Margin="10,0,0,0">
                                            <StackPanel>
                                                <TextBlock Text="Output Containers" FontWeight="Bold" Foreground="{DynamicResource TextBrush}" Margin="0,0,0,10"/>
                                                <StackPanel Margin="0,0,0,15">
                                                    <TextBlock Text="Target Format (Video)" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                    <ComboBox x:Name="Y_VFormat" SelectedIndex="0">
                                                        <ComboBoxItem>mp4</ComboBoxItem>
                                                        <ComboBoxItem>mkv</ComboBoxItem>
                                                        <ComboBoxItem>webm</ComboBoxItem>
                                                    </ComboBox>
                                                </StackPanel>
                                                <StackPanel>
                                                    <TextBlock Text="Target Format (Audio)" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                    <ComboBox x:Name="Y_AFormat" SelectedIndex="0">
                                                        <ComboBoxItem>mp3</ComboBoxItem>
                                                        <ComboBoxItem>m4a</ComboBoxItem>
                                                        <ComboBoxItem>flac</ComboBoxItem>
                                                        <ComboBoxItem>wav</ComboBoxItem>
                                                    </ComboBox>
                                                </StackPanel>
                                            </StackPanel>
                                        </Border>
                                    </Grid>
                                </Border>
                            </TabItem>

                            <TabItem Style="{StaticResource SubTabStyle}">
                                <TabItem.Header><TextBlock Text="Auth &amp; Advanced"/></TabItem.Header>
                                <Border Background="{DynamicResource BgBrush}" CornerRadius="8" Padding="20">
                                    <StackPanel>
                                        
                                        <Border Background="{DynamicResource CardBrush}" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}" Padding="15" Margin="0,0,0,15">
                                            <StackPanel>
                                                <TextBlock Text="Post-Processing" FontWeight="Bold" Foreground="{DynamicResource TextBrush}" Margin="0,0,0,10"/>
                                                <WrapPanel>
                                                    <CheckBox x:Name="Y_CheckMeta" Content="Embed Metadata &amp; Thumbnail" Margin="0,0,20,10" IsChecked="False"/>
                                                    <CheckBox x:Name="Y_CheckSubs" Content="Embed Subtitles" Margin="0,0,20,10"/>
                                                    <CheckBox x:Name="Y_CheckSponsor" Content="SponsorBlock (Skip Sponsors)" Margin="0,0,20,10"/>
                                                </WrapPanel>
                                            </StackPanel>
                                        </Border>

                                        <Border Background="{DynamicResource CardBrush}" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}" Padding="15" Margin="0,0,0,15">
                                            <StackPanel>
                                                <TextBlock Text="Authentication &amp; Bot Bypass" FontWeight="Bold" Foreground="{DynamicResource TextBrush}" Margin="0,0,0,15"/>
                                                
                                                <Grid Margin="0,0,0,15">
                                                    <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="40"/></Grid.ColumnDefinitions>
                                                    <CheckBox x:Name="Y_CheckCookie" Content="Use Cookies: " VerticalAlignment="Center" Margin="0,0,10,0"/>
                                                    <ComboBox x:Name="Y_CookieBrowser" Grid.Column="1" SelectedIndex="0" Width="100" Margin="0,0,15,0">
                                                        <ComboBoxItem>edge</ComboBoxItem>
                                                        <ComboBoxItem>chrome</ComboBoxItem>
                                                        <ComboBoxItem>firefox</ComboBoxItem>
                                                        <ComboBoxItem>opera</ComboBoxItem>
                                                        <ComboBoxItem>brave</ComboBoxItem>
                                                    </ComboBox>
                                                    <TextBox x:Name="Y_CookiePath" Grid.Column="2" Text="" ToolTip="Path to cookies.txt (Overrides Browser)" IsReadOnly="True" Cursor="Arrow"/>
                                                    <Button x:Name="Y_BtnCookie" Grid.Column="3" Content="Txt" Margin="10,0,0,0" Height="38" Background="#4B5563" Foreground="White" BorderThickness="0" Cursor="Hand" ToolTip="Select cookies.txt file manually"/>
                                                </Grid>

                                                <Grid>
                                                    <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                                    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                                                    <TextBlock Text="PO Token:" VerticalAlignment="Top" Margin="0,10,15,0" Foreground="{DynamicResource MutedBrush}" ToolTip="Proof of Origin Token (helps bypass bot blocks)"/>
                                                    <TextBox x:Name="Y_PoToken" Grid.Column="1" Height="60" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" ToolTip="Paste manual token here"/>
                                                    <CheckBox x:Name="Y_CheckAutoPoToken" Grid.Row="1" Grid.Column="1" Content="Auto-retrieve PO Token (Let yt-dlp handle it)" Margin="0,10,0,0" Foreground="{DynamicResource AccentBrush}" FontWeight="SemiBold"/>
                                                </Grid>
                                            </StackPanel>
                                        </Border>

                                        <Border Background="{DynamicResource CardBrush}" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}" Padding="15">
                                            <StackPanel>
                                                <CheckBox x:Name="Y_CheckCustomParams" Content="Add custom yt-dlp params (Live Preview)" FontWeight="Bold" Margin="0,0,0,10" IsChecked="False"/>
                                                <StackPanel x:Name="Y_CustomParamsPanel" Visibility="Collapsed">
                                                    <TextBlock Text="Final yt-dlp command preview:" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                    <TextBox x:Name="Y_ParamsPreview" IsReadOnly="True" TextWrapping="Wrap" Background="#0F172A" Foreground="#10B981" FontFamily="Consolas" FontSize="13" Padding="12" Margin="0,0,0,10" MinHeight="60"/>
                                                    <TextBlock Text="Add extra arguments (e.g. --limit-rate 5M):" FontSize="13" Foreground="{DynamicResource MutedBrush}" Margin="0,0,0,5"/>
                                                    <TextBox x:Name="Y_CustomParams"/>
                                                </StackPanel>
                                            </StackPanel>
                                        </Border>

                                    </StackPanel>
                                </Border>
                            </TabItem>
                        </TabControl>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <TabItem x:Name="TabSpecial">
                <TabItem.Header>
                    <TextBlock Text="Special" ToolTip="Advanced AI tools, repair utilities, and visualizers" />
                </TabItem.Header>
                <TabControl x:Name="SpecialSubTabs" Background="Transparent" BorderThickness="0" Margin="15">
                    
                    <TabItem Style="{StaticResource SubTabStyle}">
                        <TabItem.Header><TextBlock Text="AI Transcriber"/></TabItem.Header>
                        <Border Background="{DynamicResource CardBrush}" CornerRadius="8" Padding="20">
                            <StackPanel>
                                <TextBlock Text="AI Voice-to-Text (Whisper)" FontWeight="Bold" FontSize="18" Margin="0,0,0,15" Foreground="#8B5CF6"/>
                                
                                <TextBlock Text="1. Input Media (Drop file here)" FontSize="13" Margin="0,5,0,5"/>
                                <Grid Margin="0,0,0,15">
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="100"/></Grid.ColumnDefinitions>
                                    <TextBox x:Name="S_ScribeIn" IsReadOnly="True"/>
                                    <Button x:Name="S_BtnScribeIn" Grid.Column="1" Content="Browse" Margin="10,0,0,0" Height="40" Background="#8B5CF6" Foreground="White" BorderThickness="0" Cursor="Hand"/>
                                </Grid>
                                
                                <Grid Margin="0,0,0,15">
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                    <StackPanel Margin="0,0,10,0">
                                        <TextBlock Text="AI Model (Speed vs Quality)" FontSize="13" Margin="0,0,0,5"/>
                                        <ComboBox x:Name="S_ScribeModel" SelectedIndex="1">
                                            <ComboBoxItem>tiny (Fastest, Lowest Quality)</ComboBoxItem>
                                            <ComboBoxItem>base (Default)</ComboBoxItem>
                                            <ComboBoxItem>small</ComboBoxItem>
                                            <ComboBoxItem>medium</ComboBoxItem>
                                            <ComboBoxItem>large-v3 (Slowest, Best Quality)</ComboBoxItem>
                                        </ComboBox>
                                    </StackPanel>
                                    <StackPanel Grid.Column="1" Margin="10,0,0,0">
                                        <TextBlock Text="Language (Audio)" FontSize="13" Margin="0,0,0,5"/>
                                        <ComboBox x:Name="S_ScribeLang" SelectedIndex="0">
                                            <ComboBoxItem>Auto-Detect</ComboBoxItem>
                                            <ComboBoxItem>English</ComboBoxItem>
                                            <ComboBoxItem>German</ComboBoxItem>
                                            <ComboBoxItem>Spanish</ComboBoxItem>
                                            <ComboBoxItem>French</ComboBoxItem>
                                            <ComboBoxItem>Italian</ComboBoxItem>
                                            <ComboBoxItem>Dutch</ComboBoxItem>
                                            <ComboBoxItem>Russian</ComboBoxItem>
                                        </ComboBox>
                                    </StackPanel>
                                </Grid>

                                <Grid Margin="0,0,0,15">
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                    <StackPanel Margin="0,0,10,0">
                                        <TextBlock Text="Output Format" FontSize="13" Margin="0,0,0,5"/>
                                        <ComboBox x:Name="S_ScribeFormat" SelectedIndex="0">
                                            <ComboBoxItem>txt (Text Document)</ComboBoxItem>
                                            <ComboBoxItem>srt (Subtitles)</ComboBoxItem>
                                            <ComboBoxItem>vtt (Web Subtitles)</ComboBoxItem>
                                            <ComboBoxItem>json</ComboBoxItem>
                                        </ComboBox>
                                    </StackPanel>
                                    <StackPanel Grid.Column="1" Margin="10,0,0,0">
                                        <TextBlock Text="Task" FontSize="13" Margin="0,0,0,5"/>
                                        <ComboBox x:Name="S_ScribeTask" SelectedIndex="0">
                                            <ComboBoxItem>Transcribe (Original Language)</ComboBoxItem>
                                            <ComboBoxItem>Translate (To English)</ComboBoxItem>
                                        </ComboBox>
                                    </StackPanel>
                                </Grid>

                                <CheckBox x:Name="S_CheckBurn" Content="Burn subtitles directly into a new video file (Requires .srt or .vtt)" Margin="0,10,0,0" Foreground="{DynamicResource TextBrush}" FontWeight="SemiBold"/>
                            </StackPanel>
                        </Border>
                    </TabItem>
                    
                    <TabItem x:Name="TabSpecialUpscale" Style="{StaticResource SubTabStyle}">
                        <TabItem.Header><TextBlock Text="AI Upscaler"/></TabItem.Header>
                        <Border Background="{DynamicResource CardBrush}" CornerRadius="8" Padding="20">
                            <StackPanel>
                                <TextBlock Text="AI Image Upscaling (Upscayl)" FontWeight="Bold" FontSize="18" Margin="0,0,0,15" Foreground="#10B981"/>
                                
                                <TextBlock Text="Select Image (Drop file here)" FontSize="13" Margin="0,5,0,5"/>
                                <Grid Margin="0,0,0,15"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="100"/></Grid.ColumnDefinitions>
                                    <TextBox x:Name="S_UpscaleIn" IsReadOnly="True"/><Button x:Name="S_BtnUpscaleIn" Grid.Column="1" Content="Browse" Margin="10,0,0,0" Height="40" Background="#10B981" Foreground="White" BorderThickness="0" Cursor="Hand"/></Grid>

                                <Grid Margin="0,0,0,15">
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                    <StackPanel Margin="0,0,10,0">
                                        <TextBlock Text="AI Model" FontSize="13" Margin="0,0,0,5"/>
                                        <ComboBox x:Name="S_UpscaleModel" SelectedIndex="0">
                                            <ComboBoxItem>upscayl-standard-4x (Standard)</ComboBoxItem>
                                            <ComboBoxItem>digital-art-4x (Digital Art / Anime)</ComboBoxItem>
                                            <ComboBoxItem>remacri-4x (Highly Detailed)</ComboBoxItem>
                                            <ComboBoxItem>ultramix-balanced-4x (Balanced)</ComboBoxItem>
                                            <ComboBoxItem>ultrasharp-4x (Sharp Edges)</ComboBoxItem>
                                            <ComboBoxItem>high-fidelity-4x (High Fidelity)</ComboBoxItem>
                                            <ComboBoxItem>upscayl-lite-4x (Lite / Fast)</ComboBoxItem>
                                        </ComboBox>
                                    </StackPanel>
                                    <StackPanel Grid.Column="1" Margin="10,0,0,0">
                                        <TextBlock Text="Scale Factor" FontSize="13" Margin="0,0,0,5"/>
                                        <ComboBox x:Name="S_UpscaleScale" SelectedIndex="2">
                                            <ComboBoxItem>2x</ComboBoxItem>
                                            <ComboBoxItem>3x</ComboBoxItem>
                                            <ComboBoxItem>4x</ComboBoxItem>
                                        </ComboBox>
                                    </StackPanel>
                                </Grid>

                                <TextBlock Text="Output Directory (Defaults to 'special\upscaled')" FontSize="13" Margin="0,5,0,5"/>
                                <Grid Margin="0,0,0,10"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="100"/></Grid.ColumnDefinitions>
                                    <TextBox x:Name="S_UpscaleOutDir" IsReadOnly="True" Cursor="Arrow" Text="Select target folder..."/>
                                    <Button x:Name="S_BtnUpscaleOut" Grid.Column="1" Content="Browse" Margin="10,0,0,0" Height="40" Background="#4B5563" Foreground="White" BorderThickness="0" Cursor="Hand"/>
                                </Grid>
                            </StackPanel>
                        </Border>
                    </TabItem>
                    
                    <TabItem Style="{StaticResource SubTabStyle}">
                        <TabItem.Header><TextBlock Text="Visualizer"/></TabItem.Header>
                        <Border Background="{DynamicResource CardBrush}" CornerRadius="8" Padding="20">
                            <StackPanel>
                                <TextBlock Text="Convert Audio to Video with Visualizer" FontWeight="Bold" FontSize="18" Margin="0,0,0,15" Foreground="{DynamicResource AccentBrush}"/>
                                <TextBlock Text="1. Select Audio File (Drop file here)" FontSize="13" Margin="0,5,0,5"/>
                                <Grid Margin="0,0,0,15"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="100"/></Grid.ColumnDefinitions>
                                    <TextBox x:Name="S_VisAudio" IsReadOnly="True"/><Button x:Name="S_BtnVisAud" Grid.Column="1" Content="Browse" Margin="10,0,0,0" Height="40" Background="{DynamicResource AccentBrush}" Foreground="White" BorderThickness="0" Cursor="Hand"/></Grid>
                                
                                <TextBlock Text="2. Select Background Image (Drop optional file here)" FontSize="13" Margin="0,5,0,5"/>
                                <Grid Margin="0,0,0,15"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="100"/></Grid.ColumnDefinitions>
                                    <TextBox x:Name="S_VisImg" IsReadOnly="True"/><Button x:Name="S_BtnVisImg" Grid.Column="1" Content="Browse" Margin="10,0,0,0" Height="40" Background="#4B5563" Foreground="White" BorderThickness="0" Cursor="Hand"/></Grid>
                                
                                <TextBlock Text="Visualizer Style" FontSize="13" Margin="0,5,0,5"/>
                                <ComboBox x:Name="S_VisStyle" SelectedIndex="0" Margin="0,0,0,10">
                                    <ComboBoxItem>Waves (showwaves)</ComboBoxItem>
                                    <ComboBoxItem>Frequency Bars (showfreqs)</ComboBoxItem>
                                    <ComboBoxItem>Circular Scope (avectorscope)</ComboBoxItem>
                                </ComboBox>
                            </StackPanel>
                        </Border>
                    </TabItem>
                    
                    <TabItem Style="{StaticResource SubTabStyle}">
                        <TabItem.Header><TextBlock Text="Video Stabilizer"/></TabItem.Header>
                        <Border Background="{DynamicResource CardBrush}" CornerRadius="8" Padding="20">
                            <StackPanel>
                                <TextBlock Text="Smooth Shaky Camera Footage" FontWeight="Bold" FontSize="18" Margin="0,0,0,10" Foreground="#3B82F6"/>
                                
                                <TextBlock Text="Select Shaky Video (Drop file here)" FontSize="13" Margin="0,5,0,5"/>
                                <Grid Margin="0,0,0,15"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="100"/></Grid.ColumnDefinitions>
                                    <TextBox x:Name="S_StabIn" IsReadOnly="True"/><Button x:Name="S_BtnStabIn" Grid.Column="1" Content="Browse" Margin="10,0,0,0" Height="40" Background="#3B82F6" Foreground="White" BorderThickness="0" Cursor="Hand"/></Grid>

                                <TextBlock Text="Smoothing Level" FontSize="13" Margin="0,5,0,5"/>
                                <ComboBox x:Name="S_StabLevel" SelectedIndex="1" Margin="0,0,0,10">
                                    <ComboBoxItem>Light (Less cropping, keeps some natural motion)</ComboBoxItem>
                                    <ComboBoxItem>Normal (Balanced, standard stabilization)</ComboBoxItem>
                                    <ComboBoxItem>Heavy (Maximum smoothness, crops more edges)</ComboBoxItem>
                                </ComboBox>
                            </StackPanel>
                        </Border>
                    </TabItem>

                </TabControl>
            </TabItem>

        </TabControl>

        <StackPanel Grid.Row="2" Margin="0,20,0,0">
            <Grid Margin="0,0,0,15">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Button x:Name="BtnRun" Grid.Column="0" Content="START PROCESS" Width="160" Height="45" Background="#10B981" Foreground="White" FontWeight="Bold" FontSize="15" BorderThickness="0" Cursor="Hand" Margin="0,0,15,0" IsDefault="True"/>
                <Button x:Name="BtnCancel" Grid.Column="1" Content="CANCEL ALL" Width="110" Height="45" Background="#EF4444" Foreground="White" FontWeight="Bold" FontSize="13" BorderThickness="0" Cursor="Hand" Margin="0,0,15,0" IsEnabled="False"/>
                <Button x:Name="BtnSkip" Grid.Column="2" Content="SKIP" Width="80" Height="45" Background="#F59E0B" Foreground="White" FontWeight="Bold" FontSize="13" BorderThickness="0" Cursor="Hand" Margin="0,0,15,0" IsEnabled="False" ToolTip="Skip current file and continue queue"/>
                <Button x:Name="BtnReset" Grid.Column="3" Content="RESET" Width="90" Height="45" Background="#6B7280" Foreground="White" FontWeight="Bold" FontSize="14" BorderThickness="0" Cursor="Hand" Margin="0,0,15,0"/>
                <Button x:Name="BtnShow" Grid.Column="4" Content="Open Folder" FontSize="14" Width="120" Height="45" Background="#4B5563" Foreground="White" Visibility="Collapsed" FontWeight="Bold" BorderThickness="0" Cursor="Hand" Margin="0,0,15,0"/>
                
                <Border Grid.Column="5" Background="{DynamicResource CardBrush}" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" CornerRadius="8" Padding="15,0" HorizontalAlignment="Right" VerticalAlignment="Center" Height="45">
                    <CheckBox x:Name="CbAutoScrollLog" Content="Auto-scroll Live Log" Foreground="{DynamicResource TextBrush}" FontWeight="Bold" FontSize="13" VerticalAlignment="Center" IsChecked="True"/>
                </Border>
            </Grid>

            <Grid Margin="0,0,0,15">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <ProgressBar x:Name="PBar" Grid.Column="0" Height="12" Background="{DynamicResource BorderBrush}" Foreground="{DynamicResource AccentBrush}" BorderThickness="0" Minimum="0" Maximum="100" VerticalAlignment="Center"/>
                <TextBlock x:Name="TxtETA" Grid.Column="1" Text="ETA: --:--" TextAlignment="Right" Margin="20,0,15,0" FontSize="14" Foreground="{DynamicResource MutedBrush}" VerticalAlignment="Center" FontWeight="SemiBold"/>
                <TextBlock x:Name="StatusText" Grid.Column="2" Text="Ready." TextAlignment="Right" Margin="5,0,0,0" FontSize="14" FontWeight="Bold" VerticalAlignment="Center"/>
            </Grid>

            <Expander x:Name="ExpLog" Header="Live Log (Errors and Details)" Background="{DynamicResource CardBrush}" Foreground="{DynamicResource MutedBrush}" FontWeight="SemiBold" FontSize="13" Margin="0,5,0,0">
                <TextBox x:Name="LogBox" Background="#0F172A" Foreground="#10B981" Height="150" IsReadOnly="True" VerticalScrollBarVisibility="Visible" FontFamily="Consolas" FontSize="14" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}" Padding="15" Margin="0,10,0,0"/>
            </Expander>
        </StackPanel>
    </Grid>
</Window>
"@

    # Parse the defined XML payload into the active WPF window
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Safely bypass UIPI (Drag & Drop as Admin) by fetching the native WPF window handle once it initializes
    $window.Add_SourceInitialized({
        $helper = New-Object System.Windows.Interop.WindowInteropHelper($window)
        $handle = $helper.Handle
        [void]$script:winApi::ChangeWindowMessageFilterEx($handle, $script:WM_DROPFILES, $script:MSGFLT_ALLOW, [IntPtr]::Zero)
        [void]$script:winApi::ChangeWindowMessageFilterEx($handle, $script:WM_COPYDATA, $script:MSGFLT_ALLOW, [IntPtr]::Zero)
        [void]$script:winApi::ChangeWindowMessageFilterEx($handle, $script:WM_COPYGLOBALDATA, $script:MSGFLT_ALLOW, [IntPtr]::Zero)
    })

    # Find and map all XAML UI elements to their corresponding PowerShell variables
    $UIElements = @(
        "TaskbarProgress",
        "MainTabs", "BtnRun", "BtnShow", "BtnSettings", "BtnUpdate", "BtnCancel", "BtnSkip", "BtnReset", "StatusText", "TxtETA", "PBar", "LogBox", "CbAutoScrollLog", "TxtSubtitle",
        "TabAudio", "TabVideo", "TabImage", "TabMuxing", "TabDownload", "ExpLog",
        "A_InList", "A_OutDir", "A_BtnAdd", "A_BtnClear", "A_BtnInfo", "A_BtnOut", "A_CFormat", "A_CQual", "A_CChan", "A_CMeta", "A_CheckNorm", "A_TrimStart", "A_TrimEnd", "A_SliderTrimStart", "A_SliderTrimEnd", "A_CheckExtract", "A_CtxRemove", "A_CtxClear", "A_CheckCustomParams", "A_CustomParamsPanel", "A_ParamsPreview", "A_CustomParams",
        "V_InList", "V_OutDir", "V_OutFilename", "V_CheckSmartName", "V_BtnAdd", "V_BtnClear", "V_BtnInfo", "V_BtnOut", "V_Preset", "V_BtnSavePreset", "V_CFormat", "V_CCodec", "V_CAudio", "V_CSub", "V_CRes", "V_CFPS", "V_CVol", "V_CSpeed", "V_AudioDelay", "V_CHWAccel", "V_TrimStart", "V_TrimEnd", "V_SliderTrimStart", "V_SliderTrimEnd", "V_SliderCRF", "V_CRFText", "V_CRFDesc", "V_SubPath", "V_BtnSub", "V_CAudioTracks", "V_CheckTargetSize", "V_TargetSizeMB", "V_CtxRemove", "V_CtxClear", "V_CheckCustomParams", "V_CustomParamsPanel", "V_ParamsPreview", "V_CustomParams", "V_BtnGenPreview", "V_PreviewScroll", "V_PreviewStack",
        "I_InList", "I_OutDir", "I_BtnAdd", "I_BtnClear", "I_BtnInfo", "I_BtnOut", "I_CFormat", "I_CQual", "I_CRes", "I_CheckMeta", "I_CtxRemove", "I_CtxClear",
        "M_InVideo", "M_InAudio", "M_OutFile", "M_BtnVid", "M_BtnAud", "M_BtnOut",
        "Y_InputTabs", "Y_BatchFile", "Y_BtnBatchBrowse", "Y_Link", "Y_BtnPreview", "Y_OutDir", "Y_BtnOut", "Y_Type", "Y_Res", "Y_VFormat", "Y_AFormat", "Y_CheckMeta", "Y_CheckSubs", "Y_CheckSponsor", "Y_CheckCustomParams", "Y_CustomParamsPanel", "Y_CustomParams", "Y_ParamsPreview", "Y_CheckCookie", "Y_CookiePath", "Y_BtnCookie", "Y_CookieBrowser", "Y_PoToken", "Y_CheckAutoPoToken",
        "TabSpecial", "SpecialSubTabs", "S_VisAudio", "S_VisImg", "S_VisStyle", "S_BtnVisAud", "S_BtnVisImg", "S_StabIn", "S_BtnStabIn", "S_StabLevel",
        "S_ScribeIn", "S_BtnScribeIn", "S_ScribeLang", "S_CheckBurn", "S_ScribeFormat", "S_ScribeModel", "S_ScribeTask",
        "TabSpecialUpscale", "S_UpscaleIn", "S_BtnUpscaleIn", "S_UpscaleModel", "S_UpscaleScale", "S_UpscaleOutDir", "S_BtnUpscaleOut",
        "V_UseHandbrake"
    )
    foreach ($element in $UIElements) { Set-Variable -Name $element -Value $window.FindName($element) -Scope Script }

    # ==============================================================================
    # 5. UI BEHAVIOR & HELPER FUNCTIONS
    # ==============================================================================

    # Function to apply loaded settings (Theme, Global Default Directories) to the UI
    function Enable-Config {
        $bc = New-Object System.Windows.Media.BrushConverter
        if ($Config.Theme -eq "Dark") {
            $window.Resources["BgBrush"] = $bc.ConvertFromString("#0F172A")
            $window.Resources["CardBrush"] = $bc.ConvertFromString("#1E293B")
            $window.Resources["TextBrush"] = $bc.ConvertFromString("#F8FAFC")
            $window.Resources["MutedBrush"] = $bc.ConvertFromString("#94A3B8")
            $window.Resources["BorderBrush"] = $bc.ConvertFromString("#334155")
            $window.Resources["InputBgBrush"] = $bc.ConvertFromString("#0F172A")
        }
        else {
            $window.Resources["BgBrush"] = $bc.ConvertFromString("#F3F4F6")
            $window.Resources["CardBrush"] = $bc.ConvertFromString("#FFFFFF")
            $window.Resources["TextBrush"] = $bc.ConvertFromString("#1F2937")
            $window.Resources["MutedBrush"] = $bc.ConvertFromString("#6B7280")
            $window.Resources["BorderBrush"] = $bc.ConvertFromString("#D1D5DB")
            $window.Resources["InputBgBrush"] = $bc.ConvertFromString("#FFFFFF")
        }
        
        $window.Topmost = [bool]$Config.AlwaysOnTop

        if ($Config.DefaultOutDir -and (Test-Path $Config.DefaultOutDir)) {
            $A_OutDir.Text = $Config.DefaultOutDir
            $V_OutDir.Text = $Config.DefaultOutDir
            $I_OutDir.Text = $Config.DefaultOutDir
            $Y_OutDir.Text = $Config.DefaultOutDir
        }
    }
    Enable-Config

    # Restore Whisper AI model from config
    $modelItemIdx = 1
    for ($i = 0; $i -lt $S_ScribeModel.Items.Count; $i++) {
        if ($S_ScribeModel.Items[$i].Content -match "^$($Config.WhisperModel)") {
            $modelItemIdx = $i; break
        }
    }
    $S_ScribeModel.SelectedIndex = $modelItemIdx

    # Verify all tool paths are available, and prompt to install them if missing via WinGet
    function MissingToolsCheck {
        Find-Tools

        if ($script:State.ffmpegFound -and $script:State.handbrakeFound -and $script:isWinGetVersion -and $script:State.jsRuntimeFound) {
            $StatusText.Text = "Ready."
            $StatusText.Foreground = $window.Resources["TextBrush"]
            return
        }

        $StatusText.Text = "WinGet dependencies missing or incomplete!"
        $StatusText.Foreground = "#EF4444" 

        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        $win = New-Object System.Windows.Window
        $win.Title = "Missing System Dependencies"
        $win.SizeToContent = "WidthAndHeight"; $win.WindowStartupLocation = "CenterScreen"; $win.ResizeMode = "NoResize"; $win.Background = $window.Resources["BgBrush"]
    
        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.Margin = 20
        $tbMsg = New-Object System.Windows.Controls.TextBlock
        $tbMsg.TextWrapping = "Wrap"; $tbMsg.MaxWidth = 450; $tbMsg.Foreground = $window.Resources["TextBrush"]; $tbMsg.Margin = "0,0,0,15"
    
        if ($script:State.ytdlpFound -and -not $script:isWinGetVersion) {
            $tbMsg.Inlines.Add("A portable version of yt-dlp was found, but the official WinGet version is missing.`n`nIt is highly recommended to install the WinGet version to ensure 4K SABR support works correctly.`n`n")
        }
        else {
            $tbMsg.Inlines.Add("Required tools are missing or not installed via WinGet.`n`n")
        }

        if (-not $isAdmin) {
            $tbAdmin = New-Object System.Windows.Documents.Run
            $tbAdmin.Text = "Attention! Run as Admin to use Auto-Install.`n`n"; $tbAdmin.Foreground = "#EF4444"; $tbAdmin.FontWeight = "Bold"
            $tbMsg.Inlines.Add($tbAdmin)
        }

        if (-not $script:isWinGetVersion) { $tbMsg.Inlines.Add("- yt-dlp (WinGet version) -> Missing`n") }
        if (-not $script:State.ffmpegFound) { $tbMsg.Inlines.Add("- FFmpeg -> Missing`n") }
        if (-not $script:State.handbrakeFound) { $tbMsg.Inlines.Add("- HandBrakeCLI -> Missing`n") }
        if (-not $script:State.jsRuntimeFound) { $tbMsg.Inlines.Add("- Node.js / Deno -> Missing`n") }

        [void]$sp.Children.Add($tbMsg)
        $btnSp = New-Object System.Windows.Controls.StackPanel
        $btnSp.Orientation = "Horizontal"; $btnSp.HorizontalAlignment = "Right"

        $btnAuto = New-Object System.Windows.Controls.Button
        $btnAuto.Content = "Install WinGet Version"; $btnAuto.Width = 160; $btnAuto.Height = 35; $btnAuto.Margin = "0,0,10,0"; $btnAuto.Background = "#10B981"; $btnAuto.Foreground = "White"; $btnAuto.BorderThickness = 0; $btnAuto.Cursor = "Hand"
    
        if (-not $isAdmin) { $btnAuto.IsEnabled = $false; $btnAuto.Opacity = 0.5 } 
        else { $btnAuto.Add_Click({ $win.DialogResult = $true; $win.Close() }) }
    
        $btnCancel = New-Object System.Windows.Controls.Button
        $btnCancel.Content = "Use Existing / Close"; $btnCancel.Width = 130; $btnCancel.Height = 35; $btnCancel.Background = "#6B7280"; $btnCancel.Foreground = "White"; $btnCancel.BorderThickness = 0; $btnCancel.Cursor = "Hand"
        $btnCancel.Add_Click({ $win.DialogResult = $false; $win.Close() })

        [void]$btnSp.Children.Add($btnAuto); [void]$btnSp.Children.Add($btnCancel); [void]$sp.Children.Add($btnSp)
        $win.Content = $sp

        # If user accepts, run WinGet installs in the background
        if ($win.ShowDialog() -eq $true) {
            # Validate WinGet is installed before proceeding
            if (-not (Get-Command "winget" -ErrorAction SilentlyContinue)) {
                [void][System.Windows.MessageBox]::Show("Windows Package Manager (winget) is not installed or not recognized on your system.`n`nPlease install the 'App Installer' from the Microsoft Store, or install the missing tools manually.", "Winget Not Found", 0, 16)
                return
            }

            $StatusText.Text = "Installing WinGet dependencies..."

            if (-not $script:isWinGetVersion) {
                $ytProc = Start-Process winget -ArgumentList "install --id yt-dlp.yt-dlp --silent --force --accept-source-agreements --accept-package-agreements" -Wait -PassThru
                if ($ytProc.ExitCode -ne 0) { $ytProc.Dispose(); [void][System.Windows.MessageBox]::Show("yt-dlp installation cancelled or failed.", "Installation Aborted", 0, 48); return }
                $ytProc.Dispose()
            }

            if (-not $script:State.ffmpegFound) {
                $ffProc = Start-Process winget -ArgumentList "install --id Gyan.FFmpeg --silent --force --accept-source-agreements --accept-package-agreements" -Wait -PassThru
                if ($ffProc.ExitCode -ne 0) { $ffProc.Dispose(); [void][System.Windows.MessageBox]::Show("FFmpeg installation cancelled or failed.", "Installation Aborted", 0, 48); return }
                $ffProc.Dispose()
            }

            if (-not $script:State.handbrakeFound) {
                $hbProc = Start-Process winget -ArgumentList "install --id HandBrake.HandBrake.CLI --silent --force --accept-source-agreements --accept-package-agreements" -Wait -PassThru
                if ($hbProc.ExitCode -ne 0) { $hbProc.Dispose(); [void][System.Windows.MessageBox]::Show("HandBrakeCLI installation cancelled or failed.", "Installation Aborted", 0, 48); return }
                $hbProc.Dispose()
            }

            if (-not $script:State.jsRuntimeFound) {
                $nodeProc = Start-Process winget -ArgumentList "install --id OpenJS.NodeJS --silent --force --accept-source-agreements --accept-package-agreements" -Wait -PassThru
                if ($nodeProc.ExitCode -ne 0) { $nodeProc.Dispose(); [void][System.Windows.MessageBox]::Show("Node.js installation cancelled or failed.", "Installation Aborted", 0, 48); return }
                $nodeProc.Dispose()
            }
        
            $env:PATH = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
            Find-Tools
        
            if ($script:State.ffmpegFound -and $script:State.handbrakeFound -and $script:isWinGetVersion -and $script:State.jsRuntimeFound) {
                $StatusText.Text = "WinGet Tools Ready."; $StatusText.Foreground = "#10B981"
                [void][System.Windows.MessageBox]::Show("Dependencies successfully installed!", "Success", 0, 64)
            }
            else {
                $StatusText.Text = "Installation Incomplete."; $StatusText.Foreground = "#EF4444"
                [void][System.Windows.MessageBox]::Show("Tools installed but not detected. Restart might be required.", "Warning", 0, 48)
            }
        }
    }

    # Queue resuming logic initialized once UI finishes rendering
    $window.Add_ContentRendered({ 
            $window.Dispatcher.InvokeAsync([Action] {
        
                    MissingToolsCheck 

                    # Dynamic Hardware Acceleration Detection
                    if ($script:State.ffmpegFound) {
                        try {
                            $hwOut = & $script:State.ffmpeg -hwaccels 2>&1 | Out-String
                            if ($hwOut -notmatch "cuda|nvenc") { 
                                ($V_CHWAccel.Items[1] -as [System.Windows.Controls.ComboBoxItem]).IsEnabled = $false
                                ($V_CHWAccel.Items[1] -as [System.Windows.Controls.ComboBoxItem]).ToolTip = "NVIDIA GPU support not detected."
                            }
                            if ($hwOut -notmatch "amf") { 
                                ($V_CHWAccel.Items[2] -as [System.Windows.Controls.ComboBoxItem]).IsEnabled = $false
                                ($V_CHWAccel.Items[2] -as [System.Windows.Controls.ComboBoxItem]).ToolTip = "AMD GPU support not detected."
                            }
                            if ($hwOut -notmatch "qsv") { 
                                ($V_CHWAccel.Items[3] -as [System.Windows.Controls.ComboBoxItem]).IsEnabled = $false
                                ($V_CHWAccel.Items[3] -as [System.Windows.Controls.ComboBoxItem]).ToolTip = "Intel GPU support not detected."
                            }
                        }
                        catch {
                            Write-CrashLog "Hardware Acceleration Detection failed: $($_.Exception.Message)"
                        }
                    } 

                    # 1. Background Pre-fetch for yt-dlp supported sites
                    if ($null -eq $script:State.SupportedSitesCache) {
                        [void][System.Threading.Tasks.Task]::Run([Action] {
                                try {
                                    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                                    $rawUrl = "https://raw.githubusercontent.com/yt-dlp/yt-dlp/master/supportedsites.md"
                                    $script:State.SupportedSitesCache = Invoke-RestMethod -Uri $rawUrl -UseBasicParsing -ErrorAction Stop
                                }
                                catch { $script:State.SupportedSitesCache = "fallback_offline" }
                            })
                    }

                    # 2. Safe UI Initialization for Image Tab (Grey-out Logic)
                    if ($null -ne $I_CFormat) {
                        $I_CFormat.Add_SelectionChanged({
                                $fmt = Get-CbVal $I_CFormat
                                if ($fmt -match "JPG" -or $fmt -match "WEBP") {
                                    $I_CQual.IsEnabled = $true; $I_CQual.Opacity = 1.0
                                }
                                else {
                                    $I_CQual.IsEnabled = $false; $I_CQual.Opacity = 0.4
                                }
                            })
                        $I_CFormat.SelectedIndex = 0
                    }

                    # 3. Safe UI Initialization for all Drag & Drop TextBoxes
                    foreach ($tb in @($M_InVideo, $M_InAudio, $S_VisAudio, $S_VisImg, $S_StabIn, $S_ScribeIn, $S_UpscaleIn, $Y_BatchFile)) {
                        if ($null -ne $tb) {
                            $tb.Add_PreviewDragOver($DragEnterHandler)
                            $tb.Add_DragLeave($DragLeaveHandler)
                            $tb.Add_Drop($TextBoxDropHandler)
                        }
                    }

                    # 4. Safe UI Initialization for ALL "Special" Tab Browse Buttons
                    if ($null -ne $S_BtnVisAud) { $S_BtnVisAud.Add_Click({ $fd = New-Object System.Windows.Forms.OpenFileDialog; $fd.InitialDirectory = $ScriptDir; $fd.Filter = "Audio|*.mp3;*.wav;*.m4a;*.flac;*.ogg|All|*.*"; if ($fd.ShowDialog() -eq "OK") { $S_VisAudio.Text = $fd.FileName } }) }
                    if ($null -ne $S_BtnVisImg) { $S_BtnVisImg.Add_Click({ $fd = New-Object System.Windows.Forms.OpenFileDialog; $fd.InitialDirectory = $ScriptDir; $fd.Filter = "Images|*.jpg;*.png;*.jpeg|All|*.*"; if ($fd.ShowDialog() -eq "OK") { $S_VisImg.Text = $fd.FileName } }) }
                    if ($null -ne $S_BtnStabIn) { $S_BtnStabIn.Add_Click({ $fd = New-Object System.Windows.Forms.OpenFileDialog; $fd.InitialDirectory = $ScriptDir; $fd.Filter = "Video|*.mp4;*.mkv;*.mov;*.avi|All|*.*"; if ($fd.ShowDialog() -eq "OK") { $S_StabIn.Text = $fd.FileName } }) }
                    if ($null -ne $S_BtnScribeIn) { $S_BtnScribeIn.Add_Click({ $fd = New-Object System.Windows.Forms.OpenFileDialog; $fd.InitialDirectory = $ScriptDir; $fd.Filter = "Media Files|*.mp4;*.mkv;*.mp3;*.wav;*.m4a;*.ogg;*.flac|All|*.*"; if ($fd.ShowDialog() -eq "OK") { $S_ScribeIn.Text = $fd.FileName } }) }
                    if ($null -ne $S_BtnUpscaleIn) { $S_BtnUpscaleIn.Add_Click({ $fd = New-Object System.Windows.Forms.OpenFileDialog; $fd.InitialDirectory = $ScriptDir; $fd.Filter = "Images|*.jpg;*.png;*.jpeg;*.webp|All|*.*"; if ($fd.ShowDialog() -eq "OK") { $S_UpscaleIn.Text = $fd.FileName } }) }
                    # 5. Attempt to recover crashed/incomplete queues from the last run
                    if (Test-Path $QueueFile) {
                        $ans = [System.Windows.MessageBox]::Show("Unfinished jobs were found from a previous session.`n`nWould you like to resume processing them?", "Resume Queue", "YesNo", 32)
                        if ($ans -eq "Yes") {
                            try {
                                $rawContent = Get-Content $QueueFile -Raw
                                if ([string]::IsNullOrWhiteSpace($rawContent)) { throw "Queue file is empty." }
                                $savedQueue = $rawContent | ConvertFrom-Json
                                if ($null -eq $savedQueue) { throw "Queue file is invalid." }
                                foreach ($sq in $savedQueue) {
                                    $script:State.BatchQueue += @{
                                        Args            = [string[]]$sq.Args
                                        SafeArgs        = [string[]]$sq.SafeArgs
                                        HasCustomParams = [bool]$sq.HasCustomParams
                                        Retried         = [bool]$sq.Retried
                                        IsYtDlp         = $sq.IsYtDlp
                                        IsWhisper       = $sq.IsWhisper
                                        CustomTool      = $sq.CustomTool
                                        OutputDir       = $sq.OutputDir
                                        InputFile       = $sq.InputFile
                                        OutputFile      = $sq.OutputFile
                                        ListBox         = $null
                                        ListItem        = $null
                                    }
                                }
                                $BtnRun.IsEnabled = $false; $BtnUpdate.IsEnabled = $false; $BtnCancel.IsEnabled = $true; $BtnSkip.IsEnabled = $true
                                $LogBox.AppendText("[RESUME] Loaded $($script:State.BatchQueue.Count) jobs from previous session.`r`n")
                                ProcessNextJob
                            }
                            catch {
                                Write-CrashLog "Failed to parse/resume Queue. Corrupted JSON file? Error: $($_.Exception.Message)"
                                Remove-Item -LiteralPath $QueueFile -Force -ErrorAction SilentlyContinue
                            }
                        }
                        else {
                            Remove-Item -LiteralPath $QueueFile -Force -ErrorAction SilentlyContinue
                        }
                    }
        
                    $window.Cursor = [System.Windows.Input.Cursors]::Arrow

                }, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle)
        })

    # Helper function to extract text values cleanly from WPF ComboBox objects
    function Get-CbVal([System.Windows.Controls.ComboBox]$cb) {
        if ($cb.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) { return $cb.SelectedItem.Content.ToString().Trim() }
        if ($cb.SelectedItem) { return $cb.SelectedItem.ToString().Trim() }
        return "$($cb.Text)".Trim()
    }

    # Helper function to persist the pending job queue state to disk
    function Save-Queue {
        if ($script:State.BatchQueue.Count -eq 0 -or $script:State.CurrentJobIndex -ge $script:State.BatchQueue.Count) {
            if (Test-Path $QueueFile) { Remove-Item $QueueFile -Force -ErrorAction SilentlyContinue }
            return
        }
        $exportList = for ($i = $script:State.CurrentJobIndex; $i -lt $script:State.BatchQueue.Count; $i++) {
            $job = $script:State.BatchQueue[$i]
            @{
                Args            = $job.Args
                SafeArgs        = $job.SafeArgs
                HasCustomParams = $job.HasCustomParams
                Retried         = $job.Retried
                IsYtDlp         = $job.IsYtDlp
                IsWhisper       = $job.IsWhisper
                CustomTool      = $job.CustomTool
                OutputDir       = $job.OutputDir
                InputFile       = $job.InputFile
                OutputFile      = $job.OutputFile
            }
        }
        $exportList | ConvertTo-Json -Depth 5 | Set-Content $QueueFile -Force -ErrorAction SilentlyContinue
    }

    # ==============================================================================
    # 6. COMMAND LINE ARGUMENT BUILDERS
    # ==============================================================================

    # Function to build arguments for yt-dlp based on UI selections
    function Get-YtDlpArgs([bool]$IsPreview, [bool]$ExcludeCustom, [string]$PlaylistFlag, [string]$TargetLink) {
        $link = $TargetLink
        if ([string]::IsNullOrWhiteSpace($link) -or $link -eq "https://") { $link = "URL_HERE" }

        $outDir = $Y_OutDir.Text
        if ([string]::IsNullOrWhiteSpace($outDir) -or $outDir -match "Select target") { 
            if ($Y_Type.SelectedIndex -eq 1) { $outDir = Join-Path $ScriptDir "download\audio" }
            else { $outDir = Join-Path $ScriptDir "download\video" }
        }

        # Bypass YouTube bot checks utilizing web player clients 
        $extArgs = "youtube:player_client=web,default"
        if (-not $Y_CheckAutoPoToken.IsChecked) {
            $poToken = if ($Y_PoToken.Text) { $Y_PoToken.Text.Trim() } else { "" }
            if ($poToken) { $extArgs += ";po_token=web+$poToken" }
        }

        $argList = [System.Collections.Generic.List[string]]::new()
        $argList.AddRange([string[]]@(
                "--ffmpeg-location", $script:State.ffmpeg,
                "--clean-infojson",
                "--js-runtime", "deno",
                "--js-runtime", "node",
                "--remote-components", "ejs:github",
                "--extractor-args", $extArgs,
                "--force-overwrites",
                "--no-colors"
            ))

        $resCb = Get-CbVal $Y_Res
        if ($Y_Type.SelectedIndex -eq 1) { 
            $afmt = (Get-CbVal $Y_AFormat).ToLower()
            
            $aq = "0"
            if ($resCb -match "320") { $aq = "320K" }
            elseif ($resCb -match "256") { $aq = "256K" }
            elseif ($resCb -match "192") { $aq = "192K" }
            elseif ($resCb -match "128") { $aq = "128K" }
            elseif ($resCb -match "96") { $aq = "96K" }
            elseif ($resCb -match "64") { $aq = "64K" }
            elseif ($resCb -match "48") { $aq = "48K" }
            
            $argList.AddRange([string[]]@("-x", "--audio-format", $afmt, "--audio-quality", $aq))
        }
        else { 
            $vfmt = (Get-CbVal $Y_VFormat).ToLower()
            
            if ($resCb -match "2160" -or $resCb -match "4K") { $argList.AddRange([string[]]@("-f", "bestvideo[height<=2160]+bestaudio/best")) }
            elseif ($resCb -match "1440") { $argList.AddRange([string[]]@("-f", "bestvideo[height<=1440]+bestaudio/best")) }
            elseif ($resCb -match "1080") { $argList.AddRange([string[]]@("-f", "bestvideo[height<=1080]+bestaudio/best")) }
            elseif ($resCb -match "720") { $argList.AddRange([string[]]@("-f", "bestvideo[height<=720]+bestaudio/best")) }
            elseif ($resCb -match "480") { $argList.AddRange([string[]]@("-f", "bestvideo[height<=480]+bestaudio/best")) }
            elseif ($resCb -match "360") { $argList.AddRange([string[]]@("-f", "bestvideo[height<=360]+bestaudio/best")) }
            elseif ($resCb -match "240") { $argList.AddRange([string[]]@("-f", "bestvideo[height<=240]+bestaudio/best")) }
            elseif ($resCb -match "144") { $argList.AddRange([string[]]@("-f", "bestvideo[height<=144]+bestaudio/best")) }
            else { $argList.AddRange([string[]]@("-f", "bestvideo+bestaudio/best")) }
            
            $argList.AddRange([string[]]@("--merge-output-format", $vfmt))
        }

        if ($IsPreview) {
            if ($link -match "list=") { $argList.Add("[--yes-playlist / --no-playlist]") }
            else { $argList.Add("--no-playlist") }
        }
        else {
            if ($PlaylistFlag) { $argList.Add($PlaylistFlag) }
            else { $argList.Add("--no-playlist") }
        }

        if ($Y_CheckCookie.IsChecked -eq $true) {
            if ($Y_CookiePath.Text -and (Test-Path $Y_CookiePath.Text)) { $argList.AddRange([string[]]@("--cookies", "`"$($Y_CookiePath.Text)`"")) }
            else { $argList.AddRange([string[]]@("--cookies-from-browser", (Get-CbVal $Y_CookieBrowser))) }
        }

        if ($Y_CheckMeta.IsChecked) { $argList.AddRange([string[]]@("--embed-metadata", "--embed-thumbnail", "--convert-thumbnails", "jpg")) }
        if ($Y_CheckSubs.IsChecked) { $argList.Add("--embed-subs") }
        if ($Y_CheckSponsor.IsChecked) { $argList.AddRange([string[]]@("--sponsorblock-remove", "all")) }

        if (-not $ExcludeCustom -and $Y_CheckCustomParams.IsChecked -and -not [string]::IsNullOrWhiteSpace($Y_CustomParams.Text)) {
            $pattern = "`"[^`"]+`"|'[^']+'|[^ ]+"
            $customArgs = [System.Text.RegularExpressions.Regex]::Matches($Y_CustomParams.Text, $pattern) | ForEach-Object { $_.Value.Trim("'`"") }
            $argList.AddRange([string[]]$customArgs)
        }

        $argList.AddRange([string[]]@("-o", "$outDir\%(title)s.%(ext)s", $link))
        return $argList.ToArray()
    }

    # Function to build visual string representation for yt-dlp arguments
    function Update-YtDlpPreview {
        if (-not $Y_CheckCustomParams.IsChecked) { return }
        
        $cmdArgs = Get-YtDlpArgs -isPreview $true -ExcludeCustom $false -PlaylistFlag "" -TargetLink $Y_Link.Text.Trim()
        if ($null -eq $cmdArgs) { return }
        
        $formattedArgs = foreach ($a in $cmdArgs) {
            if ($a -match '\s' -and $a -notmatch '^\[.*\]$') { "`"$a`"" } else { $a }
        }
        if ($Y_ParamsPreview) { $Y_ParamsPreview.Text = "yt-dlp " + ($formattedArgs -join " ") }
    }

    # Function to build arguments for ffmpeg (Video tab) based on UI selections
    function Get-FfmpegArgs([bool]$IsPreview, [string]$inFile, [string]$outFile, [bool]$ExcludeCustom) {
        $argList = [System.Collections.Generic.List[string]]::new()
        $argList.AddRange([string[]]@("-hide_banner", "-y"))

        # Enable Hardware Decoding to alleviate CPU bottlenecking
        $hwCb = Get-CbVal $V_CHWAccel
        if ($hwCb -match "NVIDIA") { $argList.AddRange([string[]]@("-hwaccel", "cuda")) }
        elseif ($hwCb -match "Intel") { $argList.AddRange([string[]]@("-hwaccel", "qsv")) }
        elseif ($hwCb -match "AMD") { $argList.AddRange([string[]]@("-hwaccel", "d3d11va")) }

        # Safely parse and enforce that End Time is strictly greater than Start Time
        $startStr = $V_TrimStart.Text; $endStr = $V_TrimEnd.Text
        $startTs = [TimeSpan]::Zero; $endTs = [TimeSpan]::Zero
        [TimeSpan]::TryParse($startStr, [ref]$startTs) | Out-Null
        [TimeSpan]::TryParse($endStr, [ref]$endTs) | Out-Null

        if ($endTs.TotalSeconds -gt 0 -and $startTs.TotalSeconds -ge $endTs.TotalSeconds) {
            # Auto-swap times to prevent FFmpeg crash (-22 Invalid Argument)
            $startStr = "{0:D2}:{1:D2}:{2:D2}" -f $endTs.Hours, $endTs.Minutes, $endTs.Seconds
            $endStr = "{0:D2}:{1:D2}:{2:D2}" -f $startTs.Hours, $startTs.Minutes, $startTs.Seconds
        }

        if ($startStr -ne "00:00:00") { $argList.AddRange([string[]]@("-ss", $startStr)) }
        if ($endStr -ne "00:00:00") { $argList.AddRange([string[]]@("-to", $endStr)) }

        $delaySec = 0
        $useDualInput = $false
        # Handles custom audio syncing by delaying specific inputs
        if ([double]::TryParse($V_AudioDelay.Text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$delaySec) -and $delaySec -ne 0) {
            $argList.AddRange([string[]]@("-i", $inFile))
            $argList.AddRange([string[]]@("-itsoffset", $delaySec.ToString([System.Globalization.CultureInfo]::InvariantCulture), "-i", $inFile))
            $useDualInput = $true
        }
        else {
            $argList.AddRange([string[]]@("-i", $inFile))
        }

        $vf = [System.Collections.Generic.List[string]]::new()
        $af = [System.Collections.Generic.List[string]]::new()
        
        $resCb = Get-CbVal $V_CRes
        if ($resCb -match "1080p") { $vf.Add("scale=-2:1080") }
        elseif ($resCb -match "720p") { $vf.Add("scale=-2:720") }

        $speedCb = Get-CbVal $V_CSpeed
        if ($speedCb -match "0.5x") { $vf.Add("setpts=2.0*PTS"); $af.Add("atempo=0.5") }
        elseif ($speedCb -match "1.25x") { $vf.Add("setpts=0.8*PTS"); $af.Add("atempo=1.25") }
        elseif ($speedCb -match "1.5x") { $vf.Add("setpts=0.66667*PTS"); $af.Add("atempo=1.5") }
        elseif ($speedCb -match "2.0x") { $vf.Add("setpts=0.5*PTS"); $af.Add("atempo=2.0") }

        $fpsCb = Get-CbVal $V_CFPS
        if ($fpsCb -match "60") { $argList.AddRange([string[]]@("-r", "60")) }
        elseif ($fpsCb -match "30") { $argList.AddRange([string[]]@("-r", "30")) }
        elseif ($fpsCb -match "24") { $argList.AddRange([string[]]@("-r", "24")) }

        $volCb = Get-CbVal $V_CVol
        if ($volCb -match "150") { $af.Add("volume=1.5") }
        elseif ($volCb -match "Normalize") { $af.Add("loudnorm") }

        # Handles subtitle hardcoding via ffmpeg filters
        if (-not [string]::IsNullOrWhiteSpace($V_SubPath.Text)) {
            $safeSub = $V_SubPath.Text.Replace('\', '/').Replace(':', '\:').Replace('[', '\[').Replace(']', '\]').Replace("'", "\'").Replace(',', '\,')
            $vf.Add("subtitles='''$safeSub'''")
        }

        # Handle Hardware Acceleration Mapping
        $vCodecCb = (Get-CbVal $V_CCodec); $hwCb = (Get-CbVal $V_CHWAccel); $vCodec = "libx264"
        if ($vCodecCb -match "Copy") { $vCodec = "copy" }
        elseif ($vCodecCb -match "H.264") {
            if ($hwCb -match "NVIDIA") { $vCodec = "h264_nvenc" } elseif ($hwCb -match "AMD") { $vCodec = "h264_amf" } elseif ($hwCb -match "Intel") { $vCodec = "h264_qsv" } else { $vCodec = "libx264" }
        }
        elseif ($vCodecCb -match "H.265") {
            if ($hwCb -match "NVIDIA") { $vCodec = "hevc_nvenc" } elseif ($hwCb -match "AMD") { $vCodec = "hevc_amf" } elseif ($hwCb -match "Intel") { $vCodec = "hevc_qsv" } else { $vCodec = "libx265" }
        }
        elseif ($vCodecCb -match "AV1") {
            if ($hwCb -match "NVIDIA") { $vCodec = "av1_nvenc" } elseif ($hwCb -match "AMD") { $vCodec = "av1_amf" } elseif ($hwCb -match "Intel") { $vCodec = "av1_qsv" } else { $vCodec = "libsvtav1" }
        }

        if ($vCodec -match "nvenc|amf|qsv") {
            $vf.Add("format=yuv420p") 
            if ($vCodec -match "nvenc") { 
                $argList.AddRange([string[]]@("-preset", "p4", "-tune", "hq", "-spatial-aq", "1", "-multipass", "2")) 
            }
        }

        $aCodecCb = (Get-CbVal $V_CAudio); $aCodec = "aac"
        if ($aCodecCb -match "Copy") { $aCodec = "copy" } 
        elseif ($aCodecCb -match "EAC3") { $aCodec = "eac3" } 
        elseif ($aCodecCb -match "AC3") { $aCodec = "ac3" }

        # GIF Optimization: Force gif codec, drop audio, reduce framerate, and use RAM-safe palette generation
        if ($outFile -match "(?i)\.gif$") {
            $vCodec = "gif"
            $gifFilter = if ($vf.Count -gt 0) { ($vf -join ",") + "," } else { "" }
            # Force max 720p and 15fps to prevent Out Of Memory (-12) errors on large files
            $gifFilter += "fps=15,scale='min(720,iw)':-2:flags=lanczos,split[s0][s1];[s0]palettegen=stats_mode=diff[p];[s1][p]paletteuse=dither=bayer:bayer_scale=5:diff_mode=rectangle"
            
            $argList.AddRange([string[]]@("-c:v", $vCodec, "-an", "-vf", $gifFilter)) 
        }
        else {
            $argList.AddRange([string[]]@("-c:v", $vCodec, "-c:a", $aCodec))
            if ($vCodec -ne "copy" -and $vf.Count -gt 0) { $argList.AddRange([string[]]@("-vf", ($vf -join ","))) }
            if ($aCodec -ne "copy" -and $af.Count -gt 0) { $argList.AddRange([string[]]@("-af", ($af -join ","))) }
        }

        # Subtitle Handling Logic
        $subCb = Get-CbVal $V_CSub
        $mapSubs = if ($subCb -match "Copy All Subs") { $true } else { $false }
        if ($mapSubs) { $argList.AddRange([string[]]@("-c:s", "copy")) }

        # Audio and Subtitle Track mappings based on sync requirements
        if ($useDualInput) {
            if ($V_CAudioTracks.SelectedIndex -eq 1) { 
                $argList.AddRange([string[]]@("-map", "0:v", "-map", "1:a?"))
                if ($mapSubs) { $argList.AddRange([string[]]@("-map", "0:s?")) }
            }
            else { 
                $argList.AddRange([string[]]@("-map", "0:v:0", "-map", "1:a:0"))
                if ($mapSubs) { $argList.AddRange([string[]]@("-map", "0:s:0?")) }
            }
        }
        else {
            if ($V_CAudioTracks.SelectedIndex -eq 1) { 
                $argList.AddRange([string[]]@("-map", "0")) 
            }
            else { 
                $argList.AddRange([string[]]@("-map", "0:v:0", "-map", "0:a:0?"))
                if ($mapSubs) { $argList.AddRange([string[]]@("-map", "0:s?")) }
            }
        }

        # Target Size or CRF Video Compression Bitrate logic
        if ($vCodec -ne "copy") {
            if ($V_CheckTargetSize.IsChecked -and $script:State.ffprobeFound -and -not $IsPreview) {
                $pinfo = New-Object System.Diagnostics.ProcessStartInfo
                $pinfo.FileName = $script:State.ffprobe
                $pinfo.Arguments = "-v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 `"$inFile`""
                $pinfo.UseShellExecute = $false; $pinfo.RedirectStandardOutput = $true; $pinfo.RedirectStandardError = $true; $pinfo.CreateNoWindow = $true
                $pDur = [System.Diagnostics.Process]::Start($pinfo)
                
                $durStr = $pDur.StandardOutput.ReadToEnd().Trim()
                [void]$pDur.StandardError.ReadToEnd()
                
                if (-not $pDur.WaitForExit(3000)) { try { $pDur.Kill() } catch {} }
                $pDur.Dispose() # Clean up handle
                
                $d = 0
                if ([double]::TryParse($durStr, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d) -and $d -gt 0) {
                    $targetMB = [double]$V_TargetSizeMB.Text; $audioBR = 128
                    $videoBR = [math]::Max(100, [math]::Round((($targetMB * 8192) / $d) - $audioBR))
                    $argList.AddRange([string[]]@("-b:v", "${videoBR}k", "-maxrate", "$([math]::Round($videoBR * 1.5))k", "-bufsize", "$([math]::Round($videoBR * 2))k"))
                }
                else { $argList.AddRange([string[]]@("-crf", "$($V_SliderCRF.Value)")) }
            }
            elseif ($V_CheckTargetSize.IsChecked -and $IsPreview) { $argList.Add("[Target Size Bitrate Calculation]") }
            else {
                if ($vCodec -match "nvenc|amf|qsv") {
                    $argList.AddRange([string[]]@("-cq", "$($V_SliderCRF.Value)", "-b:v", "0"))
                }
                else {
                    $argList.AddRange([string[]]@("-crf", "$($V_SliderCRF.Value)"))
                }
            }
        }

        if (-not $ExcludeCustom -and $V_CheckCustomParams.IsChecked -and -not [string]::IsNullOrWhiteSpace($V_CustomParams.Text)) {
            $pattern = "`"[^`"]+`"|'[^']+'|[^ ]+"
            $customArgs = [System.Text.RegularExpressions.Regex]::Matches($V_CustomParams.Text, $pattern) | ForEach-Object { $_.Value.Trim("'`"") }
            $argList.AddRange([string[]]$customArgs)
        }

        $argList.Add($outFile)
        return @{ Args = $argList.ToArray(); UseDual = $useDualInput }
    }

    function Get-HandbrakeArgs([string]$inFile, [string]$outFile) {
        $argList = [System.Collections.Generic.List[string]]::new()
        $argList.AddRange([string[]]@("-i", $inFile, "-o", $outFile))

        # HandBrake Presets
        $resCb = Get-CbVal $V_CRes
        if ($resCb -match "1080p") { $argList.AddRange([string[]]@("--preset", "Fast 1080p30")) }
        elseif ($resCb -match "720p") { $argList.AddRange([string[]]@("--preset", "Fast 720p30")) }
        else { $argList.AddRange([string[]]@("--preset", "Super HQ 1080p30 Surround")) }

        # Constant Quality
        $argList.AddRange([string[]]@("-q", "$($V_SliderCRF.Value)"))

        # Framerate Handing (Force CFR to fix sync issues)
        $fpsCb = Get-CbVal $V_CFPS
        if ($fpsCb -match "60") { $argList.AddRange([string[]]@("--rate", "60", "--cfr")) }
        elseif ($fpsCb -match "30") { $argList.AddRange([string[]]@("--rate", "30", "--cfr")) }
        elseif ($fpsCb -match "24") { $argList.AddRange([string[]]@("--rate", "24", "--cfr")) }
        else { $argList.Add("--cfr") } # Always force constant framerate to fix phone video drift

        # Hardware Acceleration & Codecs
        $vCodecCb = Get-CbVal $V_CCodec
        $hwCb = Get-CbVal $V_CHWAccel
        
        $enc = "x264"
        if ($vCodecCb -match "H\.265") {
            if ($hwCb -match "NVIDIA") { $enc = "nvenc_h265" } elseif ($hwCb -match "AMD") { $enc = "vce_h265" } elseif ($hwCb -match "Intel") { $enc = "qsv_h265" } else { $enc = "x265" }
        }
        elseif ($vCodecCb -match "AV1") {
            if ($hwCb -match "Intel") { $enc = "qsv_av1" } else { $enc = "svt_av1" }
        }
        else {
            if ($hwCb -match "NVIDIA") { $enc = "nvenc_h264" } elseif ($hwCb -match "AMD") { $enc = "vce_h264" } elseif ($hwCb -match "Intel") { $enc = "qsv_h264" }
        }
        $argList.AddRange([string[]]@("-e", $enc))

        # Audio
        $aCodecCb = Get-CbVal $V_CAudio
        if ($aCodecCb -match "AC3" -or $aCodecCb -match "EAC3") { $argList.AddRange([string[]]@("-E", "ac3")) }
        else { $argList.AddRange([string[]]@("-E", "av_aac")) }

        # Format Container
        if ($outFile -match "\.mkv$") { $argList.AddRange([string[]]@("-f", "mkv")) }
        else { $argList.AddRange([string[]]@("-f", "av_mp4")) }

        # HandBrake auto-crops by default, which is great.
        return @{ Args = $argList.ToArray() }
    }
    # Smart Filename Generator
    function Get-SmartVideoFilename([string]$inFile) {
        if ([string]::IsNullOrWhiteSpace($inFile) -or $inFile -eq "[Preview_Video_Input]") { return "output_video" }
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inFile)
        if ($script:State.CustomFilenames.ContainsKey($inFile)) { return $script:State.CustomFilenames[$inFile] }

        if ($V_CheckSmartName.IsChecked) {
            $vCodecCb = Get-CbVal $V_CCodec
            $aCodecCb = Get-CbVal $V_CAudio
            
            $vTag = ""
            if ($vCodecCb -match "H\.264") { $vTag = "H264" }
            elseif ($vCodecCb -match "H\.265") { $vTag = "HEVC" }
            elseif ($vCodecCb -match "AV1") { $vTag = "AV1" }
            
            $aTag = ""
            if ($aCodecCb -match "AAC") { $aTag = "AAC" }
            elseif ($aCodecCb -match "EAC3") { $aTag = "EAC3" }
            elseif ($aCodecCb -match "AC3") { $aTag = "AC3" }

            if ($aTag) { $baseName = $baseName -replace '(?i)\b(dts-?hd(\s*ma)?|dts|truehd|flac|eac3|ac3|aac|mp3)\b', $aTag }
            if ($vTag) { $baseName = $baseName -replace '(?i)\b(hevc|x265|h265|x264|h264|avc|av1)\b', $vTag }
        }
        return $baseName
    }

    # Function to build visual string representation for FFmpeg (Video) arguments
    function Update-FfmpegPreview {
        
        $inFile = "[Preview_Video_Input]"
        if ($null -ne $V_InList.SelectedItem) { $inFile = $V_InList.SelectedItem.ToString() }
        elseif ($V_InList.Items.Count -gt 0) { $inFile = $V_InList.Items[0].ToString() }

        # Sync the UI Textbox to current item dynamically
        if ($inFile -ne "[Preview_Video_Input]") {
            $newName = Get-SmartVideoFilename $inFile
            if ($V_OutFilename.Text -ne $newName) {
                $script:State.IsAutoUpdatingFilename = $true
                $caret = $V_OutFilename.CaretIndex
                $V_OutFilename.Text = $newName
                $V_OutFilename.CaretIndex = $caret
                $script:State.IsAutoUpdatingFilename = $false
            }
        }

        if (-not $V_CheckCustomParams.IsChecked) { return }

        $outDir = $V_OutDir.Text
        if ([string]::IsNullOrWhiteSpace($outDir) -or $outDir -match "Select target") { 
            if ($inFile -ne "[Preview_Video_Input]") { $outDir = Split-Path $inFile -Parent }
            else { $outDir = Join-Path $ScriptDir "convert\video" }
        }
        $name = Get-SmartVideoFilename $inFile
        $fmt = (Get-CbVal $V_CFormat).ToLower()
        $outFile = Join-Path $outDir "$name.$fmt"
        
        $cmdArgs = (Get-FfmpegArgs -IsPreview $true -inFile $inFile -outFile $outFile -ExcludeCustom $false).Args
        $formattedArgs = foreach ($a in $cmdArgs) {
            if ($a -match '\s' -and $a -notmatch '^\[.*\]$') { "`"$a`"" } else { $a }
        }
        if ($V_ParamsPreview) { $V_ParamsPreview.Text = "ffmpeg " + ($formattedArgs -join " ") }
    }

    # Function to build arguments for ffmpeg (Audio tab) based on UI selections
    function Get-AudioFfmpegArgs([bool]$IsPreview, [string]$inFile, [string]$outFile, [bool]$ExcludeCustom) {
        $argList = [System.Collections.Generic.List[string]]::new()
        $argList.AddRange([string[]]@("-hide_banner", "-y"))
        
        $startStr = $A_TrimStart.Text; $endStr = $A_TrimEnd.Text
        $startTs = [TimeSpan]::Zero; $endTs = [TimeSpan]::Zero
        [TimeSpan]::TryParse($startStr, [ref]$startTs) | Out-Null
        [TimeSpan]::TryParse($endStr, [ref]$endTs) | Out-Null

        if ($endTs.TotalSeconds -gt 0 -and $startTs.TotalSeconds -ge $endTs.TotalSeconds) {
            $startStr = "{0:D2}:{1:D2}:{2:D2}" -f $endTs.Hours, $endTs.Minutes, $endTs.Seconds
            $endStr = "{0:D2}:{1:D2}:{2:D2}" -f $startTs.Hours, $startTs.Minutes, $startTs.Seconds
        }
        
        if ($startStr -ne "00:00:00") { $argList.AddRange([string[]]@("-ss", $startStr)) }
        if ($endStr -ne "00:00:00") { $argList.AddRange([string[]]@("-to", $endStr)) }
        
        $argList.AddRange([string[]]@("-i", $inFile, "-vn"))
        $fmt = (Get-CbVal $A_CFormat).ToLower()
        $audioCodec = if ($fmt -match "copy") { "copy" } elseif ($fmt -match "mp3") { "libmp3lame" } elseif ($fmt -match "flac") { "flac" } elseif ($fmt -match "wav") { "pcm_s16le" } else { "aac" }
        
        if ($audioCodec -eq "copy") {
            # If copying streams directly, bitrate/channels/filters cannot be applied
            $argList.AddRange([string[]]@("-c:a", "copy"))
            if ((Get-CbVal $A_CMeta) -match "Remove") { $argList.AddRange([string[]]@("-map_metadata", "-1")) }
            # If the user chose copy, assume output file extension should match the original audio stream natively. We use a generic fallback to m4a/aac container if unknown.
            if ($outFile -match "\.copy") { $outFile = $outFile -replace '\.copy \(extract original\)$', '.m4a' }
        }
        else {
            $qualCb = Get-CbVal $A_CQual
            if ($qualCb -match "128k") { $argList.AddRange([string[]]@("-b:a", "128k")) } elseif ($qualCb -match "192k") { $argList.AddRange([string[]]@("-b:a", "192k")) } elseif ($qualCb -match "320k") { $argList.AddRange([string[]]@("-b:a", "320k")) }
            
            $chanCb = Get-CbVal $A_CChan
            if ($chanCb -match "Mono") { $argList.AddRange([string[]]@("-ac", "1")) } elseif ($chanCb -match "Stereo") { $argList.AddRange([string[]]@("-ac", "2")) }
            
            if ((Get-CbVal $A_CMeta) -match "Remove") { $argList.AddRange([string[]]@("-map_metadata", "-1")) }
            if ($A_CheckNorm.IsChecked) { $argList.AddRange([string[]]@("-af", "loudnorm")) }
            
            $argList.AddRange([string[]]@("-c:a", $audioCodec))
        }
        
        if (-not $ExcludeCustom -and $A_CheckCustomParams.IsChecked -and -not [string]::IsNullOrWhiteSpace($A_CustomParams.Text)) {
            $pattern = "`"[^`"]+`"|'[^']+'|[^ ]+"
            $customArgs = [System.Text.RegularExpressions.Regex]::Matches($A_CustomParams.Text, $pattern) | ForEach-Object { $_.Value.Trim("'`"") }
            $argList.AddRange([string[]]$customArgs)
        }
        $argList.Add($outFile)
        return @{ Args = $argList.ToArray() }
    }

    # Function to build visual string representation for FFmpeg (Audio) arguments
    function Update-AudioFfmpegPreview {
        if (-not $A_CheckCustomParams.IsChecked) { return }
        
        $inFile = "[Preview_Audio_Input]"
        if ($null -ne $A_InList.SelectedItem) { $inFile = $A_InList.SelectedItem.ToString() }
        elseif ($A_InList.Items.Count -gt 0) { $inFile = $A_InList.Items[0].ToString() }

        $outDir = $A_OutDir.Text
        if ([string]::IsNullOrWhiteSpace($outDir) -or $outDir -match "Select target") { 
            if ($inFile -ne "[Preview_Audio_Input]") { $outDir = Split-Path $inFile -Parent }
            else { $outDir = Join-Path $ScriptDir "convert\audio" }
        }
        $name = if ($inFile -ne "[Preview_Audio_Input]") { [System.IO.Path]::GetFileNameWithoutExtension($inFile) } else { "output_audio" }
        $fmt = (Get-CbVal $A_CFormat).ToLower()
        $outFile = Join-Path $outDir "$name.$fmt"
        
        $cmdArgs = (Get-AudioFfmpegArgs -IsPreview $true -inFile $inFile -outFile $outFile -ExcludeCustom $false).Args
        $formattedArgs = foreach ($a in $cmdArgs) {
            if ($a -match '\s' -and $a -notmatch '^\[.*\]$') { "`"$a`"" } else { $a }
        }
        if ($A_ParamsPreview) { $A_ParamsPreview.Text = "ffmpeg " + ($formattedArgs -join " ") }
    }

    # Setup a DispatcherTimer to "debounce" UI updates (prevents lag when typing fast)
    $script:State.PreviewTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:State.PreviewTimer.Interval = [TimeSpan]::FromMilliseconds(250)
    $script:State.PreviewTimer.Add_Tick({
            $script:State.PreviewTimer.Stop()
            Update-AudioFfmpegPreview
            Update-FfmpegPreview
            Update-YtDlpPreview
        })

    # Generic function to force an update to all active tool previews
    function Update-AllPreviews {
        # Reset the timer on every keystroke/change. It only fires when the user stops for 250ms.
        $script:State.PreviewTimer.Stop()
        $script:State.PreviewTimer.Start()
    }

    # ==============================================================================
    # 7. EVENT BINDING (User Interface Interactions)
    # ==============================================================================

    # Restore from tray on double click
    $script:TrayIcon.add_DoubleClick({
        $window.ShowInTaskbar = $true
        $window.WindowState = [System.Windows.WindowState]::Normal
        $window.Activate()
    })

    # Map the window minimize button to hide the app into the tray
    $window.Add_StateChanged({
        if ($window.WindowState -eq [System.Windows.WindowState]::Minimized) {
            $window.ShowInTaskbar = $false
        }
    })

    # Update previews when user selections change
    $A_InList.add_SelectionChanged([System.Windows.Controls.SelectionChangedEventHandler] { Update-AudioFfmpegPreview })
    $V_InList.add_SelectionChanged([System.Windows.Controls.SelectionChangedEventHandler] { Update-FfmpegPreview })

    $MainTabs.add_SelectionChanged([System.Windows.Controls.SelectionChangedEventHandler] {
            if ($_.OriginalSource -eq $MainTabs) { 
                Update-AllPreviews 
            
                # Print browser detection log only when Download tab (Index 4) is selected for the first time
                if ($MainTabs.SelectedIndex -eq 4 -and -not $script:State.BrowserLogged) {
                    $browserName = Get-CbVal $Y_CookieBrowser
                    $LogBox.AppendText("`r`n[INFO] Auto-detected default browser for cookies: $browserName`r`n")
                    if ($CbAutoScrollLog.IsChecked) { $LogBox.ScrollToEnd() }
                    $script:State.BrowserLogged = $true
                }
            }
        })

    # Map settings changes to preview updates for the Audio Tab
    $A_CheckCustomParams.Add_Checked({ $A_CustomParamsPanel.Visibility = "Visible"; Update-AudioFfmpegPreview })
    $A_CheckCustomParams.Add_Unchecked({ $A_CustomParamsPanel.Visibility = "Collapsed"; $A_CustomParams.Clear(); Update-AudioFfmpegPreview })
    foreach ($ctrl in @($A_CFormat, $A_CQual, $A_CChan, $A_CMeta)) { $ctrl.Add_SelectionChanged({ Update-AudioFfmpegPreview }) }
    $A_CustomParams.Add_TextChanged({ Update-AudioFfmpegPreview })
    $A_TrimStart.Add_TextChanged({ Update-AudioFfmpegPreview })
    $A_TrimEnd.Add_TextChanged({ Update-AudioFfmpegPreview })
    $A_CheckNorm.Add_Checked({ Update-AudioFfmpegPreview })
    $A_CheckNorm.Add_Unchecked({ Update-AudioFfmpegPreview })
    $A_OutDir.Add_TextChanged({ Update-AudioFfmpegPreview })

    # Map settings changes to preview updates for the Video Tab
    $V_CheckCustomParams.Add_Checked({ $V_CustomParamsPanel.Visibility = "Visible"; Update-FfmpegPreview })
    $V_CheckCustomParams.Add_Unchecked({ $V_CustomParamsPanel.Visibility = "Collapsed"; $V_CustomParams.Clear(); Update-FfmpegPreview })
    foreach ($ctrl in @($V_CFormat, $V_CHWAccel, $V_CCodec, $V_CAudio, $V_CSub, $V_CRes, $V_CFPS, $V_CAudioTracks, $V_CVol, $V_CSpeed)) {
        $ctrl.Add_SelectionChanged({ Update-FfmpegPreview })
    }
    $V_CustomParams.Add_TextChanged({ Update-FfmpegPreview })
    $V_AudioDelay.Add_TextChanged({ Update-FfmpegPreview })
    $V_TrimStart.Add_TextChanged({ Update-FfmpegPreview })
    $V_TrimEnd.Add_TextChanged({ Update-FfmpegPreview })
    $V_CheckTargetSize.Add_Checked({ Update-FfmpegPreview })
    $V_CheckTargetSize.Add_Unchecked({ Update-FfmpegPreview })
    $V_TargetSizeMB.Add_TextChanged({ Update-FfmpegPreview })
    $V_SliderCRF.Add_ValueChanged({ Update-FfmpegPreview })
    $V_OutDir.Add_TextChanged({ Update-FfmpegPreview })
    
    $V_OutFilename.Add_TextChanged({
            if (-not $script:State.IsAutoUpdatingFilename -and $V_InList.SelectedItem) {
                $script:State.CustomFilenames[$V_InList.SelectedItem.ToString()] = $V_OutFilename.Text
                Update-FfmpegPreview
            }
        })
    $V_CheckSmartName.Add_Checked({
            if ($V_InList.SelectedItem) { $script:State.CustomFilenames.Remove($V_InList.SelectedItem.ToString()) }
            Update-FfmpegPreview
        })

    # Map settings changes to preview updates for the Downloader Tab
    $Y_CheckCustomParams.Add_Checked({ $Y_CustomParamsPanel.Visibility = "Visible"; Update-YtDlpPreview })
    $Y_CheckCustomParams.Add_Unchecked({ $Y_CustomParamsPanel.Visibility = "Collapsed"; $Y_CustomParams.Clear(); Update-YtDlpPreview })
    
    # Update quality options dynamically based on Audio vs Video download selection
    $Y_Type.Add_SelectionChanged({ 
            $Y_Res.Items.Clear()
        
            if ($Y_Type.SelectedIndex -eq 1) { 
                $Y_VFormat.IsEnabled = $false; $Y_VFormat.Opacity = 0.4
                $Y_AFormat.IsEnabled = $true; $Y_AFormat.Opacity = 1.0
            
                [void]$Y_Res.Items.Add("Best Possible (VBR 0)")
                [void]$Y_Res.Items.Add("320kbps (High CBR)")
                [void]$Y_Res.Items.Add("256kbps (Standard CBR)")
                [void]$Y_Res.Items.Add("192kbps (Medium CBR)")
                [void]$Y_Res.Items.Add("128kbps (Low CBR)")
                [void]$Y_Res.Items.Add("96kbps (Very Low CBR)")
                [void]$Y_Res.Items.Add("64kbps (Very Low)")
                [void]$Y_Res.Items.Add("48kbps (Minimum)")
            }
            else { 
                $Y_VFormat.IsEnabled = $true; $Y_VFormat.Opacity = 1.0
                $Y_AFormat.IsEnabled = $false; $Y_AFormat.Opacity = 0.4
            
                [void]$Y_Res.Items.Add("Best Possible (4K/8K if available)")
                [void]$Y_Res.Items.Add("Max 4K (2160p)")
                [void]$Y_Res.Items.Add("Max 1440p (2K)")
                [void]$Y_Res.Items.Add("Max 1080p (FHD)")
                [void]$Y_Res.Items.Add("Max 720p (HD)")
                [void]$Y_Res.Items.Add("Max 480p (SD)")
                [void]$Y_Res.Items.Add("Max 360p (Low)")
                [void]$Y_Res.Items.Add("Max 240p (Very Low)")
                [void]$Y_Res.Items.Add("Max 144p (Minimum)")
            }
        
            $Y_Res.SelectedIndex = 0
            Update-YtDlpPreview 
        })
    
    # Disable Image Quality combobox if format is not JPG or WEBP
    $I_CFormat.Add_SelectionChanged({
            $fmt = Get-CbVal $I_CFormat
            if ($fmt -match "JPG" -or $fmt -match "WEBP") {
                $I_CQual.IsEnabled = $true
                $I_CQual.Opacity = 1.0
            }
            else {
                $I_CQual.IsEnabled = $false
                $I_CQual.Opacity = 0.4
            }
        })
    # Trigger it once to set the initial state correctly
    $I_CFormat.SelectedIndex = -1
    $I_CFormat.SelectedIndex = 0

    $Y_Type.SelectedIndex = -1
    $Y_Type.SelectedIndex = 0

    # --- Bulletproof Auto-Detect Default Browser ---
    try {
        $tempFile = Join-Path $env:TEMP "mcp_browser_detect.html"
        "<html></html>" | Out-File -FilePath $tempFile -Encoding utf8 -Force
        
        $sig = '[DllImport("shell32.dll", CharSet = CharSet.Unicode)] public static extern uint FindExecutable(string lpFile, string lpDirectory, [Out] System.Text.StringBuilder lpResult);'
        if (-not ("WinApi.ShellDetect" -as [type])) { Add-Type -MemberDefinition $sig -Name "ShellDetect" -Namespace "WinApi" }
        
        $outBuf = New-Object System.Text.StringBuilder 1024
        [void][WinApi.ShellDetect]::FindExecutable($tempFile, $null, $outBuf)
        $realPath = $outBuf.ToString().ToLower()
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue

        if ($realPath -match "chrome") { $Y_CookieBrowser.SelectedIndex = 1 }
        elseif ($realPath -match "firefox") { $Y_CookieBrowser.SelectedIndex = 2 }
        elseif ($realPath -match "opera") { $Y_CookieBrowser.SelectedIndex = 3 }
        elseif ($realPath -match "brave") { $Y_CookieBrowser.SelectedIndex = 4 }
        else { $Y_CookieBrowser.SelectedIndex = 0 } # Fallback to Edge
        
        # Note: Logging removed from here to prevent it from showing at startup
    }
    catch {
        $Y_CookieBrowser.SelectedIndex = 0
    }
    $script:State | Add-Member -MemberType NoteProperty -Name "BrowserLogged" -Value $false
    
    #default to using cookies from browser with auto-detection, but allow user to uncheck if they want
    #$Y_CheckCookie.IsChecked = $true
    $Y_CheckCookie.IsChecked = $false
    # -----------------------------------------------

    $Y_Res.Add_SelectionChanged({ Update-YtDlpPreview })
    $Y_VFormat.Add_SelectionChanged({ Update-YtDlpPreview })
    $Y_AFormat.Add_SelectionChanged({ Update-YtDlpPreview })
    $Y_CookieBrowser.Add_SelectionChanged({ Update-YtDlpPreview })
    $Y_CheckMeta.Add_Checked({ Update-YtDlpPreview }); $Y_CheckMeta.Add_Unchecked({ Update-YtDlpPreview })
    $Y_CheckSubs.Add_Checked({ Update-YtDlpPreview }); $Y_CheckSubs.Add_Unchecked({ Update-YtDlpPreview })
    $Y_CheckSponsor.Add_Checked({ Update-YtDlpPreview }); $Y_CheckSponsor.Add_Unchecked({ Update-YtDlpPreview })
    $Y_CheckCookie.Add_Checked({ Update-YtDlpPreview }); $Y_CheckCookie.Add_Unchecked({ Update-YtDlpPreview })
    $Y_CheckAutoPoToken.Add_Checked({ $Y_PoToken.IsEnabled = $false; $Y_PoToken.Opacity = 0.4; Update-YtDlpPreview }) 
    $Y_CheckAutoPoToken.Add_Unchecked({ $Y_PoToken.IsEnabled = $true; $Y_PoToken.Opacity = 1.0; Update-YtDlpPreview })
    $Y_CookiePath.Add_TextChanged({ Update-YtDlpPreview })
    $Y_Link.Add_TextChanged({ Update-YtDlpPreview })
    $Y_CustomParams.Add_TextChanged({ Update-YtDlpPreview })
    $Y_OutDir.Add_TextChanged({ Update-YtDlpPreview })
    $Y_PoToken.Add_TextChanged({ Update-YtDlpPreview })

    # Button event for "Update Tools" - Generates a small WPF Window to fetch WinGet updates
    $BtnUpdate.Add_Click({
            # Validate WinGet before showing the update menu
            if (-not (Get-Command "winget" -ErrorAction SilentlyContinue)) {
                [void][System.Windows.MessageBox]::Show("Windows Package Manager (winget) is not installed on your system.`n`nUpdating dependencies automatically requires winget.", "Winget Not Found", 0, 16)
                return
            }

            [xml]$updXaml = @"
        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" 
                Title="Update Dependencies" Width="350" Height="320" WindowStartupLocation="CenterScreen" Background="{DynamicResource BgBrush}" ResizeMode="NoResize">
            <Window.Resources>
                <SolidColorBrush x:Key="BgBrush" Color="$($window.Resources["BgBrush"].Color.ToString())"/>
                <SolidColorBrush x:Key="TextBrush" Color="$($window.Resources["TextBrush"].Color.ToString())"/>
                <SolidColorBrush x:Key="AccentBrush" Color="#6366F1"/>
                <Style TargetType="CheckBox"><Setter Property="Margin" Value="0,10,0,0"/><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/><Setter Property="FontSize" Value="14"/></Style>
                <Style TargetType="TextBlock"><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/></Style>
            </Window.Resources>
            <StackPanel Margin="20">
                <TextBlock Text="Select tools to update:" FontWeight="Bold" FontSize="16" Foreground="{DynamicResource AccentBrush}"/>
                <CheckBox Name="ChkYt" Content="yt-dlp (YouTube Downloader)" IsChecked="True"/>
                <CheckBox Name="ChkFFmpeg" Content="FFmpeg (Media Converter)" IsChecked="True"/>
                <CheckBox Name="ChkHandbrake" Content="HandBrakeCLI (Video Encoder)" IsChecked="True"/>
                <CheckBox Name="ChkNode" Content="Node.js (JS Runtime)" IsChecked="True"/>
                <CheckBox Name="ChkWhisper" Content="Whisper AI (Python Package)" IsChecked="True"/>
                <CheckBox Name="ChkUpscale" Content="AI Upscaler (Upscayl)" IsChecked="True"/>
                <Button Name="BtnStartUpdate" Content="Start Update" Height="35" Margin="0,20,0,0" Background="#10B981" Foreground="White" BorderThickness="0" Cursor="Hand" FontWeight="Bold"/>
            </StackPanel>
        </Window>
"@
            $updWin = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $updXaml))
        
            $updWin.FindName("BtnStartUpdate").Add_Click({
                    $doYt = [bool]$updWin.FindName("ChkYt").IsChecked
                    $doFF = [bool]$updWin.FindName("ChkFFmpeg").IsChecked
                    $doHb = [bool]$updWin.FindName("ChkHandbrake").IsChecked
                    $doNode = [bool]$updWin.FindName("ChkNode").IsChecked
                    $doWhisper = [bool]$updWin.FindName("ChkWhisper").IsChecked
                    $doUpscale = [bool]$updWin.FindName("ChkUpscale").IsChecked
                    $updWin.Close()

                    $BtnRun.IsEnabled = $false
                    $BtnUpdate.IsEnabled = $false
                    $TaskbarProgress.ProgressState = "Normal"

                    # Setup logging to update.log in the main logs folder
                    $UpdateLogFile = Join-Path $LogDir "update.log"
                    function Write-UpdLog([string]$Message) {
                        $LogBox.AppendText("$Message`r`n")
                        Add-Content -Path $UpdateLogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
                    }

                    # --- ADDED: Extra empty line so it doesn't glue to previous logs ---
                    $LogBox.AppendText("`r`n") 
                    Write-UpdLog "[UPDATE] Starting dependency updates..."
            
                    # Helper function to force the UI to refresh visually before blocking it with Start-Process
                    function ForceUiRefresh {
                        $window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background)
                    }

                    # Helper function to properly translate WinGet and standard exit codes to the Live Log
                    function LogUpdateResult($ToolName, $Proc) {
                        if (-not $Proc) { return }
                        $code = $Proc.ExitCode
                        if ($code -eq 0) { 
                            Write-UpdLog "[UPDATE] $ToolName updated successfully." 
                        } 
                        elseif ($code -in @(-1978335189, 2316632107, -1978335188, 2316632108)) { 
                            Write-UpdLog "[INFO] $ToolName is already up-to-date." 
                        } 
                        else { 
                            Write-UpdLog "[WARNING] $ToolName update failed or cancelled (Exit Code: $code)." 
                        }
                    }

                    if ($doYt -and $script:State.ytdlpFound) {
                        $StatusText.Text = "Updating yt-dlp..."
                        $PBar.Value = 20; ForceUiRefresh
                        
                        if ($script:isWinGetVersion) {
                            $p = Start-Process winget -ArgumentList "upgrade --id yt-dlp.yt-dlp --accept-source-agreements --accept-package-agreements" -Wait -WindowStyle Normal -PassThru
                        }
                        else {
                            $p = Start-Process cmd.exe -ArgumentList "/c `"$($script:State.ytdlp)`" -U" -Wait -WindowStyle Normal -PassThru
                        }
                        LogUpdateResult "yt-dlp" $p
                    }
                    
                    if ($doFF) {
                        $StatusText.Text = "Updating FFmpeg..."
                        $PBar.Value = 40; ForceUiRefresh
                        $p = Start-Process winget -ArgumentList "upgrade --id Gyan.FFmpeg --accept-source-agreements --accept-package-agreements" -Wait -WindowStyle Normal -PassThru
                        LogUpdateResult "FFmpeg" $p
                    }

                    if ($doHb) {
                        $StatusText.Text = "Updating HandBrake..."
                        $PBar.Value = 50; ForceUiRefresh
                        $p = Start-Process winget -ArgumentList "upgrade --id HandBrake.HandBrake.CLI --accept-source-agreements --accept-package-agreements" -Wait -WindowStyle Normal -PassThru
                        LogUpdateResult "HandBrake" $p
                    }
                    
                    if ($doNode) {
                        $StatusText.Text = "Updating Node.js..."
                        $PBar.Value = 60; ForceUiRefresh
                        $p = Start-Process winget -ArgumentList "upgrade --id OpenJS.NodeJS --accept-source-agreements --accept-package-agreements" -Wait -WindowStyle Normal -PassThru
                        LogUpdateResult "Node.js" $p
                    }
                    
                    if ($doWhisper) {
                        $StatusText.Text = "Updating Whisper AI..."
                        $PBar.Value = 80; ForceUiRefresh
                        if (Get-Command "python" -ErrorAction SilentlyContinue) {
                            $tempPipLog = Join-Path $env:TEMP "whisper_update.log"
                            if (Test-Path $tempPipLog) { Remove-Item $tempPipLog -Force -ErrorAction SilentlyContinue }
                            
                            # Using PowerShell to run pip so it will "Tee" the output to a file while showing it to the user. No pause!
                            $p = Start-Process powershell.exe -ArgumentList "-NoProfile -Command `"Write-Host 'Updating Whisper AI...'; pip install -U openai-whisper | Tee-Object -FilePath '$tempPipLog'`"" -Wait -WindowStyle Normal -PassThru
                            
                            if ($p.ExitCode -eq 0) {
                                $pipOut = if (Test-Path $tempPipLog) { Get-Content $tempPipLog -Raw } else { "" }
                                
                                # Dump verbose pip output into the update.log file
                                if (-not [string]::IsNullOrWhiteSpace($pipOut)) {
                                    Add-Content -Path $UpdateLogFile -Value "`n--- Whisper Pip Verbose Log ---`n$pipOut`n-------------------------------"
                                }
                                
                                # Read the text output to see if it actually installed something new
                                if ($pipOut -match "Successfully installed") {
                                    Write-UpdLog "[UPDATE] Whisper updated successfully."
                                }
                                else {
                                    Write-UpdLog "[INFO] Whisper is already up-to-date."
                                }
                            }
                            else {
                                Write-UpdLog "[WARNING] Whisper update failed or cancelled (Exit Code: $($p.ExitCode))."
                            }
                        }
                        else {
                            Write-UpdLog "[WARNING] Python not found. Skipping Whisper update."
                        }
                    }
                    
                    if ($doUpscale) {
                        $StatusText.Text = "Updating Upscayl..."
                        $PBar.Value = 90; ForceUiRefresh
                        $p = Start-Process winget -ArgumentList "upgrade --id Upscayl.Upscayl --accept-source-agreements --accept-package-agreements" -Wait -WindowStyle Normal -PassThru
                        LogUpdateResult "Upscayl" $p
                    }

                    $StatusText.Text = "Updates completed!"
                    $PBar.Value = 100
                    $TaskbarProgress.ProgressState = "None"
                    Write-UpdLog "[UPDATE] Dependency update process finished."
                    
                    # --- ADDED: Windows System Tray / Toast Notification ---
                    Show-Toast -Title "Update Complete" -Message "Dependency update process finished."
                    # -----------------------------------------------

                    $BtnRun.IsEnabled = $true
                    $BtnUpdate.IsEnabled = $true
                    Find-Tools 
                })
            [void]$updWin.ShowDialog()
        })

    # Button event for "Settings" - Generates UI for Theme, Threads, Defaults, etc.
$BtnSettings.Add_Click({
            [xml]$setXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Settings" Width="420" Height="450" WindowStartupLocation="CenterScreen" Background="{DynamicResource BgBrush}" ResizeMode="NoResize">
    <Window.Resources>
        <SolidColorBrush x:Key="BgBrush" Color="$($window.Resources["BgBrush"].Color.ToString())"/>
        <SolidColorBrush x:Key="CardBrush" Color="$($window.Resources["CardBrush"].Color.ToString())"/>
        <SolidColorBrush x:Key="TextBrush" Color="$($window.Resources["TextBrush"].Color.ToString())"/>
        <SolidColorBrush x:Key="MutedBrush" Color="$($window.Resources["MutedBrush"].Color.ToString())"/>
        <SolidColorBrush x:Key="BorderBrush" Color="$($window.Resources["BorderBrush"].Color.ToString())"/>
        <SolidColorBrush x:Key="AccentBrush" Color="$($window.Resources["AccentBrush"].Color.ToString())"/>
        <SolidColorBrush x:Key="InputBgBrush" Color="$($window.Resources["InputBgBrush"].Color.ToString())"/>
        
        <Style TargetType="TextBlock"><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/><Setter Property="Margin" Value="0,10,0,5"/></Style>
        <Style TargetType="CheckBox"><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/><Setter Property="Margin" Value="0,10,0,0"/></Style>
        
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{Binding Foreground, RelativeSource={RelativeSource AncestorType=ComboBox}}"/>
            <Setter Property="Padding" Value="8"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="BgBorder" CornerRadius="4" Background="{TemplateBinding Background}" Margin="2">
                            <ContentPresenter Margin="{TemplateBinding Padding}" TextElement.Foreground="{TemplateBinding Foreground}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="BgBorder" Property="Background" Value="{Binding BorderBrush, RelativeSource={RelativeSource AncestorType=ComboBox}}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="BgBorder" Property="Background" Value="{DynamicResource AccentBrush}"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{DynamicResource InputBgBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton x:Name="ToggleButton" Grid.Column="2" Focusable="false"
                                          IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press"
                                          Background="{TemplateBinding Background}"
                                          BorderBrush="{TemplateBinding BorderBrush}"
                                          BorderThickness="{TemplateBinding BorderThickness}">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6">
                                            <Path x:Name="Arrow" Fill="{Binding Foreground, RelativeSource={RelativeSource AncestorType=ComboBox}}" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,10,0" Data="M0,0 L4,4 L8,0 z"/>
                                        </Border>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter x:Name="ContentSite" TextElement.Foreground="{TemplateBinding Foreground}" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" Margin="{TemplateBinding Padding}" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            
                            <Popup x:Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False">
                                <Border Background="{DynamicResource CardBrush}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="6" Margin="0,4,0,0" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="250">
                                    <ScrollViewer Margin="2" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TextBox"><Setter Property="Background" Value="{DynamicResource InputBgBrush}"/><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/><Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/><Setter Property="BorderThickness" Value="1"/><Setter Property="Padding" Value="8"/><Setter Property="VerticalContentAlignment" Value="Center"/></Style>
    </Window.Resources>
    <StackPanel Margin="20">
        <TextBlock Text="App Settings" FontSize="18" FontWeight="Bold"/>
        
        <TextBlock Text="Theme"/>
        <ComboBox x:Name="CboTheme"><ComboBoxItem>Dark</ComboBoxItem><ComboBoxItem>Light</ComboBoxItem></ComboBox>

        <TextBlock Text="CPU Thread Limit (FFmpeg)" ToolTip="Limit cores used during conversion to keep PC responsive."/>
        <ComboBox x:Name="CboThreads">
            </ComboBox>

        <TextBlock Text="Global Default Output Directory" ToolTip="Leave empty to output in the script's folder."/>
        <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="80"/></Grid.ColumnDefinitions>
            <TextBox x:Name="TxtDefaultOutDir" IsReadOnly="True" Cursor="Arrow" Margin="0,0,5,0"/>
            <Button x:Name="BtnBrowseOutDir" Grid.Column="1" Content="Browse" Height="30" Background="#4B5563" Foreground="White" BorderThickness="0" Cursor="Hand"/>
        </Grid>

        <CheckBox x:Name="ChkPlaySound" Content="Play sound when queue finishes" FontWeight="SemiBold"/>
        <CheckBox x:Name="ChkAlwaysOnTop" Content="Keep window always on top" FontWeight="SemiBold"/>
        <CheckBox x:Name="ChkAutoDelete" Content="Auto-delete original local files after successful process" FontWeight="Bold" Foreground="#EF4444"/>

        <Button x:Name="BtnSaveSet" Content="Save and Apply" Height="35" Margin="0,25,0,0" Background="#6366F1" Foreground="White" BorderThickness="0" Cursor="Hand" FontWeight="Bold"/>
    </StackPanel>
</Window>
"@
            $setWin = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $setXaml))
            
            $cboTheme = $setWin.FindName("CboTheme"); if ($Config.Theme -eq "Light") { $cboTheme.SelectedIndex = 1 } else { $cboTheme.SelectedIndex = 0 }

            # Dynamically populate thread options based on logical processor count, with "Auto" as default
            $cboThreads = $setWin.FindName("CboThreads")
            $maxThreads = [System.Environment]::ProcessorCount
            if ($maxThreads -le 0) { $maxThreads = 4 }
            
            $tAuto = New-Object System.Windows.Controls.ComboBoxItem
            $tAuto.Content = "Auto (Use All Cores)"
            [void]$cboThreads.Items.Add($tAuto)
            
            for ($i = 1; $i -le $maxThreads; $i++) {
                $tItem = New-Object System.Windows.Controls.ComboBoxItem
                $tItem.Content = "$i Thread" + $(if ($i -gt 1) { "s" }else { "" })
                [void]$cboThreads.Items.Add($tItem)
            }

            if ($Config.ThreadLimit -eq "Auto" -or $null -eq $Config.ThreadLimit) {
                $cboThreads.SelectedIndex = 0
            }
            else {
                $found = $false
                for ($idx = 1; $idx -lt $cboThreads.Items.Count; $idx++) {
                    if ($cboThreads.Items[$idx].Content -match "^$($Config.ThreadLimit)\s") {
                        $cboThreads.SelectedIndex = $idx
                        $found = $true
                        break
                    }
                }
                if (-not $found) { $cboThreads.SelectedIndex = 0 }
            }

            $setWin.FindName("TxtDefaultOutDir").Text = $Config.DefaultOutDir
            $setWin.FindName("ChkPlaySound").IsChecked = $Config.PlaySound
            $setWin.FindName("ChkAlwaysOnTop").IsChecked = $Config.AlwaysOnTop
            $setWin.FindName("ChkAutoDelete").IsChecked = $Config.AutoDelete

            $setWin.FindName("BtnBrowseOutDir").Add_Click({ 
                    $fd = New-Object System.Windows.Forms.FolderBrowserDialog
                    $fd.Description = "Select Default Output Folder"
                    $fd.SelectedPath = $ScriptDir
                    $fd.ShowNewFolderButton = $true
                
                    if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { 
                        $setWin.FindName("TxtDefaultOutDir").Text = $fd.SelectedPath
                    } 
                })

            $setWin.FindName("BtnSaveSet").Add_Click({
                    $Config.Theme = (Get-CbVal $cboTheme)
                    
                    $threadCbVal = (Get-CbVal $cboThreads)
                    if ($threadCbVal -match "Auto") { $Config.ThreadLimit = "Auto" }
                    else { $Config.ThreadLimit = ($threadCbVal -split " ")[0] }

                    $Config.DefaultOutDir = $setWin.FindName("TxtDefaultOutDir").Text
                    $Config.PlaySound = [bool]$setWin.FindName("ChkPlaySound").IsChecked
                    $Config.AlwaysOnTop = [bool]$setWin.FindName("ChkAlwaysOnTop").IsChecked
                    $Config.AutoDelete = [bool]$setWin.FindName("ChkAutoDelete").IsChecked
                    
                    $Config | ConvertTo-Json -Depth 10 | Set-Content $ConfigFile -Encoding UTF8 #UTF8 to preserve any special chars in paths
                    Enable-Config
                    $setWin.Close()
                })
            [void]$setWin.ShowDialog()
        })

    # Read custom stored video presets on application load
    if (Test-Path $PresetFile) {
        try {
            $presets = Get-Content $PresetFile | ConvertFrom-Json
            foreach ($p in $presets) {
                $item = New-Object System.Windows.Controls.ComboBoxItem
                $item.Content = $p.Name
                $item.Tag = $p
                [void]$V_Preset.Items.Add($item)
            }
        }
        catch {}
    }

    # Save custom video encode preset event handler
    $V_BtnSavePreset.Add_Click({
            $name = [Microsoft.VisualBasic.Interaction]::InputBox("Name for this preset:", "Save Preset", "My Preset")
            if ([string]::IsNullOrWhiteSpace($name)) { return }
            $newPreset = @{ Name = $name; Format = $V_CFormat.SelectedIndex; HW = $V_CHWAccel.SelectedIndex; Codec = $V_CCodec.SelectedIndex; Audio = $V_CAudio.SelectedIndex; Res = $V_CRes.SelectedIndex; Tracks = $V_CAudioTracks.SelectedIndex; CRF = $V_SliderCRF.Value }
            $presets = @()
            if (Test-Path $PresetFile) { try { $presets = @(Get-Content $PresetFile | ConvertFrom-Json) } catch {} }
            $presets += $newPreset
            $presets | ConvertTo-Json -Depth 10 | Set-Content $PresetFile -Encoding UTF8 #UTF8 to preserve any special chars in preset names
            $item = New-Object System.Windows.Controls.ComboBoxItem; $item.Content = $name; $item.Tag = (New-Object PSObject -Property $newPreset)
            [void]$V_Preset.Items.Add($item); $V_Preset.SelectedItem = $item
            [void][System.Windows.MessageBox]::Show("Preset '$name' was successfully saved!", "Info", 0, 64)
        })

    $V_Preset.Add_SelectionChanged({
            if ($V_Preset.SelectedItem -ne $null -and $V_Preset.SelectedItem.Tag) {
                $p = $V_Preset.SelectedItem.Tag
                $V_CFormat.SelectedIndex = $p.Format; $V_CHWAccel.SelectedIndex = $p.HW; $V_CCodec.SelectedIndex = $p.Codec
                $V_CAudio.SelectedIndex = $p.Audio; $V_CRes.SelectedIndex = $p.Res; $V_CAudioTracks.SelectedIndex = $p.Tracks; $V_SliderCRF.Value = $p.CRF
            }
            else {
                $idx = $V_Preset.SelectedIndex
                if ($idx -eq 1) { $V_CFormat.SelectedIndex = 0; $V_CCodec.SelectedIndex = 0; $V_CRes.SelectedIndex = 1; $V_SliderCRF.Value = 23 }
                elseif ($idx -eq 2) { $V_CFormat.SelectedIndex = 0; $V_CCodec.SelectedIndex = 0; $V_CRes.SelectedIndex = 2; $V_SliderCRF.Value = 28 }
            }
        })

    # Core function for Drag & Drop parsing: resolves files inside directories or directly checks regex match
    function Add-ToList([System.Windows.Controls.ListBox]$List, [string[]]$Paths, [string]$ExtRegex) {
        $existingItems = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($item in $List.Items) { [void]$existingItems.Add($item.ToString()) }

        $videoDetected = $false

        foreach ($p in $Paths) {
            $pStr = [string]$p
            if ([System.IO.Directory]::Exists($pStr)) { 
                try {
                    $files = [System.IO.Directory]::GetFiles($pStr, "*.*", [System.IO.SearchOption]::TopDirectoryOnly)
                    foreach ($f in $files) {
                        if ($f -match "(?i)$ExtRegex" -and $existingItems.Add($f)) { 
                            [void]$List.Items.Add($f)
                            if ($List.Name -eq "A_InList" -and $f -match "\.(mp4|mkv|avi|mov|webm)$") { $videoDetected = $true }
                        }
                    }
                }
                catch {
                    $LogBox.AppendText("`r`n[WARNING] Skipped restricted folder: $pStr`r`n")
                }
            } 
            elseif ([System.IO.File]::Exists($pStr)) { 
                if ($pStr -match "(?i)$ExtRegex" -and $existingItems.Add($pStr)) { 
                    [void]$List.Items.Add($pStr) 
                    if ($List.Name -eq "A_InList" -and $pStr -match "\.(mp4|mkv|avi|mov|webm)$") { $videoDetected = $true }
                } 
            }
        }
        
        # If we are on the Audio tab and videos were added, auto-check the box
        if ($videoDetected -and $List.Name -eq "A_InList" -and -not $A_CheckExtract.IsChecked) {
            $A_CheckExtract.IsChecked = $true
            [void][System.Windows.MessageBox]::Show("You added one or more video files to the Audio tab.`n`n'Extract Audio from Video' has been automatically checked for you.", "Video Detected", 0, 64)
        }

        if ($List.Items.Count -gt 0 -and $List.SelectedIndex -eq -1) { $List.SelectedIndex = 0 }
        Update-AllPreviews
    }

    # Setup Drag and Drop events for generic UI elements
    $DragEnterHandler = [System.Windows.DragEventHandler] { $_.Effects = [System.Windows.DragDropEffects]::Copy; $_.Handled = $true; $_.Source.Background = $window.Resources["DragBrush"] }
    $DragLeaveHandler = [System.Windows.DragEventHandler] { $_.Source.Background = $window.Resources["InputBgBrush"] }
    
    function SetupDragReorder ([System.Windows.Controls.ListBox]$lb, [string]$extRegex, [scriptblock]$checkExt) {
        $lb.Add_PreviewMouseMove({
                param($uiSender, $e)
                if ($e.LeftButton -eq [System.Windows.Input.MouseButtonState]::Pressed -and $uiSender.SelectedItem) {
                    [void][System.Windows.DragDrop]::DoDragDrop($uiSender, $uiSender.SelectedItem.ToString(), [System.Windows.DragDropEffects]::Move)
                }
            })
        $lb.Add_PreviewDragOver($DragEnterHandler)
        $lb.Add_DragLeave($DragLeaveHandler)
        $lb.Add_Drop({
                param($uiSender, $e)
                $uiSender.Background = $window.Resources["InputBgBrush"]
            
                # 1. Handle external file drop from Windows Explorer
                if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
                    $rx = if ($checkExt) { &$checkExt } else { $extRegex }
                    Add-ToList $uiSender ($e.Data.GetData([System.Windows.DataFormats]::FileDrop)) $rx
                }
                # 2. Handle internal reordering drop
                elseif ($e.Data.GetDataPresent([System.String])) {
                    $droppedData = $e.Data.GetData([System.String])
                    $targetItem = if ($e.OriginalSource.DataContext) { $e.OriginalSource.DataContext } else { $e.OriginalSource }
                
                    if ($targetItem -is [string] -and $targetItem -ne $droppedData) {
                        $targetIndex = $uiSender.Items.IndexOf($targetItem)
                        $sourceIndex = $uiSender.Items.IndexOf($droppedData)
                        if ($sourceIndex -ge 0 -and $targetIndex -ge 0) {
                            $uiSender.Items.RemoveAt($sourceIndex)
                            $uiSender.Items.Insert($targetIndex, $droppedData)
                            $uiSender.SelectedItem = $droppedData
                            Update-AllPreviews
                        }
                    }
                }
            })
    }

    # Always allow both audio and video formats to be dropped into the Audio list
    SetupDragReorder $A_InList "\.(mp3|wav|m4a|flac|ogg|aac|mp4|mkv|avi|mov|webm)$" $null
    SetupDragReorder $V_InList "\.(mp4|mkv|avi|mov|webm)$" $null
    SetupDragReorder $I_InList "\.(jpg|jpeg|png|webp|bmp|gif|heic)$" $null

    $TextBoxDropHandler = [System.Windows.DragEventHandler] {
        $_.Source.Background = $window.Resources["InputBgBrush"]
        if ($_.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
            $files = $_.Data.GetData([System.Windows.DataFormats]::FileDrop)
            if ($files.Count -gt 0) { $_.Source.Text = $files[0] }
        }
    }

    # Helper function to easily bind right-click Context Menus to ListBoxes
    function Add-ContextMenu($ListBox, $BtnRemove, $BtnClear) {
        $BtnRemove.Add_Click({
                $selected = @($ListBox.SelectedItems)
                # Clear UI selection first to prevent multi-select redraw lag
                $ListBox.SelectedItems.Clear()
                $ListBox.Dispatcher.InvokeAsync([Action] {
                        foreach ($item in $selected) { $ListBox.Items.Remove($item) }
                        Update-AllPreviews
                    }, [System.Windows.Threading.DispatcherPriority]::Background)
            })
        $BtnClear.Add_Click({ 
                $ListBox.Items.Clear()
                Update-AllPreviews
            })
    }

    Add-ContextMenu $A_InList $A_CtxRemove $A_CtxClear
    Add-ContextMenu $V_InList $V_CtxRemove $V_CtxClear
    Add-ContextMenu $I_InList $I_CtxRemove $I_CtxClear

    # Manual Add File Click handlers utilizing Windows Forms OpenFileDialog
    $A_BtnAdd.Add_Click({ 
            $fd = New-Object System.Windows.Forms.OpenFileDialog
            $fd.InitialDirectory = $ScriptDir
            $fd.Multiselect = $true
            $fd.Filter = "Media Files|*.mp3;*.wav;*.m4a;*.flac;*.ogg;*.aac;*.mp4;*.mkv;*.avi;*.mov;*.webm|All Files|*.*"
            if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Add-ToList $A_InList $fd.FileNames ".*" } 
        })

    $V_BtnAdd.Add_Click({ 
            $fd = New-Object System.Windows.Forms.OpenFileDialog
            $fd.InitialDirectory = $ScriptDir
            $fd.Multiselect = $true
            $fd.Filter = "Video|*.mp4;*.mkv;*.avi;*.mov;*.webm|All|*.*"
            if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Add-ToList $V_InList $fd.FileNames ".*" } 
        })

    $I_BtnAdd.Add_Click({ 
            $fd = New-Object System.Windows.Forms.OpenFileDialog
            $fd.InitialDirectory = $ScriptDir
            $fd.Multiselect = $true
            $fd.Filter = "Images|*.jpg;*.jpeg;*.png;*.webp;*.bmp;*.gif;*.heic|All|*.*"
            if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Add-ToList $I_InList $fd.FileNames ".*" } 
        })    

    $A_BtnClear.Add_Click({ $A_InList.Items.Clear(); Update-AudioFfmpegPreview })
    $V_BtnClear.Add_Click({ 
            $V_InList.Items.Clear()
            $V_OutFilename.Clear()
            $V_CheckSmartName.IsChecked = $false
            $V_CheckSmartName.IsEnabled = $false
            Update-FfmpegPreview 
        })
    $I_BtnClear.Add_Click({ $I_InList.Items.Clear() })

    # Helper function for assigning an Output Directory using standard FolderBrowserDialog
    function Select-OutDir([System.Windows.Controls.TextBox]$Box) {
        $fd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fd.Description = "Select Target Folder"
        $fd.SelectedPath = $ScriptDir
        $fd.ShowNewFolderButton = $true
        
        if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { 
            $Box.Text = $fd.SelectedPath
            Update-AllPreviews 
        } 
    }

    $A_BtnOut.Add_Click({ Select-OutDir $A_OutDir })
    $V_BtnOut.Add_Click({ Select-OutDir $V_OutDir })
    $I_BtnOut.Add_Click({ Select-OutDir $I_OutDir })
    $Y_BtnOut.Add_Click({ Select-OutDir $Y_OutDir })
    $S_BtnUpscaleOut.Add_Click({ Select-OutDir $S_UpscaleOutDir })

    # Muxing and Downloader Browsers
    $M_BtnVid.Add_Click({ $fd = New-Object System.Windows.Forms.OpenFileDialog; $fd.InitialDirectory = $ScriptDir; $fd.Filter = "Video|*.mp4;*.mkv;*.avi;*.webm|All|*.*"; if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $M_InVideo.Text = $fd.FileName } })
    $M_BtnAud.Add_Click({ $fd = New-Object System.Windows.Forms.OpenFileDialog; $fd.InitialDirectory = $ScriptDir; $fd.Filter = "Audio|*.mp3;*.wav;*.m4a;*.aac|All|*.*"; if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $M_InAudio.Text = $fd.FileName } })
    $M_BtnOut.Add_Click({ $sd = New-Object System.Windows.Forms.SaveFileDialog; $sd.InitialDirectory = $ScriptDir; $sd.Filter = "MP4 Video|*.mp4|MKV Video|*.mkv"; if ($sd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $M_OutFile.Text = $sd.FileName } })

    $Y_BtnCookie.Add_Click({
            $fd = New-Object System.Windows.Forms.OpenFileDialog
            $fd.InitialDirectory = $ScriptDir
            $fd.Filter = "Text Files|*.txt|All Files|*.*"
            if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $Y_CookiePath.Text = $fd.FileName
                $Y_CheckCookie.IsChecked = $true
            }
        })

    $Y_BtnBatchBrowse.Add_Click({
            $fd = New-Object System.Windows.Forms.OpenFileDialog
            $fd.InitialDirectory = $ScriptDir
            $fd.Filter = "Text Files (*.txt)|*.txt|All Files|*.*"
            if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $Y_BatchFile.Text = $fd.FileName }
        })

    $V_BtnSub.Add_Click({ 
            $fd = New-Object System.Windows.Forms.OpenFileDialog
            $fd.InitialDirectory = $ScriptDir
            $fd.Filter = "Subtitles|*.srt|All|*.*"
            if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $V_SubPath.Text = $fd.FileName } 
        })
        
    # REWRITTEN: Fast synchronous timeline generation with UI pumping (No Background Threads!)
    $V_BtnGenPreview.Add_Click({
            $file = $V_InList.SelectedItem
            if (-not $file -and $V_InList.Items.Count -gt 0) { $file = $V_InList.Items[0] }
            if (-not $file) { [void][System.Windows.MessageBox]::Show("Select a video in the queue first.", "Missing Input", 0, 48); return }

            $filePath = $file.ToString()
            $startText = $V_TrimStart.Text
            $endText = $V_TrimEnd.Text

            foreach ($child in $V_PreviewStack.Children) {
                if ($child -is [System.Windows.Controls.Image]) { $child.Source = $null }
            }
            $V_PreviewStack.Children.Clear()
            $V_PreviewScroll.Visibility = "Visible"
            $V_BtnGenPreview.IsEnabled = $false
            $V_BtnGenPreview.Content = "Generating..."
            $window.Cursor = [System.Windows.Input.Cursors]::Wait

            # Pump the UI so the user immediately sees the "Generating..." text
            $window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background)

            try {
                $startSecs = 0
                $endSecs = 0

                # Parse Start Time safely
                if ($startText -match "(\d{2}):(\d{2}):(\d{2})") {
                    $startSecs = ([int]$matches[1] * 3600) + ([int]$matches[2] * 60) + [int]$matches[3]
                }
                # Parse End Time safely
                if ($endText -match "(\d{2}):(\d{2}):(\d{2})") {
                    $endSecs = ([int]$matches[1] * 3600) + ([int]$matches[2] * 60) + [int]$matches[3]
                }

                # If end time is missing or smaller than start, fetch real duration
                if ($endSecs -le $startSecs) {
                    $pinfoDur = New-Object System.Diagnostics.ProcessStartInfo
                    $pinfoDur.FileName = $script:State.ffprobe
                    $pinfoDur.Arguments = "-v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 `"$filePath`""
                    $pinfoDur.UseShellExecute = $false; $pinfoDur.RedirectStandardOutput = $true; $pinfoDur.RedirectStandardError = $true; $pinfoDur.CreateNoWindow = $true
                    $pDur = [System.Diagnostics.Process]::Start($pinfoDur)
                    
                    $durStr = $pDur.StandardOutput.ReadToEnd().Trim()
                    [void]$pDur.StandardError.ReadToEnd()
                    
                    if (-not $pDur.WaitForExit(3000)) { try { $pDur.Kill() } catch {} }
                    $pDur.Dispose()
                    $totalSecs = 0
                    if ([double]::TryParse($durStr, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$totalSecs)) {
                        $endSecs = $totalSecs
                    }
                }

                $duration = $endSecs - $startSecs
                if ($duration -le 0) { throw "Invalid duration calculated." }

                # Phase 1: Launch all 10 FFmpeg extractions in parallel
                $procs = @()
                for ($i = 1; $i -le 10; $i++) {
                    $percent = $i * 9 
                    $targetSec = $startSecs + ($duration * ($percent / 100.0))
                    $ts = [TimeSpan]::FromSeconds($targetSec)
                    $timeStr = "{0:D2}:{1:D2}:{2:D2}" -f [int][math]::Floor($ts.TotalHours), $ts.Minutes, $ts.Seconds
                    
                    $outThumb = Join-Path $env:TEMP "thumb_preview_$i.jpg"
                    if (Test-Path $outThumb) { Remove-Item $outThumb -Force -ErrorAction SilentlyContinue }

                    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
                    $pinfo.FileName = $script:State.ffmpeg
                    $pinfo.Arguments = "-y -hide_banner -ss $timeStr -i `"$filePath`" -frames:v 1 -q:v 2 -vf scale=200:-1 `"$outThumb`""
                    $pinfo.UseShellExecute = $false; $pinfo.CreateNoWindow = $true
                    $procs += [System.Diagnostics.Process]::Start($pinfo)
                }

                # Phase 2: Wait for all extractions to complete simultaneously WITHOUT freezing the UI
                foreach ($p in $procs) { 
                    $timeout = (Get-Date).AddSeconds(15)
                    while (-not $p.HasExited) {
                        $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
                        Start-Sleep -Milliseconds 20
                        if ((Get-Date) -gt $timeout) { try { $p.Kill() } catch {}; break }
                    }
                    $p.Dispose() 
                }

                # Phase 3: Load generated images into UI safely
                for ($i = 1; $i -le 10; $i++) {
                    $percent = $i * 9
                    $targetSec = $startSecs + ($duration * ($percent / 100.0))
                    $ts = [TimeSpan]::FromSeconds($targetSec)
                    $timeStr = "{0:D2}:{1:D2}:{2:D2}" -f [int][math]::Floor($ts.TotalHours), $ts.Minutes, $ts.Seconds
                    $outThumb = Join-Path $env:TEMP "thumb_preview_$i.jpg"

                    if (Test-Path $outThumb) {
                        $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
                        $bmp.BeginInit()
                        $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                        $bmp.CreateOptions = [System.Windows.Media.Imaging.BitmapCreateOptions]::IgnoreImageCache
                        $bmp.UriSource = New-Object System.Uri($outThumb)
                        $bmp.EndInit()
                        $bmp.Freeze()
                    
                        $img = New-Object System.Windows.Controls.Image
                        $img.Source = $bmp
                        $img.Height = 100 
                        $img.Margin = New-Object System.Windows.Thickness(0, 0, 8, 0)
                        $img.ToolTip = "Position: $percent% ($timeStr)"
                    
                        [void]$V_PreviewStack.Children.Add($img)
                    }
                }
                # Force UI to draw once at the very end
                $window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background)
            }
            catch {
                [void][System.Windows.MessageBox]::Show("Timeline generation failed: $($_.Exception.Message)", "Error", 0, 16)
            }
            finally {
                $V_BtnGenPreview.IsEnabled = $true
                $V_BtnGenPreview.Content = "Generate Visual Timeline"
                $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            }
        })

    $V_SliderCRF.Add_ValueChanged({ 
            $val = [math]::Round($V_SliderCRF.Value)
            $V_CRFText.Text = $val 
            if ($V_CRFDesc) {
                if ($val -le 20) { $V_CRFDesc.Text = "High Quality (Large file)" }
                elseif ($val -le 25) { $V_CRFDesc.Text = "Balanced (Standard)" }
                else { $V_CRFDesc.Text = "Compressed (Small file)" }
            }
        })
    
    $BtnShow.Add_Click({ 
            $dir = $script:State.lastOutDir
            if (-not [string]::IsNullOrWhiteSpace($dir) -and (Test-Path -LiteralPath $dir)) { 
                [void](Start-Process "explorer.exe" -ArgumentList "`"$dir`"")
            } 
        })

    # Execute ffprobe to retrieve JSON metadata for the currently selected file and display in a popup window
    # Execute ffprobe to retrieve JSON metadata and display in a popup window
    function Show-MediaInfoDialog([string]$FilePath, [System.Windows.Controls.Button]$Btn) {
        if (-not $FilePath -or -not (Test-Path -LiteralPath $FilePath)) { [void][System.Windows.MessageBox]::Show("Please select a valid file.", "Info", 0, 48); return }
        
        $Btn.IsEnabled = $false
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background)
        
        try {
            $pinfo = New-Object System.Diagnostics.ProcessStartInfo
            $pinfo.FileName = if ($script:State.ffprobeFound) { $script:State.ffprobe } else { $script:State.ffmpeg }
            
            # Use the -f format operator to prevent PowerShell from expanding $ characters in the path
            $safePath = $FilePath.Replace('"', '\"')
            if ($script:State.ffprobeFound) { 
                $pinfo.Arguments = '-v quiet -print_format json -show_format -show_streams "{0}"' -f $safePath
            }
            else {
                $pinfo.Arguments = '-hide_banner -i "{0}"' -f $safePath
            }

            $pinfo.UseShellExecute = $false
            $pinfo.RedirectStandardError = $true
            $pinfo.RedirectStandardOutput = $true
            $pinfo.CreateNoWindow = $true
            
            $p = [System.Diagnostics.Process]::Start($pinfo)
            
            # Read streams BEFORE waiting for exit to prevent buffer deadlocks
            $stdOutTask = $p.StandardOutput.ReadToEndAsync()
            $stdErrTask = $p.StandardError.ReadToEndAsync()
            
            # Keep UI responsive using native WPF Dispatcher invocation (Avoids GC allocation spikes)
            $timeout = (Get-Date).AddSeconds(10)
            while (-not $p.HasExited) {
                [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
                Start-Sleep -Milliseconds 50
                if ((Get-Date) -gt $timeout) { break }
            }

            if (-not $p.HasExited) {
                try { $p.Kill() } catch {}
                $infoText = "Error: Tool timed out while reading file metadata."
            }
            else {
                $infoText = if ($script:State.ffprobeFound) { $stdOutTask.Result } else { $stdErrTask.Result }
            }
            $p.Dispose() # Free up system memory

            if ([string]::IsNullOrWhiteSpace($infoText)) { 
                $infoText = "No metadata could be extracted.`n`nError Output:`n$($stdErrTask.Result)" 
            }
        }
        catch {
            $infoText = "Fatal error during info retrieval: $($_.Exception.Message)"
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            $Btn.IsEnabled = $true
        }
        
        # --- UI Window Generation (Keep this as is) ---
        [xml]$infoXaml = @"
        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Media Information" Width="650" Height="500" WindowStartupLocation="CenterScreen">
            <Grid Margin="15"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                <TextBlock Name="InfoLabel" Text="File Details" FontWeight="Bold" FontSize="16" Margin="0,0,0,10"/>
                <TextBox Name="InfoTextBox" Grid.Row="1" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="13" Padding="10"/>
            </Grid>
        </Window>
"@
        $infoWindow = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $infoXaml))
        $infoWindow.Background = $window.Resources["BgBrush"]
        $infoWindow.FindName("InfoLabel").Foreground = $window.Resources["TextBrush"]
        $tb = $infoWindow.FindName("InfoTextBox")
        $tb.Background = $window.Resources["InputBgBrush"]
        $tb.Foreground = $window.Resources["TextBrush"]
        $tb.Text = $infoText
        
        [void]$infoWindow.ShowDialog()
    }
    
    $A_BtnInfo.Add_Click({ Show-MediaInfoDialog $A_InList.SelectedItem $A_BtnInfo })
    $V_BtnInfo.Add_Click({ Show-MediaInfoDialog $V_InList.SelectedItem $V_BtnInfo })
    $I_BtnInfo.Add_Click({ Show-MediaInfoDialog $I_InList.SelectedItem $I_BtnInfo })

    # --- STRICT TRIM UI & VALIDATION LOGIC ---
    
    function Set-TimeStr([string]$txt) {
        $ts = [TimeSpan]::Zero
        if ([TimeSpan]::TryParse($txt, [ref]$ts)) { return "{0:D2}:{1:D2}:{2:D2}" -f $ts.Hours, $ts.Minutes, $ts.Seconds }
        return "00:00:00"
    }

    # Audio TextBoxes sync to Sliders
    $A_TrimStart.Add_LostFocus({ 
            $A_TrimStart.Text = Set-TimeStr $A_TrimStart.Text
            $ts = [TimeSpan]::Parse($A_TrimStart.Text); if ($ts.TotalSeconds -le $A_SliderTrimStart.Maximum) { $A_SliderTrimStart.Value = $ts.TotalSeconds } 
        })
    $A_TrimEnd.Add_LostFocus({ 
            $A_TrimEnd.Text = Set-TimeStr $A_TrimEnd.Text
            $ts = [TimeSpan]::Parse($A_TrimEnd.Text); if ($ts.TotalSeconds -le $A_SliderTrimEnd.Maximum) { $A_SliderTrimEnd.Value = $ts.TotalSeconds } 
        })
    
    # Video TextBoxes sync to Sliders
    $V_TrimStart.Add_LostFocus({ 
            $V_TrimStart.Text = Set-TimeStr $V_TrimStart.Text
            $ts = [TimeSpan]::Parse($V_TrimStart.Text); if ($ts.TotalSeconds -le $V_SliderTrimStart.Maximum) { $V_SliderTrimStart.Value = $ts.TotalSeconds } 
        })
    $V_TrimEnd.Add_LostFocus({ 
            $V_TrimEnd.Text = Set-TimeStr $V_TrimEnd.Text
            $ts = [TimeSpan]::Parse($V_TrimEnd.Text); if ($ts.TotalSeconds -le $V_SliderTrimEnd.Maximum) { $V_SliderTrimEnd.Value = $ts.TotalSeconds } 
        })

    # Interactive Audio Trimmer Slider Logic
    $A_SliderTrimStart.Add_ValueChanged({
            if ($A_SliderTrimStart.Value -ge $A_SliderTrimEnd.Value -and $A_SliderTrimEnd.Value -gt 0) { $A_SliderTrimStart.Value = $A_SliderTrimEnd.Value - 1 }
            $ts = [TimeSpan]::FromSeconds($A_SliderTrimStart.Value)
            $A_TrimStart.Text = "{0:D2}:{1:D2}:{2:D2}" -f $ts.Hours, $ts.Minutes, $ts.Seconds
            Update-AllPreviews
        })
    $A_SliderTrimEnd.Add_ValueChanged({
            if ($A_SliderTrimEnd.Value -le $A_SliderTrimStart.Value -and $A_SliderTrimEnd.Value -gt 0) { $A_SliderTrimEnd.Value = $A_SliderTrimStart.Value + 1 }
            $ts = [TimeSpan]::FromSeconds($A_SliderTrimEnd.Value)
            $A_TrimEnd.Text = "{0:D2}:{1:D2}:{2:D2}" -f $ts.Hours, $ts.Minutes, $ts.Seconds
            Update-AllPreviews
        })

    # Interactive Video Trimmer Slider Logic
    $V_SliderTrimStart.Add_ValueChanged({
            if ($V_SliderTrimStart.Value -ge $V_SliderTrimEnd.Value -and $V_SliderTrimEnd.Value -gt 0) { $V_SliderTrimStart.Value = $V_SliderTrimEnd.Value - 1 }
            $ts = [TimeSpan]::FromSeconds($V_SliderTrimStart.Value)
            $V_TrimStart.Text = "{0:D2}:{1:D2}:{2:D2}" -f $ts.Hours, $ts.Minutes, $ts.Seconds
            Update-AllPreviews
        })
    $V_SliderTrimEnd.Add_ValueChanged({
            if ($V_SliderTrimEnd.Value -le $V_SliderTrimStart.Value -and $V_SliderTrimEnd.Value -gt 0) { $V_SliderTrimEnd.Value = $V_SliderTrimStart.Value + 1 }
            $ts = [TimeSpan]::FromSeconds($V_SliderTrimEnd.Value)
            $V_TrimEnd.Text = "{0:D2}:{1:D2}:{2:D2}" -f $ts.Hours, $ts.Minutes, $ts.Seconds
            Update-AllPreviews
        })

    # Core function to extract media duration using FFprobe
    function Get-MediaDuration ([string]$FilePath) {
        if (-not $script:State.ffprobeFound -or -not (Test-Path -LiteralPath $FilePath)) { return 0 }
        try {
            $pinfoDur = New-Object System.Diagnostics.ProcessStartInfo
            $pinfoDur.FileName = $script:State.ffprobe
            $pinfoDur.Arguments = "-v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 `"$FilePath`""
            $pinfoDur.UseShellExecute = $false; $pinfoDur.RedirectStandardOutput = $true; $pinfoDur.RedirectStandardError = $true; $pinfoDur.CreateNoWindow = $true
            $pDur = [System.Diagnostics.Process]::Start($pinfoDur)
            if ($null -ne $pDur) {
                $durStr = $pDur.StandardOutput.ReadToEnd().Trim()
                [void]$pDur.StandardError.ReadToEnd()
                
                if (-not $pDur.WaitForExit(3000)) { try { $pDur.Kill() } catch {} }
                $pDur.Dispose()
            }
            
            $totalSecs = 0
            if ([double]::TryParse($durStr, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$totalSecs)) {
                return $totalSecs
            }
        }
        catch {
            Write-CrashLog "Get-MediaDuration Error processing '$FilePath': $($_.Exception.Message)"
        }
        return 0
    }

    # Update Audio Sliders when a new audio file is selected
    $A_InList.add_SelectionChanged([System.Windows.Controls.SelectionChangedEventHandler] { 
            if ($A_InList.SelectedItem) {
                $totalSecs = Get-MediaDuration $A_InList.SelectedItem.ToString()
                if ($totalSecs -gt 0) {
                    $A_SliderTrimStart.IsEnabled = $true; $A_SliderTrimEnd.IsEnabled = $true
                    $A_SliderTrimStart.Maximum = $totalSecs; $A_SliderTrimEnd.Maximum = $totalSecs
                    $A_SliderTrimStart.Value = 0; $A_SliderTrimEnd.Value = $totalSecs
                }
            }
            else {
                $A_SliderTrimStart.IsEnabled = $false; $A_SliderTrimEnd.IsEnabled = $false
            }
            Update-AudioFfmpegPreview 
        })

    # Update Video Sliders and Filename UI when a new video is selected
    $V_InList.add_SelectionChanged([System.Windows.Controls.SelectionChangedEventHandler] { 
            if ($V_InList.SelectedItem) {
                $V_CheckSmartName.IsEnabled = $true
                
                $totalSecs = Get-MediaDuration $V_InList.SelectedItem.ToString()
                if ($totalSecs -gt 0) {
                    $V_SliderTrimStart.IsEnabled = $true; $V_SliderTrimEnd.IsEnabled = $true
                    $V_SliderTrimStart.Maximum = $totalSecs; $V_SliderTrimEnd.Maximum = $totalSecs
                    $V_SliderTrimStart.Value = 0; $V_SliderTrimEnd.Value = $totalSecs
                }
            }
            else {
                $V_SliderTrimStart.IsEnabled = $false; $V_SliderTrimEnd.IsEnabled = $false
                $V_CheckSmartName.IsEnabled = $false
                $V_CheckSmartName.IsChecked = $false
                $V_OutFilename.Clear()
            }
            Update-FfmpegPreview 
        })
    # --- END STRICT TRIM UI LOGIC ---

    $Y_Link.Add_GotKeyboardFocus({ if ($Y_Link.Text -eq "https://") { $Y_Link.Text = "" } })
    $Y_Link.Add_LostKeyboardFocus({ if ([string]::IsNullOrWhiteSpace($Y_Link.Text)) { $Y_Link.Text = "https://" } })

    # Validates if an inputted URL is supported by yt-dlp by referencing the markdown from GitHub
    function CheckYtDlpSupport([string]$link) {
        $isValidUrl = $false
        try {
            if ($link -match "^(https?://|www\.)") {
                $checkUri = [System.Uri]($link -replace "^www\.", "https://www.")
                if ($checkUri.Host.Contains(".")) { $isValidUrl = $true }
            }
        }
        catch {}

        if (-not $isValidUrl) {
            [void][System.Windows.MessageBox]::Show("Please enter a valid URL (e.g., https://youtube.com/...)", "Invalid URL", 0, 16)
            return $false
        }

        $isSupported = $false
        try {
            $uri = [System.Uri]($link -replace "^www\.", "https://www.")
            $hostName = $uri.Host.ToLower() -replace "^www\.", ""
            
            $domainParts = $hostName -split '\.'
            $domainPart = $hostName
            if ($domainParts.Count -ge 2) {
                if ($domainParts[-2].Length -le 3 -and $domainParts.Count -ge 3) {
                    $domainPart = $domainParts[-3] 
                }
                else {
                    $domainPart = $domainParts[-2] 
                }
            }
            if ($hostName -match "youtu\.be") { $domainPart = "youtube" }

            $commonSites = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            [void]$commonSites.UnionWith(@("youtube", "youtu", "arte", "vimeo", "twitch", "facebook", "instagram", "twitter", "x", "tiktok", "soundcloud", "dailymotion", "reddit", "kick", "rumble"))
        
            if ($commonSites.Contains($domainPart) -or $commonSites.Contains($hostName)) {
                $isSupported = $true
            }

            if (-not $isSupported) {
                # Fetch supported sites list to cache it, preventing repetitive slow HTTP requests
                if ($null -eq $script:State.SupportedSitesCache -or $script:State.SupportedSitesCache.Length -eq 0) {
                    try {
                        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                        $rawUrl = "https://raw.githubusercontent.com/yt-dlp/yt-dlp/master/supportedsites.md"
                        $script:State.SupportedSitesCache = Invoke-RestMethod -Uri $rawUrl -UseBasicParsing
                    }
                    catch { $script:State.SupportedSitesCache = "fallback_offline" }
                }
            
                if ($script:State.SupportedSitesCache -match "(?i)\b$domainPart\b") { $isSupported = $true }
            }
        }
        catch { $isSupported = $true }

        if (-not $isSupported) {
            $win = New-Object System.Windows.Window
            $win.Title = "Unsupported Site"
            $win.SizeToContent = "WidthAndHeight"; $win.WindowStartupLocation = "CenterScreen"
            $win.Background = $window.Resources["BgBrush"]; $win.ResizeMode = "NoResize"
        
            $sp = New-Object System.Windows.Controls.StackPanel
            $sp.Margin = 20
        
            $tb = New-Object System.Windows.Controls.TextBlock
            $tb.TextWrapping = "Wrap"; $tb.MaxWidth = 450; $tb.Margin = "0,0,0,15"
            $tb.Foreground = $window.Resources["TextBrush"]
            $tb.Inlines.Add("The site '$domainPart' might not be officially supported.`n`nCheck the full list here:`n")
        
            $hl = New-Object System.Windows.Documents.Hyperlink
            $hl.NavigateUri = "https://github.com/yt-dlp/yt-dlp/blob/master/supportedsites.md"
            $hl.Inlines.Add("yt-dlp Supported Sites List")
            $hl.Add_Click({ [void](Start-Process $hl.NavigateUri) })
            $tb.Inlines.Add($hl)
            $tb.Inlines.Add("`n`nDo you want to attempt it anyway? It might fail due to DRM protection or generic extractor limits.")
        
            [void]$sp.Children.Add($tb)

            $btnSp = New-Object System.Windows.Controls.StackPanel
            $btnSp.Orientation = "Horizontal"; $btnSp.HorizontalAlignment = "Right"

            $btnTry = New-Object System.Windows.Controls.Button
            $btnTry.Content = "Try Anyway"; $btnTry.Width = 100; $btnTry.Height = 35; $btnTry.Margin = "0,0,10,0"
            $btnTry.Background = "#10B981"; $btnTry.Foreground = "White"; $btnTry.BorderThickness = 0; $btnTry.Cursor = "Hand"
            $btnTry.Add_Click({ $win.DialogResult = $true; $win.Close() })

            $btnCancel = New-Object System.Windows.Controls.Button
            $btnCancel.Content = "Cancel"; $btnCancel.Width = 100; $btnCancel.Height = 35
            $btnCancel.Background = "#6B7280"; $btnCancel.Foreground = "White"; $btnCancel.BorderThickness = 0; $btnCancel.Cursor = "Hand"
            $btnCancel.Add_Click({ $win.DialogResult = $false; $win.Close() })

            [void]$btnSp.Children.Add($btnTry); [void]$btnSp.Children.Add($btnCancel)
            [void]$sp.Children.Add($btnSp)
            $win.Content = $sp

            if ($win.ShowDialog() -ne $true) { 
                return $false 
            }
        }
        return $true
    }

    # Retrieves JSON data from the entered URL to display title, duration, and channel
    $Y_BtnPreview.Add_Click({
            $link = $Y_Link.Text.Trim()
            
            if ([string]::IsNullOrWhiteSpace($link) -or $link -eq "https://") { 
                [void][System.Windows.MessageBox]::Show("Please enter a valid URL to fetch.", "Missing URL", 0, 48)
                return 
            }
            if (-not $script:State.ytdlpFound) {
                [void][System.Windows.MessageBox]::Show("yt-dlp.exe not found!", "Missing Tool", 0, 48)
                return
            }

            $Y_BtnPreview.IsEnabled = $false
            $window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background)

            if (-not (CheckYtDlpSupport $link)) {
                $Y_BtnPreview.IsEnabled = $true
                return
            }

            $window.Cursor = [System.Windows.Input.Cursors]::Wait
            $StatusText.Text = "Fetching info..."
            $StatusText.Foreground = "#6366F1"
            $window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background)

            $tempJson = Join-Path $env:TEMP "yt_info.json"
            $tempErr = Join-Path $env:TEMP "yt_info_err.log"
            if (Test-Path $tempJson) { Remove-Item $tempJson -Force -ErrorAction SilentlyContinue }
            if (Test-Path $tempErr) { Remove-Item $tempErr -Force -ErrorAction SilentlyContinue }

            try {
                $argString = "--dump-json --no-warnings --no-playlist "
                if ($Y_CheckCookie.IsChecked) {
                    if ($Y_CookiePath.Text -and (Test-Path $Y_CookiePath.Text)) { $argString += "--cookies `"$($Y_CookiePath.Text)`" " }
                    else { $argString += "--cookies-from-browser $(Get-CbVal $Y_CookieBrowser) " }
                }
                $extArgs = "youtube:player_client=web,default"
                if (-not $Y_CheckAutoPoToken.IsChecked) {
                    $poToken = if ($Y_PoToken.Text) { $Y_PoToken.Text.Trim() } else { "" }
                    if ($poToken) { $extArgs += ";po_token=web+$poToken" }
                }
                $argString += "--extractor-args `"$extArgs`" `"$link`""

                [void](Start-Process -FilePath $script:State.ytdlp -ArgumentList $argString -WindowStyle Hidden -RedirectStandardOutput $tempJson -RedirectStandardError $tempErr -Wait)

                $jsonOut = if (Test-Path $tempJson) { Get-Content $tempJson -Raw -Encoding UTF8 } else { "" }
                $errOut = if (Test-Path $tempErr) { Get-Content $tempErr -Raw -Encoding UTF8 } else { "" }

                if (-not [string]::IsNullOrWhiteSpace($jsonOut)) {
                    try {
                        $videoData = $jsonOut | ConvertFrom-Json
                        if ($videoData -is [array]) { $videoData = $videoData[0] }
                
                        $title = if ($videoData.title) { $videoData.title } else { "Unknown Title" }
                        $dur = if ($videoData.duration_string) { $videoData.duration_string } else { "Unknown Duration" }
                        $chan = if ($videoData.uploader) { $videoData.uploader } else { "Unknown Uploader" }
                
                        [void][System.Windows.MessageBox]::Show("Title: $title`nDuration: $dur`nChannel: $chan", "Video Info", 0, 64)
                    }
                    catch {
                        [void][System.Windows.MessageBox]::Show("Could not cleanly parse info.`n`nRaw Output: $( $jsonOut.Substring(0, [math]::Min($jsonOut.Length, 300)) )...", "Preview Parsing Error", 0, 48)
                    }
                }
                else {
                    $errMsg = if ([string]::IsNullOrWhiteSpace($errOut)) { "Unknown error. Link might be fully protected or invalid." } else { $errOut }
                    [void][System.Windows.MessageBox]::Show("Could not fetch info.`n`nError Details:`n$errMsg", "Preview Error", 0, 48)
                }
            }
            catch {
                [void][System.Windows.MessageBox]::Show("Process execution failed: $($_.Exception.Message)", "Error", 0, 16)
            }
            finally {
                $window.Cursor = [System.Windows.Input.Cursors]::Arrow
                $Y_BtnPreview.IsEnabled = $true
                $StatusText.Text = "Ready."
                $StatusText.Foreground = $window.Resources["TextBrush"]
            }
        })

    # ==============================================================================
    # 8. TASK SCHEDULING & PROCESS MONITORING
    # ==============================================================================
    
    # Define polling timer to read external command output and translate to UI
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(250)

    # Completely terminate a process and its child processes
    function Stop-ProcessTree($proc) {
        try {
            if ($null -ne $proc -and -not $proc.HasExited) {
                # Removed -Wait to prevent the WPF UI thread from freezing during termination
                [void](Start-Process "taskkill.exe" -ArgumentList "/PID $($proc.Id) /T /F" -WindowStyle Hidden)
            }
        } 
        catch { try { $proc.Kill() } catch {} }
    }
    
    # Core loop logic for parsing the global task queue
    function ProcessNextJob {
        if ($script:State.CurrentJobIndex -ge $script:State.BatchQueue.Count) {
            if ($timer) { $timer.Stop() }
            $BtnCancel.IsEnabled = $false
            $BtnSkip.IsEnabled = $false
            $TxtETA.Text = "ETA: --:--"
            $TaskbarProgress.ProgressState = "None"
            
            # --- Print Batch Overview for yt-dlp ---
            $dlCount = if ($script:State.YtDlpDownloadedTitles) { $script:State.YtDlpDownloadedTitles.Count } else { 0 }
            $skipCount = if ($script:State.YtDlpSkippedLinks) { $script:State.YtDlpSkippedLinks.Count } else { 0 }

            if ($dlCount -gt 1 -or $skipCount -gt 0) {
                $LogBox.AppendText("`r`n============================================================`r`n")
                $LogBox.AppendText("[BATCH COMPLETE] Successfully processed your queue!`r`n`r`n")
                if ($dlCount -gt 0) {
                    $LogBox.AppendText("Downloaded ($dlCount) videos:`r`n")
                    foreach ($title in $script:State.YtDlpDownloadedTitles) { $LogBox.AppendText(" -> $title`r`n") }
                }
                if ($skipCount -gt 0) {
                    $LogBox.AppendText("`r`nSkipped the following links:`r`n")
                    foreach ($skipped in $script:State.YtDlpSkippedLinks) { $LogBox.AppendText("$skipped`r`n") }
                }
                $LogBox.AppendText("============================================================`r`n")
                if ($CbAutoScrollLog.IsChecked) { $LogBox.ScrollToEnd() }
            }
            
            if ($LogBox.Text -match "\[ERROR\]") {
                $StatusText.Text = "Finished with errors! Check the live log."
                $StatusText.Foreground = "#EF4444"; $TaskbarProgress.ProgressState = "Error"
            }
            elseif ($LogBox.Text -match "\[CANCEL\]") {
                $StatusText.Text = "Process cancelled."; $StatusText.Foreground = "#EF4444"
            }
            else {
                $StatusText.Text = "Successfully completed! ($($script:State.BatchQueue.Count) jobs processed)"
                $StatusText.Foreground = "#10B981" 
            }
            
            $PBar.Value = 100; $BtnRun.IsEnabled = $true; $BtnUpdate.IsEnabled = $true; $BtnShow.Visibility = "Visible"
            Write-ConvertLog "=== Queue Completed ==="
            if ($Config.PlaySound) { [System.Media.SystemSounds]::Asterisk.Play() }
            
            $msg = if ($LogBox.Text -match "\[ERROR\]") { "Process finished with errors. Check the live log for details." } else { "All queued files were processed successfully!" }
            Show-Toast -Title "Batch Complete" -Message $msg
            
            Save-Queue
            return
        }

        Save-Queue
        $job = $script:State.BatchQueue[$script:State.CurrentJobIndex]
        $StatusText.Text = "Processing job $($script:State.CurrentJobIndex + 1) of $($script:State.BatchQueue.Count)..."
        $StatusText.Foreground = "#6366F1"; $PBar.Value = 0; $TxtETA.Text = "ETA: Calc..."
        $TaskbarProgress.ProgressState = "Normal"; $TaskbarProgress.ProgressValue = 0.0

        try {
            # 1. Clean up old logs
            if (Test-Path $script:State.tempLog) { Remove-Item $script:State.tempLog -Force -ErrorAction SilentlyContinue }
            if (Test-Path $script:State.tempLogErr) { Remove-Item $script:State.tempLogErr -Force -ErrorAction SilentlyContinue }
            $script:State.lastLogPos = 0; $script:State.totalDuration = 0; $script:State.totalFrames = 0

            # 2. Setup Tool Path
            $toolPath = if ($job.IsYtDlp) { $script:State.ytdlp } elseif ($job.CustomTool) { $job.CustomTool } else { $script:State.ffmpeg }

            # 3. FAST FRAME PRE-FETCH & DURATION FETCH (Prevents UI Hang on large files)
            if (-not $job.IsYtDlp -and $script:State.ffprobeFound -and $job.InputFile -and (Test-Path -LiteralPath $job.InputFile)) {
                try {
                    $pinfoF = New-Object System.Diagnostics.ProcessStartInfo
                    $pinfoF.FileName = $script:State.ffprobe
                    $pinfoF.Arguments = '-v error -select_streams v:0 -show_entries stream=nb_frames -of default=noprint_wrappers=1:nokey=1 "{0}"' -f ($job.InputFile -replace '"', '\"')
                    $pinfoF.UseShellExecute = $false; $pinfoF.RedirectStandardOutput = $true; $pinfoF.RedirectStandardError = $true; $pinfoF.CreateNoWindow = $true
                    $pFs = [System.Diagnostics.Process]::Start($pinfoF)
                    
                    $fStr = $pFs.StandardOutput.ReadToEnd().Trim()
                    [void]$pFs.StandardError.ReadToEnd()
                    
                    if (-not $pFs.WaitForExit(3000)) { try { $pFs.Kill() } catch {} }
                    else {
                        if ($fStr -match '^\d+$') { $script:State.totalFrames = [int]$fStr }
                    }
                    $pFs.Dispose() # Free up Windows handle

                    # Fetch total duration safely before FFmpeg starts to guarantee ETA calc works
                    $pinfoDur = New-Object System.Diagnostics.ProcessStartInfo
                    $pinfoDur.FileName = $script:State.ffprobe
                    $pinfoDur.Arguments = '-v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 "{0}"' -f ($job.InputFile -replace '"', '\"')
                    $pinfoDur.UseShellExecute = $false; $pinfoDur.RedirectStandardOutput = $true; $pinfoDur.RedirectStandardError = $true; $pinfoDur.CreateNoWindow = $true
                    $pDur = [System.Diagnostics.Process]::Start($pinfoDur)
                    
                    $durStr = $pDur.StandardOutput.ReadToEnd().Trim()
                    [void]$pDur.StandardError.ReadToEnd()
                    
                    if (-not $pDur.WaitForExit(3000)) { try { $pDur.Kill() } catch {} }
                    else {
                        $d = 0.0
                        if ([double]::TryParse($durStr, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d) -and $d -gt 0) {
                            $script:State.totalDuration = $d
                        }
                    }
                    $pDur.Dispose() # REQUIRED to prevent Windows handle leaks
                }
                catch { Write-CrashLog "Fast Frame & Duration pre-fetch failed: $($_.Exception.Message)" }
            }

            # 4. Build Clean Arguments
            $rawArgs = @()
            if (-not $job.IsYtDlp -and $job.CustomTool -notmatch "python.exe|upscayl") {
                if ($Config.ThreadLimit -and $Config.ThreadLimit -ne "Auto") { $rawArgs += @("-threads", $Config.ThreadLimit) }
            }
            foreach ($arg in $job.Args) { $rawArgs += $arg.ToString().Trim() }
            
            $argString = ($rawArgs | ForEach-Object {
                $a = [string]$_
                # Robustly quote arguments and escape existing quotes to prevent command injection
                if ($a -match '[ &|<>]') { "`"$($a -replace '"', '\"')`"" } else { $a }
            }) -join " "

            # 5. UI PUMP (Clears spinning icon natively without WinForms leaks)
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            $window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background)

            # 6. Logging
            $friendlyToolName = if ($job.IsYtDlp) { "yt-dlp" } elseif ($job.CustomTool) { [System.IO.Path]::GetFileNameWithoutExtension($job.CustomTool) } else { "ffmpeg" }
            $LogBox.AppendText("`r`n============================================================`r`n")
            $LogBox.AppendText("[$($friendlyToolName.ToUpper()) COMMAND USED]:`r`n")
            $LogBox.AppendText("`"$toolPath`" $argString`r`n")
            $LogBox.AppendText("============================================================`r`n")
            Write-ConvertLog "Executing: $toolPath $argString"

            # 7. Safe Execution via CMD /S /C
            $combinedLog = $script:State.tempLog
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "cmd.exe"
            $psi.Arguments = "/S /C ""`"$toolPath`" $argString > `"$combinedLog`" 2>&1"""
            $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true

            if ($job.OutputDir -and -not (Test-Path $job.OutputDir)) { [void](New-Item -ItemType Directory -Path $job.OutputDir -Force) }
            $psi.WorkingDirectory = if ($job.OutputDir -and (Test-Path $job.OutputDir)) { $job.OutputDir } else { $ScriptDir }

            $script:State.p = [System.Diagnostics.Process]::Start($psi)
            if ($null -eq $script:State.p) { throw "Process failed to start." }

            $timer.Start()
        }
        catch { 
            $errMsg = "Failed processing job $($script:State.CurrentJobIndex + 1): $($_.Exception.Message)"
            Write-CrashLog "$errMsg"
            $LogBox.AppendText("`r`n[CRITICAL ERROR] $errMsg`r`n")
            $script:State.CurrentJobIndex++; ProcessNextJob
        }
    }
    # Routine logic for polling output file of cmd wrapper and translating string matches to progress bar
    $timer.Add_Tick({
            if (Test-Path -LiteralPath $script:State.tempLog) {
                try {
                    $newText = ""
                    try {
                        # Open with full sharing permissions to prevent FFmpeg from blocking our read
                        $fs = [System.IO.File]::Open($script:State.tempLog, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                        $reader = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
                        
                        if ($script:State.lastLogPos -lt $reader.BaseStream.Length) {
                            $reader.BaseStream.Seek($script:State.lastLogPos, [System.IO.SeekOrigin]::Begin) | Out-Null
                            $newText = $reader.ReadToEnd()
                            $script:State.lastLogPos = $reader.BaseStream.Position
                        }
                    }
                    finally {
                        # Guarantee stream disposal. Disposing the reader automatically disposes the underlying stream.
                        if ($null -ne $reader) { $reader.Dispose() }
                        elseif ($null -ne $fs) { $fs.Dispose() }
                    }

                    if (-not [string]::IsNullOrEmpty($newText)) {
                        $LogBox.AppendText($newText)
                        
                        # Prevent memory leaks and UI freezing during massive FFmpeg/yt-dlp logs
                        # Keep the buffer strictly smaller to prevent massive WPF layout recalculations
                        if ($LogBox.Text.Length -gt 15000) {
                            $LogBox.Text = $LogBox.Text.Substring($LogBox.Text.Length - 10000)
                        }

                        if ($CbAutoScrollLog.IsChecked) {
                            $LogBox.ScrollToEnd()
                        }
                
                        $job = $script:State.BatchQueue[$script:State.CurrentJobIndex]

                        # Detect FFMPEG duration AND total frames for better progress tracking
                        if ($newText -match "Duration:\s*(\d{2}:\d{2}:\d{2})") {
                            try { if ($script:State.totalDuration -eq 0) { $script:State.totalDuration = [TimeSpan]::Parse($matches[1]).TotalSeconds } } catch {}
                        }
                        if ($newText -match "NUMBER_OF_FRAMES\s*:\s*(\d+)") {
                            try { if ($null -eq $script:State.totalFrames -or $script:State.totalFrames -eq 0) { $script:State.totalFrames = [int]$matches[1] } } catch {}
                        }

                        if ($job.CustomTool -match "HandBrakeCLI") {
                            $hbMatches = $script:State.Regex.HbProg.Matches($newText)
                            if ($hbMatches.Count -gt 0) {
                                try {
                                    $PBar.Value = [math]::Min(100.0, [double]::Parse($hbMatches[$hbMatches.Count - 1].Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture))
                                    $TaskbarProgress.ProgressValue = ($PBar.Value / 100)
                                    
                                    $hbEta = $script:State.Regex.HbETA.Matches($newText)
                                    if ($hbEta.Count -gt 0) { $TxtETA.Text = "ETA: " + $hbEta[$hbEta.Count - 1].Groups[1].Value }
                                } catch {}
                            }
                        }

                        if ($job.IsYtDlp) {
                            # Regex logic to capture yt-dlp percentage and string ETA
                            $percentMatches = $script:State.Regex.YtDlpProg.Matches($newText)
                            if ($percentMatches.Count -gt 0) {
                                $PBar.Value = [math]::Round([double]$percentMatches[$percentMatches.Count - 1].Groups[1].Value)
                                $TaskbarProgress.ProgressValue = ($PBar.Value / 100)
                            }
                            $etaMatches = $script:State.Regex.YtDlpETA.Matches($newText)
                            if ($etaMatches.Count -gt 0) {
                                $TxtETA.Text = "ETA: " + $etaMatches[$etaMatches.Count - 1].Groups[1].Value
                            }
                        }
                        elseif ($newText -match "(\d+\.\d+)%") {
                            # Regex Logic for Upscayl percentage 
                            $upscaleMatches = $script:State.Regex.UpscaylProg.Matches($newText)
                            try {
                                $PBar.Value = [math]::Min(100.0, [double]::Parse($upscaleMatches[$upscaleMatches.Count - 1].Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture))
                                $TaskbarProgress.ProgressValue = ($PBar.Value / 100)
                                $TxtETA.Text = "Status: Upscaling..."
                            }
                            catch {}
                        }
                        elseif ($job.IsWhisper -and $script:State.totalDuration -gt 0) {
                            # Regex logic mapping Whisper VTT output string lines back into duration
                            $whisperMatches = $script:State.Regex.WhisperTime.Matches($newText)
                            if ($whisperMatches.Count -gt 0) {
                                $wm = $whisperMatches[$whisperMatches.Count - 1]
                                try {
                                    $h = if ($wm.Groups[1].Value) { [int]$wm.Groups[1].Value } else { 0 }
                                    $m = [int]$wm.Groups[2].Value; $s = [int]$wm.Groups[3].Value; $ms = [int]$wm.Groups[4].Value
                                    $curr = ($h * 3600) + ($m * 60) + $s + ($ms / 1000.0)
                                    $prog = ($curr / $script:State.totalDuration) * 100
                                    $PBar.Value = [math]::Min([double]100.0, [double]$prog)
                                    $TaskbarProgress.ProgressValue = ($PBar.Value / 100)
                                    $TxtETA.Text = "Status: Transcribing..."
                                }
                                catch {}
                            }
                        }
                        # RESTORED & BULLETPROOFED: FFmpeg Progress Tracker with ETA Calculation
                        elseif ($script:State.totalDuration -gt 0) {
                        
                            $timeMatches = $script:State.Regex.FFmpegTime.Matches($newText)
                            $frameMatches = $script:State.Regex.FFmpegFrame.Matches($newText)
                        
                            if ($timeMatches.Count -gt 0) {
                                try {
                                    $lastTime = $timeMatches[$timeMatches.Count - 1]
                                    $h = [int]$lastTime.Groups[1].Value
                                    $m = [int]$lastTime.Groups[2].Value
                                    $s = [int]$lastTime.Groups[3].Value
                                    $ms = if ($lastTime.Groups[4].Success) { [int]($lastTime.Groups[4].Value.PadRight(3, '0').Substring(0, 3)) } else { 0 }
                                    $currentTime = ($h * 3600) + ($m * 60) + $s + ($ms / 1000.0)
                                
                                    # Pure PowerShell math (avoids strict .NET type crashes)
                                    $prog = ($currentTime / $script:State.totalDuration) * 100
                                    if ($prog -gt 100) { $prog = 100 }
                                    if ($prog -lt 0) { $prog = 0 }
                                
                                    $PBar.Value = $prog
                                    $TaskbarProgress.ProgressValue = ($prog / 100)

                                    $speedMatches = $script:State.Regex.Speed.Matches($newText)
                                    if ($speedMatches.Count -gt 0) {
                                        $speedStr = $speedMatches[$speedMatches.Count - 1].Groups[1].Value
                                        $speed = 0.0
                                    
                                        # TryParse prevents crashes if the speed text is momentarily garbled
                                        if ([double]::TryParse($speedStr, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$speed) -and $speed -gt 0) {
                                        
                                            $remainingSecs = ($script:State.totalDuration - $currentTime) / $speed
                                            if ($remainingSecs -lt 0) { $remainingSecs = 0 }
                                        
                                            $ts = [TimeSpan]::FromSeconds($remainingSecs)
                                            $hrs = [int][math]::Truncate($ts.TotalHours)
                                        
                                            if ($hrs -ge 1) {
                                                $TxtETA.Text = "ETA: {0:D2}:{1:D2}:{2:D2}" -f $hrs, $ts.Minutes, $ts.Seconds
                                            }
                                            else {
                                                $TxtETA.Text = "ETA: {0:D2}:{1:D2}" -f $ts.Minutes, $ts.Seconds
                                            }
                                        }
                                        else {
                                            $TxtETA.Text = "Status: Converting... ($([math]::Round($prog))%)"
                                        }
                                    }
                                    else { 
                                        $TxtETA.Text = "Status: Converting... ($([math]::Round($prog))%)" 
                                    }
                                }
                                catch { 
                                    # If math fails, update UI with percentage so it doesn't appear frozen
                                    $TxtETA.Text = "Status: Processing... ($([math]::Round($PBar.Value))%)"
                                }
                            }
                            # FALLBACK: If time is N/A (common with some exact stream copies), use frames
                            elseif ($frameMatches.Count -gt 0 -and $script:State.totalFrames -gt 0) {
                                try {
                                    $currentFrame = [int]$frameMatches[$frameMatches.Count - 1].Groups[1].Value
                                    $prog = ($currentFrame / $script:State.totalFrames) * 100
                                    if ($prog -gt 100) { $prog = 100 }
                                    if ($prog -lt 0) { $prog = 0 }
                                
                                    $PBar.Value = $prog
                                    $TaskbarProgress.ProgressValue = ($prog / 100)

                                    $elapsed = (Get-Date) - $job.JobStart
                                    if ($elapsed.TotalSeconds -gt 3 -and $currentFrame -gt 0) {
                                        $framesPerSec = $currentFrame / $elapsed.TotalSeconds
                                        $remainingFrames = $script:State.totalFrames - $currentFrame
                                        $remainingSecs = $remainingFrames / $framesPerSec
                                        if ($remainingSecs -lt 0) { $remainingSecs = 0 }
                                    
                                        $ts = [TimeSpan]::FromSeconds($remainingSecs)
                                        $hrs = [int][math]::Truncate($ts.TotalHours)
                                    
                                        if ($hrs -ge 1) {
                                            $TxtETA.Text = "ETA: {0:D2}:{1:D2}:{2:D2}" -f $hrs, $ts.Minutes, $ts.Seconds
                                        }
                                        else {
                                            $TxtETA.Text = "ETA: {0:D2}:{1:D2}" -f $ts.Minutes, $ts.Seconds
                                        }
                                    }
                                    else {
                                        $TxtETA.Text = "ETA: Calculating..."
                                    }
                                }
                                catch {
                                    $TxtETA.Text = "Status: Processing... ($([math]::Round($PBar.Value))%)"
                                }
                            }
                        }
                    }
                }
                catch {}
            }

            # Evaluate execution exit state logic (Ensured this is the ONLY cleanly integrated exit block)
            if ($script:State.p -and $script:State.p.HasExited) { 
                $timer.Stop()
                $job = $script:State.BatchQueue[$script:State.CurrentJobIndex]
        
                [int]$exCode = 0
                try { 
                    if ($null -ne $script:State.p.ExitCode) {
                        $exCode = [int]$script:State.p.ExitCode 
                    }
                }
                catch {}

                if ($exCode -eq 0 -or ($job.IsYtDlp -and $exCode -in @(1, 2))) {
            
                    $logText = $LogBox.Text
            
                    # Advanced parsing: sometimes tools don't return non-zero exit codes correctly on failure
                    if ($job.CustomTool -match "upscayl" -and ($logText -match "Error: Unknown model" -or $logText -match "failed to load" -or $logText -match "invalid param")) {
                        $StatusText.Text = "Failed (Model Error)"
                        $StatusText.Foreground = "#EF4444"
                        $LogBox.AppendText("`r`n[ERROR] Upscayl failed to load the AI models. The model directory might be missing or invalid.`r`n")
                        $exCode = 1
                    }
                    elseif (-not $job.IsYtDlp -and -not $job.CustomTool -and ($logText -match "Error opening input" -or $logText -match "Error binding filtergraph" -or $logText -match "Invalid argument" -or $logText -match "Cannot find an unused" -or $logText -match "Option not found" -or $logText -match "Unable to open" -or $logText -match "Error initializing filters" -or $logText -match "No such file or directory" -or $logText -match "Unrecognized option")) {
                        $StatusText.Text = "Failed (FFmpeg Error)"
                        $StatusText.Foreground = "#EF4444"
                        $LogBox.AppendText("`r`n[ERROR] FFmpeg encountered an error. The process was aborted.`r`n")
                        $exCode = 1
                    }
                    elseif ($job.IsYtDlp -and ($logText -match "ERROR:" -or $logText -match "Could not find known video or audio" -or $logText -match "Unsupported URL" -or $logText -match "This video is unavailable")) {
                        if (-not $job.Retried) {
                            $StatusText.Text = "Updating yt-dlp & Retrying..."
                            $StatusText.Foreground = "#F59E0B"
                            $LogBox.AppendText("`r`n[WARNING] Extraction failed. Attempting to auto-update yt-dlp and retry...`r`n")
                            
                            # Fire native yt-dlp updater silently to fetch newest YouTube extractors
                            try { [void](Start-Process cmd.exe -ArgumentList "/c `"$($script:State.ytdlp)`" -U" -Wait -WindowStyle Hidden) } catch {}
                            
                            $job.Retried = $true
                            $script:State.BatchQueue[$script:State.CurrentJobIndex] = $job
                            
                            # Decrement index by 1 so the queue loops back to this exact job
                            $script:State.CurrentJobIndex--
                            $exCode = 0 
                        }
                        else {
                            $StatusText.Text = "Failed (Extraction Error)"
                            $StatusText.Foreground = "#EF4444"
                            $LogBox.AppendText("`r`n[ERROR] yt-dlp could not extract media even after an update. URL might be unsupported, protected, or missing a video.`r`n")
                            $exCode = 1
                        }
                    }
                    elseif ($job.IsYtDlp -and $logText -notmatch "\[download\]" -and $logText -notmatch "\[info\]" -and $logText -notmatch "has already been downloaded") {
                        $StatusText.Text = "Failed (No Media Found)"
                        $StatusText.Foreground = "#EF4444"
                        $LogBox.AppendText("`r`n[ERROR] No valid media could be found to download. The site might not contain an extractable video/audio.`r`n")
                        $exCode = 1
                    }
                    elseif ($exCode -ne 0) {
                        $StatusText.Text = "Finished (With minor warnings)."
                        $StatusText.Foreground = "#F59E0B"
                        $LogBox.AppendText("`r`n[WARNING] Completed, but some post-processing (e.g. metadata/thumbnail) had issues.`r`n")
                    }
                    else {
                        $StatusText.Text = "Finished Successfully."
                        $StatusText.Foreground = "#10B981"

                        if ($job.IsYtDlp) {
                            try {
                                $finalFile = $job.OutputFile -replace '\.part$', '' -replace '\.ytdl$', ''
                                if (-not $finalFile -or -not (Test-Path -LiteralPath $finalFile)) {
                                    $finalFile = (Get-ChildItem -Path $job.OutputDir -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
                                }

                                if ($finalFile -and (Test-Path -LiteralPath $finalFile)) {
                                    $dlTitle = [System.IO.Path]::GetFileNameWithoutExtension($finalFile)
                                    $dlDur = "Unknown"
                                    $dlQual = "Audio Only"

                                    # --- ADDED: Save title to memory for the batch overview ---
                                    if ($null -eq $script:State.YtDlpDownloadedTitles) { $script:State.YtDlpDownloadedTitles = @() }
                                    $script:State.YtDlpDownloadedTitles += $dlTitle
                                    # ----------------------------------------------------------

                                    # Only run FFprobe if it actually exists on the system
                                    if ($script:State.ffprobeFound) {
                                        $pinfoDur = New-Object System.Diagnostics.ProcessStartInfo
                                        $pinfoDur.FileName = $script:State.ffprobe
                                        $pinfoDur.Arguments = "-v error -show_entries format=duration:stream=width,height -of json `"$finalFile`""
                                        $pinfoDur.UseShellExecute = $false; $pinfoDur.RedirectStandardOutput = $true; $pinfoDur.RedirectStandardError = $true; $pinfoDur.CreateNoWindow = $true
                                        $pDur = [System.Diagnostics.Process]::Start($pinfoDur)
                                        
                                        $jsonOut = $pDur.StandardOutput.ReadToEnd()
                                        [void]$pDur.StandardError.ReadToEnd()
                                        
                                        if (-not $pDur.WaitForExit(3000)) { try { $pDur.Kill() } catch {} }
                                        $pDur.Dispose()
                                    
                                        if ($jsonOut) {
                                            $mediaInfo = $jsonOut | ConvertFrom-Json
                                            if ($mediaInfo.format.duration) {
                                                $ts = [TimeSpan]::FromSeconds([double]$mediaInfo.format.duration)
                                                $dlDur = "{0:D2}:{1:D2}:{2:D2}" -f $ts.Hours, $ts.Minutes, $ts.Seconds
                                            }
                                            if ($mediaInfo.streams) {
                                                $videoStream = $mediaInfo.streams | Where-Object { $_.width } | Select-Object -First 1
                                                if ($videoStream) {
                                                    $dlQual = "$($videoStream.width)x$($videoStream.height)"
                                                    if ($videoStream.height -ge 2160) { $dlQual += " (4K)" }
                                                    elseif ($videoStream.height -ge 1440) { $dlQual += " (2K)" }
                                                    elseif ($videoStream.height -ge 1080) { $dlQual += " (1080p)" }
                                                    elseif ($videoStream.height -ge 720) { $dlQual += " (720p)" }
                                                }
                                            }
                                        }
                                    }

                                    $LogBox.AppendText("`r`n----------------------------------------`r`n")
                                    $LogBox.AppendText(" DOWNLOAD SUMMARY`r`n")
                                    $LogBox.AppendText(" Title:   $dlTitle`r`n")
                                    $LogBox.AppendText(" Time:    $dlDur`r`n")
                                    $LogBox.AppendText(" Quality: $dlQual`r`n")
                                    $LogBox.AppendText("----------------------------------------`r`n")
                                }
                            }
                            catch {}
                        }

                        $LogBox.AppendText("`r`n[SUCCESS] Task completed cleanly.`r`n")

                        if ($job.IsWhisper -or $job.CustomTool -match "upscayl") {
                            $LogBox.AppendText("`r`n[SUCCESS] Processing saved to: $($job.OutputDir)`r`n")
                        }
                
                        # Deletes base source file upon success if set in settings
                        if ($Config.AutoDelete -and -not $job.IsYtDlp -and $job.ListItem -and (Test-Path -LiteralPath $job.ListItem)) {
                            Remove-Item -LiteralPath $job.ListItem -Force -ErrorAction SilentlyContinue
                            $LogBox.AppendText("[CLEANUP] Deleted original file: $(Split-Path $job.ListItem -Leaf)`r`n")
                        }
                    }
                }
                else {
                    $StatusText.Text = "Failed (Exit Code: $exCode)"
                    $StatusText.Foreground = "#EF4444"
                    $LogBox.AppendText("`r`n[ERROR] Process aborted with error code $exCode.`r`n")
    
                    if (Test-Path $script:State.tempLogErr) {
                        try {
                            $errText = [System.IO.File]::ReadAllText($script:State.tempLogErr, [System.Text.Encoding]::Default)
                            if (-not [string]::IsNullOrWhiteSpace($errText)) {
                                $LogBox.AppendText("TECHNICAL ERROR LOG:`r`n$errText`r`n")
                                Write-CrashLog "Task exited with code $exCode. Output: $errText"
                            }
                        }
                        catch {}
                    }
                }

                if ($exCode -ne 0 -and $job.OutputFile -and (Test-Path -LiteralPath $job.OutputFile)) {
                    Remove-Item -LiteralPath $job.OutputFile -Force -ErrorAction SilentlyContinue
                    $LogBox.AppendText("`r`n[CLEANUP] Deleted incomplete output file: $(Split-Path $job.OutputFile -Leaf)`r`n")
                }

                if (-not [string]::IsNullOrWhiteSpace($LogBox.Text)) { Write-ConvertLog $LogBox.Text }

                if ($script:State.p) { try { $script:State.p.Dispose() } catch {}; $script:State.p = $null }

                try {
                    if ($job.ListBox -ne $null -and $job.ListItem -ne $null) {
                        $job.ListBox.Items.Remove($job.ListItem)
                    }
                }
                catch {}

                if ($CbAutoScrollLog.IsChecked) {
                    $LogBox.ScrollToEnd()
                }

                $script:State.CurrentJobIndex++
                [void][System.Windows.Threading.Dispatcher]::CurrentDispatcher.InvokeAsync({ ProcessNextJob })
            }
        })

    # Main entry point when the user clicks 'START PROCESS'. Parses inputs, creates jobs, and starts the queue.
    $BtnRun.Add_Click({
            MissingToolsCheck

            # Pre-fetch supported sites cache silently in the background to prevent UI freezing later
            if ($null -eq $script:State.SupportedSitesCache) {
                [void][System.Threading.Tasks.Task]::Run([Action] {
                        try {
                            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                            $rawUrl = "https://raw.githubusercontent.com/yt-dlp/yt-dlp/master/supportedsites.md"
                            $script:State.SupportedSitesCache = Invoke-RestMethod -Uri $rawUrl -UseBasicParsing
                        }
                        catch { 
                            $script:State.SupportedSitesCache = "fallback_offline"
                            Write-CrashLog "Failed to download supported sites: $($_.Exception.Message)"
                        }
                    })
            }
            $tabIndex = $MainTabs.SelectedIndex
    
            $script:State.BatchQueue = @()
            $script:State.CurrentJobIndex = 0
            $script:State.YtDlpDownloadedTitles = @() 
            $script:State.YtDlpSkippedLinks = @() # ADDED: Track skipped links
            $LogBox.Clear()

            $customParamText = ""
            $toolDocs = ""
            $toolName = ""
            
            # Validate custom flags before execution to ensure standard parameter behavior
            if ($tabIndex -eq 0 -and $A_CheckCustomParams.IsChecked -and -not [string]::IsNullOrWhiteSpace($A_CustomParams.Text)) { $customParamText = $A_CustomParams.Text; $toolDocs = "https://ffmpeg.org/ffmpeg.html"; $toolName = "FFmpeg" }
            elseif ($tabIndex -eq 1 -and $V_CheckCustomParams.IsChecked -and -not [string]::IsNullOrWhiteSpace($V_CustomParams.Text)) { $customParamText = $V_CustomParams.Text; $toolDocs = "https://ffmpeg.org/ffmpeg.html"; $toolName = "FFmpeg" }
            elseif ($tabIndex -eq 4 -and $Y_CheckCustomParams.IsChecked -and -not [string]::IsNullOrWhiteSpace($Y_CustomParams.Text)) { $customParamText = $Y_CustomParams.Text; $toolDocs = "https://github.com/yt-dlp/yt-dlp#usage-and-options"; $toolName = "yt-dlp" }

            if ($customParamText) {
                $isSuspicious = ($customParamText.Trim() -notmatch "^-")
                $msg = ""
                
                if ($isSuspicious) {
                    $msg = "Warning: Your custom parameters do not start with a hyphen (-). This is usually invalid syntax!`n`n"
                }
                else {
                    $msg = "Notice: You are using manual custom parameters.`n`n"
                }
                
                $msg += "If your parameters are wrong, the process will fail and will NOT auto-retry.`n`nCheck $toolName syntax here:`n$toolDocs`n`nDo you want to proceed?"
                
                $ans = [System.Windows.MessageBox]::Show($msg, "Custom Parameters Validation", "YesNo", "Warning")
                if ($ans -eq "No") { return }
            }

            # Direct parsing block for Tab: Muxing
            if ($tabIndex -eq 3) {
                if (-not $script:State.ffmpegFound) { [void][System.Windows.MessageBox]::Show("FFmpeg not found!", "Missing Tool", 0, 48); return }
                if ([string]::IsNullOrWhiteSpace($M_InVideo.Text) -or -not (Test-Path -LiteralPath $M_InVideo.Text) -or [string]::IsNullOrWhiteSpace($M_InAudio.Text) -or -not (Test-Path -LiteralPath $M_InAudio.Text)) { 
                    [void][System.Windows.MessageBox]::Show("Please select valid video and audio files.", "File Error", 0, 48)
                    return 
                }
    
                $baseOut = $M_OutFile.Text
                if ([string]::IsNullOrWhiteSpace($baseOut)) {
                    $outDir = Split-Path $M_InVideo.Text -Parent
                    $baseTargetFile = Join-Path $outDir "$([System.IO.Path]::GetFileNameWithoutExtension($M_InVideo.Text))-muxed.mp4" 
                }
                else {
                    $outDir = [System.IO.Path]::GetDirectoryName($baseOut)
                    $baseTargetFile = $baseOut
                }
                if (-not (Test-Path $outDir)) { [void](New-Item -ItemType Directory -Path $outDir -Force) }
        
                $script:State.lastOutDir = $outDir
                $targetFile = Get-UniqueFileName $baseTargetFile
        
                $argArray = @("-hide_banner", "-y", "-i", $M_InVideo.Text, "-i", $M_InAudio.Text, "-c", "copy", "-map", "0:v:0", "-map", "1:a:0", "-shortest", $targetFile)
                $script:State.BatchQueue += @{ Args = $argArray; SafeArgs = $argArray; HasCustomParams = $false; Retried = $false; IsYtDlp = $false; OutputFile = $targetFile; ListBox = $null; ListItem = $null }
            }
            # Direct parsing block for Tab: Download
            elseif ($tabIndex -eq 4) {
                if (-not $script:State.ytdlpFound) { 
                    [void][System.Windows.MessageBox]::Show("yt-dlp.exe not found!", "Missing Tool", 0, 48)
                    return 
                }

                $linksToProcess = [System.Collections.Generic.List[string]]::new()

                # Check which sub-tab the user is currently using
                if ($Y_InputTabs.SelectedIndex -eq 1) {
                    $batchPath = $Y_BatchFile.Text.Trim()
                    if ([string]::IsNullOrWhiteSpace($batchPath) -or -not (Test-Path -LiteralPath $batchPath)) {
                        [void][System.Windows.MessageBox]::Show("Please select a valid .txt file containing URLs.", "Missing Batch File", 0, 16)
                        return
                    }
                    
                    # Read text file, ignore empty lines and comments starting with #
                    try {
                        $lines = Get-Content -LiteralPath $batchPath -ErrorAction Stop
                    } catch {
                        [void][System.Windows.MessageBox]::Show("Could not read batch file. It may be open in another program.`n`n$($_.Exception.Message)", "File Error", 0, 16)
                        return
                    }
                    foreach ($line in $lines) {
                        $l = $line.Trim()
                        if (-not [string]::IsNullOrWhiteSpace($l) -and $l -notmatch "^#" -and $l -match "^(https?://|www\.)") {
                            $linksToProcess.Add($l)
                        }
                    }
                    
                    if ($linksToProcess.Count -eq 0) {
                        [void][System.Windows.MessageBox]::Show("No valid https:// links found in the text file.", "Empty File", 0, 48)
                        return
                    }
                }
                else {
                    $singleLink = $Y_Link.Text.Trim()
                    if ([string]::IsNullOrWhiteSpace($singleLink) -or $singleLink -eq "https://" -or $singleLink -notmatch "^(https?://|www\.)") { 
                        [void][System.Windows.MessageBox]::Show("Please enter a valid URL. (Format: https://website.com)", "Error", 0, 16)
                        return 
                    }
                    $linksToProcess.Add($singleLink)
                }

                $baseOut = $Y_OutDir.Text
                if ($baseOut -match "Select target" -or [string]::IsNullOrWhiteSpace($baseOut)) { 
                    $outDir = if ($Y_Type.SelectedIndex -eq 1) { Join-Path $ScriptDir "download\audio" } else { Join-Path $ScriptDir "download\video" }
                }
                else { 
                    $outDir = $baseOut
                }

                if (-not (Test-Path $outDir)) { [void](New-Item -ItemType Directory -Path $outDir -Force) }
                $script:State.lastOutDir = $outDir

                # --- BATCH PRE-SCAN FOR UNSUPPORTED DOMAINS ---
                $unsupportedDomains = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                $validLinks = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                
                # Fetch supported sites list to cache it for the batch check
                if ($null -eq $script:State.SupportedSitesCache -or $script:State.SupportedSitesCache.Length -eq 0) {
                    try {
                        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                        $rawUrl = "https://raw.githubusercontent.com/yt-dlp/yt-dlp/master/supportedsites.md"
                        $script:State.SupportedSitesCache = Invoke-RestMethod -Uri $rawUrl -UseBasicParsing
                    }
                    catch { $script:State.SupportedSitesCache = "fallback_offline" }
                }

                foreach ($chkLink in $linksToProcess) {
                    if ([string]::IsNullOrWhiteSpace($chkLink) -or $chkLink -notmatch "^(https?://|www\.)") { continue }
                    
                    $isSup = $false
                    $domainPart = "Unknown"
                    try {
                        $uri = [System.Uri]($chkLink -replace "^www\.", "https://www.")
                        $hostName = $uri.Host.ToLower() -replace "^www\.", ""
                        $domainParts = $hostName -split '\.'
                        $domainPart = $hostName
                        if ($domainParts.Count -ge 2) {
                            if ($domainParts[-2].Length -le 3 -and $domainParts.Count -ge 3) { $domainPart = $domainParts[-3] }
                            else { $domainPart = $domainParts[-2] }
                        }
                        if ($hostName -match "youtu\.be") { $domainPart = "youtube" }

                        $commonSites = @("youtube", "youtu", "arte", "vimeo", "twitch", "facebook", "instagram", "twitter", "x", "tiktok", "soundcloud", "dailymotion", "reddit", "kick", "rumble")
                        if ($commonSites -contains $domainPart -or $commonSites -contains $hostName) { $isSup = $true }

                        if (-not $isSup) {
                            if ($script:State.SupportedSitesCache -match "(?i)\b$domainPart\b") { $isSup = $true }
                        }
                    }
                    catch { $isSup = $true }

                    if ($isSup) {
                        [void]$validLinks.Add($chkLink)
                    }
                    else {
                        [void]$unsupportedDomains.Add($domainPart)
                    }
                }

                $batchIgnoreUnsupported = $false
                if ($unsupportedDomains.Count -gt 0) {
                    $domainList = ($unsupportedDomains -join ", ")
                    $msg = "Your list includes potentially unsupported pages:`n`n$domainList`n`nDo you want to attempt downloading them anyway? (They might fail, but won't break the queue)"
                    
                    $win = New-Object System.Windows.Window
                    $win.Title = "Unsupported Sites Detected"
                    $win.SizeToContent = "WidthAndHeight"; $win.WindowStartupLocation = "CenterScreen"
                    $win.Background = $window.Resources["BgBrush"]; $win.ResizeMode = "NoResize"
                    
                    $sp = New-Object System.Windows.Controls.StackPanel
                    $sp.Margin = 20
                    
                    $tb = New-Object System.Windows.Controls.TextBlock
                    $tb.TextWrapping = "Wrap"; $tb.MaxWidth = 450; $tb.Margin = "0,0,0,15"
                    $tb.Foreground = $window.Resources["TextBrush"]
                    $tb.Text = $msg
                    [void]$sp.Children.Add($tb)

                    $btnSp = New-Object System.Windows.Controls.StackPanel
                    $btnSp.Orientation = "Horizontal"; $btnSp.HorizontalAlignment = "Right"

                    $btnTry = New-Object System.Windows.Controls.Button
                    $btnTry.Content = "Yes (Try Anyway)"; $btnTry.Width = 120; $btnTry.Height = 35; $btnTry.Margin = "0,0,10,0"
                    $btnTry.Background = "#10B981"; $btnTry.Foreground = "White"; $btnTry.BorderThickness = 0; $btnTry.Cursor = "Hand"
                    $btnTry.Add_Click({ $win.DialogResult = $true; $win.Close() })

                    $btnCancel = New-Object System.Windows.Controls.Button
                    $btnCancel.Content = "No (Skip Them)"; $btnCancel.Width = 120; $btnCancel.Height = 35
                    $btnCancel.Background = "#6B7280"; $btnCancel.Foreground = "White"; $btnCancel.BorderThickness = 0; $btnCancel.Cursor = "Hand"
                    $btnCancel.Add_Click({ $win.DialogResult = $false; $win.Close() })

                    [void]$btnSp.Children.Add($btnTry); [void]$btnSp.Children.Add($btnCancel)
                    [void]$sp.Children.Add($btnSp)
                    $win.Content = $sp

                    if ($win.ShowDialog() -eq $true) {
                        $batchIgnoreUnsupported = $true
                    }
                    else {
                        [void][System.Windows.MessageBox]::Show("Unsupported links will be skipped. Only valid/supported links will be downloaded.", "Skipping", 0, 64)
                    }
                }
                # --------------------------------------------------

                $script:State.PlaylistChoice = $null
                $isBatch = ($linksToProcess.Count -gt 1)

                # --- NEW CUSTOM PLAYLIST DIALOG FUNCTION ---
                function Show-PlaylistDialog([string]$linkUrl) {
                    $win = New-Object System.Windows.Window
                    $win.Title = "Playlist Detected"
                    $win.SizeToContent = "WidthAndHeight"; $win.WindowStartupLocation = "CenterScreen"
                    $win.Background = $window.Resources["BgBrush"]; $win.ResizeMode = "NoResize"; $win.Topmost = $true
                    
                    $sp = New-Object System.Windows.Controls.StackPanel
                    $sp.Margin = 20
                    
                    $tb = New-Object System.Windows.Controls.TextBlock
                    $tb.TextWrapping = "Wrap"; $tb.MaxWidth = 450; $tb.Margin = "0,0,0,15"
                    $tb.Foreground = $window.Resources["TextBrush"]
                    $tb.Text = "A Playlist was detected in the batch with the following link:`n`n$linkUrl`n`nHow would you like to handle this for this or other playlists in the batch?"
                    [void]$sp.Children.Add($tb)

                    $btnSp = New-Object System.Windows.Controls.WrapPanel
                    $btnSp.HorizontalAlignment = "Center"

                    # Helper to cleanly generate the 5 buttons
                    function New-Btn($content, $bg) {
                        $b = New-Object System.Windows.Controls.Button
                        $b.Content = $content; $b.Height = 35; $b.Margin = "5"; $b.Padding = "12,0"
                        $b.Background = $bg; $b.Foreground = "White"; $b.BorderThickness = 0; $b.Cursor = "Hand"
                        return $b
                    }

                    # We change the visible text here, but keep the $script:plResult exact!
                    $btnYes = New-Btn "Yes (Download this playlist)" "#10B981"
                    $btnYes.Add_Click({ $script:plResult = "Yes"; $win.Close() })
                    
                    $btnYesAll = New-Btn "Yes to All (Download all playlists)" "#059669"
                    $btnYesAll.Add_Click({ $script:plResult = "YesToAll"; $win.Close() })
                    
                    $btnNo = New-Btn "No (Don't download this playlist)" "#F59E0B"
                    $btnNo.Add_Click({ $script:plResult = "No"; $win.Close() })
                    
                    $btnNoAll = New-Btn "No to All (Don't download any playlists)" "#D97706"
                    $btnNoAll.Add_Click({ $script:plResult = "NoToAll"; $win.Close() })

                    $btnCancel = New-Btn "Cancel" "#EF4444"
                    $btnCancel.Add_Click({ $script:plResult = "Cancel"; $win.Close() })

                    [void]$btnSp.Children.Add($btnYes); [void]$btnSp.Children.Add($btnYesAll)
                    [void]$btnSp.Children.Add($btnNo); [void]$btnSp.Children.Add($btnNoAll); [void]$btnSp.Children.Add($btnCancel)
                    [void]$sp.Children.Add($btnSp)
                    $win.Content = $sp

                    $script:plResult = "Cancel"
                    [void]$win.ShowDialog()
                    return $script:plResult
                }
                # -------------------------------------------

                # Iterate through all links and queue them up
                foreach ($processLink in $linksToProcess) {
                    
                    # Validate against our pre-scan results to avoid spamming popups!
                    if ([string]::IsNullOrWhiteSpace($processLink) -or $processLink -notmatch "^(https?://|www\.)") { continue }
                    
                    if ($batchIgnoreUnsupported -eq $false -and -not $validLinks.Contains($processLink)) {
                        $LogBox.AppendText("[SKIP] Ignored unsupported link: $processLink`r`n")
                        
                        # Track skipped links for the batch overview
                        if ($null -eq $script:State.YtDlpSkippedLinks) { $script:State.YtDlpSkippedLinks = @() }
                        $script:State.YtDlpSkippedLinks += $processLink
                        
                        continue 
                    }

                    # Catch playlists gracefully and remember "To All" choices
                    $playlistFlag = "--no-playlist"
                    if ($processLink -match "list=") {
                        if ($isBatch) {
                            if ($script:State.PlaylistChoice -eq "YesToAll") {
                                $playlistFlag = "--yes-playlist"
                            }
                            elseif ($script:State.PlaylistChoice -eq "NoToAll") {
                                $playlistFlag = "--no-playlist"
                            }
                            else {
                                $ans = Show-PlaylistDialog -linkUrl $processLink
                                
                                if ($ans -eq "Cancel") { 
                                    $script:State.BatchQueue = @() # Wipes the queue instantly
                                    $LogBox.AppendText("`r`n[CANCEL] Batch queue setup aborted by user.`r`n")
                                    return # Stops processing the rest of the text file immediately
                                }
                                elseif ($ans -eq "YesToAll") {
                                    $script:State.PlaylistChoice = "YesToAll"
                                    $playlistFlag = "--yes-playlist"
                                }
                                elseif ($ans -eq "NoToAll") {
                                    $script:State.PlaylistChoice = "NoToAll"
                                    $playlistFlag = "--no-playlist"
                                }
                                elseif ($ans -eq "Yes") {
                                    $playlistFlag = "--yes-playlist"
                                }
                            }
                        }
                        else {
                            $ans = [System.Windows.MessageBox]::Show("A Playlist was detected in the following link:`n`n$processLink`n`nDo you want to download the FULL playlist?`n(Selecting 'No' will download just the single video)", "Playlist Detected", "YesNoCancel", "Question")
                            
                            if ($ans -eq "Cancel") { 
                                $LogBox.AppendText("`r`n[CANCEL] Download setup aborted by user.`r`n")
                                return 
                            }
                            
                            $playlistFlag = if ($ans -eq "Yes") { "--yes-playlist" } else { "--no-playlist" }
                        }
                    }

                    $jobStart = Get-Date
                    $existing = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                    try { foreach ($f in [System.IO.Directory]::EnumerateFiles($outDir)) { [void]$existing.Add($f) } } catch {}

                    # Notice we pass $processLink here instead of leaving it empty
                    $argsSafe = Get-YtDlpArgs -isPreview $false -ExcludeCustom $true -PlaylistFlag $playlistFlag -TargetLink $processLink
                    if ($null -eq $argsSafe) { continue }

                    $argsFull = Get-YtDlpArgs -isPreview $false -ExcludeCustom $false -PlaylistFlag $playlistFlag -TargetLink $processLink
                    $hasCustom = [bool]($Y_CheckCustomParams.IsChecked -and -not [string]::IsNullOrWhiteSpace($Y_CustomParams.Text))

                    $script:State.BatchQueue += @{ 
                        Args            = $argsFull
                        SafeArgs        = $argsSafe
                        HasCustomParams = $hasCustom
                        Retried         = $false
                        IsYtDlp         = $true
                        SelectedRes     = (Get-CbVal $Y_Res)
                        OutputFile      = $null
                        OutputDir       = $outDir
                        JobStart        = $jobStart
                        ExistingFiles   = $existing
                        ActiveOutput    = $null
                        ListBox         = $null
                        ListItem        = $null 
                    }
                }

                if ($script:State.BatchQueue.Count -eq 0) { return } # Stop if all links were cancelled/invalid

                if (-not (Get-Variable -Name "Resolve-YtDlpActiveOutputForCurrentJob" -Scope Script -ErrorAction SilentlyContinue)) {
                    
                    function Get-YtDlpActiveOutputFile {
                        param(
                            [string]  $OutDir,
                            [datetime]$StartTime,
                            [string[]]$ExistingFiles
                        )
                        $candidates = Get-ChildItem -LiteralPath $OutDir -File -ErrorAction SilentlyContinue
                        foreach ($f in $candidates) {
                            if ($ExistingFiles.Contains($f.FullName)) { continue }

                            $delta = ($f.LastWriteTime - $StartTime).TotalSeconds
                            if ($delta -lt -5) { continue }

                            if ($f.Name -match '.*\.[^\.]+(\.part)?$') {
                                return $f
                            }
                        }
                        return $null
                    }

                    function Resolve-YtDlpActiveOutputForCurrentJob {
                        $idx = $script:State.CurrentJobIndex
                        if ($idx -lt 0 -or $idx -ge $script:State.BatchQueue.Count) { return }
                        $job = $script:State.BatchQueue[$idx]
                        if (-not $job.IsYtDlp) { return }
                        if ($job.ActiveOutput) { return }

                        $active = Get-YtDlpActiveOutputFile -OutDir $job.OutputDir -StartTime $job.JobStart -ExistingFiles $job.ExistingFiles
                        if ($active) {
                            $job.ActiveOutput = $active.FullName
                            $job.OutputFile = $active.FullName
                            $script:State.BatchQueue[$idx] = $job

                            $BtnCancel.IsEnabled = $true
                            $LogBox.AppendText("`r`n[INFO] Active output for current yt-dlp job: $($active.Name)`r`n")
                        }
                    }
                    Set-Variable -Name "Resolve-YtDlpActiveOutputForCurrentJob" -Scope Script -Value (Get-Command Resolve-YtDlpActiveOutputForCurrentJob -CommandType Function)
                }
            }            
            # Direct parsing block for Tab: Specials (AI and Repair tools)
            elseif ($tabIndex -eq 5) {
                if (-not $script:State.ffmpegFound) { [void][System.Windows.MessageBox]::Show("FFmpeg not found!", "Missing Tool", 0, 48); return }
                
                $subIdx = $SpecialSubTabs.SelectedIndex

                # AI Transcriber Sub-Tab
                if ($subIdx -eq 0) {
                    $whisperInstalled = $false
                    if (Get-Command "whisper" -ErrorAction SilentlyContinue) { $whisperInstalled = $true }
                    elseif (Get-Command "python" -ErrorAction SilentlyContinue) {
                        $pipCheck = & python -c "import pkgutil; print(1 if pkgutil.find_loader('whisper') else 0)" 2>$null
                        if ($pipCheck -match "1") { $whisperInstalled = $true }
                    }

                    # Automates Whisper/Python installation if missing
                    if (-not $whisperInstalled) {
                        $win = New-Object System.Windows.Window
                        $win.Title = "Missing AI Component"
                        $win.SizeToContent = "WidthAndHeight"
                        $win.WindowStartupLocation = "CenterScreen"
                        $win.ResizeMode = "NoResize"
                        $win.Background = $window.Resources["BgBrush"]
                        
                        $sp = New-Object System.Windows.Controls.StackPanel
                        $sp.Margin = 20
                        $tbMsg = New-Object System.Windows.Controls.TextBlock
                        $tbMsg.TextWrapping = "Wrap"
                        $tbMsg.MaxWidth = 450
                        $tbMsg.Foreground = $window.Resources["TextBrush"]
                        $tbMsg.Margin = "0,0,0,15"
                        
                        $tbMsg.Inlines.Add("OpenAI Whisper is required to transcribe audio, but it is not installed on your system.`n`nWould you like to install it automatically now? (This uses Python).`n")
                        [void]$sp.Children.Add($tbMsg)
                        
                        $btnSp = New-Object System.Windows.Controls.StackPanel
                        $btnSp.Orientation = "Horizontal"
                        $btnSp.HorizontalAlignment = "Right"

                        $btnAuto = New-Object System.Windows.Controls.Button
                        $btnAuto.Content = "Install Whisper AI"
                        $btnAuto.Width = 140; $btnAuto.Height = 35; $btnAuto.Margin = "0,0,10,0"
                        $btnAuto.Background = "#10B981"; $btnAuto.Foreground = "White"; $btnAuto.BorderThickness = 0; $btnAuto.Cursor = "Hand"
                        $btnAuto.Add_Click({ $win.DialogResult = $true; $win.Close() })

                        $btnCancel = New-Object System.Windows.Controls.Button
                        $btnCancel.Content = "Cancel"
                        $btnCancel.Width = 100; $btnCancel.Height = 35
                        $btnCancel.Background = "#6B7280"; $btnCancel.Foreground = "White"; $btnCancel.BorderThickness = 0; $btnCancel.Cursor = "Hand"
                        $btnCancel.Add_Click({ $win.DialogResult = $false; $win.Close() })

                        [void]$btnSp.Children.Add($btnAuto); [void]$btnSp.Children.Add($btnCancel)
                        [void]$sp.Children.Add($btnSp)
                        $win.Content = $sp

                        if ($win.ShowDialog() -eq $true) {
                            $StatusText.Text = "Checking Python & Installing Whisper..."
                            $LogBox.AppendText("`r`n[INSTALL] Starting Whisper installation...`r`n")
                            
                            $pythonCheck = Get-Command "python" -ErrorAction SilentlyContinue
                            if (-not $pythonCheck) {
                                if (-not (Get-Command "winget" -ErrorAction SilentlyContinue)) {
                                    $LogBox.AppendText("[ERROR] Winget is missing. Cannot auto-install Python.`r`n")
                                    [void][System.Windows.MessageBox]::Show("Winget is missing. Please install Python manually to use Whisper.", "Winget Not Found", 0, 16)
                                    return
                                }
                                $LogBox.AppendText("[INSTALL] Python not found. Installing Python via WinGet...`r`n")
                                $pyProc = Start-Process winget -ArgumentList "install --id Python.Python.3 --silent --force --accept-source-agreements --accept-package-agreements" -Wait -PassThru
                                if ($pyProc.ExitCode -ne 0) {
                                    $LogBox.AppendText("[WARNING] Python installation cancelled or failed.`r`n")
                                    [void][System.Windows.MessageBox]::Show("Python installation was cancelled. Whisper requires Python.", "Installation Aborted", 0, 48)
                                    return
                                }
                                $env:PATH = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
                            }

                            $LogBox.AppendText("[INSTALL] Running pip install openai-whisper...`r`n")
                            $pipProc = Start-Process cmd.exe -ArgumentList "/c echo Installing OpenAI Whisper... && pip install -U openai-whisper" -Wait -WindowStyle Normal -PassThru

                            if ($pipProc.ExitCode -ne 0) {
                                $LogBox.AppendText("[WARNING] Whisper pip installation cancelled or failed.`r`n")
                                [void][System.Windows.MessageBox]::Show("Whisper installation was cancelled or failed.", "Installation Aborted", 0, 48)
                                return
                            }

                            $LogBox.AppendText("[SUCCESS] Whisper installed successfully!`r`n")
                            $StatusText.Text = "Whisper ready."
                            [void][System.Windows.MessageBox]::Show("Whisper AI has been installed successfully!`n`nNo restart required! The process will now continue automatically.", "Installation Complete", 0, 64)
                        }
                        else {
                            return
                        }
                    }

                    if ([string]::IsNullOrWhiteSpace($S_ScribeIn.Text) -or -not (Test-Path -LiteralPath $S_ScribeIn.Text)) { 
                        [void][System.Windows.MessageBox]::Show("Please select a valid media file!", "Error", 0, 48)
                        return 
                    }
                    
                    $inFile = $S_ScribeIn.Text
                    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
                    
                    # Prevent crashes when reading files from Read-Only drives (CD/DVD/USB)
                    $baseOutDir = if ($Config.DefaultOutDir -and (Test-Path $Config.DefaultOutDir)) { $Config.DefaultOutDir } else { Join-Path $ScriptDir "convert\transcribed" }
                    $whisperOutDir = Join-Path $baseOutDir "transcribe_$ts"
                    
                    if (-not (Test-Path $whisperOutDir)) { [void](New-Item -ItemType Directory -Path $whisperOutDir -Force) }
                    
                    $script:State.lastOutDir = $whisperOutDir
                    $outFormat = (Get-CbVal $S_ScribeFormat)
                    if ($outFormat -match "txt") { $outFormat = "txt" }
                    elseif ($outFormat -match "srt") { $outFormat = "srt" }
                    elseif ($outFormat -match "vtt") { $outFormat = "vtt" }
                    elseif ($outFormat -match "json") { $outFormat = "json" }

                    $outSrt = Join-Path $whisperOutDir "$([System.IO.Path]::GetFileNameWithoutExtension($inFile)).$outFormat"
                    
                    $model = (Get-CbVal $S_ScribeModel) -split " " | Select-Object -First 1
                    $Config.WhisperModel = $model
                    
                    $scribeArgs = @("-u", "-m", "whisper", $inFile, "--model", $model, "--output_format", $outFormat, "--output_dir", $whisperOutDir, "--fp16", "False")

                    $task = Get-CbVal $S_ScribeTask
                    if ($task -match "Translate") {
                        $scribeArgs += @("--task", "translate")
                    }

                    $lang = (Get-CbVal $S_ScribeLang)
                    if ($lang -match "Auto-Detect") {
                        $LogBox.AppendText("`r`n[INFO] Auto-Detect selected. Whisper will transcribe in the original language.`r`n")
                    }
                    else {
                        $langCode = switch ($lang) {
                            "English" { "en" }
                            "German" { "de" }
                            "Spanish" { "es" }
                            "French" { "fr" }
                            "Italian" { "it" }
                            "Dutch" { "nl" }
                            "Russian" { "ru" }
                            default { "en" }
                        }
                        $scribeArgs += @("--language", $langCode)
                    }
            
                    $ffmpegDir = [System.IO.Path]::GetDirectoryName($script:State.ffmpeg)
                    if ($env:PATH -notmatch [regex]::Escape($ffmpegDir)) { $env:PATH = "$ffmpegDir;" + $env:PATH }
            
                    $script:State.BatchQueue += @{ Args = $scribeArgs; SafeArgs = $scribeArgs; HasCustomParams = $false; Retried = $false; IsYtDlp = $false; CustomTool = "python.exe"; IsWhisper = $true; OutputDir = $whisperOutDir; InputFile = $inFile; OutputFile = $outSrt; ListBox = $null; ListItem = $null }
                    
                    # Logic to immediately queue an ffmpeg job to burn the transcribed file onto the video
                    if ($S_CheckBurn.IsChecked) {
                        $ext = [System.IO.Path]::GetExtension($inFile).ToLower()
                        $audioExts = @(".mp3", ".wav", ".m4a", ".flac", ".ogg", ".aac")
                        
                        if ($audioExts -contains $ext) {
                            [void][System.Windows.MessageBox]::Show("You cannot burn subtitles into an audio-only file ($ext).`n`nThe transcription will only be saved as a text/subtitle file in your source folder.", "Audio File Detected", 0, 64)
                        }
                        elseif ($outFormat -in @("srt", "vtt")) {
                            $outFile = Get-UniqueFileName (Join-Path $whisperOutDir "SCRIBED_$([System.IO.Path]::GetFileName($inFile))")
                            $safeOutSrt = $outSrt.Replace('\', '/').Replace(':', '\:').Replace('[', '\[').Replace(']', '\]').Replace("'", "\'").Replace(',', '\,')
                            $burnArgs = @("-hide_banner", "-y", "-i", $inFile, "-vf", "subtitles='''$safeOutSrt'''", "-c:a", "copy", $outFile)
                            $script:State.BatchQueue += @{ Args = $burnArgs; SafeArgs = $burnArgs; HasCustomParams = $false; Retried = $false; IsYtDlp = $false; CustomTool = ""; IsWhisper = $false; OutputDir = $whisperOutDir; InputFile = $inFile; OutputFile = $outFile; ListBox = $null; ListItem = $null }
                        }
                    }
                }
                # Audio Visualizer Sub-Tab
                elseif ($subIdx -eq 2) {
                    if ([string]::IsNullOrWhiteSpace($S_VisAudio.Text) -or -not (Test-Path -LiteralPath $S_VisAudio.Text)) { 
                        [void][System.Windows.MessageBox]::Show("Please select a valid audio file!", "Error", 0, 48)
                        return 
                    }
                    $outDir = Split-Path $S_VisAudio.Text -Parent
                    $script:State.lastOutDir = $outDir
                    $outFile = Get-UniqueFileName (Join-Path $outDir "$([System.IO.Path]::GetFileNameWithoutExtension($S_VisAudio.Text))_visualizer.mp4")
                    
                    $filter = switch ($S_VisStyle.SelectedIndex) {
                        0 { "showwaves=s=1280x720:mode=line:colors=cyan" }
                        1 { "showfreqs=s=1280x720:mode=bar:colors=magenta" }
                        2 { "avectorscope=s=720x720:zoom=1.5:rc=0:gc=255:bc=0" }
                    }

                    if (-not [string]::IsNullOrWhiteSpace($S_VisImg.Text) -and (Test-Path -LiteralPath $S_VisImg.Text)) {
                        $argArray = @("-hide_banner", "-y", "-loop", "1", "-i", $S_VisImg.Text, "-i", $S_VisAudio.Text, "-filter_complex", "[1:a]$filter[vis];[0:v][vis]overlay=x=(main_w-overlay_w)/2:y=(main_h-overlay_h)/2:format=auto,format=yuv420p[v]", "-map", "[v]", "-map", "1:a", "-c:v", "libx264", "-tune", "stillimage", "-c:a", "aac", "-shortest", $outFile)
                    }
                    else {
                        $argArray = @("-hide_banner", "-y", "-i", $S_VisAudio.Text, "-filter_complex", "[0:a]$filter[v]", "-map", "[v]", "-map", "0:a", "-c:v", "libx264", "-pix_fmt", "yuv420p", "-c:a", "aac", $outFile)
                    }
                    $script:State.BatchQueue += @{ Args = $argArray; SafeArgs = $argArray; HasCustomParams = $false; Retried = $false; IsYtDlp = $false; OutputFile = $outFile; ListBox = $null; ListItem = $null }
                }
                # Video Stabilizer Sub-Tab
                elseif ($subIdx -eq 3) {
                    if ([string]::IsNullOrWhiteSpace($S_StabIn.Text) -or -not (Test-Path -LiteralPath $S_StabIn.Text)) { 
                        [void][System.Windows.MessageBox]::Show("Please select a valid video file to stabilize!", "Error", 0, 48)
                        return
                    }

                    $outDir = Split-Path $S_StabIn.Text -Parent
                    $script:State.lastOutDir = $outDir
                    $outFile = Get-UniqueFileName (Join-Path $outDir "STABILIZED_$([System.IO.Path]::GetFileName($S_StabIn.Text))")
                    
                    # 1. Just generate a unique filename (NO drive letter, NO colons, NO slashes)
                    # Because WorkDir is set to $env:TEMP below, FFmpeg will save/read this in the Temp folder naturally.
                    $trfName = "stab_$([guid]::NewGuid().ToString().Substring(0,8)).trf"

                    $level = $S_StabLevel.SelectedIndex
                    $shakiness = if ($level -eq 0) { 3 } elseif ($level -eq 1) { 5 } else { 8 }
                    $smoothing = if ($level -eq 0) { 10 } elseif ($level -eq 1) { 15 } else { 25 }

                    # PASS 1: Detection
                    $filter1 = "vidstabdetect=shakiness=${shakiness}:result=$trfName"
                    $argPass1 = @("-hide_banner", "-y", "-i", $S_StabIn.Text, "-vf", $filter1, "-f", "null", "-")

                    $script:State.BatchQueue += @{ Args = $argPass1; SafeArgs = $argPass1; HasCustomParams = $false; Retried = $false; IsYtDlp = $false; OutputFile = $null; ListBox = $null; ListItem = $null; WorkDir = $env:TEMP }

                    # PASS 2: Transformation
                    $filter2 = "vidstabtransform=smoothing=${smoothing}:input=$trfName"
                    $argPass2 = @("-hide_banner", "-y", "-i", $S_StabIn.Text, "-vf", $filter2, "-c:v", "libx264", "-pix_fmt", "yuv420p", "-c:a", "copy", $outFile)

                    $script:State.BatchQueue += @{ Args = $argPass2; SafeArgs = $argPass2; HasCustomParams = $false; Retried = $false; IsYtDlp = $false; OutputFile = $outFile; ListBox = $null; ListItem = $S_StabIn.Text; WorkDir = $env:TEMP }
                }
                # AI Upscaler Sub-Tab
                elseif ($subIdx -eq 1) {
                    
                    if (-not $script:State.upscaylFound) { 
                        $win = New-Object System.Windows.Window
                        $win.Title = "Missing AI Component"
                        $win.SizeToContent = "WidthAndHeight"
                        $win.WindowStartupLocation = "CenterScreen"
                        $win.ResizeMode = "NoResize"
                        $win.Background = $window.Resources["BgBrush"]
                        
                        $sp = New-Object System.Windows.Controls.StackPanel
                        $sp.Margin = 20
                        $tbMsg = New-Object System.Windows.Controls.TextBlock
                        $tbMsg.TextWrapping = "Wrap"
                        $tbMsg.MaxWidth = 450
                        $tbMsg.Foreground = $window.Resources["TextBrush"]
                        $tbMsg.Margin = "0,0,0,15"
                        
                        $tbMsg.Inlines.Add("Upscayl is required for AI upscaling, but it is not found on your system.`n`nWould you like to download and install it automatically via WinGet now?`n")
                        [void]$sp.Children.Add($tbMsg)
                        
                        $btnSp = New-Object System.Windows.Controls.StackPanel
                        $btnSp.Orientation = "Horizontal"
                        $btnSp.HorizontalAlignment = "Right"

                        $btnAuto = New-Object System.Windows.Controls.Button
                        $btnAuto.Content = "Install Upscayl"
                        $btnAuto.Width = 140; $btnAuto.Height = 35; $btnAuto.Margin = "0,0,10,0"
                        $btnAuto.Background = "#10B981"; $btnAuto.Foreground = "White"; $btnAuto.BorderThickness = 0; $btnAuto.Cursor = "Hand"
                        $btnAuto.Add_Click({ $win.DialogResult = $true; $win.Close() })

                        $btnCancel = New-Object System.Windows.Controls.Button
                        $btnCancel.Content = "Cancel"
                        $btnCancel.Width = 100; $btnCancel.Height = 35
                        $btnCancel.Background = "#6B7280"; $btnCancel.Foreground = "White"; $btnCancel.BorderThickness = 0; $btnCancel.Cursor = "Hand"
                        $btnCancel.Add_Click({ $win.DialogResult = $false; $win.Close() })

                        [void]$btnSp.Children.Add($btnAuto); [void]$btnSp.Children.Add($btnCancel)
                        [void]$sp.Children.Add($btnSp)
                        $win.Content = $sp

                        if ($win.ShowDialog() -eq $true) {
                            if (-not (Get-Command "winget" -ErrorAction SilentlyContinue)) {
                                [void][System.Windows.MessageBox]::Show("Windows Package Manager (winget) is not installed on your system.`n`nPlease install Upscayl manually.", "Winget Not Found", 0, 16)
                                return
                            }

                            $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                            if (-not $isAdmin) {
                                [void][System.Windows.MessageBox]::Show("No admin rights to install Upscayl.`n`nPlease restart the program with Administrator rights (Right-click -> Run as administrator) to continue.", "Admin Rights Required", 0, 48)
                                return
                            }

                            $StatusText.Text = "Installing Upscayl..."
                            $LogBox.AppendText("`r`n[INSTALL] Installing Upscayl via WinGet...`r`n")
                            $TaskbarProgress.ProgressState = "Indeterminate"
                            
                            $installProc = Start-Process winget -ArgumentList "install -e --id Upscayl.Upscayl --force --accept-source-agreements --accept-package-agreements" -Wait -WindowStyle Normal -PassThru
                            
                            Find-Tools
                            
                            if (-not $script:State.upscaylFound) {
                                $TaskbarProgress.ProgressState = "None"
                                if ($installProc.ExitCode -ne 0) {
                                    $LogBox.AppendText("[WARNING] Upscayl installation was cancelled or failed (Exit Code: $($installProc.ExitCode)).`r`n")
                                    [void][System.Windows.MessageBox]::Show("Installation was cancelled or encountered an error.", "Installation Aborted", 0, 48)
                                }
                                else {
                                    $LogBox.AppendText("[ERROR] Upscayl installed, but the backend binary (upscayl-bin.exe) could not be found.`r`n")
                                    [void][System.Windows.MessageBox]::Show("Upscayl installed, but the system couldn't locate the backend engine.`n`nYou may need to restart the application.", "Path Error", 0, 48)
                                }
                                return
                            }

                            $LogBox.AppendText("[SUCCESS] Upscayl backend mapped successfully!`r`n")
                            $StatusText.Text = "Upscayl ready."
                            $TaskbarProgress.ProgressState = "None"
                            [void][System.Windows.MessageBox]::Show("Upscayl has been installed successfully!`n`nThe process will now continue automatically.", "Installation Complete", 0, 64)
                        }
                        else { return }
                    }

                    if (-not $script:State.upscaylFound) { return }
                    
                    if ([string]::IsNullOrWhiteSpace($S_UpscaleIn.Text) -or -not (Test-Path -LiteralPath $S_UpscaleIn.Text)) { 
                        [void][System.Windows.MessageBox]::Show("Please select a valid media file to upscale!", "Error", 0, 48)
                        return 
                    }

                    $ext = [System.IO.Path]::GetExtension($S_UpscaleIn.Text).ToLower()
                    if ($ext -match "\.(mp4|mkv|avi|mov|webm)$") {
                        [void][System.Windows.MessageBox]::Show("The Upscayl backend currently only supports image files directly.`n`nTo upscale a video, you must extract its frames into a sequence of images, upscale the images, and then merge them back together.", "Images Only", 0, 48)
                        return
                    }

                    $esrganOutDir = if ($S_UpscaleOutDir.Text -match "Select target folder" -or [string]::IsNullOrWhiteSpace($S_UpscaleOutDir.Text)) { 
                        Join-Path $ScriptDir "special\upscaled" 
                    }
                    else { 
                        $S_UpscaleOutDir.Text 
                    }
                    if (-not (Test-Path $esrganOutDir)) { [void](New-Item -ItemType Directory -Path $esrganOutDir -Force) }
                    $script:State.lastOutDir = $esrganOutDir

                    $safeInputName = [System.IO.Path]::GetFileName($S_UpscaleIn.Text).Replace(" ", "_")
                    $outFile = Join-Path $esrganOutDir "UPSCALED_$safeInputName"
                    
                    $counter = 1
                    while (Test-Path -LiteralPath $outFile) {
                        $name = [System.IO.Path]::GetFileNameWithoutExtension($safeInputName)
                        $ext = [System.IO.Path]::GetExtension($safeInputName)
                        $outFile = Join-Path $esrganOutDir "UPSCALED_${name}_${counter}${ext}"
                        $counter++
                    }

                    $model = (Get-CbVal $S_UpscaleModel).Split(" ")[0]
                    $scale = (Get-CbVal $S_UpscaleScale).Replace("x", "")
                    
                    $exeDir = Split-Path $script:State.upscayl -Parent
                    $parentDir = Split-Path $exeDir -Parent
                    $grandParentDir = Split-Path $parentDir -Parent
                    
                    # Logic to deeply scan typical Upscayl model installations
                    $searchPaths = @(
                        $script:State.upscaylModels,
                        (Join-Path $exeDir "models"),
                        (Join-Path $parentDir "models"),
                        (Join-Path $grandParentDir "models"),
                        (Join-Path $parentDir "app.asar.unpacked\models"),
                        "C:\Program Files\Upscayl\resources\models",
                        "$env:LOCALAPPDATA\Programs\upscayl\resources\models"
                    )
                    
                    $mDir = $null
                    foreach ($p in $searchPaths) {
                        if ($p -and (Test-Path -LiteralPath (Join-Path $p "$model.param"))) {
                            $mDir = $p
                            break
                        }
                    }

                    if (-not $mDir) {
                        [void][System.Windows.MessageBox]::Show("The AI models (like '$model.param') could not be found.`n`nPlease check your Upscayl installation.", "Missing Model", 0, 48)
                        return
                    }

                    $inFilePath = $S_UpscaleIn.Text
                    try {
                        # Resolve short paths via native Win32 API to avoid COM overhead/AV blocks
                        $sb = New-Object System.Text.StringBuilder(32767)
                        if (Test-Path $mDir) { [void][WinApi.PathHelper]::GetShortPathName($mDir, $sb, $sb.Capacity); $mDir = $sb.ToString() }
                        if (Test-Path $inFilePath) { [void][WinApi.PathHelper]::GetShortPathName($inFilePath, $sb, $sb.Capacity); $inFilePath = $sb.ToString() }
                        if (Test-Path $esrganOutDir) { 
                            [void][WinApi.PathHelper]::GetShortPathName($esrganOutDir, $sb, $sb.Capacity)
                            $outFile = Join-Path $sb.ToString() (Split-Path $outFile -Leaf)
                        }
                    }
                    catch { Write-Warning "Short path resolution failed: $_" }

                    $argArray = @("-i", $inFilePath, "-o", $outFile, "-n", $model, "-s", $scale, "-m", $mDir)

                    $script:State.BatchQueue += @{ Args = $argArray; SafeArgs = $argArray; HasCustomParams = $false; Retried = $false; IsYtDlp = $false; CustomTool = $script:State.upscayl; OutputFile = $outFile; ListBox = $null; ListItem = $null; OutputDir = $esrganOutDir }
                }
            }
            # Generic parsing block for Audio, Video, and Image tabs executing loop-based batches
            else {
                if (-not $script:State.ffmpegFound) { [void][System.Windows.MessageBox]::Show("FFmpeg not found!", "Missing Tool", 0, 48); return }
                $list = if ($tabIndex -eq 0) { $A_InList } elseif ($tabIndex -eq 1) { $V_InList } else { $I_InList }
                if ($list.Items.Count -eq 0) { 
                    [void][System.Windows.MessageBox]::Show("Please select files or enter an URL before starting the process!", "No files selected or URL entered", 0, 48)
                    return 
                }

                $baseOut = if ($tabIndex -eq 0) { $A_OutDir.Text } elseif ($tabIndex -eq 1) { $V_OutDir.Text } else { $I_OutDir.Text }
                $useSourceDir = ($baseOut -match "target folder" -or [string]::IsNullOrWhiteSpace($baseOut))

                foreach ($inFile in $list.Items) {
                    
                    if ($useSourceDir) {
                        $outDir = Split-Path $inFile -Parent
                    }
                    else {
                        $outDir = $baseOut
                    }

                    if (-not (Test-Path $outDir)) { [void](New-Item -ItemType Directory -Path $outDir -Force) }
                    $script:State.lastOutDir = $outDir

                    $name = [System.IO.Path]::GetFileNameWithoutExtension($inFile)
                    $argArray = @("-hide_banner", "-y")
            
                    if ($tabIndex -eq 0) { 
                        $fmt = (Get-CbVal $A_CFormat).ToLower()
                        $outFile = Get-UniqueFileName (Join-Path $outDir "$name.$fmt")
                        
                        $argsSafe = (Get-AudioFfmpegArgs -IsPreview $false -inFile $inFile -outFile $outFile -ExcludeCustom $true).Args
                        $argsFull = (Get-AudioFfmpegArgs -IsPreview $false -inFile $inFile -outFile $outFile -ExcludeCustom $false).Args
                        $hasCustom = [bool]($A_CheckCustomParams.IsChecked -and -not [string]::IsNullOrWhiteSpace($A_CustomParams.Text))
                        
                        $script:State.BatchQueue += @{ Args = $argsFull; SafeArgs = $argsSafe; HasCustomParams = $hasCustom; Retried = $false; IsYtDlp = $false; OutputFile = $outFile; ListBox = $list; ListItem = $inFile }
                    }
                    elseif ($tabIndex -eq 1) { 
                        $fmtRaw = Get-CbVal $V_CFormat
                        $fmt = if ([string]::IsNullOrWhiteSpace($fmtRaw)) { "mp4" } else { $fmtRaw.ToLower() }
                        
                        $smartName = Get-SmartVideoFilename $inFile
                        $outFile = Get-UniqueFileName (Join-Path $outDir "$smartName.$fmt")

                        # --- HANDBRAKE ROUTING LOGIC ---
                        $vCodecCb = Get-CbVal $V_CCodec
                        if ($V_UseHandbrake.IsChecked -and $script:State.handbrakeFound -and $vCodecCb -notmatch "Copy") {
                            $argsFull = (Get-HandbrakeArgs -inFile $inFile -outFile $outFile).Args
                            $hasCustom = $false
                            $targetTool = $script:State.handbrakecli
                        }
                        else {
                            if ($V_UseHandbrake.IsChecked -and -not $script:State.handbrakeFound) {
                                [void][System.Windows.MessageBox]::Show("HandBrake CLI not found! Falling back to FFmpeg.", "Missing Tool", 0, 48)
                            }
                            $argsFull = (Get-FfmpegArgs -IsPreview $false -inFile $inFile -outFile $outFile -ExcludeCustom $false).Args
                            $hasCustom = [bool]($V_CheckCustomParams.IsChecked -and -not [string]::IsNullOrWhiteSpace($V_CustomParams.Text))
                            $targetTool = "" # Defaults to FFmpeg in the runner loop
                        }

                        $script:State.BatchQueue += @{ 
                            Args            = $argsFull
                            SafeArgs        = $argsFull
                            HasCustomParams = $hasCustom
                            Retried         = $false
                            IsYtDlp         = $false
                            CustomTool      = $targetTool
                            OutputFile      = $outFile
                            ListBox         = $list
                            ListItem        = $inFile
                            InputFile       = $inFile
                            JobStart        = Get-Date
                        }
                    }
                    elseif ($tabIndex -eq 2) { 
                        $rawFmt = Get-CbVal $I_CFormat
                        $fmt = if ($rawFmt -match "ICO") { "ico" } else { $rawFmt.ToLower() }
                        $outFile = Get-UniqueFileName (Join-Path $outDir "$name.$fmt")
                        $argArray += @("-i", $inFile)

                        if ($I_CheckMeta.IsChecked) { $argArray += @("-map_metadata", "-1") }

                        if ($fmt -eq "ico") {
                            # Ensure icon is square 256x256 with transparent padding to prevent squashing
                            $argArray += @("-vf", "scale=256:256:force_original_aspect_ratio=decrease,pad=256:256:(ow-iw)/2:(oh-ih)/2:color=black@0")
                        } 
                        else {
                            $resCb = Get-CbVal $I_CRes
                            if ($resCb -match "1920") { $argArray += @("-vf", "scale=1920:-1") }
                            elseif ($resCb -match "1280") { $argArray += @("-vf", "scale=1280:-1") }
                            elseif ($resCb -match "800") { $argArray += @("-vf", "scale=800:-1") }
                        }

                        if ($fmt -match "jpg|jpeg|webp") {
                            $qualCb = Get-CbVal $I_CQual
                            if ($fmt -match "webp") {
                                if ($qualCb -match "Medium") { $argArray += @("-qscale", "50") }
                                elseif ($qualCb -match "Low") { $argArray += @("-qscale", "20") }
                                else { $argArray += @("-qscale", "80") }
                            }
                            else {
                                if ($qualCb -match "Medium") { $argArray += @("-q:v", "6") }
                                elseif ($qualCb -match "Low") { $argArray += @("-q:v", "12") }
                                else { $argArray += @("-q:v", "2") }
                            }
                        }

                        $argArray += $outFile
                        $script:State.BatchQueue += @{ Args = $argArray; SafeArgs = $argArray; HasCustomParams = $false; Retried = $false; IsYtDlp = $false; OutputFile = $outFile; ListBox = $list; ListItem = $inFile }
                    }
                }
            }

            # Fire off queue process locally via helper function if variables check out
            if ($script:State.BatchQueue.Count -gt 0) {
                $BtnRun.IsEnabled = $false
                $BtnUpdate.IsEnabled = $false
                $BtnCancel.IsEnabled = $true
                $BtnSkip.IsEnabled = $true
                $script:State.CurrentJobIndex = 0
                ProcessNextJob
            }
        })

    # Helper script to revert fields back to generic states if user requests full wipe
    $BtnReset.Add_Click({
            # Clear Queue Lists
            @($A_InList, $V_InList, $I_InList) | ForEach-Object { if ($null -ne $_) { $_.Items.Clear() } }

            # Reset Directories
            $defaultDir = if ($Config.DefaultOutDir -and (Test-Path $Config.DefaultOutDir)) { $Config.DefaultOutDir } else { "Select target folder..." }
            @($A_OutDir, $V_OutDir, $I_OutDir, $Y_OutDir) | ForEach-Object { if ($null -ne $_) { $_.Text = $defaultDir } }
    
            # Clear Text Inputs
            @($M_InVideo, $M_InAudio, $M_OutFile, $V_SubPath, $Y_CookiePath, $Y_PoToken, $S_VisAudio, $S_VisImg, $S_StabIn, $S_ScribeIn, $S_UpscaleIn) | ForEach-Object { 
                if ($null -ne $_) { $_.Clear() } 
            }
            
            # Reset Download Link
            if ($null -ne $Y_Link) { $Y_Link.Text = "https://" }

            # Reset Special Tab Selectors
            if ($null -ne $S_VisStyle) { $S_VisStyle.SelectedIndex = 0 }
            if ($null -ne $S_StabLevel) { $S_StabLevel.SelectedIndex = 1 }
            if ($null -ne $S_ScribeLang) { $S_ScribeLang.SelectedIndex = 0 }
            if ($null -ne $S_ScribeFormat) { $S_ScribeFormat.SelectedIndex = 0 }
            if ($null -ne $S_ScribeModel) { $S_ScribeModel.SelectedIndex = 1 }
            if ($null -ne $S_ScribeTask) { $S_ScribeTask.SelectedIndex = 0 }
            if ($null -ne $S_UpscaleModel) { $S_UpscaleModel.SelectedIndex = 0 }
            if ($null -ne $S_UpscaleScale) { $S_UpscaleScale.SelectedIndex = 2 }
            if ($null -ne $S_UpscaleOutDir) { $S_UpscaleOutDir.Text = "Select target folder..." }
            if ($null -ne $S_CheckBurn) { $S_CheckBurn.IsChecked = $false }
            if ($null -ne $V_UseHandbrake) { $V_UseHandbrake.IsChecked = $false }
    
            # Reset Audio/Video/Image Settings
            if ($null -ne $A_CFormat) { $A_CFormat.SelectedIndex = 0 }
            if ($null -ne $A_CQual) { $A_CQual.SelectedIndex = 2 }
            if ($null -ne $A_CChan) { $A_CChan.SelectedIndex = 0 }
            if ($null -ne $A_CheckExtract) { $A_CheckExtract.IsChecked = $false }
            if ($null -ne $A_CheckNorm) { $A_CheckNorm.IsChecked = $false }

            # --- UPDATED VIDEO DEFAULTS (Safe / Non-Destructive) ---
            if ($null -ne $V_CFormat) { $V_CFormat.SelectedIndex = 1 } # Container: MKV (Safest for copying all streams)
            if ($null -ne $V_CCodec) { $V_CCodec.SelectedIndex = 3 } # Video Codec: Copy
            if ($null -ne $V_CAudio) { $V_CAudio.SelectedIndex = 3 } # Audio Codec: Copy
            if ($null -ne $V_CSub) { $V_CSub.SelectedIndex = 1 } # Subtitles: Copy All
            if ($null -ne $V_CAudioTracks) { $V_CAudioTracks.SelectedIndex = 1 } # Tracks: Keep All
            # -------------------------------------------------------
            
            if ($null -ne $V_CRes) { $V_CRes.SelectedIndex = 0 }
            if ($null -ne $V_Preset) { $V_Preset.SelectedIndex = 0 }
            if ($null -ne $V_CFPS) { $V_CFPS.SelectedIndex = 0 }
            if ($null -ne $V_CVol) { $V_CVol.SelectedIndex = 0 }
            if ($null -ne $V_CSpeed) { $V_CSpeed.SelectedIndex = 0 }
            if ($null -ne $V_AudioDelay) { $V_AudioDelay.Text = "0.0" }
            if ($null -ne $V_SliderCRF) { $V_SliderCRF.Value = 23 }
            if ($null -ne $V_CheckTargetSize) { $V_CheckTargetSize.IsChecked = $false }
            if ($null -ne $V_OutFilename) { $V_OutFilename.Clear() }
            if ($null -ne $V_CheckSmartName) { $V_CheckSmartName.IsChecked = $false; $V_CheckSmartName.IsEnabled = $false }

            if ($null -ne $I_CFormat) { $I_CFormat.SelectedIndex = 0 }
            if ($null -ne $I_CRes) { $I_CRes.SelectedIndex = 0 }
            if ($null -ne $I_CheckMeta) { $I_CheckMeta.IsChecked = $false }
    
            # Reset Download Settings
            if ($null -ne $Y_Type) { $Y_Type.SelectedIndex = 0 }
            if ($null -ne $Y_Res) { $Y_Res.SelectedIndex = 0 }
            if ($null -ne $Y_VFormat) { $Y_VFormat.SelectedIndex = 0 }
            if ($null -ne $Y_AFormat) { $Y_AFormat.SelectedIndex = 0 }
            if ($null -ne $Y_CookieBrowser) { $Y_CookieBrowser.SelectedIndex = 0 }
            if ($null -ne $Y_CheckMeta) { $Y_CheckMeta.IsChecked = $true }
            if ($null -ne $Y_CheckSubs) { $Y_CheckSubs.IsChecked = $false }
            if ($null -ne $Y_CheckSponsor) { $Y_CheckSponsor.IsChecked = $false }
            if ($null -ne $Y_CheckCookie) { $Y_CheckCookie.IsChecked = $false }
            if ($null -ne $Y_CheckAutoPoToken) { $Y_CheckAutoPoToken.IsChecked = $false }
            if ($null -ne $Y_PoToken) { $Y_PoToken.IsEnabled = $true; $Y_PoToken.Opacity = 1.0 }

            # Reset Progress & UI State
            if ($null -ne $V_PreviewStack) { $V_PreviewStack.Children.Clear() }
            if ($null -ne $V_PreviewScroll) { $V_PreviewScroll.Visibility = "Collapsed" }
            if ($null -ne $LogBox) { $LogBox.Clear() }
            if ($null -ne $PBar) { $PBar.Value = 0 }
            if ($null -ne $TaskbarProgress) { $TaskbarProgress.ProgressValue = 0.0; $TaskbarProgress.ProgressState = "None" }
            if ($null -ne $TxtETA) { $TxtETA.Text = "ETA: --:--" }
            if ($null -ne $StatusText) { $StatusText.Text = "Ready."; $StatusText.Foreground = $window.Resources["TextBrush"] }
            if ($null -ne $BtnShow) { $BtnShow.Visibility = "Collapsed" }

            $LogBox.AppendText("[RESET] All fields and queues have been cleared. Attention: Log will be cleared in 1 second.`r`n")
            
            # Use a script-scoped DispatcherTimer so it survives the button click ending
            if ($null -ne $script:resetTimer) { $script:resetTimer.Stop() }
            $script:resetTimer = New-Object System.Windows.Threading.DispatcherTimer
            $script:resetTimer.Interval = [TimeSpan]::FromSeconds(1) # Delay to ensure any pending UI updates are flushed before clearing the log
            $script:resetTimer.Add_Tick({
                $script:resetTimer.Stop() # Now it securely knows what to stop
                $LogBox.Clear()
            })
            $script:resetTimer.Start()
        })

    # Terminates queue safely and clears uncompleted outputs to avoid data leaks
    $BtnCancel.Add_Click({
            try {
                $BtnCancel.IsEnabled = $false
                $BtnSkip.IsEnabled = $false
                $LogBox.AppendText("`r`n[CANCEL] User requested to cancel the queue. Aborting...`r`n")
                if ($CbAutoScrollLog.IsChecked) { $LogBox.ScrollToEnd() }

                $currentJob = $null
                if ($script:State.CurrentJobIndex -lt $script:State.BatchQueue.Count) {
                    # Preserve the active job in the queue so the exit handler cleans it up safely
                    $currentJob = $script:State.BatchQueue[$script:State.CurrentJobIndex]
                    $script:State.BatchQueue = @($currentJob)
                    $script:State.CurrentJobIndex = 0
                }
                else {
                    $script:State.BatchQueue = @()
                }

                # Send kill signal instantly
                if ($script:State.p) { 
                    Stop-ProcessTree $script:State.p 
                }
                else {
                    # Force UI reset if no process is currently bound
                    $script:State.CurrentJobIndex = 1
                    [void][System.Windows.Threading.Dispatcher]::CurrentDispatcher.InvokeAsync({ ProcessNextJob })
                }

                # Run file cleanup asynchronously via hidden cmd to prevent Runspace context loss
                if ($currentJob -and $currentJob.IsYtDlp -and (Test-Path -LiteralPath $currentJob.OutputDir)) {
                    $outDir = $currentJob.OutputDir
                    $cmdArgs = "/c timeout /t 2 /nobreak >nul & del /q /f `"$outDir\*.part`" `"$outDir\*.ytdl`" `"$outDir\*.temp`""
                    [void](Start-Process cmd.exe -ArgumentList $cmdArgs -WindowStyle Hidden)
                }

                if (Test-Path $QueueFile) { Remove-Item $QueueFile -Force -ErrorAction SilentlyContinue }
            }
            catch {
                $LogBox.AppendText("`r`n[ERROR] Failed to cancel gracefully: $($_.Exception.Message)`r`n")
            }
        })

    # Aborts current process but allows the loop to trigger the next file in array
    $BtnSkip.Add_Click({
            if ($script:State.p) { 
                Stop-ProcessTree $script:State.p 
                $LogBox.AppendText("`r`n[SKIP] User skipped current process.`r`n")
            }
        })

    # Ensure application cleanly terminates child processes, saves queue, and wipes temp files
    $window.Add_Closing({ 
            Stop-ProcessTree $script:State.p
            Save-Queue 
        
            # Clean up temporary logs and thumbnails to prevent disk bloat
            @("mcp_live.log", "mcp_live_err.log", "yt_info.json", "yt_info_err.log", "whisper_update.log") | ForEach-Object {
                $f = Join-Path $env:TEMP $_
                if (Test-Path $f) { Remove-Item $f -Force -ErrorAction SilentlyContinue }
            }
        
            # Sweep all dynamically generated preview thumbnails
            Get-ChildItem -Path $env:TEMP -Filter "thumb_*.jpg" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        
            # Kill System Tray icon to prevent "ghost" icons from lingering
            if ($script:TrayIcon) { 
                $script:TrayIcon.Visible = $false
                $script:TrayIcon.Dispose() 
            }
        })
    [void]$window.ShowDialog()
}
catch {
    Write-CrashLog "Fatal Application Crash: $($_.Exception.Message)`r`n$($_.ScriptStackTrace)"
    [void][System.Windows.MessageBox]::Show("Crash! See crash.log", "Error", 0, 16)
}