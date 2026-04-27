#Clear-Host
# Check if the PS2EXE module is available, if not, asks the user to automatically install it from the PowerShell Gallery
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "The PS2EXE module is not installed. Would you like to install it now? (Y/N)" -ForegroundColor Yellow
    $response = Read-Host
    if ($response -eq 'Y' -or $response -eq 'y') {
        try {
            Install-Module -Name ps2exe -Scope CurrentUser -Force
            Write-Host "PS2EXE module installed successfully." -ForegroundColor Green
        } catch {
            Write-Host "Failed to install PS2EXE module. Please install it manually from the PowerShell Gallery." -ForegroundColor Red
            return
        }
    } else {
        Write-Host "PS2EXE module is required to compile the script. Exiting." -ForegroundColor Red
        return
    }
}
# check again if the module is now available after the installation attempt. if yes, continue, if no, exit with error message
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "PS2EXE module is still not available. Please ensure it is installed correctly and try again." -ForegroundColor Red
    return  
}


# Dynamically set the path to the current folder where this script is running
$currentDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($currentDir)) { $currentDir = Get-Location }

# Define relative paths based on the current directory
$script_path = Join-Path $currentDir "MediaConverterPro.ps1"
$output_path = Join-Path $currentDir "MediaConverterPro.exe"
$icon_path = Join-Path $currentDir "app.ico" # Uncomment and add your icon file if needed

Write-Host "--- Media Converter Pro: EXE Compiler ---" -ForegroundColor Cyan
Write-Host "Target Script: $script_path"
Write-Host "Output EXE:    $output_path"
Write-Host "----------------------------------------"

# Check if the source script exists before starting
if (-not (Test-Path $script_path)) {
    Write-Host "Error: Could not find 'MediaConverterPro.ps1' in the current folder!" -ForegroundColor Red
    return
}

Write-Host "Compiling... please wait." -ForegroundColor Yellow

# Convert PS1 to EXE using portable parameters
# -noConsole: Hides the background CMD window when the EXE starts
# -sta: Required for WPF/GUI applications
# -x64: Speeds up the launch time by targeting 64-bit architecture (optional, but recommended for modern systems)
if (Test-Path $icon_path) {
    Invoke-Ps2Exe -inputFile $script_path -outputFile $output_path -sta -x64 -icon $icon_path | Out-Null
} else {
    Invoke-Ps2Exe -inputFile $script_path -outputFile $output_path -sta -x64 | Out-Null
}

#if you dont want ico: Invoke-Ps2Exe -inputFile $script_path -outputFile $output_path -sta | Out-Null 
# no -noConsole as a workaround for potential false positives in antivirus software, but can be added if you want a cleaner user experience

# Final Verification
if (Test-Path $output_path) {
    Write-Host "SUCCESS: EXE created successfully in the current folder!" -ForegroundColor Green
    Write-Host "Location: $output_path" -ForegroundColor Gray
} else {
    Write-Host "FAILURE: The compilation failed. Ensure the PS2EXE module is installed." -ForegroundColor Red
}