<#
.SYNOPSIS
    Automates the download, installation, and activation of Office 2021 ProPlus
.DESCRIPTION
    This script performs the following actions:
    1. Downloads the Office 2021 ProPlus ISO image
    2. Mounts the ISO image
    3. Installs Office 2021 ProPlus
    4. Activates using KMS server
.NOTES
    File Name      : main.ps1
    EXE-Name       : InstallOffice2021.exe
    Requires admin privileges
    Author: https://github.com/4G0NYY
#>

#Requires -RunAsAdministrator

# Configuration
$isoUrl = "https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/de-DE/ProPlus2021Retail.img"
$isoPath = "$env:TEMP\ProPlus2021Retail.iso"
$kmsKey = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
$kmsServer = "107.175.77.7"
$kmsPort = "1688"

function Write-Status {
    param([string]$message)
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $message" -ForegroundColor Cyan
}

try {
    # Step 1: Download the ISO image
    Write-Status "Downloading Office 2021 ProPlus ISO..."
    $ProgressPreference = 'SilentlyContinue' # Speeds up download
    Invoke-WebRequest -Uri $isoUrl -OutFile $isoPath -UseBasicParsing
    $ProgressPreference = 'Continue'
    
    if (-not (Test-Path $isoPath)) {
        throw "Failed to download ISO file"
    }
    Write-Status "Download completed successfully."

    # Step 2: Mount the ISO image
    Write-Status "Mounting ISO image..."
    $mountResult = Mount-DiskImage -ImagePath $isoPath -PassThru
    $driveLetter = ($mountResult | Get-Volume).DriveLetter + ":"
    Write-Status "ISO mounted to $driveLetter"

    # Step 3: Install Office
    Write-Status "Starting Office installation..."
    $setupPath = Join-Path -Path $driveLetter -ChildPath "Office\Setup64.exe"
    $process = Start-Process -FilePath $setupPath -Wait -PassThru
    
    if ($process.ExitCode -ne 0) {
        throw "Office installation failed with exit code $($process.ExitCode)"
    }
    Write-Status "Office installation completed successfully."

    # Step 4: Activate Office
    Write-Status "Starting activation process..."

    # Find the correct Office path
    $officePaths = @(
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16",
        "$env:ProgramFiles\Microsoft Office\Office16"
    )

    $officePath = $null
    foreach ($path in $officePaths) {
        if (Test-Path $path) {
            $officePath = $path
            break
        }
    }

    if (-not $officePath) {
        throw "Could not find Office installation directory"
    }

    Write-Status "Found Office installation at $officePath"

    # Change directory to Office16
    Push-Location $officePath

    try {
        # Install licenses
        Write-Status "Installing KMS licenses..."
        $licenseFiles = Get-ChildItem "..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms" | Select-Object -ExpandProperty Name
        
        foreach ($license in $licenseFiles) {
            $licensePath = "..\root\Licenses16\$license"
            cscript ospp.vbs /inslic:"$licensePath" | Out-Null
        }

        # Configure KMS
        Write-Status "Configuring KMS settings..."
        cscript ospp.vbs /setprt:$kmsPort | Out-Null
        cscript ospp.vbs /unpkey:6F7TH | Out-Null
        cscript ospp.vbs /inpkey:$kmsKey | Out-Null
        cscript ospp.vbs /sethst:$kmsServer | Out-Null

        # Activate
        Write-Status "Activating Office..."
        $activationResult = cscript ospp.vbs /act
        
        if ($activationResult -match "Product activation successful") {
            Write-Host "Activation successful!" -ForegroundColor Green
        } else {
            Write-Host "Activation may have failed. Check output for details." -ForegroundColor Yellow
            $activationResult
        }
    }
    finally {
        Pop-Location
    }
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}
finally {
    # Cleanup - unmount ISO if it's still mounted
    if ($driveLetter -and (Test-Path $driveLetter)) {
        Write-Status "Unmounting ISO..."
        Dismount-DiskImage -ImagePath $isoPath | Out-Null
    }
    
    # Optional: Remove downloaded ISO
    if (Test-Path $isoPath) { Remove-Item $isoPath -Force }
}

Write-Status "Script completed."

try {
    # Only prompt if script was not run from existing PowerShell session
    if ([Environment]::UserInteractive) {
        Write-Host "Operation completed. Press any key to continue..." -ForegroundColor Green
        [System.Console]::ReadKey($true) | Out-Null
    }
}
catch {
    # Fallback if ReadKey fails
    Write-Host "Operation completed. Window will close automatically in 30 seconds..." -ForegroundColor Green
    Start-Sleep -Seconds 30
}