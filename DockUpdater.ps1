<#
 .Author 
  Dan Felman | HP Inc | @dan_felman 
  Gary Blok | HP Inc | @gwblok | GARYTOWN.COM
 
 .Synopsis
  HP Dock Updater Script

 .Description
  Script will to use HPCMSL to get the latest URL for the firmware.
  If you want to Bypass HPCMSL manually, because you want manually control which version is being deployed, use the parameter "BypassHPCMSL" in the command line AND
  - make sure you set the $URL to the version of the firmware you want... 
    current firmware Softpaq links hardcoded in script latest as of release of script

 .Requirements
  PowerShell must have access to the interent to download the Firmware

 .Parameter UIExperience < NonInteractive | Silent >
  Sets the UI experience for the end user.  
  Options: 
    NonInteractive - Shows the Dialog box with progress to the end user - (default) 
    Silent - completely hidden to the end user

 .Parameter BypassHPCMSL [Switch]
  IF HPCMSL is available, allows you to bypass using it so you can manually set the URL for the Firmware Version you wish to deploy - Full Control
  Softpaq url's presets as of date of script version release

 .Parameter Update [Switch]
  by default, the script does a Check of the installed firmware. The -Update switch enables the script to execute
  a firmware update if one is needed

 .Parameter Hoteling [Switch]
  when docks are used in hoteling desks, this option asks the user to allow the dock firmware update

 .ChangeLog
  2023.04.07 - First GitHub Release as DockUpdater.ps1
  2023.06.09 - added -Hoteling option to ask user permission to update dock

 .Notes
  This will create a transcription log IFF the dock is attached and it starts the process to test firmware.  If no dock is detected, no logging is created.
  Logging created by this line: Start-Transcript -Path "$OutFilePath\$SPNumber.txt" - which should be: "C:\swsetup\dockfirmware\sp144502.txt"  		

 .Example
   # Update the Dock's firmware to the latest version HPCMSL (if installed) will find completely silent to end user
   DockUpdater.ps1 -Update -UIExperience Silent

 .Example
   # Check the Dock's firmware version without CMSL (SOftpaqs URL hardcoded), or with CMSL
   DockUpdater.ps1 -BypassHPCMSL # no CMSL on device or don't want to use it, use hardcoded links
   DockUpdater.ps1               # use CMSL to find latest f/w Softpaq of attached dock
   
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)][ValidateSet('NonInteractive', 'Silent')][String]$UIExperience,
    [switch]$BypassHPCMSL,
    [switch]$Update,
    [switch]$Hoteling,
    [switch]$DebugOut
) # param

$ScriptVersion = 'DockUpdater.ps1: 01.00.12 June 29, 2023'

# check for CMSL by attempting to run one command
Try {
    $HPDeviceDetails = Get-HPDeviceDetails -ErrorAction SilentlyContinue }
catch {
    $BypassHPCMSL = $true }

$AdminRights = ([Security.Principal.WindowsPrincipal] `
            [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if ( $DebugOut ) { Write-Host "--Admin rights:"$AdminRights }

if ( $false -eq $AdminRights ) {
    Write-Host "--Elevated Admin rights are required"
    return -1
}
function Get-HPDockInfo {
    [CmdletBinding()]
    param($pPnpSignedDrivers)

    # **** Hardcode URLs in case of no CMSL used (so can't query for latest update) ****
    # ---- current supported accessories with script ----
    $Url_TBG2 = 'ftp.hp.com/pub/softpaq/sp143501-144000/sp143977.exe'   #  (as of apr 6, 2023)
    $Url_TBG4 = 'ftp.hp.com/pub/softpaq/sp143501-144000/sp143669.exe'   #  (as of apr 6, 2023)
    $Url_UniG2 = 'ftp.hp.com/pub/softpaq/sp146001-146500/sp146291.exe'  #  (as of Jun 1, 2023)
    $Url_UsbG5 = 'ftp.hp.com/pub/softpaq/sp146001-146500/sp146273.exe'  #  (as of Jun 1, 2023)
    $Url_UsbG4 = 'ftp.hp.com/pub/softpaq/sp88501-89000/sp88999.exe'     #  (as of Jun 1, 2023)
    $Url_EssG5 = 'ftp.hp.com/pub/softpaq/sp144501-145000/sp144502.exe'  #  (as of apr 6, 2023)
    $Url_E24dG4 = 'ftp.hp.com/pub/softpaq/sp145501-146000/sp145577.exe'  #  (as of Jun 1, 2023)
    $Url_E27dG4 = 'ftp.hp.com/pub/softpaq/sp145501-146000/sp145576.exe'  #  (as of Jun 1, 2023)
    $Url_Z40cG3 = 'ftp.hp.com/pub/softpaq/sp143501-144000/sp143685.exe'  #  (as of Jun 1, 2023)
    $Url_EOPOCI = 'ftp.hp.com/pub/softpaq/sp144501-145000/sp144902.exe' # (as of June 15, 2023)

    #######################################################################################
    $Dock_Attached = 0      # default: no dock found
    $Dock_ProductName = $null
    $Dock_Url = $null   
    # Find out if a Dock is connected - assume a single dock, so stop at first find
    foreach ( $iDriver in $pPnpSignedDrivers ) {
        $f_InstalledDeviceID = "$($iDriver.DeviceID)"   # analyzing current device
        if ( $f_InstalledDeviceID -match "HID\\VID_03F0" ) {
            switch -Wildcard ( $f_InstalledDeviceID ) {
                '*PID_0488*' { $Dock_Attached = 1 ; $Dock_ProductName = 'HP Thunderbolt Dock G4' ; $Dock_Url = $Url_TBG4 }
                '*PID_0667*' { $Dock_Attached = 2 ; $Dock_ProductName = 'HP Thunderbolt Dock G2' ; $Dock_Url = $Url_TBG2 }
                '*PID_484A*' { $Dock_Attached = 3 ; $Dock_ProductName = 'HP USB-C Dock G4' ; $Dock_Url = $Url_UsbG4 }
                '*PID_046B*' { $Dock_Attached = 4 ; $Dock_ProductName = 'HP USB-C Dock G5' ; $Dock_Url = $Url_UsbG5 }
                #'*PID_600A*' { $Dock_Attached = 5 ; $Dock_ProductName = 'HP USB-C Universal Dock' } # "USB\\VID_17E9"
                '*PID_0A6B*' { $Dock_Attached = 6 ; $Dock_ProductName = 'HP USB-C Universal Dock G2' ; $Dock_Url = $Url_UniG2 }
                '*PID_056D*' { $Dock_Attached = 7 ; $Dock_ProductName = 'HP E24d G4 FHD Docking Monitor' ; $Dock_Url = $Url_E24dG4 }
                '*PID_016E*' { $Dock_Attached = 8 ; $Dock_ProductName = 'HP E27d G4 QHD Docking Monitor' ; $Dock_Url = $Url_E27dG4 }
                '*PID_379D*' { $Dock_Attached = 9 ; $Dock_ProductName = 'HP USB-C G5 Essential Dock' ; $Dock_Url =  $Url_EssG5 }
                '*PID_0F84*' { $Dock_Attached = 10 ; $Dock_ProductName = 'HP Z40c G3 WUHD Curved Display' ; $Dock_Url =  $Url_Z40cG3 }
                '*PID_0380*' { $Dock_Attached = 11 ; $Dock_ProdcutName = 'HP Engage One Pro Stand Hub' ; $Dock_Url = $Url_EOPOCI }
                '*PID_0381*' { $Dock_Attached = 12 ; $Dock_ProdcutName = 'HP Engage One Pro VESA Hub' ; $Dock_Url = $Url_EOPOCI }
                '*PID_0480*' { $Dock_Attached = 13 ; $Dock_ProdcutName = 'HP Engage One Pro Advanced Fanless Hub' ; $Dock_Url = $Url_EOPOCI }
            } # switch -Wildcard ( $f_InstalledDeviceID )
        } # if ( $f_InstalledDeviceID -match "VID_03F0")
        if ( $Dock_Attached -gt 0 ) { break }
    } # foreach ( $iDriver in $gh_PnpSignedDrivers )
    #######################################################################################

    return @(
        @{Dock_Attached = $Dock_Attached ;  $Dock_ProductName = $Dock_ProductName  ;  Dock_Url = $Dock_Url}
    )
} # function Get-HPDockInfo

function Get-PackageVersion {
    [CmdletBinding()]param( $pDocknum, $pCheckFile ) # param

    $TestInfo = Get-Content -Path $pCheckFile
    if ( $pDocknum -eq 9 ) {       
        [String]$InstalledVersion = $TestInfo | Select-String -Pattern 'installed' -SimpleMatch
        $InstalledVersion = $InstalledVersion.Split(":") | Select-Object -Last 1            
    } else {
        [String]$InstalledVersion = $TestInfo | Select-String -Pattern 'Package' -SimpleMatch
        $InstalledVersion = $InstalledVersion.Split(":") | Select-Object -Last 1
    }
    return $InstalledVersion
} # function Get-PackageVersion

Function Ask_YesNo {
    [CmdletBinding()]param( $pTitle, $pContent )
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Windows.Forms
    $lResponse = [System.Windows.MessageBox]::Show($pContent,$pTitle,'YesNo')
    return $lResponse
}

#########################################################################################

#'-- Reading signed drivers list - use to scan for attached HP docks'
$PnpSignedDrivers = Get-CimInstance win32_PnpSignedDriver 

$Dock = Get-HPDockInfo $PnpSignedDrivers
if ( $DebugOut ) { Write-Host "--Dock detected:"$Dock.Dock_ProductName }
$HPFIrmwareUpdateReturnValues = @(
        @{Code = "0" ;  Message = "Success"}
        @{Code = "101" ;  Message = "Install or stage failed. One or more firmware failed to install."}
        @{Code = "102" ;  Message = "Configuration file failed to be loaded.This may be because it could not be found or that it was not properly formatted."}
        @{Code = "103" ;  Message = "One or more firmware packages specified in the configuration file could not be loaded."}
        @{Code = "104" ;  Message = "No devices could be communicated with.This could be because necessary drivers are missing to detect the device."}
        @{Code = "105" ;  Message = "Out - of - date firmware detected when running with 'check' flag."}
        @{Code = "106" ;  Message = "An instance of HP Firmware Installer is already running"}
        @{Code = "107" ;  Message = "Device not connected.This could be because PID or VID is not detected."}
        @{Code = "108" ;  Message = "Force option disabled.Firmware downgrade or re - flash not possible on this device."}
        @{Code = "109" ;  Message = "The host is not able to update firmware"}
    )
# loop for up to 10 secs in case we just powered-on, or Dock detection takes a bit of time
[int]$Counter = 0
[int]$StepAmt = 20
if ( $Dock.Dock_Attached -eq 0 ) {
    Write-Host "Waiting for Dock to be fully attached up to $WaitTimer seconds" -ForegroundColor Green
    do {
        Write-Host " Waited $Counter Seconds Total.. waiting additional $StepAmt" -ForegroundColor Gray
        $counter += $StepAmt
        Start-Sleep -Seconds $StepAmt
        $Dock = Get-HPDockInfo $PnpSignedDrivers
        if ( $counter -eq $WaitTimer ) {
            Write-Host "Waited $WaitTimer Seconds, no dock found yet..." -ForegroundColor Red
        }
    }
    while ( ($counter -lt $WaitTimer) -and ($Dock.Dock_Attached -eq "0") )
} # if ( $Dock.Dock_Attached -eq "0" )

# Tried, but couldn't find a dock (or supported monitor) attached
if ( $Dock.Dock_Attached -eq 0 ) {
    Write-Host " No dock attached" -ForegroundColor Green
} else {
    # NOW, let's get to work on the dock, if found
    if ( $BypassHPCMSL ) {
        $URL = $Dock.Dock_Url
        if ( $DebugOut ) { Write-Host "--Dock detected Url - hardcoded:"$URL }
    } else {
        $URL = (Get-SoftpaqList -Category Dock | Where-Object { $_.Name -match $dock.Dock_ProductName -and ($_.Name -match 'firmware') }).Url
        if ( $DebugOut ) { Write-Host "--Dock detected Url - via CMSL:"$URL }
    } # else if ( $BypassHPCMSL )

    $SPEXE = ($URL.Split("/") | Select-Object -Last 1)
    $SPNumber = ($URL.Split("/") | Select-Object -Last 1).replace(".exe","")
    if ( $DebugOut ) { Write-Host "--Dock detected firmware Softpaq:"$SPEXE }

    # Create Required Folders
    $OutFilePath = "$env:SystemDrive\swsetup\dockfirmware"
    $ExtractPath = "$OutFilePath\$SPNumber"
    
    Start-Transcript -Path "$OutFilePath\$SPNumber.txt"
    write-Host $ScriptVersion
    write-Host "  Running script with CMSL ="(-not $BypassHPCMSL) -ForegroundColor Gray
    if ( $Update ) {
        write-Host "  Executing a dock firmware update" -ForegroundColor Cyan
    } else {
        write-Host "  Executing a check of the dock firmware version. Use -Update to update the firmware" -ForegroundColor Cyan
    }
    try {
        [void][System.IO.Directory]::CreateDirectory($OutFilePath)
        [void][System.IO.Directory]::CreateDirectory($ExtractPath)
    } catch { 
        if ( $DebugOut ) { Write-Host "--Error creating folder"$ExtractPath }
        throw 
    }
    # Download Softpaq EXE
    if ( !(Test-Path "$OutFilePath\$SPEXE") ) { 
        try {
            $Error.Clear()
            Write-Host "  Starting Download of $URL to $OutFilePath\$SPEXE" -ForegroundColor Magenta
            Invoke-WebRequest -UseBasicParsing -Uri $URL -OutFile "$OutFilePath\$SPEXE"
        } catch {
            Write-Host "!!!Failed to download Softpaq!!!" -ForegroundColor red
            Stop-Transcript
            return -1
        }
    } else {
        Write-Host "  Softpaq already downloaded to $OutFilePath\$SPEXE" -ForegroundColor Gray
    }
    # Extract Softpaq EXE
    if ( Test-Path "$OutFilePath\$SPEXE" ) { 
        if (!(Test-Path "$OutFilePath\$SPNumber\HPFirmwareInstaller.exe")){
            Write-Host "  Extracting to $ExtractPath" -ForegroundColor Magenta
            if ( $AdminRights ) {
                $Extract = Start-Process -FilePath $OutFilePath\$SPEXE -ArgumentList "/s /e /f $ExtractPath" -NoNewWindow -PassThru -Wait
                write-Host " Softpaq extract returned: $($Extract.ExitCode)"
            } else {
                Write-Host "  Admin rights required to extract to $ExtractPath" -ForegroundColor Red
                Stop-Transcript
                return -1
            }           
        } else {
            Write-Host "  Softpaq already Extracted to $ExtractPath" -ForegroundColor Gray
        }
    } else {
        Write-Host "  Failed to find $OutFilePath\$SPEXE" -ForegroundColor Red
        Stop-Transcript
        return -1
    }

    # Get package version from downloaded Softpaq configuration file
    $ConfigFile = "$OutFilePath\$SPNumber\HPFIConfig.xml"       # All docks except Essential
    $ConfigFileEssential = "$OutFilePath\$SPNumber\config.ini"  # Essential dock
    if ( Test-Path $ConfigFile ) {
        $xmlConfigContent = [xml](Get-Content -Path $ConfigFile)
        $PackageVersion = $xmlConfigContent.SelectNodes("FirmwareCollectionPackage/PackageVersion").'#Text'
        $ModelName = $xmlConfigContent.SelectNodes("FirmwareCollectionPackage/Name").'#Text'
        Write-Host "  Extracted Softpaq Info file: $ConfigFile" -ForegroundColor Cyan
    } elseif ( Test-Path $ConfigFileEssential ) {    
        # Deal with Essential Dock especially  - different config file format and content
        $ConfigInfo = Get-Content -Path $ConfigFileEssential
        [String]$PackageVersion = $ConfigInfo | Select-String -Pattern 'PackageVersion' -CaseSensitive -SimpleMatch
        $PackageVersion = $PackageVersion.Split("=") | Select-Object -Last 1
        [String]$ModelName = $ConfigInfo | Select-String -Pattern 'ModelName' -CaseSensitive -SimpleMatch
        $ModelName = $ModelName.Split("=") | Select-Object -Last 1
        Write-Host "  Extracted Softpaq Info file: $ConfigFileEssential" -ForegroundColor Cyan
    } # elseif ( Test-Path $ConfigFileEssential )
    
    Write-Host "  Softpaq for Device: $ModelName" -ForegroundColor Gray
    Write-Host "  Softpaq Version: $PackageVersion" -ForegroundColor Gray

    if (Test-Path "$OutFilePath\$SPNumber\HPFirmwareInstaller.exe") { # Run version check - Check if Update Required
        Write-Host " Running HP Firmware Check... please, wait" -ForegroundColor Magenta
        Try {
            $Error.Clear()
            $HPFirmwareTest = Start-Process -FilePath "$OutFilePath\$SPNumber\HPFirmwareInstaller.exe" -ArgumentList "-C" -PassThru -Wait -NoNewWindow
        } Catch {
            write-Host $error[0].exception
            Stop-Transcript
            return -1
        }
        if ( $Dock.Dock_Attached -eq 9 ) {  # Essential dock found
            $VersionFile = "$OutFilePath\$SPNumber\HPFI_Version_Check.txt"
        } else {
            $VersionFile = "$PSScriptRoot\HPFI_Version_Check.txt"
        }
        switch ( $HPFirmwareTest.ExitCode ) {
            0   { 
                     Write-Host " Firmware is up to date" -ForegroundColor Green
                    $InstalledVersion = Get-PackageVersion $Dock.Dock_Attached $VersionFile
                    Write-Host " Installed Version: $InstalledVersion" -ForegroundColor Green
                } # 0
            105 {
                    if ( !($UIExperience) ) { $UIExperience = 'NonInteractive' }
                    $Mode = switch ( $UIExperience ) {
                        "NonInteractive" {"-ni"}
                        "Silent" {"-s"}
                    }
                    Write-Host " Update Required" -ForegroundColor Yellow
                    $InstalledVersion = Get-PackageVersion $Dock.Dock_Attached $VersionFile
                    Write-Host " Installed Version: $InstalledVersion" -ForegroundColor Yellow

                    if ( $Update ) {
                        $YN = 'Yes'             # default when -Hoteling option not used
                        if ( $Hoteling ) {      # let's ask user to run update now
                            $YN = Ask_YesNo 'Dock firmware Version Check' 'The connected dock requires an important firmware update. Update now?' 
                        }
                        if ( $YN -eq 'Yes' ) {
                            Write-Host " Starting Dock Firmware Update" -ForegroundColor Magenta
                            $HPFirmwareUpdate = Start-Process -FilePath "$OutFilePath\$SPNumber\HPFirmwareInstaller.exe" -ArgumentList "$mode" -PassThru -Wait -NoNewWindow
                            $ExitInfo = $HPFIrmwareUpdateReturnValues | Where-Object { $_.Code -eq $HPFirmwareUpdate.ExitCode }
                            if ($ExitInfo.Code -eq "0"){
                                Write-Host " Update Successful!" -ForegroundColor Green
                            } else {
                                Write-Host " Update Failed!" -ForegroundColor Red
                                Write-Host " Exit Code: $($ExitInfo.Code): $($ExitInfo.Message)" -ForegroundColor Gray
                            }                            
                        } else {
                            Write-Host " Update Prevented by User" -ForegroundColor Yellow
                        }
                    } # if ( $Update )
                } # 105
        }
    } # if (Test-Path "$OutFilePath\$SPNumber\HPFirmwareInstaller.exe")

    Stop-Transcript 

} # if ( $Dock.Dock_Attached -gt 0 )
<#
    $Dock.Dock_Attached results returned:

    0 - NO HP dock attached
    1 - 'HP Thunderbolt Dock G4' - VID_03F0&PID_0488
    2 - 'HP Thunderbolt Dock G2'  - VID_03F0&PID_0667    
    3 - 'HP USB-C Dock G4'  - VID_03F0&PID_484A
    4 - 'HP USB-C Dock G5'  - VID_03F0&PID_046B
    6 - 'HP USB-C Universal Dock G2' - VID_03F0&PID_0A6B
    7 - 'HP E24d G4 FHD Docking Monitor'- VID_03F0&PID_056D
    8 - 'HP E27d G4 QHD Docking Monitor'- VID_03F0&PID_016E
    9 - 'HP USB-C G5 Essential Dock' - VID_03F0&PID_379D
    10 - 'HP Z40c G3 WUHD Curved Display'
    11 - 'HP Engage One Pro Stand Hub' - VID_03F0&PID_0380
    12 - 'HP Engage One Pro VESA Hub' - VID_03F0&PID_0381
    13 - 'HP Engage One Pro Advanced Fanless Hub' - VID_03F0&PID_0480
#>
return $Dock.Dock_Attached
