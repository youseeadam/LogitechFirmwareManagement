<#
 
.SYNOPSIS
    Author: Adam Berns
    Date: 12/13/2019

    This is supplied as-is. This is not an offically supported tool from Logitech.

    This script is designed to inventory and update firmware of common Logitech conference room hardware. It has two modes:

    Inventory: Searches for Logitech Hardware and writes to a series of JSON files what is connected
    Update Firwmare: Will use the same logic, however it will not inventory, it will just update firmware.

    These log files can then be imported into Azure Log Anaylytcis Custom Logs
    
.DESCRIPTION
    Create the folder structure outlined below
    If you are not using any of the devices listed below, you can just not download or create the folder directory. The script will stil search for it.
    Place this script in that folder

    If you don't use a product then don't put the file into the folder and the system will not search for it.

    For SmartDock
         a. Download the firmware updater from Logitech Support website
         b. Place the SmartDockUpdate.exe in the Root
         c. Create a sub folder called ait, in that folder put the AitUVCExtApi.dll, do not include any other files
         d. Create a sub folder at a top level called SmartDockUpdater and put the non extracted downloaded file there for doing updates

    For MeetUp
        a. Download the latest firmware version
        b. extract the firmware, there will be a file called FWUpdateMeetUp.exe place that file in the top level
        c. Create a sub folder called MeetUpUpdater and put the downloaded, non extracted file there for doing udpates

    For RallySystem (Different than the standalone Rally Camera)
        a. Download the latest firmware version
        b. extract the firmware, there will be a file called FWUpdateMeetUp.exe place that file in the top level
        c. Create a sub folder called MeetUpUpdater and put the downloaded, non extracted file there for doing udpates 

    For RallyCamera
        a. Download the latest firmware version
        b. extract the firmware, there will be a file called FWUpdateRallyCamera.exe place that file in the top level
        c. Create a sub folder called RallyCameraUpdater and put the downloaded, non extracted file there for doing udpates 
    
    For Group
        a. Download the latest firmware version
        b. Extract the downloaded file and place all files in the sub folder LogiGroupUpdater
    
        Your folder structure should look like this. The actual version number in the Updater folders does not have to match, as long as the name starts with everything before the number.
        Group Device is different than others. Extract the
    
        [folder]Logitech Updater Files
        Get-LogitechFirmware.ps1
        SmartDockUpdate.exe 
        FWUpdateMeetUp.exe
        FWUpdateRally.exe
        FWUpdateRallyCamera.exe
        [folder]ait
            AitUVCExtApi.dll
        [folder]SmartDockUpdater
            SmartDockUpdateInstall_1.2.31.48.exe
        [folder]MeetUpUpdater
            FWUpdateMeetup_1.10.60.exe
        [folder]RallySystemUpdater
            FWUpdateRally_1.4.28.exe
        [folder]RallyCameraUpdater
            FWUpdateRallyCamera_1.4.17.exe
        [folder]LogiGroupUpdater
            [folder]$PLUGINSDIR
            devcon*.exe
            FWUpdateLauncher.exe
            FWUpdateLogiGroup.exe
            icon.ico
            uninstall.exe
.OUTPUTS

    This will generate up to 4 logs for each system as an example for Rally
    %windir%\Temp\RallySystem.log : This the log file of what was done. It is overwritten each time the script is executed
    %windir%\Temp\LogitechFirmware-RallySystem*.Json: This is the inventory output
    %windir%\Temp\LogitechFirmwareUpdate-RallySystem*.Json : This is the status of the firmware update process
    %windir%\Temp\LogitechFirmware-TAP*.Json : If there is a TAP connected and the cound

    For the other searches replace RallySystem with
    SmartDock
    MeetUp
    RallyCamera
    LogiGroupSystem

    This means you can end up with a total of 9 files to process in Azure Log Analytics (4 devices with 2 JSON Files + 1 Tap).
    The process Log file is not designed for ALA.

.NOTES
    
    Importing into Azure Log Analytics
    Using Rally System as an example, you just need to swap out the file name for each product

    1. Go to your Log Analytics Workspace > Advanced Settingds
    2. Go to Data >  Custom Logs and add a log
    3. Browse for an example file LogitechFirmware-RallySystem*.Json (you can copy it from a system you ran it on)
    4. Record Dlimiter is New Line
    5. Add the Path %windir%\Temp\LogitechFirmware-RallySystem*.Json
    6. Assign the Name LogitechFirmwareRallySystem_CL (Log Analytics does not support the -)
    7. Add the Update Log, same as above except the following
        File Name: %windir%\Temp\LogitechFirmwareUpdate-RallySystem*.json
        Log Name: LogitechFirmwareUpdateRallySystem_CL
    8. Add Tap
        File Name: %windir%\Temp\LogitechFirmware-TAP*.json
        Log Name: LogitechFirmwareTAP_CL

    I was going to add a parameter se to set what is actually searched, but the WQL searches are quick and as I started working on doing that parameter, it just go more tricky.
    This script is designed to use stand alone or with SCCM, hence why I use a String of True or False, this is an issue with how the -File is executed from a command line 

.PARAMETER Update
    Required True or False. As a String (not the boolean $True|$False), this is to support SCCM
    Set to True Updates will be applied
    Set to False inventory only

.PARAMETER Smartdock
    Optional. True or False. As a String (not the boolean $True|$False), this is to support SCCM
    If used firmware will be updated to smartdock
    I called this out seperatly since updating a smartdock may require hands on to powercycle, especially if a Flex is connected.

.PARAMETER rs
    Required if using Update
    Dynamic: RightSight will always Run
    OCS: Will only run at begining of meeting
    Off

.EXAMPLE 
    Will Just Query Systems
    Get-LogitechFirmware.ps1 -Update False

    Updates all but SmartDock
    Get-LogitechFirmware.ps1 -Update True -rs Dynamic

    Updates SmartDock Only
    Get-LogitechFirmware.ps1 -Update True -SmartDock True

    Updates Everything including SmartDock
    Get-LogitechFirmware.ps1 -Update True -rs:Dynamic -SmartDock
#>

param (
    [Parameter(Mandatory = $True,
        ParameterSetName = "Update")]
    [ValidateSet($True, $False)]
    [string[]] $update,

    [Parameter(Mandatory = $false,
        ParameterSetName = "Update")]
    [ValidateSet("Dynamic", "OCS", "Off")]
    [string[]] $rs,

    [Parameter(Mandatory = $false)]
    [ValidateSet($True, $False)]
    [string[]] $SmartDock

)

Start-Transcript C:\Windows\temp\logifirmware.log -Force

#Get the Windows Temp Directory
$TempDir = [System.Environment]::GetEnvironmentVariable('TEMP', 'Machine')
$timestamp = (get-date -Format FileDateTime)

Write-Output $TempDir
Write-Output $timestamp

#Generate the status log File
$SmartDockLog = $tempdir + "\" + "LogitechFirmware-SmartDock" + $timestamp + ".json"
$MeetupLog = $tempdir + "\" + "LogitechFirmware-MeetUp" + $timestamp + ".json"
$RallyCameraLog = $tempdir + "\" + "LogitechFirmware-RallyCamera" + $timestamp + ".json"
$RallySystemLog = $tempdir + "\" + "LogitechFirmware-RallySystem" + $timestamp + ".json"
$LogiGroupLog = $tempdir + "\" + "LogitechFirmware-LogiGroupSystem" + $timestamp + ".json"
$TAPLog = $tempdir + "\" + "LogitechFirmware-TAP" + $timestamp + ".json"

#Updater JSON Log Files
$UpdateSmartDockLog = $tempdir + "\" + "LogitechFirmwareUpdate-SmartDock" + $timestamp + ".json"
$UpdateMeetupLog = $tempdir + "\" + "LogitechFirmwareUpdate-MeetUp" + $timestamp + ".json"
$UpdateRallyCameraLog = $tempdir + "\" + "LogitechFirmwareUpdate-RallyCamera" + $timestamp + ".json"
$UpdateRallySystemLog = $tempdir + "\" + "LogitechFirmwareUpdate-RallySystem" + $timestamp + ".json"
$UpdateLogiGroupLog = $tempdir + "\" + "LogitechFirmwareUpdate-LogiGroupSystem" + $timestamp + ".json"

<###############
Check What is Connected and the quantity
################>

Function Search-LogiDevices ([string]$Query) {
    #I have to use a count since it is possible for there to be more than one camera
    Try {
        $ObjectFound = [int]((Get-WmiObject -Query $Query) | Measure-Object).Count
        Return $ObjectFound
        Write-Output $ObjectFound
    }
    Catch { Return 0 }
    $Error.Clear()
}

Function Invoke-RightSighInstall ([string]$file, $logfile) {
    $InstallFilePath = $file
    $argumentlist = "/rs-" + $rs + "=1 /S"
    write-host "  Running: $installFilePath $ArgumentList to Log File $logfile"
    #Install the RightSight vcapp
    Write-Host "  Installing Righsight"
    try {
        Start-Process -FilePath $InstallFilePath -ArgumentList $argumentlist -wait -ErrorAction Stop -NoNewWindow -RedirectStandardOutput $logfile
        Write-Host "   Installed"
        Return "Successfull"
    }
    catch {
        Write-Output "   Failed: $error[0]"
        return "Failed"
    }
    $Error.Clear()
}

Function Invoke-FirmwareUpdater  ($Hardware, $Logfile) {
    $CommonProgramFiles = ${env:CommonProgramFiles(x86)}
    $FWInstallFilePath = $CommonProgramFiles + "\LogiShrd\LogiFirmwareUpdateTool-" + $Hardware + "\FWUpdate*.exe"
    $FWInstallFilePath = get-item -path $FWInstallFilePath
    $FWargumentlist = "/silentUpdate /ForceUpdate"
    write-host "  Running: $FWinstallFilePath $FWArgumentList to Log File $logfile"
    Write-Host "  Updating Firmware"
    try {
        Start-Process -filepath $FWInstallFilePath -ArgumentList $FWargumentlist -ErrorAction Stop -NoNewWindow -RedirectStandardOutput $Logfile
        return "Successfull"
        Write-Host "   Succssesful"
    }
    catch {
        Write-Output "   Failed: $Error[0]"
        return "Failed"
    }
    $Error.Clear()
}

if (test-path ".\FWUpdateRally.exe") {
    $RallyDevice = Search-LogiDevices 'SELECT Caption FROM Win32_USBDevice WHERE Caption = "Logi Rally Table Hub"'
}
if (Test-Path ".\FWUpdateMeetUp.exe") {
    $MeetUpDevice = Search-LogiDevices 'SELECT Caption FROM Win32_PnPEntity WHERE Caption LIKE "Logi% MeetUp WinUSB" AND Manufacturer LIKE "Logi%"'
}
if (Test-Path ".\UpdateRallyCamera.exe") {
    $RallyCameraDevice = Search-LogiDevices 'SELECT Caption FROM Win32_USBDevice Where Caption = "Logi Rally Camera"'
}
if (Test-Path ".\SmartDockUpdate.exe ") {
    $SmartDockDevice = Search-LogiDevices 'SELECT Caption FROM Win32_PnPEntity WHERE Caption LIKE "Logi% SmartDock" AND Manufacturer LIKE "Logi%"'
}
if (Test-Path ".\LogiGroupUpdater\FWUpdateLogiGroup.exe") {
    $GroupDevice = Search-LogiDevices 'SELECT Caption FROM Win32_PnPEntity WHERE Caption LIKE "Logi Group Speakerphone"'
}

#No Updater yet
$TapDevice = Search-LogiDevices 'SELECT Caption FROM Win32_USBDevice WHERE Caption Like "%Tap%" AND Service = "usbvideo"'

<###############
SmartDock
################>
If ($SmartDockDevice -eq 1) {
    Write-Output "SmartDock"
    $logfile = $TempDir + "\smatdocklog.log"

    if (-not $SmartDock -and (-not $Update)) {
        #Check the Firmware
        Start-Process -FilePath ".\SmartDockUpdate.exe" -ArgumentList "-c" -RedirectStandardOutput $logfile -NoNewWindow -Wait
        write-Output "  Running: .\SmartDockUpdate.exe -c to Log File $logfile"
        Write-Output "  LogFile:$LogFile"
        #Gather the Log File gather the last one, there should only ever be one, but just in case.
        $lastlog = $logfile | Sort-Object -Property LastWriteTime | Select-Object -last 1
        Write-Output "  LastLog:$LastLog"
        #Parse the Log File
        $AIT = (Select-String -Path $lastlog -Pattern 'INFO: MAIN - AIT Device version is : ' | Select-Object -last 1 | ConvertFrom-String).P9
        Write-Output "  AIT:$AIT"
        $NXP = (Select-String -Path $lastlog -Pattern 'INFO: MAIN - NXP Device version is : ' | Select-Object -last 1 | ConvertFrom-String).P9
        Write-Output "  NXP:$NXP"
        #Write the JSON file
        New-Object psobject -Property @{devicename = "SmartDock"; AIT = $AIT; NXP = $NXP } | ConvertTo-Json -Compress | Out-File -FilePath $SmartDockLog -Encoding ascii -Append
        Write-Output "  JSON:$SmartDockLog"
        #Update Firmware
    }
    if ($SmartDock) {
        Write-Output "  Updating"
        $updater = Get-Item -Path ".\SmartDockUpdater\SmartDockUpdateInstall*.exe"
        Write-Output "  File:$updater.fullname"
        Write-Output "  Running: $updater.Fullname -s to Log File $logfile"
        Start-Process -FilePath ($updater.Fullname) -ArgumentList "-s" -RedirectStandardError $logfile -Wait -NoNewWindow
        #Gather the Log File gather the last one, there should only ever be one, but just in case.
        $lastlog = $logfile | Sort-Object -Property LastWriteTime | Select-Object -last 1
        Write-Output "  LastLog:$LastLog"
        #Parse the Log File
        $AIT = (Select-String -Path $lastlog -Pattern 'INFO: MAIN - AIT Device version is : ' | Select-Object -last 1 | ConvertFrom-String).P10
        $NXP = (Select-String -Path $lastlog -Pattern 'INFO: MAIN - NXP Device version is : ' | Select-Object -last 1 | ConvertFrom-String).P10
        if ((Select-String -Path $lastlog -Pattern 'MAIN - AIT Device Update Completed' | Select-Object -last 1 | ConvertFrom-String).P9 -like "Successfully") {
            $AITStatus = "Successfull"
        }
        else {
            $AITStatus = "Failed"
        }
        if ((Select-String -Path $lastlog -Pattern 'MAIN - NXP Device Update Completed ' | Select-Object -last 1 | ConvertFrom-String).P9 -like "Successfully") {
            $NXPStatus = "Successfull"
        }
        else {
            $NXPStatus = "Failed"
        }
        
        Write-Output "  AIT:$AIT"
        Write-Output "  NXP:$NXP"
        Write-Output "  AITStatus:$AITStatus"
        Write-Output "  NXPStatus:$NXPStatus"
        #Write the JSON file
        New-Object psobject -Property @{NXP = $NXP; NXPStatus = $NXPStatus; AITStatus = $AITStatus; devicename = "SmartDock"; AIT = $AIT } | ConvertTo-Json -Compress | Out-File -FilePath $UpdateSmartDockLog -Encoding ascii -Append
        Write-Output "  JSON:$UpdateSmartDockLog"
    }
    
}

<###############
Meetup
################>
if ($MeetUpDevice -gt 0) {
    Write-Output "MeetUp"
    $logfile = $TempDir + "\MeetUp.log"
    Switch ($Update) {
    
        $False {
                
            Write-Output "  LogFile:$LogFile"
            #Check the Firmware
            $MeetUpChecker = get-item ".\FWUpdateMeetUp*.exe"
            Write-Output "  Running: $($MeetUpChecker.FullName) /versioninfo to Log File $logfile"
            Start-Process -FilePath ($MeetUpChecker.FullName) -ArgumentList "/versioninfo" -RedirectStandardOutput $logfile -NoNewWindow -Wait
            #Gather the Log File gather the last one, there should only ever be one, but just in case.
            $lastlog = $logfile | Sort-Object -Property LastWriteTime | Select-Object -last 1
            Write-Output "  LastLog:$LastLog"
            #Parse the Log File
            [int]$videoline = (Select-String -Pattern "Info:   video" -Path $lastlog | Select-Object -first 1).linenumber
            $videostatus = (get-content -Path $lastlog -TotalCount ($videoline + 2))[-1]
            $videoversion = (([string]$videostatus | ConvertFrom-String -Delimiter "`t").p2 -split " ")[1]
            Write-Output "  Video:$videoversion"
            #Write the JSON file
            New-Object psobject -Property @{video = $videoversion; devicename = "MeetUp"; } | ConvertTo-Json -Compress | Out-File -FilePath $MeetUpLog -Encoding ascii -Append
            Write-Output "  JSON:$MeetUpLog"
        }
    
        $True {
            $MeetUpUpdater = get-item ".\MeetUpUpdater\FWUpdateMeetUp*.exe"
            #The order of these two must be as follows. The Firmware cannot be installed unless RightSight is installed first.
            $InstallRSApp = Invoke-RightSighInstall ($MeetUpUpdater.FullName) $Logfile
            Start-Sleep -Seconds 30 #Need to Sleep after firmware updater since the installer does not return an exit code for a Wait in the start-process
            if ($InstallRSApp -like "Successfull") {
                $UpdateFirmware = Invoke-FirmwareUpdater "MeetUp" $logfile
                Start-Sleep -Seconds 60 #Need to Sleep after firmware updater since the installer does not return an exit code for a Wait in the start-process
            }
            New-Object psobject -Property @{FirmwareUpdate = $UpdateFirmware; RightSightInstall = $InstallRSApp; devicename = "MeetUp"; } | ConvertTo-Json -Compress | Out-File -FilePath $UpdateMeetUpLog -Encoding ascii -Append
            Write-Output "  JSON:$UpdateMeetUpLog"
        }
    }


}

<##############
Rally Camera
################>
if (($RallyCameraDevice -gt 0) -And (-not $RallyDevice)) {
    $logfile = $TempDir + "\RallyCamera.log"
    Switch ($Update) {
    
        $False {
                
            Write-Output "  LogFile:$LogFile"
            #Check the Firmware
            $RallyCameraChecker = get-item ".\FWUpdateRallyCamera*.exe"
            Write-Output "  Running: $($RallyCameraChecker.FullName) /versioninfo to Log File $logfile"
            Start-Process -FilePath ($RallyCameraChecker.FullName) -ArgumentList "/versioninfo" -RedirectStandardOutput $logfile -NoNewWindow -Wait
            #Gather the Log File gather the last one, there should only ever be one, but just in case.
            $lastlog = $logfile | Sort-Object -Property LastWriteTime | Select-Object -last 1
            Write-Output "  LastLog:$LastLog"
            #Parse the Log File
            [int]$videoline = (Select-String -Pattern "Info:   video" -Path $lastlog | Select-Object -first 1).linenumber
            $videostatus = (get-content -Path $lastlog -TotalCount ($videoline + 2))[-1]
            $videoversion = (([string]$videostatus | ConvertFrom-String -Delimiter "`t").p2 -split " ")[1]
            Write-Output "  Video:$videoversion"
            #Write the JSON file
            New-Object psobject -Property @{video = $videoversion; devicename = "RallyCamera"; } | ConvertTo-Json -Compress | Out-File -FilePath $RallyCameraLog -Encoding ascii -Append
            Write-Output "  JSON:$RallyCameraLog"
        }
    
        $True {
            $RallyCameraUpdater = get-item ".\RallyCameraUpdater\FWUpdateRallyCamera*.exe"
            #The order of these two must be as follows. The Firmware cannot be installed unless RightSight is installed first.
            $InstallRSApp = Invoke-RightSighInstall ($RallyCameraUpdater.FullName) $Logfile
            Start-Sleep -Seconds 30 #Need to Sleep after firmware updater since the installer does not return an exit code for a Wait in the start-process
            if ($InstallRSApp -like "Successfull") {
                $UpdateFirmware = Invoke-FirmwareUpdater "RallyCamera" $logfile
                Start-Sleep -Seconds 60 #Need to Sleep after firmware updater since the installer does not return an exit code for a Wait in the start-process
            }
            New-Object psobject -Property @{FirmwareUpdate = $UpdateFirmware; RightSightInstall = $InstallRSApp; devicename = "RallyCamera"; } | ConvertTo-Json -Compress | Out-File -FilePath $UpdateRallyCameraLog -Encoding ascii -Append
            Write-Output "  JSON:$UpdateRallyCameraLog"
        }
    }
}

<###############
Rally System
################>
if ($RallyDevice -eq 1) {
    Write-Output "RallySystem"
    $logfile = $TempDir + "\RallySystem.log"
    Switch ($Update) {

        $False {
            
            Write-Output "  LogFile:$LogFile"
            #Check the Firmware
            $RallySystemChecker = get-item ".\FWUpdateRally*.exe"
            Write-Output "  Running: $($RallySystemChecker.FullName) /versioninfo to Log File $logfile"
            Start-Process -FilePath ($RallySystemChecker.FullName) -ArgumentList "/versioninfo" -RedirectStandardOutput $logfile -NoNewWindow -Wait
            #Gather the Log File gather the last one, there should only ever be one, but just in case.
            $lastlog = $logfile | Sort-Object -Property LastWriteTime | Select-Object -last 1
            Write-Output "  LastLog:$LastLog"
            #Parse the Log File
            [int]$tablehubline = (Select-String -Pattern "Info:   tablehub" -Path $lastlog | Select-Object -first 1).linenumber
            $tablehubstatus = (get-content -Path $lastlog -TotalCount ($tablehubline + 2))[-1]
            $tablehubversion = (([string]$tablehubstatus | ConvertFrom-String -Delimiter "`t").p2 -split " ")[1]
            Write-Output "  tablehub:$tablehubversion"
            #Write the JSON file
            New-Object psobject -Property @{tablehub = $tablehubversion; devicename = "RallySystem"; } | ConvertTo-Json -Compress | Out-File -FilePath $RallySystemLog -Encoding ascii -Append
            Write-Output "  JSON:$RallySystemLog"
        }

        $True {
            $RallySystemUpdater = get-item ".\RallySystemUpdater\FWUpdateRally*.exe"
            #The order of these two must be as follows. The Firmware cannot be installed unless RightSight is installed first.
            $InstallRSApp = Invoke-RightSighInstall ($RallySystemUpdater.FullName) $Logfile
            Start-Sleep -Seconds 30 #Need to Sleep after firmware updater since the installer does not return an exit code for a Wait in the start-process
            if ($InstallRSApp -like "Successfull") {
                $UpdateFirmware = Invoke-FirmwareUpdater "Rally" $logfile
                Start-Sleep -Seconds 60 #Need to Sleep after firmware updater since the installer does not return an exit code for a Wait in the start-process
            }
            New-Object psobject -Property @{FirmwareUpdate = $UpdateFirmware; RightSightInstall = $InstallRSApp; devicename = "RallySystem"; } | ConvertTo-Json -Compress | Out-File -FilePath $UpdateRallySystemLog -Encoding ascii -Append
            Write-Output "  JSON:$UpdateRallySystemLog"
        }
    }
}

<###############
Group
################>
if ($GroupDevice -eq 1) {
    Write-Output "GroupSystem"
    $logfile = $TempDir + "\GroupSystem.log"
    Switch ($Update) {

        $False {
            
            Write-Output "  LogFile:$LogFile"
            #Check the Firmware
            $LogiGroupChecker = get-item ".\LogiGroupUpdater\FWUpdateLogiGroup.exe"
            Write-Output "  Running: $($LogiGroupChecker.FullName) /versioninfo to Log File $logfile"
            Start-Process -FilePath ($LogiGroupChecker.FullName) -ArgumentList "/versioninfo" -RedirectStandardOutput $logfile -NoNewWindow -Wait
            #Gather the Log File gather the last one, there should only ever be one, but just in case.
            $lastlog = $logfile | Sort-Object -Property LastWriteTime | Select-Object -last 1
            Write-Output "  LastLog:$LastLog"
            #Parse the Log File
            [int]$videoline = (Select-String -Pattern "Info:   video" -Path $lastlog | Select-Object -first 1).linenumber
            $videostatus = (get-content -Path $lastlog -TotalCount ($videoline + 2))[-1]
            $videoversion = (([string]$videostatus | ConvertFrom-String -Delimiter "`t").p2 -split " ")[1]
            Write-Output "  video:$videoversion"
            #Write the JSON file
            New-Object psobject -Property @{video = $videoversion; devicename = "LogiGroup"; } | ConvertTo-Json -Compress | Out-File -FilePath $LogiGroupLog -Encoding ascii -Append
            Write-Output "  JSON:$LogiGroupLog"
        }

        $True {
            $StartTime = Get-Date
            $LogiGroupUpdater = get-item ".\LogiGroupUpdater\FWUpdateLogiGroup.exe"
            $FWargumentlist = "/silentUpdate /force"
            #Group Firmware is updated differently than others. It calls the Extracted file instead.
            Write-Host "  Updating Firmware"
            Write-Output "  Running: $($LogiGroupUpdater.FullName) $FWargumentlist to Log File $logfile"
            try {
                Start-Process -filepath ($LogiGroupUpdater.FullName) -ArgumentList $FWargumentlist -ErrorAction Stop -NoNewWindow -RedirectStandardOutput $Logfile
                $UpdateFirmware = "Successfull"
                Write-Host "   Succssesful"
            }
            catch {
                Write-Output "   Failed: $Error[0]"
                $UpdateFirmware = "Failed"
            }

            $LogPathItem = get-item $logfile
            $filedifftime = 0

            While (((Get-Content -Path $logfile -Tail 1) -notlike "*Info: Update*") -or ($filedifftime -lt 3)) {
                $StartTime = Get-Date
                $LogPathItem = get-item $LogPathItem
                $filedifftime = [math]::abs((New-TimeSpan -Start $StartTime -End $LogPathItem.LastWriteTime).Minutes)
                Start-Sleep -Seconds 60
                Write-Host "Updating"
            }

            write-host "  Update Completed"
            $UpdateStatus = ((Get-Content -Path $LogFile -Tail 1) | ConvertFrom-String).P3
            New-Object psobject -Property @{FirmwareUpdate = $UpdateStatus; devicename = "LogiGroup"; } | ConvertTo-Json -Compress | Out-File -FilePath $UpdateLogiGroupLog -Encoding ascii -Append
            Write-Output "  JSON:$UpdateLogiGroupLog"
        }
    }
}

<###############
TAP Device
################>
if ($TapDevice -gt 0) {
    Write-output "Tap"
    #Write the JSON file
    New-Object psobject -Property @{TapCount = $TapDevice; devicename = "TAP" } | ConvertTo-Json -Compress | Out-File -FilePath $TAPLog -Encoding ascii -Append
    Write-Output "  JSON:$TAPLog"
}

Stop-Transcript