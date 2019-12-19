# LogitechFirmwareManagement
Used for Updating and Inventory Logitech VC Equipment
I would reccomend using Sync if possible https://www.logitech.com/sync. Note that this only support Rally, MeetUp and later
The documentation within the PowerShell script has more details.

This script is used to query and update Logitech VC Equipment. It supports the following
<UL>
  <LI>Rally System</LI>
  <LI>Rally Camera</LI>
  <LI>MeetUp</LI>
  <LI>SmartDock</LI>
  <LI>Group</LI>
</UL>

# Output
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

# SCCM
You can find two Zipped files.

Is the File to import into SCCM
Contains all the files and folder structures for the latest firmware as of December 2019. How to get the files are listed in the PowerShell Script.

Imprt the Zipped file as a package, not as an application
there are 4 applications within the Package
<ul>
  <li>Check: Does inventory Only</li>
  <li>Update All: Updates all but SmartDock</li>
  <li>Update All But SmartDock: Updates everything but SmartDock</li>
  <li>Update SmartDock: Updates only SmartDock</li>
 </ul>
 
 You can then deploy each application to the desired Collection


# General Script Usage
It also requires a specific folder struvure
<BlockQuote>
       [folder]Logitech Updater Files<br />
  <BlockQuote>
        Get-LogitechFirmware.ps1<br />
        SmartDockUpdate.exe <br />
        FWUpdateMeetUp.exe<br />
        FWUpdateRally.exe<br />
        FWUpdateRallyCamera.exe<br />
  </BlockQuote>
        [folder]ait<br />
  <BlockQuote>
            AitUVCExtApi.dll<br />
    </BlockQuote>
        [folder]SmartDockUpdater<br />
    <BlockQuote>
            SmartDockUpdateInstall_1.2.31.48.exe<br />
      </BlockQuote>
        [folder]MeetUpUpdater<br />
      <BlockQuote>
            FWUpdateMeetup_1.10.60.exe<br />
        </BlockQuote>
        [folder]RallySystemUpdater<br />
        <BlockQuote>
            FWUpdateRally_1.4.28.exe<br />
          </BlockQuote>
        [folder]RallyCameraUpdater<br />
          <BlockQuote>
            FWUpdateRallyCamera_1.4.17.exe<br />
            </BlockQuote>
        [folder]LogiGroupUpdater<br />
            <BlockQuote>
            [folder]$PLUGINSDIR<br />
            devcon*.exe<br />
            FWUpdateLauncher.exe<br />
            FWUpdateLogiGroup.exe<br />
            icon.ico<br />
            uninstall.exe<br />
              </BlockQuote>
</BlockQuote>
The main script is Get-LogitechFirmware.ps1. There are more details in the document readme, but there are a few command options available.

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
  
