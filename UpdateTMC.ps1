<#
.NAME
Update TMC

.SYNOPSIS

Connects to a remote machine, tests for the installation of TMC and what version, then updates it if required.

.DESCRIPTION

This script will request a computer name to have TMC updated on.  When run it requests a computer, checks the remove computer for the version of TMC installed, and installs the new version if it's a higher version number.  It then copies a master config file over to the remote computer.
It requires the remote computer to have WinRM configured and running on it.  It also requires the powershell process to be running as an administrator.  

.PARAMETER

-RemoteComputer <string>

.EXAMPLE

Update TMC.ps1

.EXAMPLE

'.\Update TMC.ps1' Localhost

.Example

'.\Update TMC.ps1' -ComputerName Localhost
#>
#[CmdletBinding()]
#Param (
#    [string]$RemoteComputer = $(Read-Host "Name of the remote computer")
#    
#    )
#    
If($psBoundParameters['debug'])
    {
    $DebugPreference = "Continue"
    }
Else
    {
    $DebugPreference = "SilentlyContinue"
    }


# Get the ID and security principal of the current user account
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
 
# Get the security principal for the Administrator role
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
 
# Check to see if we are currently running "as Administrator"
if ($myWindowsPrincipal.IsInRole($adminRole))
   {
   # We are running "as Administrator" - so change the title and background color to indicate this
   $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)"
   $Host.UI.RawUI.BackgroundColor = "DarkBlue"
   clear-host
   }
else
   {
   # We are not running "as Administrator" - so relaunch as administrator
   
   # Create a new process object that starts PowerShell
   $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
   
   # Specify the current script path and name as a parameter
   $newProcess.Arguments = $myInvocation.MyCommand.Definition;
   
   #$newProcess.Arguments = “-NoExit -Command &’$PSCommandPath’$CommandLineAfterScriptFileSpec”
   #$newProcess.Arguments = “$myInvocation.MyCommand.Definition -NoExit -Command &’$PSCommandPath’$CommandLineAfterScriptFileSpec”

   # Indicate that the process should be elevated
   $newProcess.Verb = "runas";
   
   #Set the UseShellExecute policy to true
   #$newProcess.UseShellExecute = $False
   # Start the new process
   [System.Diagnostics.Process]::Start($newProcess);
   
   # Exit from the current, unelevated, process
   exit
   }
 
# Run your code that needs to be elevated here

[string]$RemoteComputer = $(Read-Host "Name of the remote computer")

If ((Test-Path "\\$RemoteComputer\c$\Program Files (x86)\OnGuard\CustomSolutions\TerminalManagerSetup\TerminalManager.exe") -And ($DebugPreference -ieq "Continue"))
    {
    #Test the version numbers.  If the version to install is lower then then installed, cancel.
    $TMCFolder = Test-Path "c:\tools\TMC\Temp"
    If (-not $TMCFolder )
        {
        New-Item "c:\Tools\TMC\Temp" -Type Directory
        }
    Invoke-Command -ScriptBlock { C:\Windows\SysWOW64\msiexec.exe /a '\\sgcs\software$\TMC\Terminal Manager Current\TerminalManagerSetup.msi' /qn TARGETDIR=C:\Tools\TMC\TEMP}
    $installed = (Get-Item 'C:\Program Files (x86)\OnGuard\CustomSolutions\TerminalManagerSetup\TerminalManager.exe')
    $InstalledVersion = $installed.VersionInfo.ProductVersion.tostring()
    $Installer = (Get-Item 'C:\Tools\TMC\Temp\TerminalManager.exe')
    $Installer.FileVersionInfo
    $Installerversion = $Installer.VersionInfo.ProductVersion.tostring()
    
    #Debug outputs.
    Write-Debug -Message "The installer version: $InstallerVersion"
    Write-Debug -Message "The installed version: $InstalledVersion"
    #/Debug outputs
    
    Remove-Item "c:\Tools\TMC\Temp\*" -force
    }
    
# Copy the install files to the local machine.
    
$TMCFolder = Test-Path "\\$RemoteComputer\c$\tools\TMC\Current"
If (-not $TMCFolder )
    {
    New-Item "\\$RemoteComputer\c$\Tools\TMC\Current" -Type Directory
    }
Else
    {
    #performing the uninstall here before we move the installer to a new folder.  
   	Write-Host "Attempting to uninstall the current version of TMC"
	invoke-command -computername $RemoteComputer -ScriptBlock { msiexec /x 'C:\Tools\TMC\Current\TerminalManagerSetup.msi' /L+*v 'c:\tools\TMC\installlogfile.txt' ALLUSERS=1 /qb } -Verbose
    Start-Sleep 5
    $Date = (Get-Date).ToString('ddMMyyyy_hhmmsss')
    Rename-Item -Path "\\$RemoteComputer\c$\Tools\TMC\Current" -NewName "\\$RemoteComputer\c$\Tools\TMC\$Date"
    New-Item "\\$RemoteComputer\c$\Tools\TMC\Current" -Type Directory
    }

Copy-Item -path "\\sgcs.security.local\software$\TMC\Terminal Manager Current\*.*" -Destination "\\$RemoteComputer\c$\Tools\TMC\Current\" -Verbose

Write-Debug "entering PSSesstion on $RemoteComputer"
    Enter-PSSession $RemoteComputer    
    #If Termainal Manager is running, stop the application so it can be updated.
    Try {
        #Get-Process -ComputerName $RemoteComputer -ProcessName "TerminalManager"
        Write-Debug "Attempting to stop the TerminalManager process"
        invoke-command -computername $RemoteComputer -ScriptBlock {stop-process -name "TerminalManager" -Force} -ErrorAction SilentlyContinue
        }
    Catch
        {
        Write-host "Process not found"
        }
    Exit-PSSession  
    
    #Install and set the configuration files for Terminal Manager.
	#Write-Host "Attempting to uninstall the current version of TMC"
	#invoke-command -computername $RemoteComputer -ScriptBlock { msiexec /x 'C:\Tools\TMC\Current\TerminalManagerSetup.msi' /L+*v 'c:\tools\TMC\installlogfile.txt' ALLUSERS=1 /qb } -Verbose
    #Start-Sleep 5
    Write-Host "Attempting to install the Terminal Manager software"
    invoke-command -computername $RemoteComputer -ScriptBlock { msiexec /i 'C:\Tools\TMC\Current\TerminalManagerSetup.msi' /L+*v 'c:\tools\TMC\installlogfile.txt' ALLUSERS=1 /qb } -Verbose
    Start-Sleep 5
    Copy-Item -Path "\\sgcs.security.local\software$\TMC\Terminal Manager Configs\*.*" -Destination "\\$RemoteComputer\c$\Program Files (x86)\OnGuard\CustomSolutions\TerminalManagerSetup\" -verbose
    If(Compare-Object $(Get-Content "\\$RemoteComputer\c$\Program Files (x86)\OnGuard\CustomSolutions\TerminalManagerSetup\TerminalManager.exe.config") $(Get-Content "\\failover\tmc$\Terminal Manager Configs\TerminalManager.exe.config"))
    {
        Copy-Item -Path "\\sgcs.security.local\software$\tmc\Terminal Manager Configs\*.*" -Destination "\\$RemoteComputer\c$\Program Files (x86)\OnGuard\CustomSolutions\TerminalManagerSetup\" -verbose
    }

exit

