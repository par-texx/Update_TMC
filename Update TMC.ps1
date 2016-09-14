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
[CmdletBinding()]
Param (
    [string]$RemoteComputer = $(Read-Host "Name of the remote computer")
    
    )
    
Function InstallTMC
{
    Enter-PSSession $RemoteComputer    
    #If Termainal Manager is running, stop the application so it can be updated.
    Try {
        #Get-Process -ComputerName $RemoteComputer -ProcessName "TerminalManager"
        invoke-command -computername $RemoteComputer {stop-process -name "TerminalManager"} -ErrorAction SilentlyContinue
        }
    Catch
        {
        Write-host "Process not found"
        }
    Exit-PSSession   
    
    #Install and set the configuration files for Terminal Manager.
    invoke-command -computername $RemoteComputer -ScriptBlock { msiexec /i 'C:\Tools\TMC\Current\TerminalManagerSetup.msi' /L*v 'c:\tools\TMC\installlogfile.txt' ALLUSERS=1 /qb } -Verbose
    Copy-Item -Path "\\failover.security.local\tmc$\Terminal Manager Configs\*.*" -Destination "\\$RemoteComputer\c$\Program Files (x86)\OnGuard\CustomSolutions\TerminalManagerSetup\" -verbose
    If(Compare-Object $(Get-Content "\\$RemoteComputer\c$\Program Files (x86)\OnGuard\CustomSolutions\TerminalManagerSetup\TerminalManager.exe.config") $(Get-Content "\\failover\tmc$\Terminal Manager Configs\TerminalManager.exe.config"))
    {
        Copy-Item -Path "\\failover.security.local\tmc$\Terminal Manager Configs\*.*" -Destination "\\$RemoteComputer\c$\Program Files (x86)\OnGuard\CustomSolutions\TerminalManagerSetup\" -verbose
    }
}

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
if (-not $myWindowsPrincipal.IsInRole($adminRole))
    {
       
    # We are not running "as Administrator" - so relaunch as administrator
       
    Write-Host "You must be running as an administrator. Please close and runas administrator."
     
    # Exit from the current, unelevated, process
    exit
    }

#$RemoteComputer = Read-Host -Prompt "What is the computer you wish to connect to?"

#First we check to make sure that the version being installed is a newer version then the one already installed.  
#If nothing is installed, go ahead and install it.

If (Test-Path "\\$RemoteComputer\c$\Program Files (x86)\OnGuard\CustomSolutions\TerminalManagerSetup\TerminalManager.exe")
    {
    #Test the version numbers.  If the version to install is lower then then installed, cancel.
    $TMCFolder = Test-Path "c:\tools\TMC\Temp"
    If (-not $TMCFolder )
        {
        New-Item "c:\Tools\TMC\Temp" -Type Directory
        }
    Invoke-Command -ScriptBlock { C:\Windows\SysWOW64\msiexec.exe /a '\\failover.security.local\tmc$\Teminal Manager Current\TerminalManagerSetup.msi' /qn TARGETDIR=C:\Tools\TMC\TEMP}
    $installed = (Get-Item 'C:\Program Files (x86)\OnGuard\CustomSolutions\TerminalManagerSetup\TerminalManager.exe')
    $InstalledVersion = $installed.VersionInfo.ProductVersion.tostring()
    $Installer = (Get-Item 'C:\Tools\TMC\Temp\TerminalManager.exe')
    $Installer.FileVersionInfo
    $Installerversion = $Installer.VersionInfo.ProductVersion.tostring()
    
    #Debug outputs.
    Write-Debug -Message "The installer version: $InstallerVersion"
    Write-Debug -Message "The installed version: $InstalledVersion"
    #/Debug outputs
    
    If ([Version]$InstallerVersion -LT [version]$InstalledVersion)
        {
        Write-Host "The version to be installed is lower then the current installed."
        Exit
        }
    Elseif ([Version]$InstallerVersion -EQ [version]$InstalledVersion)
        {
        Write-Host "The version to be installed is the same version the current installed."
        Exit
        }
    }
    
# Copy the install files to the local machine.
    
$TMCFolder = Test-Path "\\$RemoteComputer\c$\tools\TMC\Current"
If (-not $TMCFolder )
    {
    New-Item "\\$RemoteComputer\c$\Tools\TMC\Current" -Type Directory
    }
Else
    {
    $Date = (Get-Date).ToString('ddMMyyyy_hhmmsss')
    Rename-Item -Path "\\$RemoteComputer\c$\Tools\TMC\Current" -NewName "\\$RemoteComputer\c$\Tools\TMC\$Date"
    New-Item "\\$RemoteComputer\c$\Tools\TMC\Current" -Type Directory
    }
Copy-Item -path "\\failover.security.local\tmc$\Teminal Manager Current\*.*" -Destination "\\$RemoteComputer\c$\Tools\TMC\Current\" -Verbose
InstallTMC

exit

