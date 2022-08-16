<# 
    Microsoft courses deployement script

    Author: Mosselmans Benjamin
    Linkedin link: https://www.linkedin.com/in/mosselmansben/
    Date: Summer 2022

    Function : Seek and deploy

    Description: This script receive as parameters the id (int) of the Microsoft course and if a boolean of the computer state (Microsoft trainer computer). With these two parameters, it'll seek on a share 
    on the network (Here: \\tke-veeam\E\MOC) for the folder containing such id, get all txt extension file names in the folder containing the Virtual Machines to then find on the other directory containing all
    base drives the corresponding base drives necessary to said course. It'll then execute the unzipping executable linked to the bases. When done, it executes the scripts given by Microsoft to deploy and snapshots
    the created Virtual Machines.
    If the computer is tagged as teacher, it'll also deploy the powerpoint and onenote of the course.

    The second part of the script go inside each Virtual Machines, and if given the right circunstences (Remote allowed), it'll delete in the registry keyboard registry keys to only let English and French(Belgium).
    The IME aren't deleted.

    Parameters: -[String] Course's id
                -[Boolean] Computer state (Microsoft Trainer computer)

    Possible upgrades: -Stock all Microsoft administrator credentials in a text file to be feed to the code with iteration until it finds the right credential to use.
                       -Change the SendKeys function to one allowed by MDT directly injected in the task sequence
                       -Delete the additionnal IME
                       -Silent rearm
#>


Param (

    [string] $CourseID,
    [bool] $Teacher
)

#--------------------------------------------------------------------------------------------------------
#-------------------Set up Variables, Pathes and the Course Path-----------------------------------------
#--------------------------------------------------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms #Tentative to add System.Windows.Forms to make Sendkeys work through MDT out of local environnement
netsh advfirewall firewall set rule group="Network Discovery" new enable=Yes #Enable the Network Discovery to find the online
netsh firewall set service type=fileandprint mode=enable profile=all #Enable the File Sharing
set-executionpolicy remotesigned -Force #Allow the execution of scripts
$_BasePath = "\\tke-veeam\E\MOC\_BASE" #Path to the Base drive directory
$Folders = "\\tke-veeam\E\MOC" #Path to the Course folder
$CoursePathLocal = $CourseID.Substring(0, $CourseID.Length - 1) #Cleaning the variable $CourseID of the last letter for use on local path

#--------------------------------------------------------------------------------------------------------
#-------------------Seek the Base Drives and then unzip the executable base drive-----------------------------------
#--------------------------------------------------------------------------------------------------------

Set-Location $Folders
$CourseFolder = Get-ChildItem -Directory | Where-Object { $_.FullName -match $CourseID } #Seek the corresponding folder to the ID Course
Set-Location $CourseFolder

foreach ($Directory in (Get-ChildItem -Directory | Where-Object { $_.FullName -match "Virtual" }) ) {
    #For each folder containing any virtual machines in case of multiple sub folder of such
    Set-Location $Directory
    foreach ($File in get-ChildItem *.txt) {
        $FileNameExtension = Split-Path $File -leaf  #Takes out the path from the variable
        $FileName = $FileNameExtension -replace "$CourseID-" -replace ".txt" #Takes out the extension
        Set-Location $_BasePath\$FileName
        foreach ($File in get-ChildItem *.exe) { 
            Start-Process -FilePath $File.Fullname -Wait -ArgumentList "/S" -PassThru  #Executes the unzip executable
        }

    }
}

#--------------------------------------------------------------------------------------------------------
#-------------------Deploy the Powerpoints and Onenote for the Microsoft Trainer-------------------------
#--------------------------------------------------------------------------------------------------------

if ($Teacher) {
    Set-Location "$Folders\$CourseFolder"
    foreach ($Directory in (Get-ChildItem -Directory | Where-Object { $_.FullName -match "Trainer" })) {
        Set-Location $Directory
        Start-Process -Filepath "$Folders\$CourseFolder\$Directory\*Powerpoint.exe" -Wait -ArgumentList "/S" -PassThru
        Expand-Archive -Path "$Folders\$CourseFolder\$Directory\*TrainerHandbook*" -DestinationPath "C:\Program Files\Microsoft Learning\$CoursePathLocal\Powerpnt"
        New-Item -itemtype symboliclink -Path "C:\Users\Administrator\Desktop" -name "Powerpoint" -value "C:\Program Files\Microsoft Learning\$CoursePathLocal\Powerpnt"
    }
}

#--------------------------------------------------------------------------------------------------------
#-------------------Creation of a shortcut for Hyper-v---------------------------------------------------
#--------------------------------------------------------------------------------------------------------

Copy-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools\Hyper-V Manager.lnk" -Destination "C:\Users\Administrator\Desktop"

#--------------------------------------------------------------------------------------------------------
#-------------------Unzip the executable of the ID Course------------------------------------------------
#--------------------------------------------------------------------------------------------------------

Set-Location "$Folders\$CourseFolder"
foreach ($Directory in (Get-ChildItem -Directory | Where-Object { $_.FullName -match "Virtual" }) ) {
    Set-Location $Directory
    foreach ($File in get-ChildItem *.exe) {
        Start-Process -FilePath $File.Fullname -Wait -ArgumentList '/S' -PassThru
    }

}

#--------------------------------------------------------------------------------------------------------
#-------------------Executes the Microsoft Course scripts with corresponding Inputs----------------------
#--------------------------------------------------------------------------------------------------------

$PathScript = "D:\Program Files\Microsoft Learning\$CoursePathLocal\Drives"
Set-Location $PathScript
.\CreateVirtualSwitches.ps1 -Wait
[System.Windows.Forms.SendKeys]::SendWait("C{ENTER}D{ENTER}{ENTER}") | .\VM-Pre-Import*.ps1 -Wait
[System.Windows.Forms.SendKeys]::SendWait("D{ENTER}") | .\*_ImportVirtualMachines.ps1 -Wait 
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}") | .\TakeVMSnapshot.ps1 -Wait

#--------------------------------------------------------------------------------------------------------
#-------------------Deleting Keyboards keys registry and rearm of the Virtual Machines-------------------
#--------------------------------------------------------------------------------------------------------

<#
    Setup of known credentials for Microsoft Virtual Machine Administrator accounts
#>
$username1 = "Adatum\Administrator"
$password1 = ConvertTo-SecureString "Pa`$`$w0rd" -AsPlainText -Force
$psCred1 = New-Object System.Management.Automation.PSCredential -ArgumentList ($username1, $password1)
$password2 = ConvertTo-SecureString "Pa55w.rd" -AsPlainText -Force
$psCred2 = New-Object System.Management.Automation.PSCredential -ArgumentList ($username1, $password2)
$username2 = "Adatum\Admin"
$psCred3 = New-Object System.Management.Automation.PSCredential -ArgumentList ($username2, $password2)
$username3 = "Admin"
$psCred4 = New-Object System.Management.Automation.PSCredential -ArgumentList ($username3, $password2)

<#
The next try and catch uses each of the setup credentials and test them until one works. It then send the command to search and delete registry keys.
#>

foreach ($VM in (Get-VM).Name) {

    Start-VM -Name $VM
    try {
        Invoke-Command -VMName $VM -Credential $psCred1 -Scriptblock {
            New-Item -Path "C:\Users\Administrator.ADATUM\Desktop" -Name "Coucou" -ItemType "Directory"
            New-PSDrive -PSProvider Registry -Name HKU -Root hkey_users
        
            foreach ($user in (Get-ChildItem -Path "HKU:\" -Name)) {
                for ($i = 0; $i -lt 26; $i++) {
    
                    Remove-ItemProperty -Path "HKU:\$user\Keyboard Layout\Preload" -Name "$i"
    
                }
                New-ItemProperty -Path "HKU:\$user\Keyboard Layout\Preload" -Name "1" -Value "0000080c"
                New-ItemProperty -Path "HKU:\$user\Keyboard Layout\Preload" -Name "2" -Value "00000409"
            }
    
            foreach ($name in (Get-Item -Path 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes' | Select-Object -ExpandProperty Property)) {
                if ($name -notlike "(Default)") {
                    Remove-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes" -Name $name
                }
    
            }
    
            New-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes" -Name "0000080c" -Value "be"
            New-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes" -Name "00000409" -Value "us"
    
            $KeyInfo = Get-ChildItem -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layouts" -Name
    
            foreach ($Key in $KeyInfo) {
                if ( ($Key -notlike "00000409") -and ($Key -notlike "0000080c")) {
                    Remove-Item -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layouts\$Key"
                }
            }

        }     
        -ErrorAction Stop
    }
    catch {
        Invoke-Command -VMName $VM -Credential $psCred2 -Scriptblock {

            New-PSDrive -PSProvider Registry -Name HKU -Root hkey_users
        
            foreach ($user in (Get-ChildItem -Path "HKU:\" -Name)) {
                for ($i = 0; $i -lt 26; $i++) {
    
                    Remove-ItemProperty -Path "HKU:\$user\Keyboard Layout\Preload" -Name "$i"
    
                }
                New-ItemProperty -Path "HKU:\$user\Keyboard Layout\Preload" -Name "1" -Value "0000080c"
                New-ItemProperty -Path "HKU:\$user\Keyboard Layout\Preload" -Name "2" -Value "00000409"
            }
    
            foreach ($name in (Get-Item -Path 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes' | Select-Object -ExpandProperty Property)) {
                if ($name -notlike "(Default)") {
                    Remove-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes" -Name $name
                }
    
            }
    
            New-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes" -Name "0000080c" -Value "be"
            New-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes" -Name "00000409" -Value "us"
    
            $KeyInfo = Get-ChildItem -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layouts" -Name
    
            foreach ($Key in $KeyInfo) {
                if ( ($Key -notlike "00000409") -and ($Key -notlike "0000080c")) {
                    Remove-Item -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layouts\$Key"
                }
            }
        }

    }

    try {
        Invoke-Command -VMName $VM -Credential $psCred3  -Scriptblock {

            New-PSDrive -PSProvider Registry -Name HKU -Root hkey_users
            
            foreach ($user in (Get-ChildItem -Path "HKU:\" -Name)) {
                for ($i = 0; $i -lt 26; $i++) {
        
                    Remove-ItemProperty -Path "HKU:\$user\Keyboard Layout\Preload" -Name "$i"
        
                }
                New-ItemProperty -Path "HKU:\$user\Keyboard Layout\Preload" -Name "1" -Value "0000080c"
                New-ItemProperty -Path "HKU:\$user\Keyboard Layout\Preload" -Name "2" -Value "00000409"
            }
        
            foreach ($name in (Get-Item -Path 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes' | Select-Object -ExpandProperty Property)) {
                if ($name -notlike "(Default)") {
                    Remove-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes" -Name $name
                }
        
            }
        
            New-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes" -Name "0000080c" -Value "be"
            New-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes" -Name "00000409" -Value "us"
        
            $KeyInfo = Get-ChildItem -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layouts" -Name
        
            foreach ($Key in $KeyInfo) {
                if ( ($Key -notlike "00000409") -and ($Key -notlike "0000080c")) {
                    Remove-Item -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layouts\$Key"
                }
            }
        }
        -ErrorAction Stop
    }
    catch {
        Invoke-Command -VMName $VM -Credential $psCred4 -Scriptblock {

            New-PSDrive -PSProvider Registry -Name HKU -Root hkey_users
            
            foreach ($user in (Get-ChildItem -Path "HKU:\" -Name)) {
                for ($i = 0; $i -lt 26; $i++) {
        
                    Remove-ItemProperty -Path "HKU:\$user\Keyboard Layout\Preload" -Name "$i"
        
                }
                New-ItemProperty -Path "HKU:\$user\Keyboard Layout\Preload" -Name "1" -Value "0000080c"
                New-ItemProperty -Path "HKU:\$user\Keyboard Layout\Preload" -Name "2" -Value "00000409"
            }
        
            foreach ($name in (Get-Item -Path 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes' | Select-Object -ExpandProperty Property)) {
                if ($name -notlike "(Default)") {
                    Remove-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes" -Name $name
                }
        
            }
        
            New-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes" -Name "0000080c" -Value "be"
            New-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layout\DosKeybCodes" -Name "00000409" -Value "us"
        
            $KeyInfo = Get-ChildItem -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layouts" -Name
        
            foreach ($Key in $KeyInfo) {
                if ( ($Key -notlike "00000409") -and ($Key -notlike "0000080c")) {
                    Remove-Item -Path "HKLM:\SYSTEM\ControlSet001\Control\Keyboard Layouts\$Key"
                }
            }
        }
    }

} 

<#
    Restarts two time the machine to update the keyboards, otherwise they do not appear.
#>
for ($i = 0; $i -lt 1; $i++) {
   
    foreach ($VM in (Get-VM).Name) {
    
        restart-VM -Name $VM -Force
    
    }
    

}
