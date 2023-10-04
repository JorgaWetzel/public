<#PSScriptInfo
.VERSION 1.0.0
.GUID d4bf7f0a-794e-4c16-bd3f-86ea7074cae9
.AUTHOR AndrewTaylor
.DESCRIPTION   Displays a GUI during ESP with a selection of troubleshooting tools
.COMPANYNAME 
.COPYRIGHT GPL
.TAGS Intune Autopilot
.LICENSEURI https://github.com/andrew-s-taylor/public/blob/main/LICENSE
.PROJECTURI https://github.com/andrew-s-taylor/public
.ICONURI 
.EXTERNALMODULEDEPENDENCIES
.REQUIREDSCRIPTS 
.EXTERNALSCRIPTDEPENDENCIES
.RELEASENOTES
#>
<#
.SYNOPSIS
  Displays a GUI during ESP with a selection of troubleshooting tools
.DESCRIPTION
  Displays a GUI during ESP with a selection of troubleshooting tools

.INPUTS
None required
.OUTPUTS
GridView
.NOTES
  Version:        1.0.0
  Author:         Andrew Taylor
  Twitter:        @AndrewTaylor_2
  WWW:            andrewstaylor.com
  Creation Date:  03/08/2022
  Purpose/Change: Initial script development

  
.EXAMPLE
N/A
#>


##Create the Form
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$AutopilotMenu             = New-Object system.Windows.Forms.Form
$AutopilotMenu.ClientSize  = New-Object System.Drawing.Point(396,500)
$AutopilotMenu.text        = "Autopilot Tools"
$AutopilotMenu.TopMost     = $false
$AutopilotMenu.BackColor   = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Powered bei JJW oneICT AG"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(7,480)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',8)

$eventvwr                       = New-Object system.Windows.Forms.Button
$eventvwr.text                  = "Event Viewer"
$eventvwr.width                 = 157
$eventvwr.height                = 56
$eventvwr.location              = New-Object System.Drawing.Point(21,18)
$eventvwr.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

$scriptrun                         = New-Object system.Windows.Forms.Button
$scriptrun.text                    = "Troubleshooting Script"
$scriptrun.width                   = 157
$scriptrun.height                  = 56
$scriptrun.location                = New-Object System.Drawing.Point(219,16)
$scriptrun.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

$regedit                           = New-Object system.Windows.Forms.Button
$regedit.text                      = "Regedit"
$regedit.width                     = 157
$regedit.height                    = 56
$regedit.location                  = New-Object System.Drawing.Point(22,131)
$regedit.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

$explorer                          = New-Object system.Windows.Forms.Button
$explorer.text                     = "File Explorer"
$explorer.width                    = 157
$explorer.height                   = 56
$explorer.location                 = New-Object System.Drawing.Point(222,132)
$explorer.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

$log1                           = New-Object system.Windows.Forms.Button
$log1.text                      = "SetupAct Log"
$log1.width                     = 157
$log1.height                    = 56
$log1.location                  = New-Object System.Drawing.Point(22,245)
$log1.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

$log2                          = New-Object system.Windows.Forms.Button
$log2.text                     = "Intune Mgmt Log"
$log2.width                    = 157
$log2.height                   = 56
$log2.location                 = New-Object System.Drawing.Point(222,245)
$log2.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

# Neuer Button für IntuneManagementExtensionDiagnostics
$IntuneManagementExtensionDiagnostics = New-Object system.Windows.Forms.Button
$IntuneManagementExtensionDiagnostics.text = "IntuneManagementExtensionDiagnostics"
$IntuneManagementExtensionDiagnostics.width = 157
$IntuneManagementExtensionDiagnostics.height = 56
$IntuneManagementExtensionDiagnostics.location = New-Object System.Drawing.Point(21,360)
$IntuneManagementExtensionDiagnostics.Font = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

# Aktion für den neuen Button definieren
$IntuneManagementExtensionDiagnostics.Add_Click({
    Start-Process -FilePath "powershell.exe" -ArgumentList '-nologo -noprofile -executionpolicy bypass -command {
        cd C:\ProgramData;
        Set-ExecutionPolicy bypass -Scope Process;
        Save-Script Get-IntuneManagementExtensionDiagnostics -Path ./;
        ./Get-IntuneManagementExtensionDiagnostics.ps1;
        ./Get-IntuneManagementExtensionDiagnostics.ps1 -Online;
    }' -Wait
})

# Hinzufügen der Buttons zum Formular, inklusive des neuen Buttons
$AutopilotMenu.controls.AddRange(@($Label1,$scriptrun,$eventvwr,$regedit,$explorer,$log1, $log2, $IntuneManagementExtensionDiagnostics))


##Launch Autopilot Diagnostics in new window (don't autoclose)
$scriptrun.Add_Click({ 
start-process powershell.exe -argument '-nologo -noprofile -noexit -executionpolicy bypass -command  Get-AutopilotDiagnostics ' -Wait
   
 })

 ##Launch Event Viewer
$eventvwr.Add_Click({ 
    start-process -filepath "C:\Windows\System32\eventvwr.exe"
 })

 ##Launch Regedit
$regedit.Add_Click({ 
    start-process -filepath "C:\Windows\regedit.exe"
 })

 ##Launch Windows Explorer
$explorer.Add_Click({ 
    start-process -filepath "C:\Windows\explorer.exe"
 })

 ##Launch CMTrace on SetupAct.log
$log1.Add_Click({ 
    start-process -filepath "C:\ProgramData\ServiceUI\cmtrace.exe" -argumentlist c:\windows\panther\setupact.log
 })

    ##Launch CMTrace on IntuneMgmt.log
$log2.Add_Click({ 
    start-process -filepath "C:\ProgramData\ServiceUI\cmtrace.exe" -argumentlist C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\intunemanagementextension.log 
})


[void]$AutopilotMenu.ShowDialog()

## Install-Module IntuneStuff -Force
#Install-Module IntuneStuff -Force

## show Win32App data (source will be client Intune log file combined with registry data)
#Get-IntuneWin32AppLocally | Out-GridView

# show Win32App data (source will be just client Intune log file)
#Get-IntuneLogWin32AppData
#Get-IntuneLogWin32AppReportingResultData