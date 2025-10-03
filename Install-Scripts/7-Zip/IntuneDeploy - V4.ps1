# ===== Einfaches Intune-Upload & Assignment (Gruppennamen) =====
# Struktur:
#   Commands\Install.ps1, Uninstall.ps1, Detection.ps1, App.params.psd1
#   Icon\*.png (optional)
#   Source\  (wird befüllt)
#   Output\  (wird geleert)

$ErrorActionPreference = 'Stop'

# --- Mandant (App-Only) ---
$TenantId     = "6ac80ced-0455-4575-86c7-777085e7f076"
$ClientId     = "33461a5c-9560-49d6-8f4a-b05c2d36897a"
$ClientSecret = "h8ZHivI5YKBcT6Lb4Kam8IMvPqrZwohukJcRLeE6Vrc="

# --- Module ---
Import-Module SvRooij.ContentPrep.Cmdlet -ErrorAction SilentlyContinue
Import-Module IntuneWin32App             -ErrorAction SilentlyContinue
if (-not (Get-Module SvRooij.ContentPrep.Cmdlet)) { Install-Module SvRooij.ContentPrep.Cmdlet -Scope CurrentUser -Force; Import-Module SvRooij.ContentPrep.Cmdlet -ErrorAction Stop }
if (-not (Get-Module IntuneWin32App))   { Install-Module IntuneWin32App   -Scope CurrentUser -Force -AcceptLicense; Import-Module IntuneWin32App -ErrorAction Stop }

# --- Pfade ---
$Root       = $PSScriptRoot
$Cmd        = Join-Path $Root 'Commands'
$IconDir    = Join-Path $Root 'Icon'
$SourceDir  = Join-Path $Root 'Source'
$OutDir     = Join-Path $Root 'Output'

# --- Pflichtdateien ---
$InstallPs1   = Join-Path $Cmd 'Install.ps1'
$UninstallPs1 = Join-Path $Cmd 'Uninstall.ps1'
$DetectPs1    = Join-Path $Cmd 'Detection.ps1'
$ParamFile    = Join-Path $Cmd 'App.params.psd1'
foreach ($f in @($InstallPs1,$UninstallPs1,$DetectPs1,$ParamFile)) { if (-not (Test-Path $f)) { throw "Fehlt: $f" } }

# --- Parameter laden ---
$app = Import-PowerShellDataFile -Path $ParamFile
$DisplayName       = $app.DisplayName       ; if ([string]::IsNullOrWhiteSpace($DisplayName))       { throw "Param 'DisplayName' fehlt" }
$Description       = $app.Description       ; if ([string]::IsNullOrWhiteSpace($Description))       { throw "Param 'Description' fehlt" }
$Publisher         = $app.Publisher         ; if ([string]::IsNullOrWhiteSpace($Publisher))         { throw "Param 'Publisher' fehlt" }
$AppVersion        = $app.Version           ; if ([string]::IsNullOrWhiteSpace($AppVersion))        { throw "Param 'Version' fehlt" }
$InstallExperience = $app.InstallExperience ; if ([string]::IsNullOrWhiteSpace($InstallExperience)) { $InstallExperience='system' }

# Gruppen-Parameter (Namen statt GUIDs)
$InstallGroupName   = $app.InstallGroupName
$InstallIntent      = if ($app.InstallIntent)      { $app.InstallIntent }      else { 'required' }  # 'available' oder 'required'
$UninstallGroupName = $app.UninstallGroupName
$UninstallIntent    = if ($app.UninstallIntent)    { $app.UninstallIntent }    else { 'uninstall' }
$ClearAssignments   = if ($app.RemoveExistingAssignments -is [bool]) { $app.RemoveExistingAssignments } else { $true }  # Standard: aufräumen

# Icon (optional)
$Icon = $null
$IconPath = Get-ChildItem -Path $IconDir -Filter *.png -File -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
if ($IconPath) { $Icon = [Convert]::ToBase64String([IO.File]::ReadAllBytes($IconPath)) }

# --- Build-Verzeichnisse ---
if (Test-Path $OutDir)    { Remove-Item $OutDir -Recurse -Force }
if (Test-Path $SourceDir) { Remove-Item $SourceDir -Recurse -Force }
New-Item -ItemType Directory -Path $OutDir,$SourceDir | Out-Null

# PS1 1:1 ins Paket kopieren
Copy-Item $InstallPs1   -Destination (Join-Path $SourceDir 'Install.ps1')   -Force
Copy-Item $UninstallPs1 -Destination (Join-Path $SourceDir 'Uninstall.ps1') -Force

# --- Intune verbinden (liefert Header für Graph) ---
$GraphHeader = Connect-MSIntuneGraph -TenantID $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
if ($GraphHeader) { $Global:AuthenticationHeader = $GraphHeader }

# --- Helper: GroupName -> GroupId (exakter Match, mit startsWith-Fallback) ---
function Resolve-GroupId {
  param([Parameter(Mandatory)][string]$GroupName)
  $escaped = $GroupName.Replace("'", "''")

  $uri1 = "https://graph.microsoft.com/v1.0/groups`?$filter=displayName eq '$escaped'&`$select=id,displayName"
  $res1 = Invoke-RestMethod -Uri $uri1 -Headers $Global:AuthenticationHeader -Method GET -ErrorAction Stop
  if ($res1.value -and $res1.value.Count -ge 1) { return $res1.value[0].id }

  $uri2 = "https://graph.microsoft.com/v1.0/groups`?$filter=startsWith(displayName,'$escaped')&`$select=id,displayName"
  $res2 = Invoke-RestMethod -Uri $uri2 -Headers $Global:AuthenticationHeader -Method GET -ErrorAction Stop
  if ($res2.value -and $res2.value.Count -ge 1) { return $res2.value[0].id }

  throw "Gruppe '$GroupName' nicht gefunden."
}

# --- .intunewin bauen ---
$SetupFile = 'Install.ps1'
New-IntuneWinPackage -SourcePath $SourceDir -SetupFile $SetupFile -DestinationPath $OutDir
$IntuneWinFile = Get-ChildItem -Path $OutDir -Recurse -Filter *.intunewin | Sort-Object LastWriteTime -Desc | Select-Object -First 1
if (-not $IntuneWinFile) { throw ".intunewin wurde nicht erstellt." }

# --- DetectionRule aus Script ---
$DetectionRule = New-IntuneWin32AppDetectionRuleScript -ScriptFile $DetectPs1 -EnforceSignatureCheck:$false -RunAs32Bit:$false

# --- Upload (einfach) ---
Write-Output "Adding app to Intune"
$InstallCommandLine   = 'powershell.exe -ExecutionPolicy Bypass -File Install.ps1'
$UninstallCommandLine = 'powershell.exe -ExecutionPolicy Bypass -File Uninstall.ps1'

$addParams = @{
  FilePath              = $IntuneWinFile.FullName
  DisplayName           = $DisplayName
  Description           = $Description
  Publisher             = $Publisher
  AppVersion            = $AppVersion
  InstallExperience     = $InstallExperience
  RestartBehavior       = 'suppress'
  DetectionRule         = $DetectionRule
  InstallCommandLine    = $InstallCommandLine
  UninstallCommandLine  = $UninstallCommandLine
  Verbose               = $true
}
if ($Icon) { $addParams.Icon = $Icon }

# >>> WICHTIG: Objekt direkt abgreifen
$Win32App = Add-IntuneWin32App @addParams
Write-Output "App added"

# Fallback: falls das Cmdlet aus deiner Modulversion nichts zurückgibt, bis zu 15x kurz nachfassen
if (-not $Win32App -or -not $Win32App.id) {
  Write-Output "Resolving created app by name..."
  for ($i=1; $i -le 15 -and (-not $Win32App -or -not $Win32App.id); $i++) {
    Start-Sleep -Seconds 2
    $Win32App = (Get-IntuneWin32App -DisplayName $DisplayName -Verbose |
                 Sort-Object createdDateTime -Descending | Select-Object -First 1)
  }
}
if (-not $Win32App -or -not $Win32App.id) {
  throw "App nicht gefunden (Get-IntuneWin32App -DisplayName '$DisplayName')."
}

# --- optional: vorhandene Assignments vorher löschen ---
if ($ClearAssignments) {
  Write-Output "Clearing existing assignments"
  Remove-IntuneWin32AppAssignment -Id $Win32App.id -ErrorAction SilentlyContinue | Out-Null
  Write-Output "Existing assignments cleared"
}

# --- Assignments (Gruppennamen -> ID -> zuweisen) ---
if ($InstallGroupName) {
  Write-Output "Assigning install group (by name)"
  $installId = Resolve-GroupId -GroupName $InstallGroupName
  Add-IntuneWin32AppAssignmentGroup -Include -Id $Win32App.id -GroupId $installId -Intent $InstallIntent -Notification "showAll" -Verbose
  Write-Output "Install group assigned: $InstallGroupName ($InstallIntent)"
}

if ($UninstallGroupName) {
  Write-Output "Assigning uninstall group (by name)"
  $uninstallId = Resolve-GroupId -GroupName $UninstallGroupName
  Add-IntuneWin32AppAssignmentGroup -Include -Id $Win32App.id -GroupId $uninstallId -Intent $UninstallIntent -Notification "showAll" -Verbose
  Write-Output "Uninstall group assigned: $UninstallGroupName ($UninstallIntent)"
}

