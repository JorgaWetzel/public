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
  param([Parameter(Mandatory)][string]$GroupName, [string]$PreferredId)

  # Joker: Wenn im Param-File eine GUID steht, die direkt verwenden
  if ($PreferredId) { return $PreferredId }

  # Basis-Header aus Connect-MSIntuneGraph
  $hdr = @{}
  $Global:AuthenticationHeader.GetEnumerator() | ForEach-Object { $hdr[$_.Key] = $_.Value }

  # 1) Simple OData-Filter (kein $count) -> robust gegen 400er
  $nameEsc = $GroupName.Replace("'","''")  # OData-Escaping
  $select  = "id,displayName,securityEnabled,mailEnabled,groupTypes"
  $uri     = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$nameEsc'&`$select=$select&`$top=999"

  try {
    $res = Invoke-RestMethod -Uri $uri -Headers $hdr -Method GET -ErrorAction Stop
  } catch {
    # 2) Fallback: $search (braucht ConsistencyLevel)
    $hdr['ConsistencyLevel'] = 'eventual'
    $uri = "https://graph.microsoft.com/v1.0/groups?`$search=`"displayName:$nameEsc`"&`$select=$select&`$top=999"
    $res = Invoke-RestMethod -Uri $uri -Headers $hdr -Method GET -ErrorAction Stop
  }

  $items = @($res.value)
  if ($items.Count -eq 0) { throw "Gruppe '$GroupName' nicht gefunden." }

  # Bevorzugung: klassische Security-Groups, nicht mail-enabled, nicht dynamisch
  $pref = $items | Where-Object {
    $_.securityEnabled -eq $true -and $_.mailEnabled -eq $false -and ($_.groupTypes -notcontains 'DynamicMembership')
  }

  if ($pref.Count -eq 1) {
    Write-Host "• Gruppenauflösung: '$($pref[0].displayName)' -> $($pref[0].id)"
    return $pref[0].id
  }

  if ($items.Count -eq 1) {
    Write-Host "• Gruppenauflösung: '$($items[0].displayName)' -> $($items[0].id)"
    return $items[0].id
  }

  # exakte Groß/Kleinschreibung als weiterer Versuch
  $case = $pref | Where-Object { $_.displayName -ceq $GroupName }
  if ($case.Count -eq 1) { return $case[0].id }

  # Mehrdeutig -> Kandidaten zeigen und abbrechen
  Write-Warning "Mehrere Gruppen mit '$GroupName' gefunden. Bitte im Param-File InstallGroupId setzen."
  ($pref + $items | Select-Object -Unique) |
    Select-Object @{n='id';e={$_.id}}, @{n='name';e={$_.displayName}},
                  @{n='sec';e={$_.securityEnabled}}, @{n='mail';e={$_.mailEnabled}},
                  @{n='types';e={$_.groupTypes -join ','}} |
    Format-Table -AutoSize | Out-String | Write-Host
  throw "Auflösung von '$GroupName' ist mehrdeutig."
}

# Available für alle Benutzer (optional, aus Param-Datei gesteuert)
function Set-AvailableToAllUsers {
  param([Parameter(Mandatory)][string]$AppId, [string]$Notification = 'showAll')

  $hdr  = $Global:AuthenticationHeader
  $base = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps/$AppId/assignments"

  # exists?
  $ex = Invoke-RestMethod -Uri $base -Headers $hdr -Method GET -ErrorAction Stop
  $exists = $false
  foreach ($a in $ex.value) {
    if ($a.intent -eq 'available' -and $a.target.'@odata.type' -like '*allLicensedUsersAssignmentTarget') { $exists = $true }
  }

  if (-not $exists) {
    $body = @{
      "@odata.type" = "#microsoft.graph.mobileAppAssignment"
      intent        = "available"
      target        = @{ "@odata.type" = "#microsoft.graph.allLicensedUsersAssignmentTarget" }
      settings      = @{
        "@odata.type" = "#microsoft.graph.win32LobAppAssignmentSettings"
        notifications = $Notification
        deliveryOptimizationPriority = "notConfigured"
      }
    } | ConvertTo-Json -Depth 5

    Invoke-RestMethod -Uri $base -Headers $hdr -Method POST -ContentType "application/json" -Body $body -ErrorAction Stop | Out-Null
    Write-Host "• 'Available' an Alle Benutzer hinzugefügt"
  } else {
    Write-Host "• 'Available' an Alle Benutzer bereits vorhanden – übersprungen"
  }
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
Write-Output "Assigning install group (by name)"
$installId = if ($app.InstallGroupId) { $app.InstallGroupId } else { Resolve-GroupId -GroupName $InstallGroupName }

# PS 5.1-kompatibel: Fallback auf 'showAll', wenn nichts gesetzt
$installNotif = 'showAll'
if ($app.ContainsKey('InstallNotification') -and -not [string]::IsNullOrWhiteSpace($app.InstallNotification)) {
    $installNotif = $app.InstallNotification
}

Add-IntuneWin32AppAssignmentGroup -Include -Id $Win32App.id -GroupId $installId -Intent $InstallIntent -Notification $installNotif -Verbose


if ($UninstallGroupName) {
  Write-Output "Assigning uninstall group (by name)"
  $uninstallId = Resolve-GroupId -GroupName $UninstallGroupName
  Add-IntuneWin32AppAssignmentGroup -Include -Id $Win32App.id -GroupId $uninstallId -Intent $UninstallIntent -Notification "showAll" -Verbose
  Write-Output "Uninstall group assigned: $UninstallGroupName ($UninstallIntent)"
}

# Required an DEV-WIN-Standard ist bereits gesetzt ...
# --- Assignments (nach deinem Required an DEV-WIN-Standard) ---

# Optional: Available für alle Benutzer
$addAvail = $false
if ($app -and $app.ContainsKey('AvailableToAllUsers')) { $addAvail = [bool]$app.AvailableToAllUsers }

if ($addAvail) {
  $notif = 'showAll'
  if ($app.ContainsKey('AvailableNotification') -and -not [string]::IsNullOrWhiteSpace($app.AvailableNotification)) {
    $notif = $app.AvailableNotification
  }
  Set-AvailableToAllUsers -AppId $Win32App.id -Notification $notif
}
