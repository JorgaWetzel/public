<# Intune-Upload für Andrew-Taylor-"Deploy-Application"-Ordner (z. B. 7-Zip)

Voraussetzungen:
- App-Registrierung mit Application Permission: DeviceManagementApps.ReadWrite.All (Admin-Consent).
- Zertifikat in CurrentUser/LocalMachine Zertifikatsspeicher (Thumbprint).
- PowerShell 7+ empfohlen.

Das Script:
1) Baut .intunewin mit SvRooij.ContentPrep.Cmdlet
2) Erstellt Win32LobApp (IntuneWin32App-Modul)
3) Legt (optional) eine "required"-Zuweisung auf eine AAD-Gruppe an
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory)] [string] $AppFolder,
  [Parameter(Mandatory)] [string] $TenantId,
  [Parameter(Mandatory)] [string] $ClientId,

  # Eingabe als Thumbprint ODER direkt als Zertifikatsobjekt
  [Alias('CertificateThumbprint','Thumbprint')]
  [string] $CertThumbprint,
  [Alias('ClientCertificate','Cert','Certificate')]
  [System.Security.Cryptography.X509Certificates.X509Certificate2] $ClientCert,

  [string] $DisplayName = "7-Zip (Deploy-Application)",
  [string] $Publisher   = "7-Zip",
  [string] $Description = "Bereitgestellt via Deploy-Application-Vorlage",
  [ValidateSet('system','user')] [string] $RunAs = 'system',
  [string] $AssignGroupDisplayName
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# --- Zertifikat auflösen (Thumbprint -> Objekt) ---
if (-not $ClientCert) {
  if (-not $CertThumbprint) { throw "Bitte -CertThumbprint oder -ClientCert angeben." }
  $ClientCert = Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My |
                Where-Object Thumbprint -eq $CertThumbprint |
                Select-Object -First 1
  if (-not $ClientCert) { throw "Zertifikat mit Thumbprint $CertThumbprint nicht gefunden." }
}

# --- Module laden (nur die nötigen) ---
$mods = @(
  'IntuneWin32App',              # Upload/Rules/Assignments
  'SvRooij.ContentPrep.Cmdlet',  # .intunewin bauen
  'MSAL.PS'                      # wird von Connect-MSIntuneGraph genutzt
)
foreach($m in $mods){
  if(-not (Get-Module -ListAvailable -Name $m)){ Install-Module $m -Scope CurrentUser -Force -AllowClobber }
  Import-Module $m -ErrorAction Stop
}

# --- AUTH: exakt EIN Parameterset → ClientCert (kein Secret/DeviceCode/Interactive) ---
# Baut $Global:AuthenticationHeader mit gültigem ExpiresOn, wie vom Modul erwartet.
Connect-MSIntuneGraph -TenantId $TenantId -ClientId $ClientId -ClientCert $ClientCert -Refresh | Out-Null
Test-AccessToken | Out-Null   # optionaler Check – wirft sonst die „expired“-Warnung

# --- .intunewin bauen (SetupFile = Deploy-Application.ps1)
$Out = Join-Path $env:TEMP ("out_" + [guid]::NewGuid().ToString('N'))
New-Item $Out -ItemType Directory | Out-Null
New-IntuneWinPackage -SourcePath $AppFolder -SetupFile "Deploy-Application.ps1" -DestinationPath $Out | Out-Null   # ContentPrep-Cmdlet
$intuneWin = Get-ChildItem $Out -Filter *.intunewin | Select-Object -First 1

# --- Detection Rule
# Falls ein vorhandenes detect.ps1 im Ordner liegt, nimm das; sonst Registry-basierte 7-Zip-Erkennung
$detectPs1 = Join-Path $AppFolder 'detect.ps1'
if(-not (Test-Path $detectPs1)){
  @'
# Fallback-Detection: prüft 7-Zip-Registrypfad & EXE
$ok = $false
try{
  $path = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\7-Zip" -Name "Path" -ErrorAction Stop
  if($path -and (Test-Path (Join-Path $path "7z.exe"))){ $ok = $true }
}catch{}
exit ($(if($ok){0}else{1}))
'@ | Set-Content -Path ( $detectPs1 ) -Encoding UTF8
}
$DetectionRule   = New-IntuneWin32AppDetectionRuleScript -ScriptFile $detectPs1 -EnforceSignatureCheck:$false -RunAs32Bit:$false
$RequirementRule = New-IntuneWin32AppRequirementRule -Architecture "All" -MinimumSupportedWindowsRelease "W10_20H2"


# --- Install-/Uninstall-Command (Deploy-Application-Vorlage)
$InstallCmd   = 'powershell.exe -ExecutionPolicy Bypass -NoProfile -File .\Deploy-Application.ps1 -DeploymentType Install -DeployMode Silent'
$UninstallCmd = 'powershell.exe -ExecutionPolicy Bypass -NoProfile -File .\Deploy-Application.ps1 -DeploymentType Uninstall -DeployMode Silent'


# --- Upload als Win32LobApp
$app = Add-IntuneWin32App -FilePath $intuneWin.FullName `
        -DisplayName $DisplayName `
        -Description $Description `
        -Publisher $Publisher `
        -InstallExperience $RunAs `
        -RestartBehavior 'suppress' `
        -DetectionRule $DetectionRule `
        -RequirementRule $RequirementRule `
        -InstallCommandLine $InstallCmd `
        -UninstallCommandLine $UninstallCmd `
        -Verbose

$AppId = $app.id
Write-Host "App erstellt: $DisplayName  (Id: $AppId)"

# --- Optional: Assignment (required) für Gruppe
if($AssignGroupDisplayName){
  Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $cert.Thumbprint -NoWelcome | Out-Null
  $grp = Get-MgGroup -Filter "displayName eq '$($AssignGroupDisplayName.Replace("'","''"))'" -ConsistencyLevel eventual -Count cnt
  if($grp){
    $target   = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = $grp[0].Id }
    $settings = @{ '@odata.type' = '#microsoft.graph.win32LobAppAssignmentSettings' }
    New-MgDeviceAppManagementMobileAppAsWin32LobAppAssignment -MobileAppId $AppId -Intent 'required' -Target $target -Settings $settings | Out-Null
    Write-Host "Assignment erstellt → $AssignGroupDisplayName (required)"
  } else {
    Write-Warning "Gruppe '$AssignGroupDisplayName' nicht gefunden – Assignment übersprungen."
  }
}

# Cleanup
Remove-Item -LiteralPath $Out -Recurse -Force -ErrorAction SilentlyContinue
