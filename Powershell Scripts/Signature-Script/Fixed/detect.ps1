# Definiere den Registrierungspfad fÃ¼r Outlook-Profile.
$OutlookRegistryPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\Outlook";

# ÃœberprÃ¼fe, ob der Registrierungspfad existiert.
if (Test-Path $OutlookRegistryPath) {

    # ÃœberprÃ¼fe, ob die Standard-Signatur vorhanden ist.
    $defaultSignature = Get-ItemPropertyValue -Path "$OutlookRegistryPath\9375CFF0413111d3B88A00104B2A6676\00000002" -Name "New Signature" -ErrorAction SilentlyContinue;
    if ($defaultSignature) {
        Write-Output "Die Standard-Signatur ist vorhanden.";
        Exit 0
    }
    else {
        Write-Output "Die Standard-Signatur ist nicht vorhanden.";
        Exit 1
    }

}
else {
    Write-Output "Der Registrierungspfad existiert nicht.";
    Exit 0
}
