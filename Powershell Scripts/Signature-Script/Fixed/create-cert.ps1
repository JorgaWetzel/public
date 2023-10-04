#What is your Tenant URL
$turl = "customer.sharepoint.com"

# Login to Azure AD PowerShell With Admin Account
Connect-AzureAD 

# Create the self signed cert
$currentDate = Get-Date
$endDate = $currentDate.AddYears(1)
$notAfter = $endDate.AddYears(1)
$pwd = "P@ssw0rd"
$domain = $turl
$thumb = (New-SelfSignedCertificate -CertStoreLocation cert:\localmachine\my -DnsName $turl -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter $notAfter).Thumbprint
$pwd = ConvertTo-SecureString -String $pwd -Force -AsPlainText
Export-PfxCertificate -cert "cert:\localmachine\my\$thumb" -FilePath c:\temp\cert.pfx -Password $pwd


# Load the certificate
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate("C:\temp\cert.pfx", $pwd)
$keyValue = [System.Convert]::ToBase64String($cert.GetRawCertData())


# Create the Azure Active Directory Application
$application = New-AzureADApplication -DisplayName "pshellsignature" -IdentifierUris $turl
New-AzureADApplicationKeyCredential -ObjectId $application.ObjectId -CustomKeyIdentifier "pshellsignature" -StartDate $currentDate -EndDate $endDate -Type AsymmetricX509Cert -Usage Verify -Value $keyValue

# Create the Service Principal and connect it to the Application
$sp=New-AzureADServicePrincipal -AppId $application.AppId

# Give the Service Principal Reader access to the current tenant (Get-AzureADDirectoryRole)
#Add-AzureADDirectoryRoleMember -ObjectId 2bcf4db5-eac2-4ff6-9940-a3e692f9f2f1 -RefObjectId $sp.ObjectId

# Give the Service Principal Reader access to the current tenant (Get-AzureADDirectoryRole)
$roleId = (Get-AzureADDirectoryRole |? {$_.DisplayName -eq "Directory Readers"}).ObjectId
Add-AzureADDirectoryRoleMember -ObjectId $roleId -RefObjectId $sp.ObjectId

##################################
### Print the secret to upload to user secrets or a key vault
##################################
Write-Host 'Service principal:'
Write-Host $application.ObjectId
Write-Host '-----'
Write-Host 'CertificateThumbprint:'
Write-Host $thumb
Write-Host '-----'