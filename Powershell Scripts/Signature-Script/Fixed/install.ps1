<#PSScriptInfo
.VERSION 1.0
.GUID 729ebf90-26fe-4795-92dc-ca8f570cdd22
.AUTHOR AndrewTaylor
.DESCRIPTION Sets user outlook signature using details from AzureAD
.COMPANYNAME 
.COPYRIGHT GPL
.TAGS intune outlook signature azureAD
.LICENSEURI https://github.com/andrew-s-taylor/public/blob/main/LICENSE
.PROJECTURI https://github.com/andrew-s-taylor/public
.ICONURI 
.EXTERNALMODULEDEPENDENCIES azureAD
.REQUIREDSCRIPTS 
.EXTERNALSCRIPTDEPENDENCIES 
.RELEASENOTES
#>
<#
.SYNOPSIS
  Creates user outlook signature from hosted template using Azure AD details
.DESCRIPTION
Creates user outlook signature from hosted template using Azure AD details

.INPUTS
None required
.OUTPUTS
Within Azure
.NOTES
  Version:        1.0
  Author:         Andrew Taylor
  Twitter:        @AndrewTaylor_2
  WWW:            andrewstaylor.com
  Creation Date:  05-11-2021
  Purpose/Change: Initial script development
  
.EXAMPLE
N/A
#>

#################################################### EDIT THESE SETTINGS ####################################################

#Custom variables
$certurl = "https://github.com/andrew-s-taylor/public/raw/main/Powershell%20Scripts/Signature-Script/template.docx"
$templateurl = "https://github.com/andrew-s-taylor/public/raw/main/Powershell%20Scripts/Signature-Script/template.docx"
$SignatureName = 'Staefa-Hombrechtikon' #insert the company name (no spaces) - could be signature name if more than one sig needed
$SignatureVersion = "1" #Change this if you have updated the signature. If you do not change it, the script will quit after checking for the version already on the machine
$ForceSignature = '0' #Set to 1 if you don't want the users to be able to change signature in Outlook
$Address =      ""
$Tel =          ""
$TenantId     = "c77ac78f-946a-4a81-bb8a-94e6a56b0b0d"
$ClientId     = "6d355681-f1fe-4156-bd0f-94e843f0f6d4"
$thumb        = "1F474441752CC5F232044E9AACAADD3FF70DE959"
$pwd          = "SuperS3cur3P@ssw0rd"


#################################################### DO NOT EDIT BELOW THIS LINE ####################################################

# Alle Ordner auf Deutsch
<#
Start-Process "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" -ArgumentList "/resetfoldernames"
Start-Sleep -Seconds 4
Stop-process -name OUTLOOK -Force
#>
Copy-Item .\cert.pfx C:\Service\cert.pfx
$certPath = "C:\Service\cert.pfx"
$certPassword = ConvertTo-SecureString -String $pwd -Force -AsPlainText

# Importiere das Zertifikat in den "Eigene Zertifikate"-Speicher
Import-PfxCertificate -FilePath $certPath -CertStoreLocation Cert:\CurrentUser\My -Password $certPassword
Remove-Item $certPath
Connect-AzureAD -TenantId $TenantId -ApplicationId  $ClientId -CertificateThumbprint $thumb


$appdataf = $env:APPDATA
$outlookf = $appdataf+"\Microsoft"

if((Test-Path -Path $outlookf )){
    #New-Item -ItemType directory -Path $TARGETDIR

#Environment variables
$AppData=(Get-Item env:appdata).value
$SigPath = '\Microsoft\Signatures' #Path to signatures. Might need to be changed to correspond language, i.e Swedish '\Microsoft\Signaturer'
$LocalSignaturePath = $AppData+$SigPath
$SigSource = $SignatureName+'.docx' #Path to the *.docx file, i.e "c:\temp\template.docx"
$RemoteSignaturePathFull = $SigSource

#Copy version file
If (-not(Test-Path -Path $LocalSignaturePath\$SignatureVersion))
{
New-Item -Path $LocalSignaturePath\$SignatureVersion -ItemType Directory
$path = "C:\templates"
If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
}

#Create Temp folder to store template
$random = Get-Random -Maximum 1000 
$random = $random.ToString()
$date =get-date -format ddMMyymmss
$date = $date.ToString()
$path2 = $random + "-"  + $date
$path = "c:\temp\" + $path2 + "\"

New-Item -ItemType Directory -Path $path


##Download Template File

$output = $path + "template.docx"
Invoke-WebRequest -Uri $templateurl -OutFile $output -Method Get


#Download Certificate
$output = $path + "cert.pfx"
Invoke-WebRequest -Uri $certurl -OutFile $output -Method Get

$sourceDirectory  = $output



$destinationDirectory = "C:\templates\"
Copy-item -Force -Recurse -Verbose $sourceDirectory -Destination $destinationDirectory
}
Elseif (Test-Path -Path $LocalSignaturePath\$SignatureVersion)
{
Write-Output "Latest signature already exists"
break
}

#Check signature path (needs to be created if a signature has never been created for the profile
if (-not(Test-Path -path $LocalSignaturePath)) {
	New-Item $LocalSignaturePath -Type Directory
}

$currentuser = $env:userdnsdomain + "@" + $env:username
get-azureaduser -Filter "UserPrincipalName eq '$currentuser'"

$userPrincipalName = whoami -upn
$userObject = Get-AzureADUser -ObjectId $userPrincipalName


#Get AzureAD information for current user
$DisplayName = $userObject.DisplayName
$Fax = $userObject.Fax
$Personal = $ADUser.PersonalTitle
$Surname = $userObject.Surname
$FirstName = $userObject.GivenName
$EmailAddress = $userObject.Mail
$Title = $userObject.JobTitle
$Description = $userObject.description
$TelePhoneNumber = $userObject.TelephoneNumber
$Mobile = $userObject.Mobile
$StreetAddress = $userObject.StreetAddress
$City = $userObject.City
$CustomAttribute1 = $userObject.extensionAttribute1

{

function RemoveLineCarriage($object)
{
	$result = [System.String] $object;
	$result = $result -replace "`t","";
	$result = $result -replace "`n","";
	$result = $result -replace "`r","";
	$result = $result -replace " ;",";";
	$result = $result -replace "; ",";";
	
	$result = $result -replace [Environment]::NewLine, "";
	
	$result;
}


$fullname = $FirstName.TrimEnd("`r?`n")+" "+$Surname.TrimEnd("`r?`n")        
$signame2 = $Fax

$signame2 = $signame2.Substring(0,$signame2.Length-6)
$signame3 = $signame2.Split("{,}")

$signame4 = $signame3[1]
$signame4 = $signame4.Substring(1)
$signame5 = $signame3[0]
$fullname = $signame4+" "+$signame5
}


#Copy signature templates from source to local Signature-folder
Write-Output "Copying Signatures"
Copy-Item "$Sigsource" $LocalSignaturePath -Recurse -Force
$ReplaceAll = 2
$FindContinue = 1
$MatchCase = $False
$MatchWholeWord = $True
$MatchWildcards = $False
$MatchSoundsLike = $False
$MatchAllWordForms = $False
$Forward = $True
$Wrap = $FindContinue
$Format = $False
	
#Insert variables from Active Directory to rtf signature-file
$MSWord = New-Object -ComObject word.application
$fullPath = $LocalSignaturePath+'\'+$SignatureName+'.docx'
$MSWord.Documents.Open($fullPath)
	
#User Name $ Designation 
$FindText = "DisplayName" 
$Designation = $CustomAttribute1 #designations in Exchange custom attribute 1
If ($Designation -ne '') { 
	$Name = $Fax
	$ReplaceText = $Name+', '+$Designation
}
Else {
	$ReplaceText = $Fax
}
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)	

#Title		
$FindText = "Title"
$ReplaceText = $Title
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
	
#Description
   If ($Description -ne '') { 
   	$FindText = "Description"
   	$ReplaceText = $Title
}
Else {
	$FindText = " | Description "
   	$ReplaceText = ""
}
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
#$LogInfo += $NL+'Description: '+$ReplaceText


#Description

	$FindText = "TelephoneNumber"
   	$ReplaceText = $TelePhoneNumber

$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
#$LogInfo += $NL+'Description: '+$ReplaceText
   	
#Street Address
If ($StreetAddress -ne '') { 
       $FindText = "StreetAddress"
    $ReplaceText = $StreetAddress
   }
   Else {
    $FindText = "StreetAddress"
    $ReplaceText = $DefaultAddress
    }
	$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)

#City
If ($City -ne '') { 
    $FindText = "City"
       $ReplaceText = $City
   }
   Else {
    $FindText = "City"
    $ReplaceText = $DefaultCity 
   }
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)


#Email
$MSWord.Selection.Find.Execute("E-Mail")
$MSWord.ActiveDocument.Hyperlinks.Add($MSWord.Selection.Range, "mailto:"+$ADEmailAddress, $missing, $missing, "E-Mail")

If ($EmailAddress -ne '') { 
    $FindText = "E-Mail"
       $ReplaceText = $EmailAddress
   }
   Else {
    $FindText = "E-Mail"
    $ReplaceText = ""
   }

$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)

#Fullname
If ($fullname -ne '') { 
    $FindText = "FullName"
       $ReplaceText = $DisplayName
   }
   Else {
    $FindText = "FullName"
    $ReplaceText = "" 
   }
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)



#FirstName
If ($FirstName -ne '') { 
        $FindText = "FirstName"
       $ReplaceText = $forename
   }
   Else {
    $FindText = "FirstName"
    $ReplaceText = "" 
   }
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)

#Surname
If ($Surname -ne '') { 
    $FindText = "Surname"
       $ReplaceText = $Surname
   }
   Else {
    $FindText = "Surname"
    $ReplaceText = "" 
   }
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
	
#Telephone
If ($TelephoneNumber -ne "") { 
	$FindText = "TelephoneNumber"
	$ReplaceText = " | Ext:  "+$TelephoneNumber
   }
Else {
	$FindText = "TelephoneNumber"
    $ReplaceText = $DefaultTelephone
	}
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
	
#Mobile
If ($Mobile -ne "") { 
	$FindText = "MobileNumber"
	$ReplaceText = $Mobile
   }
Else {
	$FindText = "| Mob MobileNumber "
    $ReplaceText = ""
	}
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)

#Save new message signature 
Write-Output "Saving signatures"
#Save HTML
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatHTML");
$path = $LocalSignaturePath+'\'+$SignatureName+" ("+ $userPrincipalName +").htm"
$MSWord.ActiveDocument.saveas([ref]$path, [ref]$saveFormat)
    
#Save RTF 
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatRTF");
$path = $LocalSignaturePath+'\'+$SignatureName+" ("+ $userPrincipalName +").rtf"
$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$saveFormat)
	
#Save TXT    
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatText");
$path = $LocalSignaturePath+'\'+$SignatureName+" ("+ $userPrincipalName +").txt"
$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$SaveFormat)
$MSWord.ActiveDocument.Close()
$MSWord.Quit()
	

#Office 2010
If (Test-Path HKCU:'\Software\Microsoft\Office\14.0')
{
If ($ForceSignature -eq '1')
    {
    Write-Output "Setting signature for Office 2010 as forced"
    New-ItemProperty HKCU:'\Software\Microsoft\Office\14.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName -PropertyType 'String' -Force
    New-ItemProperty HKCU:'\Software\Microsoft\Office\14.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force
    }
else
{
Write-Output "Setting Office 2010 signature as available"

$MSWord = New-Object -comobject word.application
$EmailOptions = $MSWord.EmailOptions
$EmailSignature = $EmailOptions.EmailSignature
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries

}
}



#Office 2013 signature

If (Test-Path HKCU:Software\Microsoft\Office\15.0)

{
Write-Output "Setting signature for Office 2013"

If ($ForceSignature -eq '0')

{
Write-Output "Setting Office 2013 as available"

$MSWord = New-Object -ComObject word.application
$EmailOptions = $MSWord.EmailOptions
$EmailSignature = $EmailOptions.EmailSignature
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries

}

If ($ForceSignature -eq '1')
{
Write-Output "Setting signature for Office 2013 as forced"
    If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings') { } 
    Else { 
    New-ItemProperty HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force 
    } 
    If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings') { } 
    Else { 
    New-ItemProperty HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName -PropertyType 'String' -Force
    } 
        If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings') { } 
    Else { 
    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force 
    } 
    If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings') { } 
    Else { 
    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName -PropertyType 'String' -Force
    } 
}
}

#Office 2016 signature

If (Test-Path HKCU:Software\Microsoft\Office\16.0)

{
Write-Output "Setting signature for Office 2016"

If ($ForceSignature -eq '0')

{
Write-Output "Setting Office 2016 as available"

$MSWord = New-Object -ComObject word.application
$EmailOptions = $MSWord.EmailOptions
$EmailSignature = $EmailOptions.EmailSignature
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries


Function Get-OutlookProfiles
{
    [cmdletbinding()]

    Param
    (
        [Parameter(Mandatory=$true)][string]$Path
    )
    # Get Outlook profiles names.
    $OutlookDefaultProfiles = (Get-ChildItem -Path $Path).PSChildName;
    # More than one profile.
    If($OutlookDefaultProfiles -eq $null -or $OutlookDefaultProfiles.Count -ne 1)
    {
        # Set profile path.
        $OutlookProfilePath = 'HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings';
    }
    # Default profile.
    Else
    {
        # Set default profile path.
        $OutlookProfilePath = ("HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\{0}\9375CFF0413111d3B88A00104B2A6676\00000002" -f $OutlookDefaultProfiles);
    }
    # Return path.
    Return $OutlookProfilePath;
}


            
$OutlookRegistryPath = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles';

# Get Outlook profile registry path.
$OutlookProfilePath = Get-OutlookProfiles -Path $OutlookRegistryPath;
$OutlookNewSignature = $SignatureName+" ("+ $userPrincipalName +")"
$SignatureNewName   = $SignatureName+" ("+ $userPrincipalName +")"
$SignatureReplyName = $SignatureName+" ("+ $userPrincipalName +")"
#$SignatureNewName   = "$OutlookNewSignature ($CurrentUser)"
#$SignatureReplyName = "$OutlookNewSignature ($CurrentUser)"
Get-Item -Path $OutlookProfilePath | New-Itemproperty -Name "New Signature" -value $SignatureNewName -Propertytype string -Force | Out-Null;
Get-Item -Path $OutlookProfilePath | New-Itemproperty -Name "Reply-Forward Signature" -value $SignatureReplyName -Propertytype string -Force | Out-Null;
}

If ($ForceSignature -eq '1')
{
Write-Output "Setting signature for Office 2016 as forced"
    If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings') { } 
    Else { 
    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force 
    } 
    If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings') { } 
    Else { 
    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName -PropertyType 'String' -Force
    } 
        If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings') { } 
    Else { 
    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force 
    } 
    If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings') { } 
    Else { 
    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName -PropertyType 'String' -Force
    } 
}
}

}

Write-Host "" 
Write-Host "Daten:" 
Write-Host "DisplayName:" $DisplayName 
Write-Host "Fax:" $Fax
Write-Host "Personal:" $Personal
Write-Host "Surname:" $Surname
Write-Host "FirstName:" $FirstName
Write-Host "EmailAddres:" $EmailAddress
Write-Host "Title:" $Title
Write-Host "Description:" $Description
Write-Host "TelePhoneNumber:" $TelePhoneNumber
Write-Host "Mobile:" $Mobile
Write-Host "StreetAddress:" $StreetAddress
Write-Host "City:" $City
Write-Host "CustomAttribute1:" $CustomAttribute1