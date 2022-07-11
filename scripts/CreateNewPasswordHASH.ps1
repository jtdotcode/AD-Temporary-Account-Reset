# This script is used for creating a hashed password file for the smtp account for sending temp account reset emails. 
# Use this script when you change smtp account passwords or move scripts to a new server.
# The output file doesnt contain account information only a hashed password, so when using the hashed out you must make sure the username matches the password that was hashed. 
# script use example 
#.\CreateNewPasswordHASH.ps1 -SmtpUsername "user@mail.com.au" -FileName "stmp_password_@yourdomain.com" -TestRecipientEmailAddress "joetester@test.com.au" -SmtpHostAddress "smtp.gmail.com"

[CmdletBinding()]
param (
    [Parameter(Mandatory=$True)]
    [string]$SmtpUsername = $(throw "-smtp username is required."),
    [SecureString]$SmtpPassword = $( Read-Host -asSecureString "Input password for Smtp Account" ),
    [string]$OutputPath = $(Get-Location),
    [string]$FileName = "smtp_password_hash.txt",
    [Parameter(Mandatory=$True)]    
    [string]$TestRecipientEmailAddress,
    [Parameter(Mandatory=$True)]
    [string]$SmtpHostAddress
)

$hashFile = Join-Path $OutputPath $FileName

# export the password hash file to the scripts folder

$smtpPassword | ConvertFrom-SecureString | Out-File $hashFile

# test the password hash by sending a file.

# check if file exists 
if (-not(Test-Path -Path $hashFile -PathType Leaf)) {
    throw "Something went wrong, unable to find hash file"
}

try{
    Write-Host "checking hash file"
    Get-Content $hashFile | ConvertTo-SecureString -ErrorAction stop
   
   }catch{
   
   Write-Host "Unable to decrypt password hash file, please make sure it was exported for the user that is running this script"
   break

   }


Write-Host "Creating Credential Object.."
$SmtpHostCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SmtpUsername, (Get-Content $hashFile | ConvertTo-SecureString)

Write-Host "Sending Test Email.."
try {
    Send-Mailmessage -smtpServer $SmtpHostAddress -From "IT Support <$SmtpUsername>" -To $TestRecipientEmailAddress -subject "test for password hash" -body "if you have recieved this the hashed password is working" -UseSsl -Port 587 -Credential $SmtpHostCredential
}
catch {

    Write-Host $_.exception.Message
    
}

