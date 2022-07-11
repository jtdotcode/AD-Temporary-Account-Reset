#Created by John Thompson under GPL license with no warranty or support use at your own risk.
# The accounts your want to reset they must be the AD SamAccount Name

# to send emails from this script you must have a valid password hash file
# this can be created by using the CreateNewPasswordHASH this must be done when moving the script to a new server

#test example $PSScriptRoot
# ./resetTempAccounts.ps1 -SmtpServer "smtp.gmail.com" -SentFrom "IT Department" -Logging -Testing -TestRecipientEmail "joetester@test.com" -OuSearchBase "OU=Casual Relief Staff,OU=Staff,DC=subdomain,DC=domain,DC=com" -BodyTemplateFileName "emailbodyTemplate.html" -SmtpUsername "user@gmail.com" -UseSSL -HashPasswordFileName "stmp_password_@device.txt" -CSVResetAccountsFileName "example-acccounts.csv"

#live example
# ./resetTempAccounts.ps1 -SmtpServer "smtp.gmail.com" -SentFrom "IT Department" -Logging -EmailRecipient "joetester@test.com" -OuSearchBase "OU=Casual Relief Staff,OU=Staff,DC=subdomain,DC=domain,DC=com" -BodyTemplateFileName "emailbodyTemplate.html" -SmtpUsername "user@gmail.com" -UseSSL -HashPasswordFileName "stmp_password_@device.txt" -CSVResetAccountsFileName "example-acccounts.csv"


param(
    # $smtpServer Enter Your SMTP Server Hostname or IP Address
    [Parameter(Mandatory = $True)]
    [string]$SmtpServer,
    [Parameter(Mandatory = $True)]
    [string]$SentFromText,
    [Parameter(Mandatory = $True)]
    [switch]$Logging,
    #[string]$LogFilePath = $PSScriptRoot,
    [switch]$Testing,
    [string]$TestRecipientEmail,
    [string]$ResetDay = "Sunday",
    [Int]$ResetFrequencyDays = 7,
    [Int]$TotalPasswordlength = 7,
    [Parameter(Mandatory = $True)]
    [string]$OuSearchBase,
    [Parameter(Mandatory = $True)]
    [string]$BodyTemplateFileName,
    [string]$EmailRecipient,
    [Parameter(Mandatory = $True)]
    [string]$SmtpUsername,
    [switch]$UseSSL,
    [string]$SmtpPort = 587,
    [Parameter(Mandatory = $True)]
    [string]$HashPasswordFileName,
    [string]$OverridePath,
    [Parameter(Mandatory = $True)]
    [string]$CSVResetAccountsFileName,
    [string]$CSVHeaderName = "Username"

)
###################################################################################################################

$ScriptRoot = $PSScriptRoot

if ($OverridePath) {
    $ScriptRoot = $OverridePath
}

Write-Host $ResetFrequencyDays

if ($Testing -and [string]::IsNullOrEmpty($TestRecipientEmail)) {
    throw "In testing mode -TestRecipientEmail cannot be empty"
}

$SmtpUser = $SmtpUsername

$passwordSymbols = @("!", "@", "#", "$", "%", "*")

$minPasswordLength = 4

#$csvFile = $CSVResetAccountsFilePath + "\" + $CSVResetAccountsFileName

$csvFile = Join-Path $ScriptRoot $CSVResetAccountsFileName

$hashFile = Join-Path $ScriptRoot $HashPasswordFileName

if (-not(Test-Path -Path $csvFile -PathType Leaf)) {

    throw "csv accounts file doesn't exist"
}

if (-not(Test-Path -Path $hashFile -PathType Leaf)) {

    throw "hash file doesn't exist"
}
else {

    try {
        Write-Host "checking hash file"
        Get-Content $hashFile | ConvertTo-SecureString -ErrorAction stop
       
    }
    catch {
       
        Write-Host "Unable to decrypt password hash file, please make sure it was exported for the user that is running this script"
        break
    
    }

}

    



Write-Host $hashFile


$accounts = Import-Csv -Path $csvFile



$SmtpHostCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SmtpUsername, (Get-Content $hashFile | ConvertTo-SecureString)

$smtpPort = $SmtpPort

$from = "$SentFromText <$SmtpUser>"

$toEmailAddress = $EmailRecipient

Write-Host $toEmailAddress

$textEncoding = [System.Text.Encoding]::UTF8

#$logFile = $LogFilePath + '\' + "log_file.txt"

$logFile = Join-Path $ScriptRoot "log_file.txt"

$dateTime = (get-date).AddDays($ResetFrequencyDays - 1).AddHours(11) | Get-Date -f g
$day = (Get-Date).AddDays($ResetFrequencyDay - 1).AddHours(11).DayOfWeek
$date = $day, $dateTime, -Join "" 


# $dateTime = (get-date).AddDays(7) | Get-Date -f g
# $day = (Get-Date).DayOfWeek
# $date = $day, $dateTime, -Join "" 

$messageNumber = 1

$bodyTemplate = Join-Path $ScriptRoot $BodyTemplateFileName

$randomFileName = -join ((65..90) + (97..122) | Get-Random -Count 12 | ForEach-Object { [char]$_ })

#$currentPathLocation = $(Get-Location).ToString()

$bodyOutFile = Join-Path $ScriptRoot "$randomFileName.html"




Function Get-RandomPassword {

    # this creates a sort of random string 
    # its limited one symbol, number and uppercase as the crt staff are incapable of anything else

    if ($TotalPasswordlength -le $minPasswordLength) {
    
        throw "min password character length exceeded"
    }

    $passwordLength = $TotalPasswordlength

    $upperCase = -join ((65..90) | Get-Random -Count 1 | ForEach-Object { [char]$_ })

    $passwordLength -= 1

    $symbol = @($passwordSymbols | Get-Random -Count 1 | ForEach-Object { [char]$_ })

    $passwordLength -= 1

    $number = -join ((48..57) | Get-Random -Count 1 | ForEach-Object { [char]$_ })

    $passwordLength -= 1

    $lowerCase = -join ((97..122) | Get-Random -Count $passwordLength | ForEach-Object { [char]$_ })

    $password = $lowerCase + $symbol + $number


    $newPassword = -join ($password.ToCharArray() | Get-Random -Count $password.Length) 

    $newPassword = $upperCase + $newPassword

    return $newPassword

}




foreach ($account in $accounts) {

    $csvUser = $account.$CSVHeaderName

    $user = Get-ADUser -SearchBase $OuSearchBase -Properties EmailAddress, GivenName, Surname, SamAccountName -Filter { sAMAccountName -eq $csvUser } 

    $userName = $user.SamAccountName
    $emailAddress = $user.EmailAddress
    
    
    # generate new random password for account
    $newPassword = Get-RandomPassword

  
    
    

    # check and see that the password have been generated and that the AD accounts match the accounts which are going to be changed

    if ( $newPassword.Length -eq $TotalPasswordlength -and $csvUser -eq $userName ) {
    
    (Get-Date).DateTime + " - Resetting password for - " + $user.SamAccountName | Out-File -FilePath $logFile -Append

        #change account password
    
        if ($testing) {
            Write-Host "*TestingMode* - Changing password for: $userName new password is: $newPassword - no password have been reset"
        }
        else {
            Set-ADAccountPassword -Identity $userName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $newPassword -Force) | Out-File -FilePath $logFile -Append
        
            Set-ADAccountExpiration -Identity $userName  -DateTime $dateTime
    
        }
    


        # subject for email
        $subject = "New details for $userName - Password expires on $date "

 

        # hash table of keys and values for replacement of variables in the html file bodyTemplate.html
        $varList = @{
            '[USERNAME]' = $userName
            '[PASSWORD]' = $newPassword
            '[EMAIL]'    = $emailAddress
        }


        try {

            # check if body template exists
            # if it doesnt exit loop

            if ( -Not (Test-Path $bodyTemplate)) {

                Write-Output "body template not found"

(Get-Date).DateTime + " - body template not found" | Out-File -FilePath $logFile -Append 

                throw "No body template please check path"


            }

            Copy-Item -Path $bodyTemplate -Destination $bodyOutFile

            $varList.keys | ForEach-Object {
                $message = '{0} = {1}' -f $_, $varList[$_]
                Write-Output $message
        (Get-Date).DateTime + " - " + $message | Out-File -FilePath $logFile -Append 

        
                #search html template for variables and replace with new account information


        (Get-Content -Path $bodyOutFile -Raw).Replace($_, $varList[$_]) | Set-Content $bodyOutFile

                $body = Get-Content -Path $bodyOutFile | Out-String

            }
            #end forEach Loop

        }
        catch {

            $errorMessage = $_.exception.Message
            if ($Logging) {

         (Get-Date).DateTime, $errorMessage | Out-File -FilePath $logFile -Append 

            }

        }
        # End try for html create 


        try {
    
            if ($testing) {
                $toEmailAddress = $TestRecipientEmail

                Write-Output "Sending Message" $messageNumber, "of" $accounts.Length 

                if ($UseSSL) {
                    Send-Mailmessage -smtpServer $smtpServer -from $from -to $TestRecipientEmail -subject $subject -body $body -bodyasHTML -priority High -Encoding $textEncoding -ErrorAction Stop -UseSsl -Port $smtpPort -Credential $SmtpHostCredential
                }

                if (!($UseSSL)) {
                    Write-Host "sending without ssl"

                    Send-Mailmessage -smtpServer $smtpServer -from $from -to $TestRecipientEmail -subject $subject -body $body -bodyasHTML -priority High -Encoding $textEncoding -ErrorAction Stop -Port $smtpPort -Credential $SmtpHostCredential
            
                }
       
        (Get-Date).DateTime + " - Sending Message " + $messageNumber + " of " + $accounts.Length  | Out-File -FilePath $logFile -Append 
            }

   
    
   

            if (!($testing)) {
                Write-Output "Sending Message" $messageNumber, "of" $accounts.Length 

                if ($UseSSL) {
                    Send-Mailmessage -smtpServer $smtpServer -from $from -to $toEmailAddress -subject $subject -body $body -bodyasHTML -priority High -Encoding $textEncoding -ErrorAction Stop -UseSsl -Port $smtpPort -Credential $SmtpHostCredential
                }
    
                if (!($UseSSL)) {
                    Send-Mailmessage -smtpServer $smtpServer -from $from -to $toEmailAddress -subject $subject -body $body -bodyasHTML -priority High -Encoding $textEncoding -ErrorAction Stop -Port $smtpPort -Credential $SmtpHostCredential
                }

       (Get-Date).DateTime + " - Sending Message " + $messageNumber + " of " + $accounts.Length  | Out-File -FilePath $logFile -Append 

            }
   
     

    
    
        }
        catch {

            $errorMessage = $_.exception.Message

            Write-Host $errorMessage
            if ($Logging) {
           

         (Get-Date).DateTime + " - " + $errorMessage | Out-File -FilePath $logFile -Append 
            }


        }
        #end try for email send 

        $messageNumber++ 
        

    }
    else {
    
    
        Write-host "Password length or account mismatch"

     (Get-Date).DateTime + " - " + "Password length or account mismatch" | Out-File -FilePath $logFile -Append 

    
    }
    # end if check for account and password mismatch



}
#end foreach loop 

#clean up

try {
    Remove-Item $bodyOutFile
}
catch {
        
        (Get-Date).DateTime + " - " + "Unable to delete bodyout put file" | Out-File -FilePath $logFile -Append 
}
    

 