#these can be set by command line, but defaults can also be changes
param (
        
        [int]$firstnotifydays = 14, #when to send out the first notification
        [int]$dailynotifydays = 5, #when to start sending daily notifications
        [string]$mailserver = '0.0.0.0',  #SMTP relay address
        [int]$mailserverport = 25,  #SMTP relay port (probably 25)
        [string]$mailfrom = 'no-reply@contoso.com',  #notification email sent from address
        [string]$statusto = 'gary@contoso.com',  #address to send stats to
        [string]$csvbasename = 'UsrExpiringPass',
        [switch]$nostatusemail, #prevents status email from being sent
        [switch]$sendemails,  #you must specify on command line to send out emails, this prevents oopses
        [string]$dateformat = 'M/d/yyyy hh:mm tt', #date format for emails and console output
        [string]$attachment = '.\HowToChangeMyPassword.pdf' #set the attachment to send
        )

$ErrorActionPreference = 'stop'


Import-Module ActiveDirectory



#list of OUs to check
$userOUs = @(
                "DC=contoso,DC=corp"
            )






$todaysdate = (Get-Date -Hour 0 -Minute 0 -Second 0 -Millisecond 0)

$firstnotifydate = $todaysdate.AddDays($firstnotifydays + 1)  #date to do first notifcation on
$nonotifybefore = $todaysdate.AddDays($firstnotifydays)  #end range for first notification

$dailynotifydate = $todaysdate.AddDays($dailynotifydays + 1) #date to do daily notifications on

$totaluserscount = 0  #total users with expiring passwords for stats
$expireduserscount = 0 #total users with already expired passwords
$nonotifyuserscount = 0 #total excluded from the notifications for expiring to soon
$notifyusers = @()  #array containing users objects


#loop through all OUs in array $userOUs
foreach ($OU in $userOUs) {

    ""

    "Getting:  " + $OU
    $users = Get-ADUser -SearchBase $OU -Filter "(Enabled -eq 'True') -And (PasswordNeverExpires -eq 'False')" -Properties EmailAddress,msDS-UserPasswordExpiryTimeComputed,Title,CanonicalName
    
    
    #Sorts out special cases like null value and single user returned instead of array these problems throw errors when running $users.Count if not resolved
    if (-not $users)  {
        "Count:          0 group empty"
    } elseif(($users.GetType()).name -eq 'ADUser') {
        "Count:          1"
        $totaluserscount += 1
    } else {
        "Count:          " + $users.Count
        $totaluserscount += $users.Count
    }


    $ExpiringInOUcount = 0 #calculates how many are expiriong in each OU, good for sanity checks
    $ExpiredInOUcount = 0 #already expired in the OU
    $nonotifyusersInOUCount = 0 #users in OU who wont be notified because they are expiring too soon
    $notexpiringinOU = 0 #A count of users for which no action is taken, as sanity check

    #loop through users for each OU
    foreach ($user in $users ){
        try {
            $PasswordExpire = (Get-Date $user.'msDS-UserPasswordExpiryTimeComputed').AddYears(1600)  #calculate the users password expiration date, add 1600 years because AD date format weirdness
        } catch {
            "Calculating expiration for user " + $user.name + " failed: " + $_.Exception.Message + "`n"
            continue
        }

        #select for users whos passwords have not expired already, but there password will expire before $notifydate, but not after $nonotifybefore
        if ($PasswordExpire -le $todaysdate) {  #check if already expired and increment the variables
            $ExpiredInOUcount++
            
        } elseif ( (($PasswordExpire -le $firstnotifydate) -and ($PasswordExpire -ge $nonotifybefore)) -or ($PasswordExpire -le $dailynotifydate) ) {   #check if users are about to expire but not in the no notify period

            #add data to the userobject for export to CSV
            $user.PasswordExpire = $PasswordExpire
            $user.DaysToExpire = ($PasswordExpire - $todaysdate).Days.toString()

            #add user properties to array for later processing, and change password expire dates from array object to value
            $notifyusers += $user | Select-Object -Property Name,@{N="Password Exp Date";E={$_.PasswordExpire[0]}},@{N="Days To Exp";E={$_.DaysToExpire[0]}},EmailAddress,Title,CanonicalName,SamAccountName

            $ExpiringInOUcount++
        } elseif ( $PasswordExpire -lt $nonotifybefore ) { #captures users who wont be notified because password expires too soon
            $nonotifyusersInOUcount++
            
        } else {
            $notexpiringinOU++
        }


    }

    $nonotifyuserscount += $nonotifyusersInOUcount
    $expireduserscount += $ExpiredInOUcount
    "Expired         " + $ExpiredInOUcount
    "Expiring Soon:  " + $ExpiringInOUcount
    "No Notify:      " + $nonotifyusersInOUcount
    "Not Expiring:   " + $notexpiringinOU
    
}

#uncomment to print users in output
#notifyusers


#output users with expiring passwords to CSV, add date
$csvpath = (Get-Location).Path + "\" + $csvbasename +".csv"
$notifyusers | Select-Object "Name", "Days To Exp", "Password Exp Date", "EmailAddress", "Title", "SamAccountName", "CanonicalName" | Sort-Object -Property CanonicalName | Export-Csv -Path $csvpath -NoTypeInformation 

""
"Users Output to:  " + $csvpath

#select only users with email addresses in AD
$emailusers = $notifyusers | where {$_.EmailAddress}

""
"Sending email to " + $emailusers.Count + " users."

$failures = @()  #keep track of any failures
foreach ($user in $emailusers) {
    
    $rcpt = $user.EmailAddress
    $subject = "Please change your password, it will expire in " + $user.'Days To Exp' + " days"

    $body = $user.Name +",`n`n    Our records indicate that your password will expire in " + 
    $user.'Days To Exp' + 
    " days on " + 
    $user.'Password Exp Date'.ToString($dateformat) + 
    ".  If your password is not changed in time you may be locked out of the VPN, and have other problems.  Please see the attached document on how to change your password when working remotely.`n`nThanks,`nIT Operations"

    if( $sendemails ) {

        "Sending mail to:  " + $rcpt

        #try sending the email capture any failures
        try {
            Send-MailMessage -To $rcpt -From $mailfrom -SmtpServer $mailserver -Port $mailserverport -Subject $subject -Body $body -Attachments $attachment
        } catch {
            $user
        
            $failures += New-Object -TypeName psobject -Property @{Name = $user.Name; PassExpDate = $user.'Password Exp Date'; DaysToExp = $user.'Days To Exp'; EmailAddress=$user.EmailAddress; Title=$user.Title; CanonicalName = $user.CanonicalName; Error= $_.Exception.Message; SamAccountName = $user.SamAccountName }
        
        
        }
    } else {
        '-sendemails not used skipping email to:  ' + $rcpt
    }
}

#output failures to CSV
$failcsvpath = (Get-Location).Path + "\" + $csvbasename + "_EmailFails.csv"
$failures | Select-Object "Name", "DaysToExp", "PassExpDate", "EmailAddress", "Title", "SamAccountName", "CanonicalName", "Error" | Sort-Object -Property CanonicalName | Export-Csv -Path $failcsvpath -NoTypeInformation 


""
""

#create notification email for IT department
$body = "First Notification:`t" + $firstnotifydate.ToString($dateformat) + "
Daily Notification:`t" + $dailynotifydate.ToString($dateformat) + "

Expired:`t`t" + $expireduserscount + "
User Count:`t`t" + $totaluserscount + "

Total Expiring:`t`t" + ($notifyusers.Count +$nonotifyuserscount)  + "
Excluded by date:`t" + $nonotifyuserscount + "   #these users exluded because password expires between first notify and daily notify date
Expiring to notify:`t" + $notifyusers.Count + "
Expiring w/o Email:`t" + ($notifyusers.Count - $emailusers.Count) + "
Email Attempts:`t`t" + $emailusers.Count + "
Email Failures:`t`t" + $failures.Count + "

Users Output to:`t`t" + $csvpath + "
Email Failures Output to:`t" + $failcsvpath + "

***OU(s) scanned***`n"


foreach ($OU in $userOUs ) { $body += $OU + "`n" }  #add all OUs to $body


if (-not $sendemails) {  #add a warning that we are not actually sending emails to the notification and output.

    $noemailwarning = "`n`n!!!WARNING '-sendemails' ARGUMENT NOT PASSED, NO PASSWORD EXPIRATION NOTICE EMAILS WERE ACTUALLY SENT!!!`n`n"
    $body = $noemailwarning + $body + $noemailwarning
}

$body  #output status to console

#send stats to IT department
if (-not $nostatusemail) {  #check for switch to avoid sending email

    ""
    "Sending status email to:  " + $statusto

    $subject = 'Password expiration notification report (' + $firstnotifydays + " days / " + $dailynotifydays + " days)"

    Send-MailMessage -To $statusto -From $mailfrom -SmtpServer $mailserver -Port $mailserverport -Subject $subject -Body $body -Attachments @($csvpath, $failcsvpath)
}

