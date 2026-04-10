
#:: VERSION 1.0.1

#:: PROGRAM SETTINGS / GLOBAL VARIABLES
  #:: MODULE SETTINGS
    $checkForADUpdates = $false # do you want the script to check for updates to the ActiveDirectory module?

  #:: EMAIL SETTINGS
    $email_FromAddress = 'FromAddress@company.com' # who is this email going to be sent from?
    #* $email_ToAddress is set right before the Send-MailMessage command near the end of the script.
      # You can replace this with your own email for debugging, but just remember to put `$UserEmailAddress` back when you're done.
    $email_SmtpServer = '0.0.0.0' # IP Address of our SMTP Relay server.
    $email_PortNumber = '25' # as information from these email is not sensitive by default, it will be sent over Port 25 by default. This means everything is sent in plain-text (no encryption), but does not require authentication.
    $email_Subject = 'Your Password Expires Soon' # email subject.
    $email_Priority = "High" # Choose the priority that will show in the sent email. Options are "High", "Normal", or "Low".1203
    # $email_Body grabs the $cssStyle variable to make the information in the email more readable.

#:: FUNCTIONS
function FindDepartment([string]$path) {
  switch ($path) {
    #:: OU TEMPLATE
    "full/ou/path1" {
      return "
      <p> Put your custom HTML message here for users in this OU path! </p>
      "
    }
    "full/ou/path2" {
      return "
      <p> Put a different message here for users in THIS ou path! </p>
      "
    }
    "full/ou/path3" {
      return $false #? For OUs that you don't want to send emails to for now.
    }
    #:: ELSE
    default { # if there are no other matches. (Basically an else-like statement)
      # You can add a message here for any user not in any above OU, but by default the program will choose to not send an email.
      return $false
    }
  }
}

#:: SCRIPT START
try {
  Import-Module ActiveDirectory
}
catch {
  Write-Host "[!] Something went wrong importing the ActiveDirectory module" -ForegroundColor Red
  Throw $_
}

if ($checkForADUpdates -eq $true){
  try {
      Write-Host "[+] Checking for updates on the ExchangeOnlineManagement module..." -ForegroundColor Yellow
      Write-Host '    If you want to disable checking for updates, change the $checkForADUpdates variable to $false' -ForegroundColor Gray
      Update-Module -Name ExchangeOnlineManagement -ErrorAction Stop
      Write-Host "    Done" -ForegroundColor Green
  }
  catch {
      Write-Host "[!] There was an error updating the ExchangeOnlineManagement module..." -ForegroundColor Red
      Write-Host "    Continuing script..." -ForegroundColor Yellow
  }
}

# get today's date - convert to a yyyy-MM-dd string format
$TodaysDate = (Get-Date).ToString("yyyy-MM-dd")

# get all enabled AD users who have passwords that will expire
$AllActiveADUsers = Get-ADUser -Filter {enabled -eq $true -and PasswordNeverExpires -eq $false} | Select-Object -ExpandProperty SamAccountName

Write-Host "[+] Beginning to send emails..." -ForegroundColor Cyan

foreach ($user in $AllActiveADUsers) {
  # convert string into a yyyy-MM-dd format
  $ExpiryDateStr = Get-ADUser -Identity $user -Properties msDS-UserPasswordExpiryTimeComputed | Select-Object Name, @{Name="Password Expiry Date";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} | Select-Object -ExpandProperty "Password Expiry Date"
  $ExpiryDate = $ExpiryDateStr.ToString("yyyy-MM-dd")

  # get number of days from today's date and the user's password expiry date
  $ts = New-TimeSpan -Start $TodaysDate -End $ExpiryDate
  $daysRemaining = $ts.Days

  # if password expires in 1-7 or 14 days
  if ($daysRemaining -in 1..7 -or $daysRemaining -eq 14) {
    # find if number of days is singular or plural
    if ($daysRemaining -eq 1) {
      $singleOrPlural = "day"
    } else {
      $singleOrPlural = "days"
    }

    # if days =< 3, make banner color red. Else, make yellow
    if ($daysRemaining -le 3) {
      $bannerColor = '#c72828'
    } elseif ($daysRemaining -ge 14) {
      $bannerColor = '#f7dd4c'
    } else {
      $bannerColor = '#edaa1a'
    }
    
    # get user info
    $UserInfo = Get-ADUser -Identity $user -Properties GivenName, Name, CanonicalName, EmailAddress | Select-Object GivenName, Name, CanonicalName, EmailAddress
    $UserFirstName = $UserInfo.GivenName
    $UserFullName = $UserInfo.name
    $ouPath = $UserInfo.CanonicalName
    $UserEmailAddress = $UserInfo.EmailAddress

    # find which OU $user is in, and send a respective custom message
    try {
      $message = FindDepartment($ouPath)
      if ($message -eq $False) {
        Write-Host "    Skipping $user since FindDepartment function returned False." -ForegroundColor Magenta
        continue
      }
    }
    catch {
      Write-Host "    Something went wrong figuring out $user's department" -ForegroundColor Red
    }

    # HTML code for email notification body
    $email_Body = "
      <html>
      <head>
      <style>
        /* Import 'Roboto' font */
        @import url('https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,100..900;1,100..900&display=swap');
        
        /* VARIABLES */
        :root {
          --red: #c72828;
          --white: #fff;
          --light-gray: #eee;
          --gray: #777;
        }

        /* css Styling */
        
        html{
          font-family: 'Roboto', Arial, Helvetica, sans-serif;
          background-color: var(--light-gray);
          padding: 1em;
        }
        body {
          margin-left: 1em;
          margin-right: 1em;
        }
        .banner{
          background-color: $bannerColor;
          padding: 5px;
          border-top-left-radius: 5px;
          border-top-right-radius: 5px;
          text-wrap: balance;
        }
        .container{
          background-color: var(--white);
          border-radius: 6px;
          max-width: 35em;
            margin-left: auto;
            margin-right: auto;
          border: 1px solid $bannerColor; 
        }
        .message{
          padding: 1em;
        }
        h1, h2, h3, h4, footer{
          text-align: center;
          text-wrap: balance;
        }
        .footer{
          text-wrap: balance;
          text-align: center;
          max-width: fit-content;
          margin-left: auto;
          margin-right: auto;
          color: var(--gray);
          font-size: 12px;
        }
        .footer, .subtext{
          max-width: fit-content;
          margin-left: auto;
          margin-right: auto;
          color: #777;
          font-size: 12px;
        }
        .bigSpan {
          /* display: flex; */
          border-radius: 5px;
        }
        .littleSpan {
          max-width: 300px;
          /* text-align: center;
          padding-left: 1em;
          padding-right: 1em; */
        }
      </style>
      </head>
      <body>
      <div class='container'>
        <div class='banner'></div>
        <div class='message'>
          <h3>Your [company] network password </br>is going to expire in $daysRemaining $singleOrPlural.</h3>
          <p>Hi $userFirstName,</p>
          <p>Your current [company] network password is going to expire.</p>

          <p class='subtext'>This is the password used for logging into most [company] computers/laptops/VPNs. This is not regarding web services like Outlook, [app], [app], etc.</p>
          $message
          <p>Please update your password promptly to avoid any workflow disruptions.</p>
        </div>
      </div>
      <div class='footer'>
        <p>This is an automated notification from [company]'s Windows Active Directory network.</p>
      </div>
      <div class='footer'>
        <p>If you have any questions, please <a href='mailto:ITSupport@company.com'>submit a support ticket</a> to the [company] IT Department.</p>
      </div>
      </body>
      </html>
    "

    $email_ToAddress = $UserEmailAddress # Sets the "To Address" to the current user's email address

    #:: SEND THE EMAIL
    try {
      Send-MailMessage -To $email_ToAddress -From $email_FromAddress -Subject $email_Subject -Body $email_Body -BodyAsHtml -Priority $email_Priority -SmtpServer $email_SmtpServer -Port $email_PortNumber -ErrorAction Stop
      Write-Host "    Email sent to $user" -ForegroundColor Green -NoNewline
      Write-Host " - ouPath = $ouPath"
    }
    catch {
      Write-Host "[!] Something went wrong with the 'Send-MailMessage' command for $user" -ForegroundColor Red
    }
  }
  #* script moves onto the next $user in $AllActiveADUsers
}
#* Script is complete
exit
