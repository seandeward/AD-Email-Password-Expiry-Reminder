
#:: VERSION 1.1

#:: PROGRAM SETTINGS / GLOBAL VARIABLES
  #:: MODULE SETTINGS
    $DebugMode = $true # Do you want to display extra Write-Host messages that explain the script's process?
      $AdminEmailForDebugging = "" # When $DebugMode is set to $true, what email address will receive all of the emails? (This is where you'd put your email while you're configuring the script for your organization!)
    
    $checkForADUpdates = $false # do you want the script to check for updates to the ActiveDirectory module?

  #:: EMAIL SETTINGS
    $email_FromAddress = 'FromAddress@company.com' # who is this email going to be sent from?
    if ($DebugMode -eq $true) {
      $email_ToAddress = $AdminEmailForDebugging
      if ($email_ToAddress -eq "") {
        while ($email_ToAddress -eq ""){
          $email_ToAddress = Read-Host "Please enter your email address. (This will be used for debugging purposes. Instead of sending an email to a user for their password's approaching expiration date, this script will send this email all of the emails instead.) "
        }
      }
    }
    #* If $DebugMode is $false, the $email_ToAddress is declared right before sending the Send-MailMessage at the end of the script.
    $email_SmtpServer = '0.0.0.0' # IP Address of our SMTP Relay server.
    $email_PortNumber = '25' # as information from these email is not sensitive by default, it will be sent over Port 25 by default. This means everything is sent in plain-text (no encryption), but does not require authentication.
    $email_Subject = 'Your Password Expires Soon' # email subject.
    $email_Priority = "High" # Choose the priority that will show in the sent email. Options are "High", "Normal", or "Low".1203
    # $email_Body grabs the $cssStyle variable to make the information in the email more readable.

if ($DebugMode -eq $true) {
  Write-Host "[DEBUG] Loading Functions..." -ForegroundColor Magenta
}

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

if ($DebugMode -eq $true) {
  Write-Host "    Done" -ForegroundColor Green
  Write-Host "[DEBUG] Importing `Active Directory` Module... " -ForegroundColor Magenta
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
if ($DebugMode -eq $true) {
  Write-Host "[DEBUG] Today's Date = $TodaysDate" -ForegroundColor Magenta
}

# get all enabled AD users who have passwords that will expire
$AllActiveADUsers = Get-ADUser -Filter {enabled -eq $true -and PasswordNeverExpires -eq $false} | Select-Object -ExpandProperty SamAccountName
if ($DebugMode -eq $true) {
  Write-Host "[DEBUG] All Active AD Users:" -ForegroundColor Magenta
  Write-Host "$AllActiveADUsers" -ForegroundColor Gray
  Write-Host "" # for terminal readability
  Write-Host "[DEBUG] Beginning to send emails..." -ForegroundColor Magenta
}

foreach ($user in $AllActiveADUsers) {
  if ($DebugMode -eq $true) {
    Write-Host "" # for terminal readability by separating each user's information visually
  }
  # convert string into a yyyy-MM-dd format
  $ExpiryDateStr = Get-ADUser -Identity $user -Properties msDS-UserPasswordExpiryTimeComputed | Select-Object Name, @{Name="Password Expiry Date";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} | Select-Object -ExpandProperty "Password Expiry Date"
  $ExpiryDate = $ExpiryDateStr.ToString("yyyy-MM-dd")
  if ($DebugMode -eq $true) {
    Write-Host "$user" -NoNewline
    Write-Host " | Expiration date: $ExpiryDate" -ForegroundColor Gray -NoNewline
  }

  # get number of days from today's date and the user's password expiry date
  $ts = New-TimeSpan -Start $TodaysDate -End $ExpiryDate
  $daysRemaining = $ts.Days
  if ($DebugMode -eq $true) {
    Write-Host " | Days Remaining: $daysRemaining" -ForegroundColor Gray -NoNewline
  }

  # if password expires in 1-7 or 14 days
  if ($daysRemaining -in 1..7 -or $daysRemaining -eq 14) {
    # find if number of days is singular or plural
    if ($daysRemaining -eq 1) {
      $singleOrPlural = "day"
    } else {
      $singleOrPlural = "days"
    }
    if ($DebugMode -eq $true) {
      Write-Host " | Single/Plural: $singleOrPlural" -ForegroundColor Gray -NoNewline
    }

    # if days =< 3, make banner color red. Else, make yellow
    if ($daysRemaining -le 3) {
      $bannerColor = '#c72828'
    } elseif ($daysRemaining -ge 14) {
      $bannerColor = '#f7dd4c'
    } else {
      $bannerColor = '#edaa1a'
    }
    if ($DebugMode -eq $true) {
      Write-Host " | BannerColor: $bannerColor" -ForegroundColor Gray -NoNewline
    }
    
    # get user info
    $UserInfo = Get-ADUser -Identity $user -Properties GivenName, Name, CanonicalName, EmailAddress | Select-Object GivenName, Name, CanonicalName, EmailAddress
    $UserFirstName = $UserInfo.GivenName
    $UserFullName = $UserInfo.name
    $ouPath = $UserInfo.CanonicalName
    $UserEmailAddress = $UserInfo.EmailAddress
    if ($DebugMode -eq $true) {
      Write-Host " | User Info: $UserFirstName, $userFullName, $UserEmailAddress" -ForegroundColor Gray # $ouPath is printed in the terminal when an email is sent to the user.
    }

    # find which OU $user is in, and send a respective custom message
    try {
      $message = FindDepartment($ouPath)
      if ($message -eq $False) {
        if ($DebugMode -eq $true) {
          Write-Host "    | Skipping $user since FindDepartment function returned False." -ForegroundColor Yellow -NoNewline
        } else {
          Write-Host "    Skipping $user since FindDepartment function returned False." -ForegroundColor Yellow -NoNewline
        }
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

    if ($DebugMode -eq $true) {
      Write-Host ""
      Write-Host "[DEBUG] $UserEmailAddress would normally get passed to $email_ToAddress here. Skipping... " -ForegroundColor Magenta
    } else {
      $email_ToAddress = $UserEmailAddress # actually make the $email_ToAddress = $UserEmailAddress
    }

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
if ($DebugMode -eq $true) {
  Write-Host "" # for terminal readability
  Write-Host "" # for terminal readability
  Read-Host "[DEBUG] Script is complete! Please press Enter to close this terminal." -ForegroundColor Magenta # add this so that the terminal doesn't automatically close when the script is complete
}
exit
