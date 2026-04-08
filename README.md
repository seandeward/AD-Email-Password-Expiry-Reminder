# PowerShell Email Expiration Notification

<div style="text-align: center;"><b>Table of Contents</b></div>
[Requirements](#Requirements) | [Installation/Configuration](#Installation/Configuration) | [Contributing](#Contributing) | [Future Feature Wishlist](#Future Feature Wishlist)

## Overview

### Requirements
1. You must have a SMTP Relay server that can send all of these emails.
2. All users must have their email in their AD User Account to work.
	- For example, in `Active Directory Users and Computers`, if you open a user's account under `Properties`, a user must have their email entered in the `E-mail` field to receive the notification.
3. A little knowledge on basic HTML elements (for customizing the email body itself).

### What is it?

This is a simple PowerShell script makes a list of all of your users, checks if their password expires within 14 days, and then sends then an email to let them know ahead of time.

### Why does it matter?

Windows _does_ give a toast notification to a user when their password is going to expire, but its really easy to ignore. It might pop up in the corner when a user logs in, but it also goes away after a few seconds. **Plus, Windows doesn't give users _any_ instruction on how to actually reset their password.** 
- This usually leads to user downtime (and thus frustration and confusion), and more work for your IT Department supporting your end users.

An **email** notification is much harder to ignore since it appears in someone's inbox. They have to see it or act on it.

### How does it work?

1. Declares some configurable Global Variables + Functions.
2. Import the `ActiveDirectory` module (It also checks for updates here, if you choose it to).
3. Get today's date.
4. Get a list of all `enabled` AD users whose `PasswordNeverExpires` is set to `false`.
5. For each of those users, get the date their password expires, and see if it's up to 14 days away from today.
6. If it's more than 14 days away, skip that user. If it's 14 days or fewer, use the `FindDepartment` function to determine which OU they're in. _(This is where you configure custom messages for your AD users, depending on their OU!)_
7. Send the email with a customized string, specific to their name, remaining days until the password expires, department-specific instructions, and so on until each user in the list from Step 4 has been iterated over.

Then the script is done!

## Installation/Configuration

### Installation

Installation is pretty simple. I'm personally learning the basics of `git` commands, but you're welcome to fork this repo and customize it to your company's needs.

### Configuration

**For this script to work in your company's particular AD environment, there's a few settings you will have to configure:**
1. The Global Variables - starting at Line 6
	- These settings are crucial to getting the script to work. This is where you'll enter what IP your SMTP Relay server is, what emails to send it from/to, etc.
2. The `FindDepartment` function - at Line 19
	- This is the function that gives a user a department-specific message by matching them with a OU that you list in the `switch` statement on line 20. More detailed instructions are in the script itself.
3. The `$email_Body` string variable - starting at line 120
	- This is the body of the email itself. The script as is provides a template that gives you something to work off of.

## Contributing

(I'm currently working on a `CONTRIBUTING.MD` file. For the time being, you are more than welcome to create Issues or Pull Requests if you have any great ideas!)

## Development process

When making this app, I had 3 big questions to answer:
1. How to I actually send an email in PowerShell?
2. Can I pass HTML code through PowerShell, and include PowerShell variables inside of the HTML code?
3. If I can do both of those individual, how do I send an HTML body (with some CSS styling to make it look pretty) of code through an email?

Thankfully, these were easy to answer after some research! The `Send-MailMessage` was essential to making this script work- that's the command that sends the email. Then I figured out that the `Send-MailMessage` command has a parameter called `-body`, which is the body of the email itself. 

Since emails are basically HTML code, I created the template in `main.ps1` and was able to pass it to the `-body` parameter as a string. Meaning, HTML technically isn't processed inside of the script, but rather the email client that an end user is using.
- The only downside with this approach is that, while CSS codes works great on a lot of clients, it may not work well on some others. **The only client this has been tested with is the new Outlook Windows app and web-app.** At this time, I'm not sure how emails will appear in other clients.

While I could research more into how to make the email body more consistent across email clients, my primary focus is learning how to code in a variety of languages. If you find a solution, feel free to submit a Pull Request!

## Future Feature Wishlist

- **Add admin email alert feature if a script-ending bug occurs.**
	- For example, if the `ActiveDirectory` module isn't loaded, or if a list of users couldn't be obtained, then the script doesn't work. This 
