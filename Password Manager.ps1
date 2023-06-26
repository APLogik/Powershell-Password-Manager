$script:words = Import-CSV "C:\dict.csv"
$script:EmailSender = "ExampleSupport <examplesupport@example.com>"
$script:ITOrgName = "Example IT Org Name" #to be used in subject line of emails
$script:SMTPServer = "192.168.1.1"
$script:EmailBodyTop = "An IT Staff member has created a new user account or set a password and wanted to share the details."
$LogPath = "C:\PasswordManagementLog.txt"
$AdminAccount = $env:USERNAME

function MakeLogEntry ([string]$LogText) {
    # Get the current date and time in the specified format
    $TodaysDate = Get-Date -Format "yyyyMMdd-HHmm"

    # Create the log entry by combining the date/time and log text
    $LogEntry = "$($TodaysDate) $($LogText)"

    # Append the log entry to the log file
    Add-Content -Path $LogPath -Value $LogEntry
}

function IsValidEmail {
    param([string]$EmailAddress)

    try {
        # Attempt to create a MailAddress object using the provided email address
        $null = [mailaddress]$EmailAddress

        # The email address is valid
        return $true
    }
    catch {
        # An exception occurred, indicating the email address is invalid
        return $false
    }
}

Function GeneratePw {
    $word1 = Get-Random -InputObject $words  # Select random word from the $words array, store it in $word1
    $word2 = Get-Random -InputObject $words  # Select random word from the $words array, store it in $word2
    $number = Get-Random -Minimum 100 -Maximum 999  # Generate a random three-digit number, store it in $number
    $password = $word1.word + $word2.word + $number  # Combine the two words and the number to form the password
    $password = $password.substring(0,1).toupper() + $password.substring(1).tolower()  # Capitalize the first letter of the password and convert the rest to lowercase
    $SecurePassword = ConvertTo-SecureString $password -AsPlainText -Force  # Convert the password to a secure string
    $PasswordPackage = ,@($password, $SecurePassword)  # Create a package containing both the password as a string and the secure password
    return $PasswordPackage  # Return the password package
}

function GetAdUserInformation {
    [CmdletBinding()]
    param(
        [string]$UserName
    )

    # Retrieve user information from Active Directory
    $script:UserObject = Get-ADUser -Identity $UserName -Properties EmailAddress, SamAccountName, GivenName, Surname, Description, AccountExpirationDate, displayName, lockedout, msDS-UserPasswordExpiryTimeComputed, Enabled

    # Extract individual properties from the user object
    $script:FullName = $UserObject.displayName
    $script:UsersEmail = $UserObject.EmailAddress
    $script:FirstName = $UserObject.GivenName
    $script:LastName = $UserObject.Surname
    $script:FullDescription = $UserObject.Description
    $script:ExpirationDate = $UserObject.AccountExpirationDate
    $script:LockedOutStatus = $UserObject.lockedout

    # Determine the enabled status of the user
    if ($UserObject.Enabled) {
        $script:EnabledStatus = "Enabled"
    } else {
        $script:EnabledStatus = "Disabled"
    }

    $script:UserName = $UserObject.SamAccountName

    # Convert the password expiry time to a readable format
    $script:PasswordExpiryLong = [datetime]::FromFileTime($UserObject."msDS-UserPasswordExpiryTimeComputed")

    # Check if the password was changed
    if ($script:PasswordString) {
        # Password string exists
    } else {
        $script:PasswordString = "Password was not changed"
    }

    # Get the email address of the current context
    $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
    $script:AdminEmail = $searcher.FindOne().Properties.mail

    # Check if the user's email address is blank
    if ($UsersEmail) {
        # UsersEmail is not blank
    } else {
        while ($Ready -ne "TRUE") {
            # Prompt the user to enter the email address
            $script:UsersEmail = Read-Host "User's email address appears to be blank, please enter it now"

            if ($UsersEmail) {
                # Validate the entered email address
                $EmailTest = IsValidEmail $UsersEmail # Returns true if email is valid

                if ($EmailTest) {
                    $Ready = "TRUE"
                    # Set the user's email address in AD and the script
                    Set-ADUser -EmailAddress $UsersEmail -Identity $UserName
                    Write-Host "Email address has been set for this user's account" -ForegroundColor Yellow -BackgroundColor DarkGreen
                } else {
                    Write-Host "Invalid entry, try again"
                    $Ready = "False"
                    continue
                }
            } else {
                Write-Host "Invalid entry, try again"
            }
        }
        $Ready = "FALSE" # Resetting for next use
    }
}

function CheckForExistingAccount {
    cls

    # Check for existing accounts
    write-host "`n`nExisting Account Search Function`n"

    # Set initial values
    $script:UseExistingUser = "TRUE"
    $DoneSearchingForUsers = "FALSE"
    $script:selection = $null

    # Continue searching for users until done
    while ($DoneSearchingForUsers -ne "TRUE") {
        # Prompt for a search term
        $searchterm = Read-Host -Prompt "ENTER A SEARCH TERM that would for sure be in their account, first, or last name`n`n"

        # Search for users matching the search term
        $users = Get-Aduser -Filter "anr -like '$searchterm'"

        if ($users.count -gt 1) {
            # Multiple users found
            write-host "`nMultiple users were found" -Foregroundcolor Yellow -Backgroundcolor DarkGreen

            # Display each user's information
            for ($i = 0; $i -lt $users.count; $i++) {
                Write-Host "$($i): $($users[$i].SamAccountName) | $($users[$i].Name)" -Foregroundcolor White -Backgroundcolor Blue
            }

            # Prompt for user selection
            while ([string]::IsNullOrWhiteSpace($script:selection)) {
                while ($Ready -ne "TRUE") {
                    $UserNumber = Read-Host "`nIf you want to continue with an existing user, enter its number`nPress q to stop searching"

                    # Check if user wants to quit searching
                    if ($UserNumber -eq "q") {
                        $QuitNow = "TRUE"
                    }

                    # Check if user entered a valid user number
                    if ($UserNumber) {
                        $Ready = "TRUE"
                        $script:UseExistingUser = "TRUE"
                    } else {
                        write-host "Invalid entry, try again"
                    }
                }

                $Ready = "FALSE" # Resetting for next use
                $script:selection = $users[$UserNumber].SamAccountName

                # Check if user selection is valid
                if ([string]::IsNullOrWhiteSpace($script:selection)) {
                    write-host "Invalid entry, try again"
                } else {
                    $DoneSearchingForUsers = "TRUE"
                }

                # Check if user wants to quit searching
                if ($QuitNow -eq "TRUE") {
                    $script:selection = "quit searching"
                    $DoneSearchingForUsers = "TRUE"
                }
            }

            cls
            write-host "`nYou have selected $selection" -Foregroundcolor Yellow -Backgroundcolor DarkGreen

            # Perform actions based on user selection
            if ($script:selection -eq "quit searching") {
                $script:UseExistingUser = "FALSE"
                write-host "`nExiting Search Function`n" -Foregroundcolor Yellow -Backgroundcolor DarkGreen
            } else {
                $script:UseExistingUser = "TRUE"
                GetAdUserInformation -Username $script:selection
            }
        } else {
            # Single user found
            if ([string]::IsNullOrWhiteSpace($users)) {
                write-host "`nNothing Found" -Foregroundcolor Yellow -Backgroundcolor DarkGreen

                # Prompt to stop searching or continue
                $StopSearching = Read-Host "Press q to stop searching, press anything else to continue"

                # Check if user wants to quit searching
                if ($StopSearching -eq "q") {
                    $script:UseExistingUser = "FALSE"
                    break
                }
                continue
            } else {
                write-host "`nSingle user found - $users.SamAccountName" -Foregroundcolor Yellow -Backgroundcolor DarkGreen
                $script:selection = $users.SamAccountName

                # Prompt for user confirmation
                while ($Ready -ne "TRUE") {
                    $Ready = Read-Host "`nDo you want to continue with this existing user? [y/n]"

                    # Check if user wants to quit searching
                    if ($Ready -eq "n") {
                        $Ready = "TRUE"
                        $DoneSearchingForUsers = "TRUE"
                        $script:selection = "quit searching"
                    }

                    # Check if user wants to continue with existing user
                    if ($Ready -eq "y") {
                        $Ready = "TRUE"
                        $DoneSearchingForUsers = "TRUE"
                        $script:UseExistingUser = "TRUE"
                    }
                }

                cls
                write-host "`nYou have selected $script:selection" -Foregroundcolor Yellow -Backgroundcolor DarkGreen

                # Perform actions based on user selection
                if ($script:selection -eq "quit searching") {
                    $script:UseExistingUser = "FALSE"
                    write-host "`nExiting Search Function`n" -Foregroundcolor Yellow -Backgroundcolor DarkGreen
                } else {
                    $script:UseExistingUser = "TRUE"
                    GetAdUserInformation -Username $script:selection
                }
            }
        }
    }
}

function DoEmailing {
    # Now let's e-mail this to the customer and other people
    Write-Host "`n`n"

    # Check if UserName is "None" and get AD user information if needed
    if ($script:UserName -eq "None") {
        CheckForExistingAccount
        GetAdUserInformation -Username $script:UserName
    }

    # Set a default value for the PasswordString if it is empty
    if ([string]::IsNullOrWhiteSpace($script:PasswordString)) {
        $script:PasswordString = "Password was not changed"
    }

    # Prepare an array to store attachments
    $Attachments = @()
    $Attachments += 'C:\Welcome.jpg' # example of adding attachments

    # Ask the technician if Materials Package 1 should be included
    $IncludeMaterialsPackage1 = Read-Host "Include Materials Package 1? [y/n]"
    if ($IncludeMaterialsPackage1 -eq 'y') {
        $MaterialsPackage1Info = "Example text describing MaterialsPackage1."
        $Attachments += 'C:\UserGuide1.docx'
    } else {
        $MaterialsPackage1Info = ""
    }

    # Ask the technician if Materials Package 2 should be included
    $IncludeMaterialsPackage2 = Read-Host "Include Materials Package 2? [y/n]"
    if ($IncludeMaterialsPackage2 -eq 'y') {
        $MaterialsPackage2Info = "Example text describing MaterialsPackage2."
        $Attachments += 'C:\UserGuide2.docx'
    } else {
        $MaterialsPackage2Info = ""
    }

    # Prepare the email subject and body
    $EmailSubject = "$script:ITOrgName - New Account Information"
    $EmailBody = "$EmailBodyTop`n`n
Username = $UserName`n
Password = $script:PasswordString`n`n
$MaterialsPackage1Info`n
$MaterialsPackage2Info"

    # Set the parameters for sending the email
    $SendMailParameters = @{
        From       = $EmailSender
        To         = $UsersEmail
        Subject    = $EmailSubject
        Body       = $EmailBody
        SMTPServer = $SMTPServer
        Attachments = $Attachments
    }
	# Ask if the email should be sent to the customer's primary email
    $EmailCustomer = Read-Host "Send to this User's Primary Email? [y/n]"
    if ($EmailCustomer -eq 'y') {
        # Send the email to the customer
        Send-MailMessage @SendMailParameters
		MakeLogEntry "$AdminAccount - Sent account info to $UsersEmail"
	}
    # Ask if the email should be sent to others
    $EmailOthers = Read-Host "To Others? [y/n]"
    if ($EmailOthers -eq 'y') {
        while ($DoneEmailing -ne "TRUE") {
            Write-Host
            $OtherEmail = Read-Host -Prompt "Enter an email address to send to`n"
            $SendMailParameters.To = $OtherEmail
            # Send the email to the specified email address
            Send-MailMessage @SendMailParameters
			MakeLogEntry "$AdminAccount - Sent account info to $OtherEmail"
			$EmailOthers = Read-Host "Email someone else? [y/n]"
			if ($EmailOthers -eq 'n') {
                $DoneEmailing = "TRUE"
            }
		}
	}
}

function SetUserPassword {
	cls
	# Check if the UserName variable is set to "None"
	if ($script:UserName -eq "None") {
		CheckForExistingAccount
		GetAdUserInformation -Username $script:UserName
	}

	# Generate a password
	$script:password = GeneratePw
	$script:PasswordString = $password[0]
	Write-Host

	# Retrieve the account expiration date
	$ExpirationDate = (Get-ADUser -Identity $UserName -Properties accountexpirationdate).accountexpirationdate

	# Check if the expiration date is null or empty
	if ([string]::IsNullOrWhiteSpace($ExpirationDate)) {
		$ExpirationDate = "None"
	}

	Write-Host "Account Expiration Date = $ExpirationDate`n" -ForegroundColor Yellow -BackgroundColor DarkGreen

	Write-Host "$UserName's password will be set to $PasswordString" -ForegroundColor Yellow -BackgroundColor DarkGreen

	# Prompt for password regeneration
	while ($Regen -ne 'n') {
		$Regen = Read-Host -Prompt "Re-Generate the password? [y/n]"
		if ($Regen -eq 'y') {
			$script:password = GeneratePw
			$script:PasswordString = $password[0]
			Write-Host "$UserName's password will be set to $PasswordString" -ForegroundColor Yellow -BackgroundColor DarkGreen
		}
	}

	Write-Host
	
	# Prompt for password confirmation
	while ($Ready -ne "TRUE") {
		$SetPassword = Read-Host "Press y to set this password and email it to yourself, n to cancel"
		if ($SetPassword -eq 'y') {
			$Ready = "TRUE"
			clear-host

			# Set the account password
			Set-ADAccountPassword -Identity $UserName -Reset -NewPassword $password[1]
			MakeLogEntry "$AdminAccount - Set password on $UserName"
			
			# Get the current context's email address
			$searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
			$script:AdminEmail = $searcher.FindOne().Properties.mail

			# Compose the email body and subject
			$PWEmailBody = "A password has been set on this account.`n`nUsername = $UserName`nPassword = $PasswordString`n`n"
			$PWEmailSubject = "$script:ITOrgName - Account Password Changed"

			# Send the email with the new password
			Send-MailMessage -From $EmailSender -To $AdminEmail -Subject $PWEmailSubject -Body $PWEmailBody -SMTPServer $SMTPServer
			MakeLogEntry "$AdminAccount - Sent account info to $AdminEmail"
			Write-Host "`n`n Account password has been set to $PasswordString" -ForegroundColor Yellow -BackgroundColor DarkGreen
		}
		if ($SetPassword -eq 'n') {
			clear-host
			Write-Host "No password was changed" -ForegroundColor Yellow -BackgroundColor DarkGreen
			$Ready = "TRUE"
		}
	}
}

function UnlockUser {
    # Clear the screen
    cls
    
    # Check if the UserName variable is set to "None" and prompt for an existing account if needed
    if ($script:UserName -eq "None") {
        CheckForExistingAccount
    }
    
    # Unlock the Active Directory account for the specified UserName
    Unlock-ADAccount -Identity $script:UserName
    
    MakeLogEntry "$AdminAccount - Unlocked $UserName"
    
    Write-Host "$UserName has been unlocked" -Foregroundcolor Yellow -Backgroundcolor DarkGreen
    
    # Refresh the account information before returning
    GetAdUserInformation -Username $script:UserName
}

function Menu {
	# Check if UserName is null or empty, set it to "None" if true
    if([string]::IsNullOrWhiteSpace($script:UserName)){
        $script:UserName = "None"
    }

    # Display the main menu
	do {
		Write-Host "`n`n================ Main Menu ================`n" -Foregroundcolor Yellow -Backgroundcolor DarkGreen
		Write-Host "Currently Selected User: $UserName`n" -Foregroundcolor Yellow -Backgroundcolor DarkGreen
		Write-Host "Account Expiration Date: $ExpirationDate" -Foregroundcolor Yellow -Backgroundcolor DarkGreen
		Write-Host "Account Status: $EnabledStatus" -Foregroundcolor Yellow -Backgroundcolor DarkGreen
		Write-Host "Password Expiration Date: $PasswordExpiryLong" -Foregroundcolor Yellow -Backgroundcolor DarkGreen
		Write-Host "Password Lockout Status: $LockedOutStatus`n" -Foregroundcolor Yellow -Backgroundcolor DarkGreen

		
		Write-Host "1: Press '1' to search and select an account." -Foregroundcolor White -Backgroundcolor Blue
		Write-Host "2: Press '2' to set account password." -Foregroundcolor White -Backgroundcolor Blue
		Write-Host "3: Press '3' to unlock account." -Foregroundcolor White -Backgroundcolor Blue
		Write-Host "4: Press '4' to email account information." -Foregroundcolor White -Backgroundcolor Blue
		Write-Host "5: Press '5' to refresh account information." -Foregroundcolor White -Backgroundcolor Blue
		Write-Host "Q: Press 'Q' to exit.`n" -Foregroundcolor White -Backgroundcolor Blue

		$menuselection = Read-Host "Please make a selection"
		switch ($menuselection){
			'1' {CheckForExistingAccount;}
			'2' {SetUserPassword;}
			'3' {UnlockUser;}
			'4' {DoEmailing;}
			'5' {if($script:UserName -eq "None"){CheckForExistingAccount};GetAdUserInformation -Username $script:UserName;cls;Write-Host "`nAccount info refreshed" -Foregroundcolor Yellow -Backgroundcolor DarkGreen}
		}
	}
until ($menuselection -eq 'q')
}

#clear the screen and present the menu
cls
Menu;