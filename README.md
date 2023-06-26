# Powershell-Password-Manager
A command line based, menu-driven password manager intended for roles where access is limited to unlocking and resetting passwords.

This script provides a set of functions for managing user passwords, unlocking accounts, and email account information to users.

## Prerequisites

- Active Directory PowerShell module
- CSV file containing a dictionary of words (e.g., "dict.csv")

## Functions

The script includes the following functions:

### 1. Search and Select a User

This function allows you to search for existing user accounts based on a specified search term (e.g., first name, last name, or any term that matches the account). If multiple users are found, a list is displayed, and you can select the user to continue with. The selected user's information is then retrieved from Active Directory.

### 2. Generate and Set Password

This function generates a random password for a user account. It selects two random words from a dictionary file and combines them with a random three-digit number. The password is then set for the selected user account in Active Directory.

### 3. Unlock Account

This function unlocks the user account. It checks the locked-out status of the selected user and unlocks the account if necessary.

### 4. Email Account Information

This function sends an email to the user's primary email address with the account details, including the username and password. The email can also include attachments such as user guides or other materials. You can choose whether to send the email to the customer and/or other recipients.

### 5. Refresh Account Information

This function retrieves updated information for the selected user account from Active Directory. It fetches details such as the user's full name, email address, account expiration date, enabled status, and password expiry time.

## Usage

To use this script, follow these steps:

1. Import the required CSV file containing the dictionary of words.
2. Set the email sender, IT organization name, SMTP server, and email body in the script variables.
3. Run the desired function based on the menu options:
   - **Search and Select a User**: Search for and select a user account to perform actions on.
   - **Generate and Set Password**: Generate a random password and set it for the selected user account.
   - **Unlock Account**: Unlock the selected user account if it is locked out.
   - **Email Account Information**: Send an email to the user's primary email address with account details.
   - **Refresh Account Information**: Retrieve updated information for the selected user account.

Please ensure that you have the necessary permissions and prerequisites in place before running this script.

## Log File

The script creates a log file at the specified path (`C:\PasswordManagementLog.txt`). Each log entry contains a timestamp and relevant information about the action performed, such as the account name running the script, account information sent, or any errors encountered.
