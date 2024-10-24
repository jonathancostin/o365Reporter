# User Activity and MFA Reporting Script

This PowerShell script connects to Microsoft Graph and Exchange Online to generate detailed reports on user activity, assigned licenses, Multi-Factor Authentication (MFA) status, and enrolled devices. It categorizes users into active and inactive based on their last sign-in date and exports the data into CSV files.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Script Details](#script-details)
  - [Connections](#connections)
  - [Data Retrieval](#data-retrieval)
  - [Reporting](#reporting)
- [Error Handling](#error-handling)
- [Notes](#notes)
- [License](#license)

## Overview

The script performs the following actions:

1. **Establishes Connections**: Connects to Microsoft Graph and Exchange Online with the necessary scopes and permissions.
2. **Data Collection**: Retrieves user information, including sign-in activity, licenses, MFA status, enrolled devices, and last sent email date for inactive users.
3. **Classification**: Categorizes users as active or inactive based on their last sign-in date compared to one month ago.
4. **Reporting**: Exports the collected data into `ActiveUsers.csv` and `InactiveUsers.csv` files for further analysis.

## Features

- **User Activity Analysis**: Determines active and inactive users based on sign-in activity and last email sent.
- **MFA Status Reporting**: Retrieves detailed MFA status, including the type of MFA method used.
- **License Reporting**: Lists all assigned licenses for each user with SKU part numbers.
- **Device Enrollment**: Provides information on devices enrolled by each user.
- **CSV Export**: Outputs the results into easy-to-use CSV files for reporting and compliance purposes.

## Prerequisites

- **PowerShell 5.1** or higher.
- **Modules**:
  - [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation).
  - [Exchange Online Management Module](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/).
- **Permissions**:
  - Appropriate admin permissions to access user data and mailboxes.
- **Azure AD Sign-In Logs**: Enabled and retained for accurate sign-in activity data.

## Installation

1. **Install Required Modules**:

   ```powershell
   Install-Module mggraph -Force
   Install-Module ExchangeOnlineManagement -Force
   ```

2. **Clone the Repository**:

   ```bash
   git clone https://github.com/jonathancostin/o365Reporter.git
   ```

3. **Navigate to the Script Directory**:

   ```bash
   cd o365Reporter
   ```

## Usage

1. **Open PowerShell as Administrator**.

2. **Run the Script**:

   ```powershell
   .\o365Reporter.ps1
   ```

3. **Authenticate**:

   - The script will prompt you to sign in for Microsoft Graph and Exchange Online.
   - Use an account with the necessary permissions.

4. **Wait for Completion**:

   - The script processes each user and displays progress in the console.
   - Upon completion, it generates `ActiveUsers.csv` and `InactiveUsers.csv` in a `Results` folder.

## Script Details

### Connections

- **Microsoft Graph**: Connects with scopes for reading user information, devices, directories, authentication methods, and audit logs.

  ```powershell
  Connect-MgGraph -Scopes "User.Read.All, DeviceManagementManagedDevices.Read.All, Directory.Read.All, User.ReadBasic.All, UserAuthenticationMethod.Read.All, AuditLog.Read.All"
  ```

- **Exchange Online**: Connects to retrieve mailbox information.

  ```powershell
  Connect-ExchangeOnline
  ```

### Data Retrieval

- **Date Calculations**: Determines the current date and the date one month ago for comparison.

- **SKU Mapping**: Retrieves all subscribed SKUs and builds a hashtable mapping SKU IDs to SKU Part Numbers.

- **User Data**: Fetches all users with selected properties, including sign-in activity and assigned licenses.

- **Functions**:

  - `Get-LastSignInDate`: Retrieves the last sign-in date for a user.
  - `Get-LastSentMessageDate`: Retrieves the date of the last sent email for a user.
  - `Get-AssignedLicenses`: Maps assigned license IDs to their SKU part numbers.
  - `Get-MfaStatus`: Determines if a user has MFA enabled and the type of MFA method.
  - `Get-UserEnrolledDevices`: Retrieves devices enrolled by the user.

### Reporting

- **Classification**:

  - **Inactive Users**: Users who have never signed in or have not signed in within the last month.
  - **Active Users**: Users with sign-in activity within the last month.

- **CSV Export**:

  - Creates a `Results` directory in the current location.
  - Exports inactive users to `InactiveUsers.csv`.
  - Exports active users to `ActiveUsers.csv`.

## Error Handling

- **Mailbox Access**:

  - If the script cannot access a user's mailbox, it logs an error and continues processing.

- **MFA Retrieval**:

  - If MFA methods cannot be retrieved, the script logs a warning and sets MFA status to "Error" or "Unknown".

- **Device Retrieval**:

  - If device information cannot be retrieved, it logs a warning and continues.

- **General Exceptions**:

  - The script uses try-catch blocks to handle exceptions and ensure that one failure doesn't halt the entire process.

## Notes

- **Execution Policy**:

  - Ensure that your PowerShell execution policy allows running scripts:

    ```powershell
    Set-ExecutionPolicy RemoteSigned -Scope Process
    ```

- **Performance**:

  - The script may take time to run, depending on the number of users.
  - Running during off-peak hours is recommended for large environments.

- **Data Accuracy**:

  - Sign-in activity depends on Azure AD sign-in logs; ensure logs are retained as per your organization's policy.

- **Logging**:

  - Progress and error messages are displayed in the console for monitoring.

## License

This project is licensed under the [MIT License](LICENSE). You are free to use, modify, and distribute this script as per the license terms.

---

**Disclaimer**: This script is provided "as-is" without any warranty. Use it at your own risk. Always test scripts in a controlled environment before deploying them in production.

---
