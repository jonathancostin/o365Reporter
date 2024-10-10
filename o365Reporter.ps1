# TODO
# Add mfa reports into desktop folder
#
#
# Connections
Connect-MgGraph -Scopes "User.Read.All, DeviceManagementManagedDevices.Read.All, Directory.Read.All, User.ReadBasic.All, UserAuthenticationMethod.Read.All, AuditLog.Read.All" -NoWelcome
Connect-ExchangeOnline

# Get Dates for use in main loop
$CurrentDate = Get-Date
$OneMonthAgo = $CurrentDate.AddMonths(-1)

# Get all subscribed SKUs once
$SubscribedSkus = Get-MgSubscribedSku

# Build a hashtable of SkuId to SkuPartNumber
$SkuMap = @{}
foreach ($Sku in $SubscribedSkus)
{
  $SkuMap[$Sku.SkuId] = $Sku.SkuPartNumber
}

# Function to get the last sign-in date
function Get-LastSignInDate
{
  param (
    [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser]$User
  )
  # Use the existing signInActivity from the user object
  $LastSignInDate = $User.signInActivity.lastSignInDateTime

  if ($null -eq $LastSignInDate)
  {
    $LastSignInDate = "Never"
  } else
  {
    $LastSignInDate = [datetime]$LastSignInDate
  }

  return $LastSignInDate
}

# Function to get the last sent message date
function Get-LastSentMessageDate
{
  param (
    [string]$MailId
  )
    
  # Initialize $LastSentTime
  $LastSentTime = $null

  # Check if the user has a mailbox
  $HasMailbox = Get-Recipient -Identity $MailId -ErrorAction SilentlyContinue

  if ($null -ne $HasMailbox)
  {
    try
    {
      # Use Get-MailboxFolderStatistics to get the Sent Items folder statistics
      $MailboxStats = Get-MailboxFolderStatistics -Identity $MailId -FolderScope SentItems -IncludeOldestAndNewestItems -ResultSize 5 -ErrorAction Stop

      # Get the 'Sent Items' folder
      $SentItemsFolder = $MailboxStats | Where-Object { $_.FolderType -eq 'SentItems' }

      if ($null -ne $SentItemsFolder)
      {
        $LastSentTime = $SentItemsFolder.NewestItemReceivedDate

        if ($null -eq $LastSentTime)
        {
          $LastSentTime = "No Sent Emails"
        }
      } else
      {
        $LastSentTime = "No Sent Items Folder"
      }
    } catch
    {
      Write-Host "Error accessing mailbox for '$MailId': $($_.Exception.Message)"
      $LastSentTime = "Error"
    }
  } else
  {
    Write-Host "User $MailId does not have a mailbox."
    $LastSentTime = "No Mailbox"
  }

  return $LastSentTime
}

# Function to get assigned licenses
function Get-AssignedLicenses
{
  param (
    [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser]$User,
    [Hashtable]$SkuMap
  )

  # Define AssignedLicenses
  $AssignedLicenses = $User.AssignedLicenses

  if ($AssignedLicenses -and $AssignedLicenses.Count -gt 0)
  {
    # Extract the SKU IDs from the assigned licenses
    $skuIds = $AssignedLicenses | Select-Object -ExpandProperty SkuId

    # Map SKU IDs to SKU Part Numbers using the hashtable
    $skuPartNumbers = $skuIds | ForEach-Object {
      $skuId = $_
      $skuPartNumber = $SkuMap[$skuId]
      if ($null -eq $skuPartNumber)
      {
        $skuPartNumber = $skuId  # Use the SKU ID if Part Number not found
      }
      $skuPartNumber
    }
    return ($skuPartNumbers -join "; ")
  } else
  {
    return "No Licenses Assigned"
  }
}

# Function to get MFA status for a user
function Get-MfaStatus
{
  param (
    [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser]$User
  )

  try
  {
    # Retrieve the Microsoft Authenticator methods for the user
    $authenticatorMethods = Get-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $User.Id

    # Retrieve the Software OATH methods for the user
    $oauthMethods = Get-MgUserAuthenticationSoftwareOathMethod -UserId $User.Id

    # Check if the user has the Microsoft Authenticator app or Software OATH tokens registered
    $hasMfa = if ($authenticatorMethods -or $oauthMethods)
    { "Yes" 
    } else
    { "No" 
    }

    # Determine MFA Type
    if ($authenticatorMethods -and $oauthMethods)
    {
      $mfaType = "App/Token"
    } elseif ($oauthMethods)
    {
      $mfaType = "Token"
    } elseif ($authenticatorMethods)
    {
      $mfaType = "App"
    } else
    {
      $mfaType = "SMS/None"
    }

  } catch
  {
    Write-Host "Failed to retrieve MFA methods for user: $($User.UserPrincipalName)"
    $hasMfa = "Error"
    $mfaType = "Unknown"
  }

  return @{
    HasMfa  = $hasMfa
    MFAType = $mfaType
  }
}


# Function to get Enrolled Devices
function Get-UserEnrolledDevices
{
  param(
    [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser]$User
  )

  # Initialize an array to hold the user's device names
  $DeviceNames = @()

  try
  {
    # Get all managed devices where the user is the primary user
    $Devices = Get-MgDeviceManagementManagedDevice -Filter "userPrincipalName eq '$($User.UserPrincipalName)'"

    foreach ($Device in $Devices)
    {
      # Add the device name to the array
      $DeviceNames += $Device.DeviceName
    }

  } catch
  {
    Write-Warning "Failed to retrieve devices for user $($User.UserPrincipalName): $_"
  }

  # Join the device names into a single string
  $DeviceList = $DeviceNames -join "; "

  return $DeviceList
}

# Initialize arrays to hold the results
$InactiveResults = @()
$ActiveResults = @()

# Get all users
$Users = Get-MgUser -All -Select "id,displayName,userPrincipalName,signInActivity,assignedLicenses"

# Main foreach loop to get user data
foreach ($User in $Users)
{
  Write-Host "Processing user: $($User.DisplayName) ($($User.UserPrincipalName))"
    
  # Get the assigned licenses for the user
  $Licenses = Get-AssignedLicenses -User $User -SkuMap $SkuMap

  # Get Last Sign-In Date
  $LastSignInDate = Get-LastSignInDate -User $User

  # Get UserPrincipalName
  $MailId = $User.UserPrincipalName

  # Get MFA Status
  $MfaStatus = Get-MfaStatus -User $User
  $HasMfa = $MfaStatus.HasMfa
  $MFAType = $MfaStatus.MFAType
  
  # Get Devices
  $Devices = Get-UserEnrolledDevices -User $User

  if ($LastSignInDate -eq "Never" -or $LastSignInDate -lt $OneMonthAgo)
  {
    # Get Last Sent Message Date
    $LastSentTime = Get-LastSentMessageDate -MailId $MailId

    # Create a custom object with user info for inactive users
    $UserInfo = [PSCustomObject]@{
      DisplayName       = $User.DisplayName
      UserPrincipalName = $MailId
      LastSignInDate    = $LastSignInDate
      LastSentMessage   = $LastSentTime
      Licenses          = $Licenses
      HasMfa            = $HasMfa
      MFAType           = $MFAType
    }

    # Add to inactive results
    $InactiveResults += $UserInfo
  } else
  {
    # Create a custom object with user info for active users (exclude LastSignInDate and LastSentMessage)
    $UserInfo = [PSCustomObject]@{
      DisplayName       = $User.DisplayName
      UserPrincipalName = $MailId
      Licenses          = $Licenses
      HasMfa            = $HasMfa
      MFAType           = $MFAType
      Devices           = $Devices
    }

    # Add to active results
    $ActiveResults += $UserInfo
  }
}
# Sort inactive results by sign in date
$InactiveResults = $InactiveResults | Sort-Object -Property {
  if ($_.LastSignInDate -eq "Never")
  {
    [datetime]::MinValue  # Treat "Never" as the earliest date
  } elseif ($_.LastSignInDate -is [datetime])
  {
    $_.LastSignInDate
  } else
  {
    [datetime]::MinValue
  }
} -Descending

# Results path creation
$currentlocation = Get-Location
$resultdir = "\Results"
$reportlocation = [string]::Concat($currentlocation, $resultdir)
mkdir $reportlocation

# Export the inactive users to a CSV file
$InactiveResults | Export-Csv -Path "$reportlocation\InactiveUsers.csv" -NoTypeInformation

# Export the active users to a separate CSV file
$ActiveResults | Export-Csv -Path "$reportlocation\ActiveUsers.csv" -NoTypeInformation

Write-Host "Reports generated: InactiveUsers.csv and ActiveUsers.csv"
