# Connect to services
Connect-MgGraph -Scopes "User.Read.All", "DeviceManagementManagedDevices.Read.All", "Directory.Read.All", "User.ReadBasic.All", "UserAuthenticationMethod.Read.All", "AuditLog.Read.All", "Policy.Read.All" -NoWelcome
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

# Function to get assigned licenses
function Get-AssignedLicenses
{
  param (
    [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser]$User,
    [hashtable]$SkuMap
  )

  $LicenseNames = @()

  foreach ($License in $User.AssignedLicenses)
  {
    $SkuId = $License.SkuId
    if ($SkuMap.ContainsKey($SkuId))
    {
      $LicenseNames += $SkuMap[$SkuId]
    } else
    {
      $LicenseNames += $SkuId
    }
  }

  return $LicenseNames -join ", "
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

  $LastSentTime = $null

  $HasMailbox = Get-Recipient -Identity $MailId -ErrorAction SilentlyContinue

  if ($null -ne $HasMailbox)
  {
    try
    {
      $MailboxStats = Get-MailboxFolderStatistics -Identity $MailId -FolderScope SentItems -IncludeOldestAndNewestItems -ResultSize 5 -ErrorAction Stop

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


# Function to get enrolled devices
function Get-UserEnrolledDevices
{
  param(
    [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser]$User
  )

  $DeviceNames = @()

  try
  {
    $Devices = Get-MgDeviceManagementManagedDevice -Filter "userPrincipalName eq '$($User.UserPrincipalName)'"

    foreach ($Device in $Devices)
    {
      $DeviceNames += $Device.DeviceName
    }

  } catch
  {
    Write-Warning "Failed to retrieve devices for user $($User.UserPrincipalName): $_"
  }

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
  
  if ($User.UserPrincipalName -like "*#EXT#*")
  {
    Write-Host "Skipping external user: $($User.DisplayName) ($($User.UserPrincipalName))"
    continue
  } 
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

  if ($LastSignInDate -eq "Never" -or ($LastSignInDate -is [datetime] -and $LastSignInDate -lt $OneMonthAgo))
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
    # Create a custom object with user info for active users
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

# Sort inactive results by sign-in date
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
$reportlocation = Join-Path -Path $currentlocation -ChildPath "Results"

if (-not (Test-Path -Path $reportlocation))
{
  New-Item -Path $reportlocation -ItemType Directory | Out-Null
}

# Export the inactive users to a CSV file
$InactiveReportPath = Join-Path -Path $reportlocation -ChildPath "InactiveUsers.csv"
$InactiveResults | Export-Csv -Path $InactiveReportPath -NoTypeInformation

# Export the active users to a separate CSV file
$ActiveReportPath = Join-Path -Path $reportlocation -ChildPath "ActiveUsers.csv"
$ActiveResults | Export-Csv -Path $ActiveReportPath -NoTypeInformation

Write-Host "Reports generated: InactiveUsers.csv and ActiveUsers.csv. They are in your results folder of your current working directory."
