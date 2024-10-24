# Connect to services
Connect-MgGraph -Scopes "User.Read.All", "DeviceManagementManagedDevices.Read.All", "Directory.Read.All", "User.ReadBasic.All", "UserAuthenticationMethod.Read.All", "AuditLog.Read.All", "Policy.Read.All" -NoWelcome
Connect-ExchangeOnline

# Get Dates for use in main loop
$CurrentDate = Get-Date
$OneMonthAgo = $CurrentDate.AddMonths(-1)

# Function to get assigned licenses
function Get-UserLicenseInfo
{
  param (
    [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser]$User
  )
  # Making sure variables are grabbed only once here 
  if (-not $script:SubscribedSkus)
  {
    Write-Host "Retrieving Subscribed SKU"
    $script:SubscribedSkus = Get-MgSubscribedSku

    # Build the SkuMap
    $script:SkuMap = @{}
    foreach ($Sku in $script:SubscribedSkus)
    {
      $script:SkuMap[$Sku.SkuId] = $Sku.SkuPartNumber
    }
  }
  
  # Custom sku map
  $customSKUMap = @{
    "SBP" = "Business Premium"
    "FLOW_FREE" = "Power Automate Free"
    "O365_BUSINESS_PREMIUM" = "Business Premium"
    "O365_BUSINESS_ESSENTIALS" = "Business Essentials"
    "ENTERPRISEPACK" = "Office E3"
    "AAD_PREMIUM_P2" = "Azure AD Premium P2"
  }
  # Get assigned licenses for the user
  $AssignedLicenses = $User.AssignedLicenses

  # Extract the SKU IDs from the assigned licenses
  $skuIds = $AssignedLicenses | Select-Object -ExpandProperty SkuId

  # Map SKU IDs to SKU Part Numbers 
  $friendlyLicenseName = $skuIds | ForEach-Object {
    $skuId = $_
    $skuPartNumber = $script:SkuMap[$skuId]
    if ($null -eq $skuPartNumber)
    {
      $skuPartNumber = $skuId  # Use the SKU ID if Part Number not found
    }
    $friendlyLicenseName = $customSKUMap[$skuPartNumber]
    if ($null -ne $friendlylicenseName)
    {
      $skuPartNumber = $friendlyLicenseName
    }
    $friendlyLicenseName
  }

  # Format License table
  $License = ($friendlyLicenseName -join "; ")

  if ($null -eq $License -or $License -eq "")
  {
    $License = "No License"
  } else
  {
    return $License
  }
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


# Function to convert bytes to a readable size/
function Convert-BytesToReadableSize
{
  param (
    [int64]$Bytes
  )

  switch ($Bytes)
  {
    {$_ -ge 1PB}
    { "{0:N2} PB" -f ($Bytes / 1PB); break 
    }
    {$_ -ge 1TB}
    { "{0:N2} TB" -f ($Bytes / 1TB); break 
    }
    {$_ -ge 1GB}
    { "{0:N2} GB" -f ($Bytes / 1GB); break 
    }
    {$_ -ge 1MB}
    { "{0:N2} MB" -f ($Bytes / 1MB); break 
    }
    {$_ -ge 1KB}
    { "{0:N2} KB" -f ($Bytes / 1KB); break 
    }
    default
    { "{0:N2} Bytes" -f $Bytes 
    }
  }
}

function Get-MailboxSize
{
  param (
    [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser]$User
  )
  # Get primary mailbox statistics
  $PrimaryStats = Get-MailboxStatistics -Identity $User.UserPrincipalName
  $sizeInMB = [math]::Round($PrimaryStats.TotalItemSize.Value.ToMB(), 2)
  return $sizeInMB

  # Extract total size of the primary mailbox
  $PrimaryMailboxSizeString = $PrimaryStats.TotalItemSize.Value.ToString()

  # Extract bytes from PrimaryMailboxSizeString, handling commas
  if ($PrimaryMailboxSizeString -match '\(([\d,]+) bytes\)')
  {
    $BytesString = $Matches[1]
    $BytesStringClean = $BytesString -replace ',', ''
    $PrimaryMailboxBytes = [int64]$BytesStringClean
  } else
  {
    $PrimaryMailboxBytes = 0
  }

  # Initialize archive variables
  $ArchiveEnabled = $false
  $ArchiveMailboxSizeString = "N/A"
  $ArchiveMailboxBytes = 0

  # Check if archive is enabled
  if ($Mailbox.ArchiveStatus -eq "Active")
  {
    $ArchiveEnabled = $true

    # Get archive mailbox statistics
    $ArchiveStats = Get-MailboxStatistics -Identity $Mailbox.Identity -Archive

    # Extract total size of the archive mailbox
    $ArchiveMailboxSizeString = $ArchiveStats.TotalItemSize.Value.ToString()

    # Extract bytes from ArchiveMailboxSizeString, handling commas
    if ($ArchiveMailboxSizeString -match '\(([\d,]+) bytes\)')
    {
      $BytesString = $Matches[1]
      $BytesStringClean = $BytesString -replace ',', ''
      $ArchiveMailboxBytes = [int64]$BytesStringClean
    } else
    {
      $ArchiveMailboxBytes = 0
    }
  }
}
# Calculate Total Mailbox Bytes
$TotalMailboxBytes = $PrimaryMailboxBytes + $ArchiveMailboxBytes

# Convert bytes to readable sizes
$PrimaryMailboxSizeReadable = Convert-BytesToReadableSize -Bytes $PrimaryMailboxBytes
$ArchiveMailboxSizeReadable = if ($ArchiveEnabled)
{ Convert-BytesToReadableSize -Bytes $ArchiveMailboxBytes 
} else
{ "N/A" 
}
$TotalMailboxSizeReadable   = Convert-BytesToReadableSize -Bytes $TotalMailboxBytes

# Add data to the report
return [PSCustomObject]@{
  PrimaryMailboxSizeReadable = $PrimaryMailboxSizeReadable
  ArchiveEnabled             = $ArchiveEnabled
  ArchiveMailboxSizeReadable = $ArchiveMailboxSizeReadable
  TotalMailboxSizeReadable   = $TotalMailboxSizeReadable
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
  $Licenses = Get-UserLicenseInfo -User $User -SkuMap $SkuMap

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
      PrimaryMailboxSizeReadable = $PrimaryMailboxSizeReadable
      ArchiveEnabled             = $ArchiveEnabled
      ArchiveMailboxSizeReadable = $ArchiveMailboxSizeReadable
      TotalMailboxSizeReadable   = $TotalMailboxSizeReadable


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
    Get-MailboxSize -User $user
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
