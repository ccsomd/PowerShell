# Authenticate with 365
Connect-ExchangeOnline -UserPrincipalName #<username>

# List all RoomLists
Get-DistributionGroup -ResultSize Unlimited | Where {$_.RecipientTypeDetails -eq "RoomList"} | Format-Table DisplayName,Identity,PrimarySmtpAddress –AutoSize

# List all Rooms
Get-Mailbox -ResultSize Unlimited -Filter {RecipientTypeDetails -eq "RoomMailbox"} | Select-Object Name, Alias, PrimarySmtpAddress, DisplayName, RecipientTypeDetails, ResourceCapacity


# New Room List
New-Distributiongroup -Name 'Room List Name' -RoomList -PrimarySmtpAddress "Room-Name@windows.onmicrosoft.com"

# New Room
New-Mailbox -Name "New Room Name" -Room

# Add Room to Room List
Add-DistributionGroupMember -Identity <RoomListName> -Member <RoomMailboxUPN>
