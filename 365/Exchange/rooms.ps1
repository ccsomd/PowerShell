# Authenticate with 365
Connect-ExchangeOnline -UserPrincipalName <Username>

# List all RoomLists
Get-DistributionGroup -ResultSize Unlimited | Where {$_.RecipientTypeDetails -eq "RoomList"} | Format-Table DisplayName,Identity,PrimarySmtpAddress –AutoSize

# List all Rooms
Get-Mailbox -ResultSize Unlimited -Filter {RecipientTypeDetails -eq "RoomMailbox"} | Select-Object Name, Alias, PrimarySmtpAddress, DisplayName, RecipientTypeDetails, ResourceCapacity


# New Room List
New-Distributiongroup -Name <RoomListName> -RoomList -PrimarySmtpAddress <RoomListEmail>

# New Room
New-Mailbox -Name <RoomName> -Room

# Add Room to Room List
Add-DistributionGroupMember -Identity <RoomListName> -Member <RoomMailboxUPN>
