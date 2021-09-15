Get-CimInstance -Class Win32_Product -Filter "IdentifyingNumber='{36446A3B-187C-4698-80D2-1310B72F51A6}'" | Invoke-CimMethod -MethodName Uninstall
