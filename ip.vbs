strQuery = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE MACAddress > ''"

Set objWMIService = GetObject( "winmgmts://./root/CIMV2" )
Set colItems      = objWMIService.ExecQuery( strQuery, "WQL", 48 )

For Each objItem In colItems
    If IsArray( objItem.IPAddress ) Then
        If UBound( objItem.IPAddress ) = 0 Then
            strIP = "IP Address: " & objItem.IPAddress(0)
        Else
            strIP = "IP Addresses: " & Join( objItem.IPAddress, "," )
        End If
    End If
Next

WScript.Echo strIP