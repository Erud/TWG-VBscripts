Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand =   CreateObject("ADODB.Command")
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000

strUser = "LOPEZA"

objCommand.CommandText = "<LDAP://DC=chi,DC=warranty,DC=local>;(&(objectCategory=person)(sAMAccountName=" & strUser & ")); name,objectSid;subtree"
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
Set oGroup = objCommand.Execute


arrbytSID = oGroup.Fields("objectSid")

' Convert byte array to hex string.
strHexSID = OctetToHexStr(arrbytSID)

' Convert hex SID values to decimal format.
strDecSID = HexSIDToDec(strHexSID)
Wscript.Echo "SID " & strDecSID

Function OctetToHexStr(ByVal arrbytOctet)
    ' Function to convert OctetString (byte array) to Hex string.

    Dim k

    OctetToHexStr = ""
    For k = 1 To Lenb(arrbytOctet)
        OctetToHexStr = OctetToHexStr _
            & Right("0" & Hex(Ascb(Midb(arrbytOctet, k, 1))), 2)
    Next

End Function

Function HexSIDToDec(ByVal strSID)
    ' Function to convert most hex SID values to decimal format.

    Dim arrbytSID, lngTemp, j

    ReDim arrbytSID(Len(strSID)/2 - 1)
    For j = 0 To UBound(arrbytSID)
        arrbytSID(j) = CInt("&H" & Mid(strSID, 2*j + 1, 2))
    Next

    If (UBound(arrbytSID) = 11) Then
        HexSIDToDec = "S-" & arrbytSID(0) & "-" _
            & arrbytSID(1) & "-" & arrbytSID(8)

        Exit Function
    End If

    If (UBound(arrbytSID) = 15) Then
        HexSIDToDec = "S-" & arrbytSID(0) & "-" _
            & arrbytSID(1) & "-" & arrbytSID(8)

        lngTemp = arrbytSID(15)
        lngTemp = lngTemp * 256 + arrbytSID(14)
        lngTemp = lngTemp * 256 + arrbytSID(13)
        lngTemp = lngTemp * 256 + arrbytSID(12)

        HexSIDToDec = HexSIDToDec & "-" & CStr(lngTemp)

        Exit Function
    End If

    HexSIDToDec = "S-" & arrbytSID(0) & "-" _
        & arrbytSID(1) & "-" & arrbytSID(8)

    lngTemp = arrbytSID(15)
    lngTemp = lngTemp * 256 + arrbytSID(14)
    lngTemp = lngTemp * 256 + arrbytSID(13)
    lngTemp = lngTemp * 256 + arrbytSID(12)

    HexSIDToDec = HexSIDToDec & "-" & CStr(lngTemp)

    lngTemp = arrbytSID(19)
    lngTemp = lngTemp * 256 + arrbytSID(18)
    lngTemp = lngTemp * 256 + arrbytSID(17)
    lngTemp = lngTemp * 256 + arrbytSID(16)

    HexSIDToDec = HexSIDToDec & "-" & CStr(lngTemp)

    lngTemp = arrbytSID(23)
    lngTemp = lngTemp * 256 + arrbytSID(22)
    lngTemp = lngTemp * 256 + arrbytSID(21)
    lngTemp = lngTemp * 256 + arrbytSID(20)

    HexSIDToDec = HexSIDToDec & "-" & CStr(lngTemp)

    If (UBound(arrbytSID) > 23) Then
        lngTemp = arrbytSID(27)
        lngTemp = lngTemp * 256 + arrbytSID(26)
        lngTemp = lngTemp * 256 + arrbytSID(25)
        lngTemp = lngTemp * 256 + arrbytSID(24)

        HexSIDToDec = HexSIDToDec & "-" & CStr(lngTemp)
    End If

End Function

