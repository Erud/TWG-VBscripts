' Current User Drives 
' Copyright 2016 Edward Rudakov. All rights reserved.
' Copyright 2012 Clare Thompson. All rights reserved.

'Option Explicit
'on error resume next
Const HKEY_USERS = &H80000003
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20
Const ADS_SCOPE_SUBTREE = 2

' ********** Dim variables
Dim objFolder
Dim objWMIService
Dim strComputer
Dim colSettings
Dim objComputer
Dim strMsg, strMsgBoxTitle, strDirectory, strStatus
Dim objExplorer
Dim objOperatingSystem
Dim objFSO
Dim objWBem
Dim objFileList
Dim objFile
Dim objRegistry
Dim regpath2enumerate
Dim colProc, oProcess, subkey, regpath, regentry, dapath
Dim sUser, sDomain, arrRegKeys, arrKeyNames, lngRtn, strKey, objSID, arrComputers
Dim strSourceList, strSaveToFile, strDrive, strFolder

' ********** User Variables
strMsgBoxTitle = "Drives and Printers"
strDrive = "C"
strFolder = "temp\scripts"
strDirectory = strDrive & ":\" & strFolder
strSourceList = "PCtoTest2.txt"
'strSourceList = "comp1.txt"
strSaveToFile = "PCtoTest2_out.txt"
strUser = "d0c0rudakove"
strPassword = "@batman01"

' ********** Blank the report message
strMsg = ""

' ********** Check for the existence of the "SysInfoCheck" folder and previous txt file
Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FolderExists(strDirectory) Then

	' ********** Open the list
	Const ForReading = 1
	Const ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	' ********** Check the existence of the source file
	If objFSO.FileExists(strDirectory & "\" & strSourceList) Then

		' ********** Delete current TSV file or create new one if not present
		If objFSO.FileExists(strDirectory & "\" & strSaveToFile) Then
    		objFSO.DeleteFile(strDirectory & "\" & strSaveToFile)
		End if
		set objFile = objFSO.OpenTextFile(strDirectory & "\" & strSaveToFile, ForAppending, True)

		' ********** Start processing each text file line

		Set objFileList = objFSO.OpenTextFile(strDirectory & "\" & strSourceList, ForReading, False)
		Do while objFileList.AtEndOfStream = false
    			strComputer = objFileList.ReadLine

		'	WScript.Echo "Processing " & strComputer 
			
			' ********** Set computer object 
			Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
			numErr = 0
			On Error Resume Next 
			Set objSWbemServiceD = objSWbemLocator.ConnectServer(strComputer, "root\default", strUser, strPassword)
'			WScript.Echo "Error # " & CStr(Err.Number) & " " & Err.Description
			numErr = Err.Number
			Err.Clear   ' Clear the error.
			On Error GoTo 0 
		

			' ******** If the computer doesn't respond to WMI then note in output
			Select Case numErr
				Case 462
					WScript.Echo   "======================== " & strComputer & " Unable to connect." 
				Case -2147217394
					WScript.Echo   "======================== " & strComputer & " Invalid Namespace." 
				Case -2147023174
					WScript.Echo   "======================== " & strComputer & " The RPC server is unavailable." 	
				Case 70
					WScript.Echo   "======================== " & strComputer & " Access Denied." 
    			Case Else
					
					Set objRegistry = objSWbemServiceD.Get("StdRegProv")
					Set objWMIService = objSWbemLocator.ConnectServer(strComputer, "root\cimv2", strUser, strPassword)
					objWMIService.Security_.ImpersonationLevel = 3 

					' ********** Get the current user 

					Set Items = objWMIService.ExecQuery("Select USername From Win32_ComputerSystem")
					For Each obj in Items
						strUserN = obj.USername
						nVar = VarType(strUserN) 
    					If nVar = vbString Then aUser = Split(strUserN, "\") 
					Next
				If nVar = vbString Then 	
					sUser = aUser(1)
					sDom  = aUser(0)
					If UCase(sDom) = UCase(strComputer) Then 
						WScript.Echo  "Processing " & strComputer & " local account " & sUser
						Set colLocUsers = objWMIService.ExecQuery("SELECT * FROM Win32_UserAccount WHERE name='" & sUser & "' ", "WQL", _
	                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)
	                    For Each objLocUser In colLocUsers
	                    	strDecSID = objLocUser.SID
	                    Next
					Else
						' get user SID from AD
						WScript.Echo  "Processing " & strComputer 
						Set objConnection = CreateObject("ADODB.Connection")
						objConnection.Provider = "ADsDSOObject"
						objConnection.Open "Active Directory Provider"
						Set objCommand =   CreateObject("ADODB.Command")
						Set objCommand.ActiveConnection = objConnection
						objCommand.Properties("Page Size") = 1000
						
						objCommand.CommandText = "<LDAP://DC=chi,DC=warranty,DC=local>;(&(objectCategory=person)(sAMAccountName=" & sUser & ")); name,objectSid;subtree"
						objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
						Set oGroup = objCommand.Execute
						
						arrbytSID = oGroup.Fields("objectSid")
						' Convert byte array to hex string.
						strHexSID = OctetToHexStr(arrbytSID)
						' Convert hex SID values to decimal format.
						strDecSID = HexSIDToDec(strHexSID)
					End If
					' ********** Loop through the HKEY_USERS hive until the currently logged on user is matched
					lngRtn = objRegistry.EnumKey(HKEY_USERS, "", arrRegKeys)
					
					For Each strKey In arrRegKeys
						If UCase(strKey) = UCase(strDecSID) Then
							regpath2enumerate = strkey & "\Network" 'strkey is the SID
							lngRtn = objRegistry.enumkey (HKEY_USERS, regpath2enumerate, arrKeyNames)
							nVarKey = VarType(arrKeyNames)
							If nVarKey > 1 Then
								For Each subkey In arrkeynames
									regpath = strkey & "\Network\" & subkey
									regentry = "RemotePath"
									objRegistry.getstringvalue hkey_users, regpath, regentry, dapath
									strMsg = strComputer & vbTab & sDom & vbTab & sUser & vbTab & subkey & ":" & vbTab & dapath 
									objFile.WriteLine(strMsg)
								Next
							End If
						End If
					Next 
				End If	
			End Select

			' ********** Append to the file
'			objFile.WriteLine(strMsg)
			strMsg = ""
			strComputer = ""
			err.number = 0

		Loop

		objFileList.Close
		objFile.close
		

	else
		strAnsTxt = "The source list file cannot be found." & VbCrLF & "Please place " & strSourceList & " into the " & strDirectory & " folder."
    		strAnswer=MsgBox(strAnsTxt, 48, strMsgBoxTitle)

	End If

else
	strAnsTxt = "The " & strDirectory & " folder doesn't exist and will be created." & VbCrLF & "You will need to place the " & strSourceList & " file into the folder."
	strAnswer=MsgBox(strAnsTxt, 48, strMsgBoxTitle)
	Set objFolder = objFSO.CreateFolder(strDirectory)

End If

' subs & functions
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

