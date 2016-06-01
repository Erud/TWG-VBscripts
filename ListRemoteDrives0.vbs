' Current User Drives
' Copyright 2016 Edward Rudakov. All rights reserved.
' Copyright 2012 Clare Thompson. All rights reserved.

'Option Explicit
'on error resume next
Const HKEY_USERS = &H80000003
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20
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
strSourceList = "comp1.txt"
strSaveToFile = "comp_out.txt"
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
			Set objFile = objFSO.CreateTextFile(strDirectory & "\" & strSaveToFile,True)
    			
		Else
    			Set objFile = objFSO.CreateTextFile(strDirectory & "\" & strSaveToFile,True)
    		
		End if


		' ********** Start processing each text file line

		Set objFileList = objFSO.OpenTextFile(strDirectory & "\" & strSourceList, ForReading, False)
		Do while objFileList.AtEndOfStream = false
    			strComputer = objFileList.ReadLine

			WScript.Echo "Processing " & strComputer 
			
			' ********** Set computer object 
			Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
			Set objSWbemServiceD = objSWbemLocator.ConnectServer(strComputer, "root\default", strUser, strPassword)
			Set objRegistry = objSWbemServiceD.Get("StdRegProv")
			Set objWMIService = objSWbemLocator.ConnectServer(strComputer, "root\cimv2", strUser, strPassword)
			objWMIService.Security_.ImpersonationLevel = 3 

			' ******** If the computer doesn't respond to WMI then note in output
			Select Case err.number
				Case 462
					strMsg = strMsg & "======================== " & strComputer & " ========================" & VbCrLf & VbCrLf
					strMsg = strMsg & "Unable to connect." & VbCrLf & VbCrLf
					err.number = 0
				Case -2147217394
					strMsg = strMsg & "======================== " & strComputer & " ========================" & VbCrLf & VbCrLf
					strMsg = strMsg & "Invalid Namespace." & VbCrLf & VbCrLf
					err.number = 0
				Case 70
					strMsg = strMsg & "======================== " & strComputer & " ========================" & VbCrLf & VbCrLf
					strMsg = strMsg & "Access Denied." & VbCrLf & VbCrLf
					err.number = 0
    				Case Else
					strMsg = strMsg & "======================== " & strComputer & " ========================" & VbCrLf & VbCrLf


					' ********** Get the current user from Explorer  

					Set colProc = objWmiService.ExecQuery("Select * from Win32_Process Where Name='explorer.exe'", "WQL", _
														wbemFlagReturnImmediately + wbemFlagForwardOnly)
					If colProc.Count > 0 Then
						For Each oProcess In colProc
							oProcess.GetOwner sUser, sDomain
						Next
					End If

					' ********** Print user and computer
					strMsg = strMsg & "    User: " & sUser & VbCrLf
					strMsg = strMsg & "Computer: " & strComputer & VbCrLf & VbCrLf
					

					' ********** Loop through the HKEY_USERS hive until the currently logged on user is matched
					lngRtn = objRegistry.EnumKey(HKEY_USERS, "", arrRegKeys)
					
					For Each strKey In arrRegKeys
						If UCase(strKey) = ".DEFAULT" Or UCase(Right(strKey, 8)) = "_CLASSES" Then
						Else
							Set objSID = objWbem.Get("Win32_SID.SID='" & strKey & "'")
							If objSID.accountname = sUser Then 
								regpath2enumerate = strkey & "\Network" 'strkey is the SID
								objRegistry.enumkey hkey_users, regpath2enumerate, arrkeynames
								
								If Not (IsEmpty(arrkeynames)) Then
									For Each subkey In arrkeynames
										regpath = strkey & "\Network\" & subkey
										regentry = "RemotePath"
										objRegistry.getstringvalue hkey_users, regpath, regentry, dapath
										strMsg = strMsg & subkey & ":" & vbTab & dapath & VbCrLf
									Next
								End If
							End If
						End If
					Next
					
					strMsg = strMsg & VbCrLf & VbCrLf 
					
			End Select

			' ********** Append to the file
			set objFile = objFSO.OpenTextFile(strDirectory & "\" & strSaveToFile, ForAppending, True)
			objFile.WriteLine(strMsg)
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

End if