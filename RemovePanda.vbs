strIn = "C:\temp\scripts\servers.txt"
strOut = "C:\temp\scripts\servers_log.txt"

'-----------------------------------------------------------------------
Const ForReading = 1
Const ForAppending = 8

strAccount = InputBox("Enter account") 
strPass = InputBox("Enter password") 

If (strPass = "") Or ( strAccount = "" ) Then WScript.Quit

Set objFSOl = CreateObject("Scripting.FileSystemObject")
Set objInFile = objFSOl.OpenTextFile(strIn , ForReading)
'If objFSOl.FileExists(strOut) Then objFSOl.DeleteFile(strOut)
Set objOutFile = objFSOl.OpenTextFile(strOut , ForAppending, True)

Do Until objInFile.AtEndOfStream 
	strComputer = objInFile.ReadLine
	
	Wscript.Echo ">> Working on " & strComputer
	objOutFile.WriteLine("Working on " & strComputer ) 
'	call PingIt(strComputer)' ping it 
'	If goodPing Then 
		Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
		Set objWMIService = objSWbemLocator.ConnectServer _
		    (strComputer, "root\cimv2", strAccount, strPass)
		objWMIService.Security_.ImpersonationLevel = 3 
		Call uninstall(strComputer)
		Call Delfolder(strComputer)
		Call Delfolder1(strComputer)
'	End If 
Loop
objOutFile.Close

' ------------------------ SUBS ----------------------------
Sub PingIt(strComputer)
	Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
        ExecQuery("select * from Win32_PingStatus where address = '"_
            & strComputer & "'")
    goodPing = True        
    For Each objStatus in objPing
        If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then
        	goodPing = False
        	objOutFile.WriteLine("!! Computer " & strComputer & " is not reachable") 
            WScript.Echo("Computer " & strComputer & " is not reachable") 
        End If
    Next
End Sub

Sub uninstall(strComputer)  
	Set colSoftware = objWMIService.ExecQuery _
	    ("Select * from Win32_Product")
	goodUins = False 
	For Each objSoftware in colSoftware
		If Left(objSoftware.Caption, 5) = "Panda" Then  
		    WScript.Echo "Unistalling from " & strComputer & " " & _
		    objSoftware.Name & vbtab & _
		    objSoftware.Vendor & vbtab & _
		    objSoftware.Version
		    objOutFile.WriteLine("Unistalling from " & strComputer & " " & _
		    objSoftware.Name & vbtab & _
		    objSoftware.Vendor & vbtab & _
		    objSoftware.Version)
		    objSoftware.Uninstall()
		    goodUins = True
		End if 
	Next 
	If goodUins Then 
	Else
		objOutFile.WriteLine("!! No Panda software found on " & strComputer )
		WScript.Echo "No Panda software found on " & strComputer   
    End if
End Sub

Sub Delfolder(strComputer)
	Set colFolders = objWMIService.ExecQuery _
	    ("Select * from Win32_Directory where Name = 'c:\\Program Files (x86)\\Panda Security'")
	goodDelete = False
	For Each objFolder in colFolders
	
	    errResults = objFolder.Delete
	    goodDelete = True
	    objOutFile.WriteLine("Panda folder (x86) deleted " & strComputer & ". Error code=" & errResults )
		WScript.Echo "Panda folder (x86) deleted " & strComputer & ". Error code=" & errResults
	Next
	If goodDelete Then 
	Else
		objOutFile.WriteLine("!! No Panda (x86) folder found on " & strComputer )
		WScript.Echo "No Panda (x86) folder found on " & strComputer   
    End if
End Sub

Sub Delfolder1(strComputer)
	Set colFolders = objWMIService.ExecQuery _
	    ("Select * from Win32_Directory where Name = 'c:\\Program Files\\Panda Security'")
	goodDelete = False
	For Each objFolder in colFolders
	
	    errResults = objFolder.Delete
	    goodDelete = True
	    objOutFile.WriteLine("Panda folder deleted " & strComputer & ". Error code=" & errResults )
		WScript.Echo "Panda folder deleted " & strComputer & ". Error code=" & errResults
	Next
	If goodDelete Then 
	Else
		objOutFile.WriteLine("!! No Panda folder found on " & strComputer )
		WScript.Echo "No Panda folder found on " & strComputer   
    End if
End Sub