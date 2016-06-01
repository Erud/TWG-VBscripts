' List the Subfolders of a Folder on the Local Computer

strComputer = "."
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set colDrives = objFSO.Drives

For Each objDrive in colDrives
	If objDrive.DriveType = 2 Then 
	
	strDrive = objDrive.DriveLetter & ":\."
	Set objFolder = objFSO.GetFolder(strDrive)
	Set colSubfolders = objFolder.Subfolders
	 
	For Each objSubfolder in colSubfolders
		
		
		If objSubfolder.Attributes = 16 Then
	    
	    	WScript.Echo objDrive.DriveLetter & ":\" & objSubfolder.Name
	     
	    End If 
	Next
	End if
Next

' list shares 

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")

	Wscript.Echo " "
	Wscript.Echo "Caption " & vbTab & "Name " & vbTab &  "Path " 
	
For each objShare in colShares
    If objShare.Type = 0 Then ' Disk share
    	Wscript.Echo objShare.Caption & vbTab & objShare.Name & vbTab & objShare.Path   
    End if
Next