' List Installed Software


'Set objFSO = CreateObject("Scripting.FileSystemObject")
'If objFSO.FileExists("C:\temp\scripts\software.tsv" ) Then objFSO.DeleteFile("C:\temp\scripts\software.tsv" )
'Set objTextFile = objFSO.CreateTextFile("C:\temp\scripts\software.tsv", True)

strComputer = "uschi2k8admon"

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objSWbemLocator.ConnectServer _
    (strComputer, "root\cimv2", "d0c0rudakove", "@batman01")
objWMIService.Security_.ImpersonationLevel = 3   
    
Set colSoftware = objWMIService.ExecQuery _
    ("Select * from Win32_Product")

'objTextFile.WriteLine "Caption" & vbtab & _
'    "Description" & vbtab & "Identifying Number" & vbtab & _
'    "Install Date" & vbtab & "Install Location" & vbtab & _
'    "Install State" & vbtab & "Name" & vbtab & _ 
'    "Package Cache" & vbtab & "SKU Number" & vbtab & "Vendor" & vbtab _
'        & "Version" 

For Each objSoftware in colSoftware
	If Left(objSoftware.Caption, 4) = "Pand" Then  
'	    objTextFile.WriteLine objSoftware.Caption & vbtab & _
	   
	    WScript.Echo "Unistalling from " & strComputer & " " & _
	    objSoftware.Name & vbtab & _
	    objSoftware.Vendor & vbtab & _
	    objSoftware.Version
	    objSoftware.Uninstall()
    End if
Next
'objTextFile.Close