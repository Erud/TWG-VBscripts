' Delete a Folder


strComputer = "."

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objSWbemLocator.ConnectServer _
    (strComputer, "root\cimv2", "d0c0rudakove", "@batn")
objWMIService.Security_.ImpersonationLevel = 3 

Set colFolders = objWMIService.ExecQuery _
    ("Select * from Win32_Directory where Name = 'c:\\Program Files (x86)\\Panda Security'")

For Each objFolder in colFolders
	WScript.Echo objFolder.name
    'errResults = objFolder.Delete
    WScript.Echo errResults
Next
