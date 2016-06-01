On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array("1004-MORAF")
For Each strComputer In arrComputers

   Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
	Set objWMIService = objSWbemLocator.ConnectServer(strComputer, "root\CIMV2", "d0c0rudakove", "@batman01")
   name = "openscape"
'   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_UserAccount WHERE name='" & name & "' ", "WQL", _
'	                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)
'Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_UserAccount ", "WQL",
Set WMIUser = objWMIService.Get("Win32_UserAccount.Domain='" & strComputer & "'" & ",Name='" & name & "'")

 Result = WMIUser.SID
 
   For Each objItem In colItems
     
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "Description: " & objItem.Description
     
      WScript.Echo "Domain: " & objItem.Domain
      WScript.Echo "FullName: " & objItem.FullName
      
      WScript.Echo "Name: " & objItem.Name
     
      WScript.Echo "SID: " & objItem.SID
      WScript.Echo "SIDType: " & objItem.SIDType
      
      WScript.Echo
   Next
Next
