'On Error Resume Next
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array("0219-MAYERSKYL")
For Each strComputer In arrComputers

	Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
	Set objWMIService = objSWbemLocator.ConnectServer(strComputer, "root\CIMV2", "d0c0rudakove", "@batman01")
	objWMIService.Security_.ImpersonationLevel = 3
	
	Set Items = objWMIService.ExecQuery("Select USername From Win32_ComputerSystem")
	For Each obj in Items
    	aUser = Split(obj.USername, "\") 
	Next
	Wscript.Echo strComputer & " " & aUser(1) 
	
	Set objAccount = objWMIService.Get ("Win32_UserAccount.Name='mayerskyl',Domain='chi.warranty.local'")
 Wscript.Echo objAccount.SID
	'------------------------------------------------------------------------
	
	Set ColItems = objWMIService.ExecQuery("Select * from Win32_UserAccount Where LocalAccount = True")

   For Each objItem In colItems
      WScript.Echo "Name: " & objItem.Name & " SID: " & objItem.SID
      
   Next
Next

