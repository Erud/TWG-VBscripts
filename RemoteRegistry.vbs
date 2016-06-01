Const HKEY_USERS = &H80000003
strUser = "d0c0rudakove"
strPassword = "@batman01"
strComputer = "2827-LOPEZA"

Set objSWbemLocator = CreateObject("wbemScripting.SwbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, "root\default", strUser, strPassword)
Set objRegistry = objSWbemServices.Get("StdRegProv")
lngRtn = objRegistry.EnumKey(HKEY_USERS, "", arrRegKeys)
Wscript.Echo lngRtn

objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strEntryName, strValue
 Wscript.Echo strValue