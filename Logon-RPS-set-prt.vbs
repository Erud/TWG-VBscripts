Option Explicit
Dim objADSystemInfo, objUser, objMemberOf, objGroup, objNetwork, objPrinter 
dim arGroup, strUser, aCN, aParam, aPrinter, strPrinter
Dim colGroups, dictGroups

on error resume next 
Set dictGroups = CreateObject("Scripting.Dictionary") 
Set objNetwork = CreateObject("Wscript.Network") 
 
'Get current user info from active directory 
Set objADSystemInfo = CreateObject("ADSystemInfo")
'bind to current user in active directory
strUser =  objADSystemInfo.UserName
set objUser = GetObject("LDAP://" & objADSystemInfo.UserName) 

'Get an array of group names that the user is a member of 
Set colGroups = objUser.Groups

For Each objGroup in colGroups
    If Not dictGroups.Exists(objGroup.distinguishedName) Then
	    dictGroups.Add objGroup.distinguishedName, objGroup.CN
   		GetNested(objGroup)
    End If
Next
objMemberOf = dictGroups.Keys

'objMemberOf = objUser.MemberOf 
for Each objGroup in objMemberOf  
	arGroup = Split(objGroup,",OU=")
    'Test to see if it is a printer group. all printer groups should be in the same OU 
    If (UBound(arGroup) > 0) Then 
    	If (arGroup(1) = "Printer Groups") Then
    		aCN = Split(arGroup(0), "=")
    		aParam = Split(aCN(1), "\;")
    		aPrinter = Split(aParam(0),"\\")
    		strPrinter = "\\" & aPrinter(2) & "\" & aPrinter(3)
    		ReDim Preserve aParam(3)
'    		WScript.Echo strPrinter & " 1 " & aParam(1) & " 2 " & aParam(2)
 	 	    objNetwork.AddWindowsPrinterConnection strPrinter 
' 	        objNetwork.SetDefaultPrinter strPrinter  'set default printer
    	End If
    End If 
next 

' CN=\\\\rpsushc1ps02vp\\RPSFLTP2C4000\;\;C,OU=Printer Groups,OU=RPS Groups,OU=RPS,DC=amerdev,DC=ajgcodev,DC=int

' ======================================================================================================
Function GetNested(objGroup)
	Dim colMembers, strMember, strPath, objNestedGroup 
    On Error Resume Next
    colMembers = objGroup.GetEx("memberOf")
    For Each strMember in colMembers
        strPath = "LDAP://" & strMember
        Set objNestedGroup = GetObject(strPath)
        dictGroups.Add objNestedGroup.distinguishedName, objNestedGroup.CN
        GetNested(objNestedGroup)
    Next
End Function