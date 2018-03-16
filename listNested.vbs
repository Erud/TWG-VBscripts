Option Explicit
Dim objADSystemInfo, objUser, objMemberOf, objGroup, objNetwork, objPrinter 
dim arGroup, strUser, aCN, aParam, aPrinter, strPrinter
Dim colGroups, dictGroups, i
'on error resume next 
 
Set objNetwork = CreateObject("Wscript.Network") 
Set dictGroups = CreateObject("Scripting.Dictionary")
 
'Get current user info from active directory 
Set objADSystemInfo = CreateObject("ADSystemInfo")
'bind to current user in active directory
strUser =  objADSystemInfo.UserName
set objUser = GetObject("LDAP://" & objADSystemInfo.UserName) 

'Get an array of group names that the user is a member of 
'objMemberOf = objUser.MemberOf 
Set colGroups = objUser.Groups

For Each objGroup in colGroups
    Wscript.Echo objGroup.distinguishedName
    If Not dictGroups.Exists(objGroup.distinguishedName) Then
	    dictGroups.Add objGroup.distinguishedName, objGroup.CN
   		GetNested(objGroup)
    End If
Next
WScript.Echo "=== dict ==="
objMemberOf = dictGroups.Keys
For i = 0 To dictGroups.Count -1 ' Iterate the array.
      WScript.Echo objMemberOf(i)
Next

' ======================================================================================================
Function GetNested(objGroup)
	Dim colMembers, strMember, strPath, objNestedGroup 
    On Error Resume Next
    colMembers = objGroup.GetEx("memberOf")
    For Each strMember in colMembers
    	'WScript.Echo "nested <>",strMember
        strPath = "LDAP://" & strMember
        Set objNestedGroup = GetObject(strPath)
        WScript.Echo "nested ",objNestedGroup.distinguishedName
        dictGroups.Add objNestedGroup.distinguishedName, objNestedGroup.CN
        GetNested(objNestedGroup)
    Next
End Function