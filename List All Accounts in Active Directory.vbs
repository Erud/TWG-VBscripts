' List All Computer Accounts in Active Directory


Const ADS_SCOPE_SUBTREE = 2
Const ADS_UF_ACCOUNTDISABLE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCOmmand.ActiveConnection = objConnection
'objCommand.CommandText = _
'    "Select Name, Location from 'LDAP://DC=chi,DC=warranty,DC=local' " _
'        & "Where objectClass='person'"  
strBase = "<LDAP://DC=chi,DC=warranty,DC=local>"

 ' Filter on user objects.
 strFilter = "(&(objectCategory=person)(objectClass=user))"

 ' Comma delimited list of attribute values to retrieve.
 strAttributes = "sAMAccountName,cn,distinguishedName,ProfilePath,ScriptPath,HomeDirectory,HomeDrive,userAccountControl"

 ' Construct the LDAP syntax query.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
objCommand.CommandText = strQuery
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
	intUAC=objRecordSet.Fields("userAccountControl")
    If intUAC AND ADS_UF_ACCOUNTDISABLE Then 
    	strDisabled = "Y"
    Else
    	strDisabled = ""
    End If
        
    Wscript.Echo objRecordSet.Fields("sAMAccountName").Value & vbTab & objRecordSet.Fields("cn").Value & vbTab &_
    objRecordSet.Fields("distinguishedName").Value & vbTab & objRecordSet.Fields("ProfilePath").Value &_
    vbTab & objRecordSet.Fields("ScriptPath").Value & vbTab & objRecordSet.Fields("HomeDirectory").Value & vbTab & objRecordSet.Fields("HomeDrive").Value & vbTab & strDisabled
    
    objRecordSet.MoveNext
Loop
objConnection.Close