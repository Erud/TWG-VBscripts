Dim objConnection, objRootDSE, objRecordSet
Dim strFilter

strFilter = "(NETBIOSName=*)"

Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objRootDSE = GetObject("LDAP://RootDSE")
Set objRecordSet = objConnection.Execute( _
  "<LDAP://" & objRootDSE.Get("configurationNamingContext") & ">;" & strFilter & ";dnsroot,ncname;subtree")
Set objRootDSE = Nothing

While Not objRecordSet.EOF
  WScript.Echo Join(objRecordSet.Fields("dnsroot").Value)
  WScript.Echo objRecordSet.Fields("ncname").Value
  objRecordSet.MoveNext
WEnd