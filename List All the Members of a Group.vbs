' List All the Members of a Group

'On Error Resume Next
 
Set objGroup = GetObject _
  ("LDAP://cn=vsc.CorporateLaw,ou=Groups,ou=CorporateLaw,dc=chi,dc=warranty,dc=local")

objGroup.GetInfo
 
arrMemberOf = objGroup.GetEx("member")
 
For Each strMember in arrMemberOf
    WScript.echo strMember
Next