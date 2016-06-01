'List All Computers in AD with the Description field 
 
Dim strDescription 
 
Const ADS_SCOPE_SUBTREE = 1000 
 
Set objConnection = CreateObject("ADODB.Connection") 
Set objCommand =   CreateObject("ADODB.Command") 
objConnection.Provider = "ADsDSOObject" 
objConnection.Open "Active Directory Provider" 
 
Set objCOmmand.ActiveConnection = objConnection 
objCommand.CommandText = _ 
    "Select Name, Location, Description from 'LDAP://chi.warranty.local/DC=chi,DC=warranty,DC=local' Where objectClass='computer'" 
objCommand.Properties("Page Size") = 1000 
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE  
Set objRecordSet = objCommand.Execute 
objRecordSet.MoveFirst 
 
Do Until objRecordSet.EOF 
 
                
    WScript.Echo objRecordSet.Fields("Name").Value 
   
 
    objRecordSet.MoveNext 
Loop 
 
