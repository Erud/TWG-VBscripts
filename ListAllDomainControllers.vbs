' Create log file
'
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLog = objFSO.CreateTextFile ("C:\Temp\log.txt")
' Get Forest’s root
'
Set objRoot = GetObject("LDAP://rootDSE")
' Get root’s Configuration
'
Set objConfig = GetObject("LDAP://" & objRoot.Get("ConfigurationNamingContext"))
' Search for the Partitions container in root’s Configuration
'
objConfig.Filter = Array("crossRefContainer")
For Each objPartition in objConfig
    strPartition = "LDAP://" & objPartition.Get("distinguishedName")
Next
' Search in Partitions for all domains in Forest
'
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.CommandText = "<" & strPartition & ">;(&(systemFlags=3));nCName,systemFlags;subTree"
Set objRecordset = objCommand.Execute
' List all domains in Forest
'
Do While Not objRecordset.EOF
	' List all DCs in one domain
	'
	objLog.WriteLine("=================================================")
	objLog.WriteLine(objRecordset.Fields(0))
	objLog.WriteLine("=================================================")
	Set objDCs = GetObject("GC://OU=Domain Controllers," & objRecordset.Fields(0))
	For Each objDC in objDCs
		objLog.WriteLine(objDC.Get("Name"))
	Next
	objLog.WriteLine("————————————————-")
	objRecordset.MoveNext
Loop
' We are done!
'
objLog.Close
MsgBox "The End"