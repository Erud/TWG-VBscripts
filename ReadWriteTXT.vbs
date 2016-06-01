Const ForReading = 1
Const ForAppending = 8

strIn = "C:\temp\scripts\servers.txt"
strOut = "C:\temp\scripts\servers_log.txt"

Set objFSOl = CreateObject("Scripting.FileSystemObject")
Set objInFile = objFSOl.OpenTextFile(strIn , ForReading)
If objFSOl.FileExists(strOut) Then objFSOl.DeleteFile(strOut)
Set objOutFile = objFSOl.OpenTextFile(strOut , ForAppending, True)

Do Until objInFile.AtEndOfStream 
	strNextLine = objInFile.ReadLine
	
	

	Wscript.Echo strNextLine
	objOutFile.WriteLine(strNextLine & vbTab) 
	

Loop
objOutFile.Close

