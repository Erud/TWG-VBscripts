Const ForReading = 1

strFolder = "C:\temp\scripts"
strFname  = "folders.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(strFolder & "\" & strFname , ForReading)

Do Until objTextFile.AtEndOfStream 
	strNextLine = objTextFile.ReadLine
	
	Set objFolder = objFSO.GetFolder(strNextLine)

	Wscript.Echo strNextLine & vbTab & objFolder.Size

Loop


