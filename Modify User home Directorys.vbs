' Modify User home Directory
' Copyright 2016 Edward Rudakov. All rights reserved.

Const ForReading = 1

strReadFile = "C:\temp\scripts\ADuserschange.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFileIn = objFSO.OpenTextFile(strReadFile, ForReading, False)

Do While objFileIn.AtEndOfStream = False
	strUser = "LDAP://" & objFileIn.ReadLine
	Set objUser = GetObject(strUser)
	strhomeDirectory = objUser.Get("homeDirectory")
	WScript.Echo "Processing " & strUser & " " & strhomeDirectory
	strhomeDirectory = Replace(LCase(strhomeDirectory), "uschi2k8fs01", "CHI\TWG")
	objUser.Put "homeDirectory", strhomeDirectory
'	objUser.SetInfo
	WScript.Echo "Put " & strUser & " " & strhomeDirectory
Loop