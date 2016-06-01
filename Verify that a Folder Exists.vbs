' Verify that a Folder Exists


Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FolderExists("\\uschi2k8admon\c$\Program Files (x86)\Panda Security") Then
    Wscript.Echo "Folder exist."
Else
    Wscript.Echo "Folder does not exist."
End If