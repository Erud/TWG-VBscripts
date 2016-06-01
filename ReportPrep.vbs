' Read Dumsec reports from a folder and produce Prep reports
strFolder = "C:\temp\reports_DFS_shares\uschi2k8fs01\data"


Set objFSOo = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSOo.GetFolder(strFolder)
Set objFiles = objFolder.Files

'all files in a folder
For Each objFile In objFiles
    strFileN = objFile.Name
    Call makeOut(strFileN) 
Next
'--------------------------------------------------------------- 
Sub makeOut(strFname)
Const ForReading = 1
Const ForAppending = 8

'Dim aPage(200000, 500)
Dim aPage(481000, 150)
WScript.Echo strFname
'strFname = "e_actuarial"
Set objDictionary = CreateObject("Scripting.Dictionary")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(strFolder & "\" & strFname , ForReading)
If objFSO.FileExists("C:\temp\reports_DFS_shares\" & strFname ) Then objFSO.DeleteFile("C:\temp\reports_DFS_shares\" & strFname )
Set objTextFileW = objFSO.OpenTextFile("C:\temp\reports_DFS_shares\" & strFname , ForAppending, True)
i = 1
ii = 1
Do Until objTextFile.AtEndOfStream 
    strNextLine = objTextFile.ReadLine
    If Right(strNextLine, 16) = "==>access denied" Then
'    	l = Len(strNextLine) - 16
'    	strLine = Trim(Left(strNextLine, l)) 
'    	strNextLine = strLine & vbTab & "==>access denied" & vbtab & "*"  
    	aLine = Split(strNextLine, vbTab)
    	strNextLine = aLine(0) & vbTab & "==>access denied" & vbtab & "*"
    End If
	    aLine = Split(strNextLine, vbTab) 
	    If UBound(aline) > 1 Then
		    z = Len(aLine(0)) - Len("(disktree)") + 1
		    If Mid(aLine(0), z) <> "(disktree)" Then 
			  '  WScript.Echo aLine(0) & vbTab & aLine(1) & vbTab & aLine(2)
			  	For u = 0 To i
			  		If aPage(0,u) = RTrim(aLine(1)) Then Exit For ' Uid in array
			  	Next
			  	If u <= i Then 
			  	Else
			    	aPage(0,i) = RTrim(aLine(1))         ' add  ID to array
			    	u = i
			    	i = i + 1
			    End If
			    If aPage(ii - 1,0) = RTrim(aLine(0)) Then
			    	'aPage(ii - 1,i) = aLine(2)
			    	strLine = Trim(UCase(Replace(aLine(2), "  ", "")))
				    Select Case strLine
				    	Case "RWXD RWXD" 		aPage(ii - 1,u) = "M"
				    	Case "RWXRWX" 			aPage(ii - 1,u) = "M"
				      	Case "ALLALL"   		aPage(ii - 1,u) = "F"
				      	Case "ALL"   			aPage(ii - 1,u) = "F"
				      	Case "R XR X"   		aPage(ii - 1,u) = "R"
				      	Case "R X"   			aPage(ii - 1,u) = "R"
				      	Case "*"	   			aPage(ii - 1,u) = "X"
				      	Case "AX00000004"   	aPage(ii - 1,u) = "4"
				      	Case "AX00000002"   	aPage(ii - 1,u) = "2"
				      	Case Else      			aPage(ii - 1,u) = "?? " &  strLine
				   	End Select 
			    else
				    aPage(ii,0) = RTrim(aLine(0))
				    ' process permissions
				    strLine = Trim(UCase(Replace(aLine(2), "  ", "")))
				    'aPerm = Split(strLine, " ")
				    Select Case strLine
				    	Case "RWXD RWXD" 	aPage(ii,u) = "M"
				    	Case "RWXRWX" 		aPage(ii,u) = "M"
				      	Case "ALLALL"   	aPage(ii,u) = "F"
				      	Case "ALL"   		aPage(ii,u) = "F"
				      	Case "R XR X"   	aPage(ii,u) = "R"
				      	Case "R X"   		aPage(ii,u) = "R"
				      	Case "*"	   		aPage(ii,u) = "X"
				      	Case "AX00000004"   aPage(ii,u) = "4"
				      	Case "AX00000002"   aPage(ii,u) = "2"
				      	Case Else       	aPage(ii,u) = "?? " &  strLine
				   	End Select
				    ii = ii + 1
			    End If
	    	End If
	    End If
	   
Loop
For m = 0 To ii 	
	For n = 0 To i
		If n = 0 Then 
			strLine = aPage(m, n) 
		Else
			strLine = strLine & vbTab & aPage(m, n) 
		End If
	next    
    objTextFileW.WriteLine(strLine & vbTab)
   ' WScript.Echo strLine
Next
objTextFileW.Close
End Sub