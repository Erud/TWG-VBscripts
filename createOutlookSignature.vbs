'Option Explicit
On Error Resume Next

   Dim strSigName
   Dim strFullName, strTitle, strCompany, strTel, strFax, strEmail, strWeb, strCorpEmail
   Dim boolUpdateStyle

'==========================================================================
' Some script variables
'==========================================================================

'  Name signature
   strSigName  = "Standard Signature"
'  If signature exists, overwrite (true) or leave alone (false)?
   boolUpdateStyle = true

'==========================================================================
' Set some static information
'==========================================================================

'  Company information
   strCompany  = "Your Company Name"
   strTel      = "+12 1234 567890"
   strFax      = "+12 9876 543210"
   strWeb      = "http://www.yourdomain.com"

'  Fallback email address when no address is found
   strCorpEmail = "contact@yourdomain.com"

'==========================================================================
' Read User's Active Directory information
'==========================================================================
   Dim objSysInfo, objUser

   Set objSysInfo = CreateObject("ADSystemInfo")
   Set objUser    = GetObject("LDAP://" & objSysInfo.Username)

   strFullName = objUser.displayname
   strTitle    = objUser.title
   strEmail    = objuser.emailaddress

   If Trim(strTitle) = "" Then strTitle = "_"
   If Trim(strEmail) = "" Then strEmail = strCorpEmail

   Set objUser    = Nothing
   Set objSysInfo = Nothing


'==========================================================================
' Get Signature Folder
'==========================================================================
   Dim objShell
   Set objShell = CreateObject("WScript.Shell")
   strSigFolder = ObjShell.ExpandEnvironmentStrings("%appdata%") & "\Microsoft\Signatures\"
   Set objShell = Nothing


'==========================================================================
' Get Signature Folder
'==========================================================================
   Dim objFSO, objFile
   Set objFSO   = CreateObject("Scripting.FileSystemObject")

   If Not (objFSO.FolderExists(strSigFolder)) Then
      Call objFSO.CreateFolder(strSigFolder)
   End If

   strHTMFile = strSigFolder & strSigName & ".htm"
   strRTFFile = strSigFolder & strSigName & ".rtf"
   strTXTFile = strSigFolder & strSigName & ".txt"


'==========================================================================
' Create HTM File
'==========================================================================
'chr(47) = /

   Err.Clear
   Set objFile = objFSO.CreateTextFile(strHTMFile, boolUpdateStyle, False)
   If Err.Number = 0 Then
      objFile.Write "<html> <head> <title> <" & Chr(47) & "title> <" & Chr(47) & "head> <body>"&vbCrLf
      objFile.Write "<p><span style=""FONT-SIZE: 10pt; COLOR:#1F497D; FONT-FAMILY: Calibri"">"&vbCrLf
      objFile.Write strFullName & "<br />"&vbCrLf
      objFile.Write strTitle & "<br />"&vbCrLf
      objFile.Write strCompany & "<br />"&vbCrLf
      objFile.Write "T:&nbsp;" & strTel & " &nbsp; F:&nbsp;" & strFax & "<br />"&vbCrLf
      objFile.Write "E:&nbsp;<a href=""mailto:" & strEmail & """ style=""FONT-SIZE: 10pt; COLOR:#1F497D; 

FONT-FAMILY: Calibri"""
      objFile.Write ">" & strEmail & "<" & Chr(47) & "a>"&vbCrLf
      objFile.Write " &nbsp; <a href=""" &strWeb & """ style=""FONT-SIZE: 10pt; COLOR:#1F497D; FONT-FAMILY: 

Calibri"""
      objFile.Write ">" & strWeb & "<" & Chr(47) & "a>"&vbCrLf
      objFile.Write "<" & Chr(47) & "p>"&vbCrLf
      objFile.Write "<span style=""FONT-SIZE: 10pt; COLOR: Green; FONT-FAMILY: Calibri"">"&vbCrLf
      objFile.Write "<span style=""FONT-SIZE: 18pt; COLOR: Green; FONT-FAMILY: Webdings"">P<" & Chr(47) & 

"span>"&vbCrLf
      objFile.Write "Please consider the environment - do you really need to print this email?<br />"&vbCrLf
      objFile.Write "<" & Chr(47) & "span><" & Chr(47) & "p>"&vbCrLf
      objFile.Write "<" & Chr(47) & "body> <" & Chr(47) & "html> "&vbCrLf
      objFile.close
   End If


'==========================================================================
' Create TXT File
'==========================================================================
   Err.Clear
   Set objFile = objFSO.CreateTextFile(strTXTFile, boolUpdateStyle, False)
   If Err.Number = 0 Then
      objFile.Write strFullName & vbCrLf
      objFile.Write strTitle & vbCrLf
      objFile.Write strCompany & vbCrLf
      objFile.Write "T: " & strTel   & "   F: " & strFax & vbCrLf
      objFile.Write "E: " & strEmail & vbCrLf
      objFile.Write strWeb & vbCrLf & vbCrLf
      objFile.Write "Please consider the environment - do you really need to print this email?" & vbCrLf
      objFile.close
   End If


'==========================================================================
' Create RTF File
'==========================================================================
   Err.Clear
   Set objFile = objFSO.CreateTextFile(strRTFFile, boolUpdateStyle, False)
   If Err.Number = 0 Then
      objfile.write "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fswiss\fprq2\fcharset0 Calibri;}

{\f1\froman\fprq2\fcharset2 Webdings;}}" & vbCrLF
      objfile.write "{\colortbl;\red031\green073\blue125;\red0\green0\blue255;\red0\green128\blue0;}" & 

vbCrLF
      objfile.write "{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1\pard

\sb100\sa100\cf1\lang2057\f0\fs20 " & strFullName & "\line "
      objfile.write strTitle & "\line " & strCompany & "\line T: " & strTel & "   F: " & strFax & "\line E: "
      objfile.write "{\field{\*\fldinst{HYPERLINK ""mailto:" & strEmail & """}}{\fldrslt{\ul " & strEmail & 

"}}}\ulnone\f0\fs20    "
      objfile.write "{\field{\*\fldinst{HYPERLINK """ & strWeb & """}}{\fldrslt{\ul " & strWeb & "}}}\ulnone

\f0\fs20\par" & vbCrLF
      objfile.write "\cf3\f1\fs36 P\fs20  \f0 Please consider the environment - do you really need to print 

this email?\par" & vbCrLF
      objfile.write "\pard\cf1\lang1033\par" & vbCrLF
      objfile.write "}" & vbCrLF
      objFile.close
   End If


'==========================================================================
' Tidy-up
'==========================================================================
   set objFile = Nothing
   set objFSO  = Nothing

