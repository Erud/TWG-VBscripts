' ListNetworkSharesbyDiffUser.vbs
' Sample VBScript by WMI to find and list Network Shares.
' Author: http://www.morgantechspace.com/
' ------------------------------------------------------

strComputer = "uschi2k8fs01"
' you can change your own Computer or Server.

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator") 
Set objWMIService= objSWbemLocator.ConnectServer(strComputer , "root\cimv2","Administrator","YourPassword")

Set colShares = objWMIService.ExecQuery("Select * from Win32_Share Where Type=0")

For each objShare in colShares

    Wscript.Echo "Name: " & objShare.Name  & vbCrLf _
                       & "Path: " & objShare.Path & vbCrLf  _ 
                       & "Type: " & objShare.Type  & vbCrLf _
                       & "Caption: " & objShare.Caption   
Next
