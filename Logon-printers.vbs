Option Explicit
Dim objADSystemInfo, objUser, objMemberOf, objGroup, objGroupEnum, objNetwork, objPrinter 
dim i, bTroubleFlag, strUser

on error resume next 
 
Set objNetwork = CreateObject("Wscript.Network") 
 
'Get current user info from active directory 
Set objADSystemInfo = CreateObject("ADSystemInfo")
'bind to current user in active directory
strUser =  objADSystemInfo.UserName
set objUser = GetObject("LDAP://" & objADSystemInfo.UserName) 
 
'if objuser.description = "printer.trouble" then 
'    bTroubleFlag = true 
'    msgbox "Troubleshooting Printer Logon Script" 
'end if 
 
Set objPrinter = objNetwork.EnumPrinterConnections 'Test to see if we have any printers mapped 
If objPrinter.Count > 0 Then 
    'The Printer array is Printer name, printer path that is why it is step 2 
    for i=1 to objPrinter.Count Step 2 
        'test to make sure it is a network printer 
        strUser = objPrinter.Item(i)
        strUser = objPrinter.Item(i+1)
        if instr(objPrinter.Item(i),"\\") <> 0 then 
            if bTroubleFlag then 
                msgbox "Deleting:" & vbcrlf & bjPrinter.Item(i) 
            end if 
'            objNetwork.RemovePrinterConnection objPrinter.Item(i),true,true 
        end if 
    next 
end if 
 
 
'Get an array of group names that the user is a member of 
objMemberOf = objUser.MemberOf 
for Each objGroup in objMemberOf 
    'Test to see if it is a printer group. all printer groups should be in the same OU 
    if (instr(objGroup,"OU=Printer-Groups") <> 0) then 
        'Bind to the group to get is description. The description contain the path to the printer 
        set objGroupEnum = GetObject("LDAP://" & objGroup) 
        if bTroubleFlag then 
            msgbox "Adding:" & vbcrlf & "[" & objGroupEnum.name & "]" & vbcrlf & objGroupEnum.description 
        end if 
'        objNetwork.AddWindowsPrinterConnection objGroupEnum.description 
        set objGroupEnum = nothing 
    end if 
next 
'Repeat as above for the default printer 
for Each objGroup in objMemberOf 
    if (instr(objGroup,"OU=Printer-Default-Groups") <> 0) then 
        set objGroupEnum = GetObject("LDAP://" & objGroup) 
        if bTroubleFlag then 
            msgbox "Setting Default:" & vbcrlf & "[" & objGroupEnum.name & "]" & vbcrlf & objGroupEnum.description 
        end if 
'        objNetwork.SetDefaultPrinter objGroupEnum.description 
        set objGroupEnum = nothing 
    end if 
next 
 
if bTroubleFlag then 
    msgbox "Printer Logon Script Finished" 
end if 