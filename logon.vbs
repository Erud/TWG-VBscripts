'==========================================================================
'
' NAME: LogonScript.vbs
'
' AUTHOR:  Mark D. MacLachlan, The Spider's Parlor
' URL   : http://www.thespidersparlor.com
' Copyright (c) 2003-2007
' DATE  : 4/10/2003
'
' COMMENT: Enumerates current users' group memberships in given domain.
'          Maps and disconnects drives and printers
'
'==========================================================================


ON ERROR RESUME NEXT

Dim WSHShell, WSHNetwork, objDomain, DomainString, UserString, UserObj, Path


Set WSHShell = CreateObject("WScript.Shell")
Set WSHNetwork = CreateObject("WScript.Network")
'Automatically find the domain name
Set objDomain = getObject("LDAP://rootDse")
DomainString = objDomain.Get("dnsHostName")
'Find the Windows Directory
WinDir = WshShell.ExpandEnvironmentStrings("%WinDir%")

'Grab the user name
UserString = WSHNetwork.UserName
'Bind to the user object to get user name and check for group memberships later
Set UserObj = GetObject("WinNT://" & DomainString & "/" & UserString)

'Grab the computer name for use in add-on code later
strComputer = WSHNetwork.ComputerName


'Disconnect any drive mappings as needed.
WSHNetwork.RemoveNetworkDrive "F:", True, True

'Disconnect ALL mapped drives
Set clDrives = WshNetwork.EnumNetworkDrives
For i = 0 to clDrives.Count -1 Step 2
    WSHNetwork.RemoveNetworkDrive clDrives.Item(i), True, True
Next

'Give the PC time to do the disconnect, wait 300 milliseconds
wscript.sleep 300

'Map drives needed by all
'Note the first command uses the user name as a variable to map to a user share.
WSHNetwork.MapNetworkDrive "H:", "\\server\users\" & UserString,True
WSHNetwork.MapNetworkDrive "U:", "\\server\users",True
WSHNetwork.MapNetworkDrive "X:", "\\server\executables",True

'Now check for group memberships and map appropriate drives
'Note that this checks Global Groups and not domain local groups.
For Each GroupObj In UserObj.Groups
'Force upper case comparison of the group names, otherwise this is case sensitive.
    Select Case UCase(GroupObj.Name)
    'Check for group memberships and take needed action
    'In this example below, ADMIN and WORKERB are groups.
    'Note the use of all upper case letters as mentioned above.
    'Note also that the groups must be Global Groups.
        Case "ADMIN"
            WSHNetwork.MapNetworkDrive "w:", "\\Server\Admin Stuff",True
        Case "WORKERB"
            WSHNetwork.MapNetworkDrive "w:", "\\Server\Shared Documents",True
            'Below is an example of how to set the default printer
            WSHNetwork.SetDefaultPrinter "\\ServerName\PrinterName"
    End Select
Next

'Remove ALL old printers
'Enumerate all printers first, after that you can select the printers you want by performing some string checks
Set WSHPrinters = WSHNetwork.EnumPrinterConnections
For LOOP_COUNTER = 0 To WSHPrinters.Count - 1 Step 2
'To remove only networked printers use this If Statement
    If Left(WSHPrinters.Item(LOOP_COUNTER +1),2) = "\\" Then
      WSHNetwork.RemovePrinterConnection WSHPrinters.Item(LOOP_COUNTER +1),True,True
    End If
'To remove all printers incuding LOCAL printers use this statement and comment out the If Statement above
'WSHNetwork.RemovePrinterConnection WSHPrinters.Item(LOOP_COUNTER +1),True,True
Next

'Remove a specific printer
WSHNetwork.RemovePrinterConnection "\\ServerOld\HP5si",True,True
                                
'Install Printers
WSHNetwork.AddWindowsPrinterConnection "\\Server\HP5si"

'Add On Code goes below this line
'=====================================



'=====================================
'Add On Code goes above this line

'Clean Up Memory We Used
set UserObj = Nothing
set GroupObj = Nothing
set WSHNetwork = Nothing
set DomainString = Nothing
set WSHSHell = Nothing
Set WSHPrinters = Nothing
  

'Quit the Script
wscript.quit