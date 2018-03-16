Option Explicit
Dim objADSystemInfo, objUser, objMemberOf, objGroup, objGroupEnum, objNetwork, objPrinter 
dim i, bTroubleFlag, strUser

Set objNetwork = CreateObject("Wscript.Network") 
 

 
Set objPrinter = objNetwork.EnumPrinterConnections 'Test to see if we have any printers mapped 
If objPrinter.Count > 0 Then 
    'The Printer array is Printer name, printer path that is why it is step 2 
    for i=0 to objPrinter.Count-2 Step 2 
        'test to make sure it is a network printer 
        strUser = objPrinter.Item(i+1)
        if instr(objPrinter.Item(i+1),"\\") <> 0 then 
             WScript.Echo objPrinter.Item(i+1) 
'            objNetwork.RemovePrinterConnection objPrinter.Item(i+1),true,true 
        end if 
    next 
end if 
 
 
