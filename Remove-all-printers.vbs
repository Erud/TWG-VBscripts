ON ERROR RESUME NEXT

Dim WSHNetwork

Set WSHNetwork = CreateObject("WScript.Network")
'Remove ALL old printers
'Enumerate all printers first, after that you can select the printers you want by performing some string checks
Set WSHPrinters = WSHNetwork.EnumPrinterConnections
For LOOP_COUNTER = 0 To WSHPrinters.Count - 1 Step 2
    If Left(WSHPrinters.Item(LOOP_COUNTER +1),2) = "\\" Then
      WSHNetwork.RemovePrinterConnection WSHPrinters.Item(LOOP_COUNTER +1),True,True
    End If
Next