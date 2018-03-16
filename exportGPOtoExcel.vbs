Set xmlDoc = CreateObject("Microsoft.XMLDOM")
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set oShell = WScript.CreateObject ("WScript.shell")
Dim CurCat, Row
CurCat = ""
Do Until ValidFile = "OK"
 XMLFile = InputBox("Location of the GPO XML report file.", "XML File", oShell.CurrentDirectory)
 If XMLFile = "" Then
  popup = msgbox ("Would you like to stop?", vbYesNo, "Quit?")
  If PopUp = vbYes Then
    WScript.Quit
  End If   
 Else
   If Not (FSO.FileExists(XMLFile)) Then
    popup = msgbox ("File " & XMLFile & " not found.")
    ValidFile = "ERROR"
   Else
    ValidFile = "OK"
   End If
 End If
Loop
'File Path
FilePath = Split(XMLFile, "\")
FilePath(UBound(FilePath,1)) = ""
FilePath = Join(FilePath, "\")
 
 '**************************************
 'New Excel workbook
 '**************************************
 Set XLapp = CreateObject("Excel.Application")
 Set wb = XLapp.Workbooks.Add
 XLapp.Visible = True
 XLapp.AlertBeforeOverwriting = True
 wb.Activate
 Set wsUser = wb.Worksheets(1)
 Set wsComputer = wb.Worksheets(2)
 wb.Worksheets(1).Name = "User Settings"
 wb.Worksheets(2).Name = "Computer Settings"
  
 '*****************************
 'Column heads
 '*****************************
 wsUser.Cells(5,2).Value = "Snap-in Location"
 wsUser.Cells(5,3).Value = "Setting Name"
 wsUser.Cells(5,4).Value = "Requires at Least"
 wsUser.Cells(5,5).Value = "Set to"
 wsUser.Cells(5,6).Value = "Value"
 wsUser.Cells(5,7).Value = "Explain"
 wsUser.Cells(7,1).Value = "User Configuration"
 wsUser.Cells(8,1).Value = "Administrative Templates"
 wsUser.Cells(2,2).Value = "GPO Name"
 wsUser.Cells(1,1).Value = "Group Policy Object Settings"
 wsUser.Cells(1,1).Font.Size = 16
 
 wsUser.Cells(5,2).Font.Bold = True
 wsUser.Cells(5,3).Font.Bold = True
 wsUser.Cells(5,4).Font.Bold = True
 wsUser.Cells(5,5).Font.Bold = True
 wsUser.Cells(5,6).Font.Bold = True
 wsUser.Cells(5,7).Font.Bold = True
 wsUser.Cells(7,1).Font.Bold = True
 wsUser.Cells(8,1).Font.Bold = True
 wsUser.Cells(2,2).Font.Bold = True
 wsUser.Cells(1,1).Font.Bold = True
 
 wsUser.Range(wsUser.Cells(1,1), wsUser.Cells(5,7)).Interior.ColorIndex = 15
 wsUser.Range(wsUser.Cells(1,1), wsUser.Cells(1,7)).Merge
 wsUser.Range(wsUser.Cells(1,1), wsUser.Cells(1,1)).HorizontalAlignment = 3
 wsUser.Columns("B:B").ColumnWidth = 18
 
 wsComputer.Cells(5,2).Value = "Snap-in Location"
 wsComputer.Cells(5,3).Value = "Setting Name"
 wsComputer.Cells(5,4).Value = "Requires at Least"
 wsComputer.Cells(5,5).Value = "Set to"
 wsComputer.Cells(5,6).Value = "Value"
 wsComputer.Cells(5,7).Value = "Explain"
 wsComputer.Cells(7,1).Value = "Computer Configuration"
 wsComputer.Cells(8,1).Value = "Administrative Templates"
 wsComputer.Cells(2,2).Value = "GPO Name"
 wsComputer.Cells(1,1).Value = "Group Policy Object Settings"
 wsComputer.Cells(1,1).Font.Size = 16
 
  
 wsComputer.Cells(5,2).Font.Bold = True
 wsComputer.Cells(5,3).Font.Bold = True
 wsComputer.Cells(5,4).Font.Bold = True
 wsComputer.Cells(5,5).Font.Bold = True
 wsComputer.Cells(5,6).Font.Bold = True
 wsComputer.Cells(5,7).Font.Bold = True
 wsComputer.Cells(7,1).Font.Bold = True
 wsComputer.Cells(8,1).Font.Bold = True
 wsComputer.Cells(2,2).Font.Bold = True
 wsComputer.Cells(1,1).Font.Bold = True
 
 wsComputer.Range(wsComputer.Cells(1,1), wsComputer.Cells(5,7)).Interior.ColorIndex = 15
 wsComputer.Range(wsComputer.Cells(1,1), wsComputer.Cells(1,7)).Merge 
 wsComputer.Range(wsComputer.Cells(1,1), wsComputer.Cells(1,1)).HorizontalAlignment = 3
 wsComputer.Columns("B:B").ColumnWidth = 18

xmlDoc.async="false"
xmlDoc.load(XMLFile)
 
for each x in xmlDoc.documentElement.childNodes
 
 Row = 10
 
 'GPO Name
 If x.nodename = "Name" Then
  GPOnaam = x.text
  wsComputer.Cells(2,3).Value = GPOnaam
  wsUser.Cells(2,3).Value = GPOnaam
  wsComputer.Cells(2,3).Font.Bold = True
  wsUser.Cells(2,3).Font.Bold = True
 End If
 
 If x.nodename = "Computer" Then
  For Each y In x.childnodes
   if y.Nodename = "ExtensionData" Then
    For Each z In y.childnodes
     If z.Nodename = "Extension" Then
      For Each setting In z.childnodes
       DocumentComputerPolicy(Setting)
      Next
     End If
    Next
   End If
  Next
 
 ElseIf x.nodename = "User" Then
  For Each y In x.childnodes
   if y.Nodename = "ExtensionData" Then
    For Each z In y.childnodes
     If z.Nodename = "Extension" Then
      For Each setting In z.childnodes
       DocumentUserPolicy(Setting)
      Next
     End If
    Next
   End If
  Next
 
 End If
Next
wsUser.Columns("C:C").EntireColumn.Autofit
wsComputer.Columns("C:C").EntireColumn.Autofit
MsgBox "GPO Export done!!", vbInformation, "Done"
XLapp.Visible = True
Set xmlDoc = Nothing
Set fso = Nothing
Set oShell = Nothing
WScript.Quit

Public Function DocumentComputerPolicy(Setting)
 
 'Ajust the name of the node
 If InStr(setting.nodename,":") Then
  Node = setting.nodename
  Node = Right(Node,Len(Node)-InStr(Node,":"))
 End If
 If not Node = "Policy" Then
  Exit Function
 End If
   
 For Each Value In Setting.Childnodes
  'Ajust the name of the node
  If InStr(Value.nodename,":") Then
   NodeName = Value.nodename
   NodeName = Right(NodeName,Len(NodeName)-InStr(NodeName,":"))
  End If
 
  If NodeName = "Category" Then
   If CurCat = "" Then
    wsComputer.Cells(10,2).Value = value.text
    CurCat = value.text
    Category = value.text
   End If
   
   If Not CurCat = value.text AND CurCat <> "" Then
    Category = value.text
    CurCat = value.text
    Row = Row + 1
   End If
  End If
  
  If NodeName = "Name" Then
   Name = value.text
  End If
  
  If NodeName = "State" Then
   State = value.text
  End If
  
  If NodeName = "Supported" Then
   Supported = value.text
  End If
  
  If NodeName = "Explain" Then
   'Explain = value.text
  End If
  
  If NodeName = "EditText" Then
   EditText = value.text
   EditText = Replace(EditText, "Enabled", " ")
  End If
  
  If NodeName = "ListBox" Then
   EditText = value.text
  End If
  
  
  
 Next
 
 If isnull(Setting.childnodes) Then
  For Each node In Setting.childnodes
   DocumentPolicy(node)
  Next
 End If
  
  wsComputer.Cells(Row,2).Value = Category
  wsComputer.Cells(Row,2).Font.Bold = True
     Row = Row + 1
  wsComputer.Cells(Row,3).Value = Name
  wsComputer.Cells(Row,5).Value = State
  wsComputer.Cells(Row,4).Value = Supported
  wsComputer.Cells(Row,6).Value = EditText
  wsComputer.Cells(Row,7).Value = Explain
  
End Function
 
 

Public Function DocumentUserPolicy(Setting)
 'Ajust the name of the node
 If InStr(setting.nodename,":") Then
  Node = setting.nodename
  Node = Right(Node,Len(Node)-InStr(Node,":"))
 End If
 
 If not Node = "Policy" Then
  Exit Function
 End If
 
   
 For Each Value In Setting.Childnodes
  'Ajust the name of the node
  If InStr(Value.nodename,":") Then
   NodeName = Value.nodename
   NodeName = Right(NodeName,Len(NodeName)-InStr(NodeName,":"))
  End If
 
  If NodeName = "Category" Then
   If CurCat = "" Then
    wsUser.Cells(10,2).Value = value.text
    CurCat = value.text
    Category = value.text
   End If
   
   If Not CurCat = value.text AND CurCat <> "" Then
    Category = value.text
    CurCat = value.text
    Row = Row + 1
   End If
  End If
  
  If NodeName = "Name" Then
   Name = value.text
  End If
  
  If NodeName = "State" Then
   State = value.text
  End If
  
  If NodeName = "Supported" Then
   Supported = value.text
  End If
  
  If NodeName = "Explain" Then
   Explain = value.text
  End If
  
  If NodeName = "EditText" Then
   EditText = value.text
   EditText = Replace(EditText, "Enabled", " ")
  End If
  
  If NodeName = "ListBox" Then
   EditText = value.text
  End If
  
  
  
 Next
 
 If isnull(Setting.childnodes) Then
  For Each node In Setting.childnodes
   DocumentUserPolicy(node)
  Next
 End If
  
  wsUser.Cells(Row,2).Value = Category
  wsUser.Cells(Row,2).Font.Bold = True
  Row = Row + 1
  wsUser.Cells(Row,3).Value = Name
  wsUser.Cells(Row,5).Value = State
  wsUser.Cells(Row,4).Value = Supported
  wsUser.Cells(Row,6).Value = EditText
  wsUser.Cells(Row,7).Value = Explain
  
End Function