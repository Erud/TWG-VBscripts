Sub FixHyperlinks()
    Dim wks As Worksheet
    Dim hl As Hyperlink
    Dim sOld As String
    Dim sNew As String

    Set wks = ActiveSheet
    sOld = "D:\Audit Data" 
    sNew = "\\amer.ajgco.com\corp\MSEngineeringDomainData"
    For Each hl In wks.Hyperlinks
        hl.Address = Replace(hl.Address, sOld, sNew)
    Next hl
End Sub