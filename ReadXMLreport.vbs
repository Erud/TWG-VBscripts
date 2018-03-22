
Dim oXML
Set oXML = CreateObject("MSXML2.DOMDocument")

'Load the XML file
oXML.Load("C:\Temp\LAX Redirect Chapman Policy.xml")
compExt = "No"
Set nodes = oXML.selectNodes("/GPO/Computer")

For Each node In nodes
	For each elem in node.childNodes
		If elem.tagName = "ExtensionData" Then 
			compExt = "Yes"
			If LCase(elem.text) = "registry" Then compExt = "No"
		End If
		If elem.tagName = "Enabled" Then
			If elem.text = "true" Then
				compEna = "Yes"
			Else
				compEna = "No"
			End If
		End If
    Next
Next
compExt = compEna & "/" & compExt
userExt = "No"
Set nodes = oXML.selectNodes("/GPO/User")

For Each node In nodes
	For each elem in node.childNodes
		If elem.tagName = "ExtensionData" Then  userExt = "Yes"
		If elem.tagName = "Enabled" Then
			If elem.text = "true" Then
				userEna = "Yes"
			Else
				userEna = "No"
			End If
		End If
    Next
Next

Set nodes = oXML.getElementsByTagName("FilterName")

If nodes.length > 0 then
	For each x in nodes
		WMIfilter =  x.text
	Next
Else
		WMIfilter = ""
End If

userExt = userEna & "/" & userExt
WScript.Echo compExt & userExt & WMIfilter