' List Registry Subkeys
' HKEY_CLASSES_ROOT\CLSID\{A9B8E64D-3F7E-4D32-8FC9-E391DEE67D75} 
' C:\Program Files (x86)\Panda Security\WAC\PSUAShell.dll --------- RG_SZ

Const HKEY_CLASSES_ROOT = &H80000000 
Dim strComputer, strRegKey, strArrKeyes, strKeyPath, strSubKey
strKey = "Panda Security"
strComputer = "uschi2k8admon"

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objService = objSWbemLocator.ConnectServer _
    (strComputer, "root\default", "d0c0rudakove", "@batman01")
objService.Security_.ImpersonationLevel = 3   
Set objRegistry = objService.Get("StdRegProv")

objRegistry.EnumKey HKEY_CLASSES_ROOT, "", arrSubkeys

For each strRegKey in arrSubkeys
    enumSubkeys HKEY_CLASSES_ROOT, strRegKey
Next

Sub enumSubkeys(HKEY_CLASSES_ROOT, strKeyPath) 
    objRegistry.EnumKey HKEY_CLASSES_ROOT, strKeyPath, arrSubkeys 
    If IsArray(arrSubkeys) Then 
       For Each strSubkey In arrSubkeys
            strTestKey = strKeyPath & "\" & strSubKey
           
            'If Left(strTestKey, 5) = "CLSID" Then
	            objRegistry.EnumValues HKEY_CLASSES_ROOT, strTestKey, arrValueNames, arrTypes
	            If IsArray(arrValueNames) Then 
					For i = LBound(arrValueNames) To UBound(arrValueNames)
						strValueName = arrValueNames(i)
						If arrTypes(i) = 1 Then ' Show a REG_SZ value
							objRegistry.GetStringValue HKEY_CLASSES_ROOT, strTestKey, strValueName, strValue
							If InStr(strValue, "Panda Security") > 0 Then 
								wscript.echo strTestKey & " " & strValueName & " (REG_SZ) = " & strValue
							End If
						End If
					Next
				End If
	            '	WScript.Echo strTestKey
            '	WScript.Quit
            'End If	
            enumSubkeys HKEY_CLASSES_ROOT, strKeyPath & "\" & strSubkey 
       Next
    End If 
    'objRegistry.DeleteKey HKEY_CLASSES_ROOT, strKeyPath
    
End Sub