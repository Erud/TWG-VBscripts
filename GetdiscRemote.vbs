'On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array("0219-MAYERSKYL")
For Each strComputer In arrComputers

	Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
	Set objWMIService = objSWbemLocator.ConnectServer(strComputer, "root\cimv2", "d0c0rudakove", "@batman01")
	objWMIService.Security_.ImpersonationLevel = 3 

   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDisk ", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
     
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "Compressed: " & objItem.Compressed
      
      WScript.Echo "CreationClassName: " & objItem.CreationClassName
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "DeviceID: " & objItem.DeviceID
      WScript.Echo "DriveType: " & objItem.DriveType
      
      WScript.Echo "FileSystem: " & objItem.FileSystem
      WScript.Echo "FreeSpace: " & objItem.FreeSpace
      
      
      WScript.Echo "MediaType: " & objItem.MediaType
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "NumberOfBlocks: " & objItem.NumberOfBlocks
      WScript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
      
      
      WScript.Echo "ProviderName: " & objItem.ProviderName
      WScript.Echo "Purpose: " & objItem.Purpose
      WScript.Echo "QuotasDisabled: " & objItem.QuotasDisabled
      WScript.Echo "QuotasIncomplete: " & objItem.QuotasIncomplete
      WScript.Echo "QuotasRebuilding: " & objItem.QuotasRebuilding
      WScript.Echo "Size: " & objItem.Size
      WScript.Echo "Status: " & objItem.Status
      WScript.Echo "StatusInfo: " & objItem.StatusInfo
      WScript.Echo "SupportsDiskQuotas: " & objItem.SupportsDiskQuotas
      WScript.Echo "SupportsFileBasedCompression: " & objItem.SupportsFileBasedCompression
      WScript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
      WScript.Echo "SystemName: " & objItem.SystemName
      WScript.Echo "VolumeDirty: " & objItem.VolumeDirty
      WScript.Echo "VolumeName: " & objItem.VolumeName
      WScript.Echo "VolumeSerialNumber: " & objItem.VolumeSerialNumber
      WScript.Echo
   Next
Next

