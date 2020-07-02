strComputer ="." 

Set oWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colComputerSystemItems = oWMIService.ExecQuery ("SELECT * FROM Win32_ComputerSystem",,48)
Set colComputerSystemProductItems = oWMIService.ExecQuery ("SELECT * FROM Win32_ComputerSystemProduct",,48) 
Set colSMBIOSItems = oWMIService.ExecQuery( "Select * from Win32_BIOS where PrimaryBIOS = true", , 48 )
Set colNetworkAdapterItems = oWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled = 1",,48)

' On Error Resume Next 
For Each oComputerSystem In colComputerSystemItems 
	Wscript.Echo "Make.............................: " & oComputerSystem.Manufacturer
	WScript.Echo "Model............................: " & oComputerSystem.Model
Next 

For Each oComputerSystemProduct In colComputerSystemProductItems 
	Wscript.Echo "UUID.............................: " & oComputerSystemProduct.UUID
Next
	
For Each oSMBIOS In colSMBIOSItems 
	Wscript.Echo "SMBIOS Version...................: " & oSMBIOS.SMBIOSBIOSVersion
	Wscript.Echo "SMBIOS Serial Number.............: " & oSMBIOS.SerialNumber
Next 

For Each NetworkAdapter in colNetworkAdapterItems  
  	'If Left(NetworkAdapter.Manufacturer,5) = "Intel" or Left(NetworkAdapter.Manufacturer,8) = "Broadcom" Then  
  Wscript.Echo "Network Adapter Name.............: " & NetworkAdapter.Description
  WScript.Echo "Network Adapter MACAddress.......: " & NetworkAdapter.MACAddress
    'End If   
Next  

