Const	HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."

Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")
	
strKeyPath = "SYSTEM\ControlSet001\services\Tcpip\Parameters\Interfaces\{AD09EC28-9879-4CAA-8860-F90FF6BDBAF8}"
objReg.CreateKey HKEY_LOCAL_MACHINE,strKeyPath

strEntryName = "NameServer"
strValue = "10.168.11.249,192.168.1.4,10.168.12.249"
objReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strEntryName,strValue


Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")
	
strKeyPath = "SYSTEM\ControlSet001\services\Tcpip\Parameters"
objReg.CreateKey HKEY_LOCAL_MACHINE,strKeyPath

strEntryName = "SearchList"
strValue = "eevertech.com,etron.com.tw,eys3d.com"
objReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strEntryName,strValue