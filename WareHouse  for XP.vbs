Set WshNetwork = CreateObject("WScript.Network") 
WshNetwork.AddWindowsPrinterConnection "\\etr_print\WareHouse"
WScript.Echo "已經完成新增印表機!!"