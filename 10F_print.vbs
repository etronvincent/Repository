Set WshNetwork = CreateObject("WScript.Network") 
WshNetwork.AddWindowsPrinterConnection "\\tpe_print\TPE 10F HP LaserJet 4250"
WshNetwork.SetDefaultPrinter "\\tpe_print\TPE 10F HP LaserJet 4250"
WshNetwork.RemovePrinterConnection "\\tpe_print\TPE_10F FX DocuPrint 305-AP", true, true
WScript.Echo "已經完成新增印表機!!"