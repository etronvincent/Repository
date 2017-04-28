Set objShell = CreateObject("WScript.Shell")
Set objNetwork = CreateObject("WScript.Network")
Const HIDE_WINDOW = 1
On Error Resume Next
FileServer1="eever.eevertech.com"

  objNetwork.MapNetworkDrive "Z:", "\\" & FileServer1 & "\eEver_Dept\SOC Common\System Design"
  

