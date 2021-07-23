Set WshShell = WScript.CreateObject("WScript.Shell")

'Path of Cisco AnyConnect Client'
WshShell.Run """%PROGRAMFILES(x86)%\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe""" 

WScript.Sleep 1500

WshShell.AppActivate "Cisco AnyConnect Secure Mobility Client"

WshShell.SendKeys "{ENTER}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
'VPN URL'
WshShell.SendKeys "127.0.0.1"
WshShell.SendKeys "{ENTER}"

WScript.Sleep 1500
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1500
WshShell.SendKeys "PASSWORD"

WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"
