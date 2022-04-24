set WshShell = WScript.CreateObject("WScript.Shell")
set shell = WScript.CreateObject("Shell.Application")

WshShell.Run("powershell.exe")
WScript.Sleep 500

If WScript.Arguments.Length = 0 Then
  Set ObjShell = CreateObject("Shell.Application")
  ObjShell.ShellExecute "wscript.exe" _
    , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
  WScript.Quit
End if

WshShell.SendKeys "cd $home\Desktop"
WshShell.SendKeys "{ENTER}"

intAnswer =_
	Msgbox("Do you want to export app list?",_
		vbYesNo, "Export app list")
If intAnswer = vbYes Then
	WScript.Sleep 100
	WshShell.SendKeys "%{Tab}"
	WshShell.SendKeys "winget export --output app_list.json"
	WshShell.SendKeys "{ENTER}"
	WScript.Sleep 1000
Else
	WScript.Sleep 500
	WshShell.SendKeys ""
End If

WshShell.SendKeys "winget upgrade"
WScript.Sleep 100
WshShell.SendKeys "{ENTER}"
WScript.Sleep 10000
intAnswer =_
	Msgbox("Do you want to update all programs?",_
		vbYesNo, "Update Applications")
If intAnswer = vbYes Then
	WScript.Sleep 100
	WshShell.SendKeys "winget upgrade --all"
	WshShell.SendKeys "{ENTER}"
	WScript.Sleep 500
	WshShell.SendKeys "%{Tab}"
Else
	WScript.Sleep 500
	WshShell.SendKeys "exit"
	WshShell.SendKeys "{ENTER}"
	WScript.Sleep 500
	WshShell.Run "taskkill /im powershell.exe", , True
End If

	WScript.Sleep 500
	WshShell.Run "taskkill /im powershell.exe", , True
