Set objShell = CreateObject("Wscript.Shell")
Set objArguments = WScript.Arguments

If (objArguments.Count = 1) Then
	objShell.Run "%windir%\explorer.exe ms-settings-connectabledevices:devicediscovery"
	Wscript.Sleep(1000)
	objShell.SendKeys "+{Tab}"
	objShell.SendKeys "{Tab}"
	objShell.SendKeys "{Enter}"
	Wscript.Sleep(1000)
	objShell.SendKeys objArguments(0)
	Wscript.Sleep(1000)
	objShell.SendKeys "{Tab}"
	objShell.SendKeys "{Enter}"
	objShell.SendKeys "{Esc}"
	objShell.SendKeys "{Esc}"
Else
  MsgBox "Please start the script with the display name to use as the first parameter.", VBOKOnly, "Missing Parameter"
End If
