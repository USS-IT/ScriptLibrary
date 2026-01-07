' Displays a message when the defined address is unreachable.

Const strAddress = "google.com"
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set WshShell = CreateObject("WScript.Shell")

Set colItems = objWMIService.ExecQuery("Select * From win32_PingStatus where address='" & strAddress & "'")
Dim intResult
intResult = 0
For Each objItem in colItems
	If IsNumeric(objItem.statuscode) And objItem.statuscode = 0 And IsNumeric(objItem.ResponseTime) Then
		intResult = 1
	End If
Next

If intResult = 0 Then
	WshShell.Popup "Cannot reach [" & strAddress & "]. System may be missing MAC registration.", 28800, "Address Unreachable", 16+4096
End If

