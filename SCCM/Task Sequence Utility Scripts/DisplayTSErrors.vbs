' Displays a message when task sequence errors are detected in the smsts.log file located at the given argument path.
' Recommended to check for custom %OSDStatus% or built-in %_SMTSLastActionSucceeded% TS variable.
' Message will automatically exit after several hours.
' MJC 2-7-23

Const WshRunning = 0
Const WshFinished = 1
Const WshFailed = 2
dim strLogPath: strLogPath = WScript.Arguments.Item(0)

' VBScript regexp is really slow when parsing line by line. Use find instead.
strCommand = "find ""<![LOG[Failed to run the action:"" """ & strLogPath & "\smsts.log"""

Set WshShell = CreateObject("WScript.Shell")
Set WshShellExec = WshShell.Exec(strCommand)

While WshShellExec.Status = WshRunning : WScript.Sleep 50 : Wend
Select Case WshShellExec.Status
   Case WshFinished
       strOutput = WshShellExec.StdOut.ReadAll
   Case WshFailed
       strOutput = WshShellExec.StdErr.ReadAll
End Select

' Parse what we need from the output.
Dim oRE
Set oRE = New RegExp
With oRE
	.Pattern = "<!\[LOG\[(Failed to run the action:[^\]]+)]"
	.Global = True
	.IgnoreCase = True
End With

' Output to console and show pop-up message.
' Message should automatically exit after 1 hour inactivity.
Dim strMsg
Dim numCount
strMsg = ""
numCount = 0
For Each m In oRE.Execute(strOutput)
	WScript.Echo m.Submatches(0)
	numCount = numCount + 1
	If Len(strMsg) = 0 Then
		strMsg = numCount & ") " & m.Submatches(0)
	Else
		strMsg = strMsg & vbNewLine & numCount & ") " & m.Submatches(0)
	End If
Next
If Len(strMsg) > 0 Then
	WshShell.Popup strMsg, 28800, "Task Sequence Errors", 16+4096
End If

