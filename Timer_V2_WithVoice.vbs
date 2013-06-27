dim time
dim reason
dim command
dim speak
time = inputbox("Welcome to the Computer Timer. Please input your time limit (in minutes)","Computer Timer")
if time > 0 Then
	reason = inputbox("Please input your reminder message.","Computer Timer")
	time = time * 60000
	WScript.Sleep time
    Set Speak=CreateObject("sapi.spvoice")
    Speak.Speak reason
	msgbox(reason)
	command = inputbox("Enter Command (Leave blank if none)", "Computer Timer")
	if command = "sleep" Then
		RunDll32.exe powrprof.dll,SetSuspendState
	ElseIf command = "shutdown" Then
		strCmd = "shutdown -s -t 0 -f"
		set objShell = CreateObject("WScript.Shell")
		objShell.Run strCmd
	ElseIf command = "log off" Then
		Set wshShell = WScript.CreateObject ("WSCript.shell")
		wshshell.run "shutdown /l /t 0"
	ElseIf command = "lock" Then
		Set WshShell = WScript.CreateObject("WScript.Shell")
		WshShell.Run "rundll32 user32.dll,LockWorkStation"
	else
		WScript.Sleep (1000)
	End If
else
	msgbox("That isn't a valid time.")
End If