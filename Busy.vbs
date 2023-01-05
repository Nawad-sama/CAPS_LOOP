Dim vbsToKill 
vbsToKill = "Busy.vbs"

Dim wmiObj, objKill, task, counter 

Set wmiObj = GetObject("winmgmts://./root/cimv2")
Set objKill = wmiObj.ExecQuery( _
"Select * from Win32_Process " & _
"Where Name = 'wscript.exe'" & _
"AND commandline LIKE '%"& vbsToKill &"%'" _
)

For Each task in objKill
	counter = counter + 1
  Next

If counter >= 2 Then

KURWIX

Else 

BUSY

End If


Function KURWIX

Dim vbsToKill 
vbsToKill = "Busy.vbs"

Dim wmiObj, objKill, task

Set wmiObj = GetObject("winmgmts://./root/cimv2")
Set objKill = wmiObj.ExecQuery( _
"Select * from Win32_Process " & _
"Where Name = 'wscript.exe'" & _
"AND commandline LIKE '%"& vbsToKill &"%'" _
)

For Each task in objKill
	task.Terminate()
  Next

End Function

Function BUSY

set shell = WScript.CreateObject("WScript.Shell")
Do While True
shell.SendKeys "{CAPSLOCK}"
WScript.Sleep 500
Loop

End Function
