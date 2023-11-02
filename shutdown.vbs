result = MsgBox("Would you like to shutdown your computer?",4+64,"Shutdown computer")
If result = 6 Then
   set shell = CreateObject("WScript.Shell")
   shell.run "shutdown.exe -s -t 0",0 ,false
End If