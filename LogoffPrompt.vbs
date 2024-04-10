Option Explicit

Dim objShell, desktopPath, myShortcut
Set objShell = CreateObject("WScript.Shell")
desktopPath = objShell.SpecialFolders("Desktop")

' Prompt the user
Dim answer
answer = MsgBox("Are you sure you want to log off?", vbYesNo + vbQuestion, "Log Off Prompt")
If answer = vbYes Then
    ' Log off the computer
    objShell.Run "shutdown /l"
End If
