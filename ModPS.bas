Attribute VB_Name = "ModPS"
Option Explicit

Dim debugmode As Integer

Const logfile = "log.txt"

Public Sub RunPowershellCommand(ByVal filename As String)
log ShellExt("cmd /c start powershell -NoProfile -ExecutionPolicy unrestricted -NoExit -File " + filename)
'Shell "@powershell -NoProfile -ExecutionPolicy unrestricted -NoLogo -NonInteractive -WindowStyle Hidden -File " + filename , vbHide
End Sub
