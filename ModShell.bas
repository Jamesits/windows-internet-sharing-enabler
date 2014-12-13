Attribute VB_Name = "ModShell"
Option Explicit

Private Declare Function IsUserAnAdmin Lib "shell32" Alias "#680" () As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ALIVE = &H103

Public Function hasAdmin() As Boolean
hasAdmin = IsUserAnAdmin
End Function

' Run a program, wait for it to exit and get the output
Public Function ShellExt(ByVal path As String, Optional mode = vbNormal)
Dim pid As Long
Dim outfile As String
outfile = App.path & "\out.tmp"
pid = Shell("cmd /c " & path & " >>" & outfile, mode)
Dim hProcess, ExitCode
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
Do
Call GetExitCodeProcess(hProcess, ExitCode)
Loop While ExitCode = STILL_ALIVE
Call CloseHandle(hProcess)
ShellExt = File2Text(outfile)
Shell "cmd /c rm " & outfile, vbHide
End Function
