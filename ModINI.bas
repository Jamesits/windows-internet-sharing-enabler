Attribute VB_Name = "ModReadINI"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function ��ȡ(ByVal A As String, ByVal B As String) As String
Dim ret As Long
Dim buff As String
buff = String(255, 0)
ret = GetPrivateProfileString(A, B, "", buff, 256, ·�� & "Config.ini")
��ȡ = buff
End Function

Public Function д��(ByVal A As String, ByVal B As String, ByVal C As String) As String
On Error Resume Next
Dim success As Long
success = WritePrivateProfileString(A, B, C, ·�� & "Config.ini")
End Function
Function ·��() As String
On Error Resume Next
Dim A As String
A = App.Path
If Right(A, 1) = "\" Then
Else
A = App.Path & "\"
End If
·�� = A
End Function
