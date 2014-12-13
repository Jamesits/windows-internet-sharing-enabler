Attribute VB_Name = "ModReadINI"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function 读取(ByVal A As String, ByVal B As String) As String
Dim ret As Long
Dim buff As String
buff = String(255, 0)
ret = GetPrivateProfileString(A, B, "", buff, 256, 路径 & "Config.ini")
读取 = buff
End Function

Public Function 写入(ByVal A As String, ByVal B As String, ByVal C As String) As String
On Error Resume Next
Dim success As Long
success = WritePrivateProfileString(A, B, C, 路径 & "Config.ini")
End Function
Function 路径() As String
On Error Resume Next
Dim A As String
A = App.Path
If Right(A, 1) = "\" Then
Else
A = App.Path & "\"
End If
路径 = A
End Function
