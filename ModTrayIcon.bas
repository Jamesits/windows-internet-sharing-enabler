Attribute VB_Name = "托盘模块"
Option Explicit
    '更换托盘图标 Picture1.Picture
    '气泡提示 Text1.Text, Text2.Text, 信息图标, Picture1.Picture
    
    
 '  Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim lMsg As Single
   ' lMsg = X / Screen.TwipsPerPixelX
   ' If MinFlag = True Then
   '     Select Case lMsg
  '      Case 513    'WM_LBUTTONUP 鼠标左键点击图标
            '鼠标左键单击
  '      Case 515
            '鼠标左键双击
  '          MinFlag = False
  '          Me.WindowState = vbNormal
  '          Me.Show
  '          气泡提示 "弹出窗口。"
  '      Case 517    'WM_RBUTTONUP 鼠标右键点击图标
  '          PopupMenu 右键菜单    '如果是在系统托盘图标上点右键，则弹出菜单
  '      Case 518
   '         '鼠标右键双击
  '      End Select
 '   End If
'End Sub

'Private Sub Form_Resize()
'    If Me.WindowState = 1 And MinFlag = False Then
    '    MinFlag = True
   '     Me.Hide    '隐藏窗口
  '      气泡提示 "在此次双击鼠标左键，将弹出窗口。" & vbCrLf & "在此次点击鼠标右键，将弹出菜单。"
 '   End If
'End Sub
    
    
    
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Const NIM_ADD = &H0    '在任务栏中增加一个图标
Private Const NIM_DELETE = &H2    '删除任务栏中的一个图标
Private Const NIM_MODIFY = &H1    '修改任务栏中个图标信息
Private Const NIM_SETFOCUS = &H3
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONUP = &H205
Private Type NOTIFYICONDATA
    cbSize As Long    '该数据结构的大小
    hwnd As Long    '处理任务栏中图标的窗口句柄
    uID As Long    '定义的任务栏中图标的标识
    uFlags As Long    '任务栏图标功能控制，可以是以下值的组合（一般全包括）
    uCallbackMessage As Long    '任务栏图标通过它与用户程序交换消息，处理该消息的窗口由hWnd决定
    hIcon As Long    '任务栏中的图标的控制句柄
    szTip As String * 128    '图标的提示信息。若要产生气泡提示信息，则一定要128才性，为64则无法生成气泡，其它功能都正常，原因不明
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256    '气泡提示内容
    uTimeout As Long    '气泡提示显示时间
    szInfoTitle As String * 64    '气泡提示标题
    dwInfoFlags As Long    '气泡提示类型，见 NIIF_*** 部分
End Type
Public Enum ico '气泡提示类型
    无图标 = &H0      '  NIIF_NONE = &H0
    信息图标 = &H1    '  NIIF_INFO = &H1
    警告图标 = &H2    '  NIIF_WARNING = &H2
    错误图标 = &H3    '  NIIF_ERROR = &H3
    托盘图标 = &H4    '  NIIF_GUID = &H4
End Enum
Private IconData As NOTIFYICONDATA

Public Sub 开启托盘(窗口 As Form, Optional 托盘提示信息 As String = "默认为无提示", Optional ByVal 托盘图标 = 0)
    '生成系统托盘图标
    With IconData
        .cbSize = Len(IconData)
        .hwnd = 窗口.hwnd
        .uID = 0
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE    '响应鼠标事件 'WM_LBUTTONDOWN

        If 托盘图标 = 0 Then
            .hIcon = 窗口.Icon    '默认为窗口的图标
        Else
            .hIcon = 托盘图标
        End If

        If 托盘提示信息 <> "默认为无提示" Then
            .szTip = 托盘提示信息 & vbNullChar
        End If
    End With
    Shell_NotifyIcon NIM_ADD, IconData    '增加托盘图标
End Sub

Public Sub 气泡提示(Optional ByVal 提示标题 As String = "气泡提示", Optional 提示内容 As String = "系统托盘气泡提示文字", Optional ByVal 提示类型 As ico, Optional ByVal 托盘图标 = 0)
    With IconData
        .szInfoTitle = 提示标题 & Chr(0)
        .szInfo = 提示内容 & Chr(0)
        .dwInfoFlags = 提示类型
        If 托盘图标 <> 0 Then
        .hIcon = 托盘图标    '更换托盘图标
        End If
    End With
    Shell_NotifyIcon NIM_MODIFY, IconData    '修改托盘图标及相关信息
End Sub

Public Sub 更换托盘图标(Optional ByVal 托盘图标 = 0)
    With IconData
         .szInfoTitle = Chr(0)
        .szInfo = Chr(0)
        If 托盘图标 <> 0 Then
            .hIcon = 托盘图标
        End If
    End With
    Shell_NotifyIcon NIM_MODIFY, IconData    '更换托盘图标
End Sub

Public Sub 关闭托盘()
    Shell_NotifyIcon NIM_DELETE, IconData    '卸载托盘图标
End Sub

