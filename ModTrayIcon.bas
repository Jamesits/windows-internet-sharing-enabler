Attribute VB_Name = "����ģ��"
Option Explicit
    '��������ͼ�� Picture1.Picture
    '������ʾ Text1.Text, Text2.Text, ��Ϣͼ��, Picture1.Picture
    
    
 '  Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim lMsg As Single
   ' lMsg = X / Screen.TwipsPerPixelX
   ' If MinFlag = True Then
   '     Select Case lMsg
  '      Case 513    'WM_LBUTTONUP ���������ͼ��
            '����������
  '      Case 515
            '������˫��
  '          MinFlag = False
  '          Me.WindowState = vbNormal
  '          Me.Show
  '          ������ʾ "�������ڡ�"
  '      Case 517    'WM_RBUTTONUP ����Ҽ����ͼ��
  '          PopupMenu �Ҽ��˵�    '�������ϵͳ����ͼ���ϵ��Ҽ����򵯳��˵�
  '      Case 518
   '         '����Ҽ�˫��
  '      End Select
 '   End If
'End Sub

'Private Sub Form_Resize()
'    If Me.WindowState = 1 And MinFlag = False Then
    '    MinFlag = True
   '     Me.Hide    '���ش���
  '      ������ʾ "�ڴ˴�˫�������������������ڡ�" & vbCrLf & "�ڴ˴ε������Ҽ����������˵���"
 '   End If
'End Sub
    
    
    
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Const NIM_ADD = &H0    '��������������һ��ͼ��
Private Const NIM_DELETE = &H2    'ɾ���������е�һ��ͼ��
Private Const NIM_MODIFY = &H1    '�޸��������и�ͼ����Ϣ
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
    cbSize As Long    '�����ݽṹ�Ĵ�С
    hwnd As Long    '������������ͼ��Ĵ��ھ��
    uID As Long    '�������������ͼ��ı�ʶ
    uFlags As Long    '������ͼ�깦�ܿ��ƣ�����������ֵ����ϣ�һ��ȫ������
    uCallbackMessage As Long    '������ͼ��ͨ�������û����򽻻���Ϣ���������Ϣ�Ĵ�����hWnd����
    hIcon As Long    '�������е�ͼ��Ŀ��ƾ��
    szTip As String * 128    'ͼ�����ʾ��Ϣ����Ҫ����������ʾ��Ϣ����һ��Ҫ128���ԣ�Ϊ64���޷��������ݣ��������ܶ�������ԭ����
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256    '������ʾ����
    uTimeout As Long    '������ʾ��ʾʱ��
    szInfoTitle As String * 64    '������ʾ����
    dwInfoFlags As Long    '������ʾ���ͣ��� NIIF_*** ����
End Type
Public Enum ico '������ʾ����
    ��ͼ�� = &H0      '  NIIF_NONE = &H0
    ��Ϣͼ�� = &H1    '  NIIF_INFO = &H1
    ����ͼ�� = &H2    '  NIIF_WARNING = &H2
    ����ͼ�� = &H3    '  NIIF_ERROR = &H3
    ����ͼ�� = &H4    '  NIIF_GUID = &H4
End Enum
Private IconData As NOTIFYICONDATA

Public Sub ��������(���� As Form, Optional ������ʾ��Ϣ As String = "Ĭ��Ϊ����ʾ", Optional ByVal ����ͼ�� = 0)
    '����ϵͳ����ͼ��
    With IconData
        .cbSize = Len(IconData)
        .hwnd = ����.hwnd
        .uID = 0
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE    '��Ӧ����¼� 'WM_LBUTTONDOWN

        If ����ͼ�� = 0 Then
            .hIcon = ����.Icon    'Ĭ��Ϊ���ڵ�ͼ��
        Else
            .hIcon = ����ͼ��
        End If

        If ������ʾ��Ϣ <> "Ĭ��Ϊ����ʾ" Then
            .szTip = ������ʾ��Ϣ & vbNullChar
        End If
    End With
    Shell_NotifyIcon NIM_ADD, IconData    '��������ͼ��
End Sub

Public Sub ������ʾ(Optional ByVal ��ʾ���� As String = "������ʾ", Optional ��ʾ���� As String = "ϵͳ����������ʾ����", Optional ByVal ��ʾ���� As ico, Optional ByVal ����ͼ�� = 0)
    With IconData
        .szInfoTitle = ��ʾ���� & Chr(0)
        .szInfo = ��ʾ���� & Chr(0)
        .dwInfoFlags = ��ʾ����
        If ����ͼ�� <> 0 Then
        .hIcon = ����ͼ��    '��������ͼ��
        End If
    End With
    Shell_NotifyIcon NIM_MODIFY, IconData    '�޸�����ͼ�꼰�����Ϣ
End Sub

Public Sub ��������ͼ��(Optional ByVal ����ͼ�� = 0)
    With IconData
         .szInfoTitle = Chr(0)
        .szInfo = Chr(0)
        If ����ͼ�� <> 0 Then
            .hIcon = ����ͼ��
        End If
    End With
    Shell_NotifyIcon NIM_MODIFY, IconData    '��������ͼ��
End Sub

Public Sub �ر�����()
    Shell_NotifyIcon NIM_DELETE, IconData    'ж������ͼ��
End Sub

