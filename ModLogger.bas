Attribute VB_Name = "ModLogger"
Option Explicit

Public Sub log(s As String)
FrmMain.TxtMsg.Text = FrmMain.TxtMsg.Text + vbCrLf + s
End Sub
