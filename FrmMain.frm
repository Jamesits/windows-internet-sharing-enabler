VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox TxtMsg 
      Height          =   1455
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1320
      Width           =   3975
   End
   Begin VB.CommandButton CmdEnableSharing 
      Caption         =   "Start"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdEnableSharing_Click()
EnableICS
End Sub

Private Sub Form_Load()
If hasAdmin = False Then
    log "Caution: this program needs administrator function to run correctly."
End If
End Sub
