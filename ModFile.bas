Attribute VB_Name = "ModFile"
Option Explicit

Public Function File2Text(filename$) As String
    Dim handle As Integer
    handle = FreeFile
    Open filename$ For Input As #handle
    File2Text = Input$(LOF(handle), handle)
    Close #handle
End Function
