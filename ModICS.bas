Attribute VB_Name = "ModICS"
Option Explicit

Public Function CheckICSAvailablity()

End Function

Public Function EnableICS()
RunPowershellCommand (App.Path + "\enablesharing.ps1")
End Function
