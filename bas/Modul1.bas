Attribute VB_Name = "Modul1"
Option Explicit

Sub Umrechnen()

    Dim x As Integer
    x = Cells(3, 2).Value
    
    Dim y As String
    y = ""
    
    Do
    
    If x = 1 Then
        y = y + "1"
    Else
        Dim z As String
        
        z = x Mod 2
        y = y + z
    End If
    
    Loop Until x > 0
    
    
End Sub
