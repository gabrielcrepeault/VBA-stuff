Attribute VB_Name = "General"
' Proc�dure d'optimisation de macro
Public Sub Optimisation(trigger As Boolean)
    With Application
    If trigger Then
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    Else
        .EnableEvents = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .Calculate
    End If
    End With
End Sub


