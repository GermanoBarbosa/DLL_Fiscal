Attribute VB_Name = "mLogger"
Option Explicit

Public gLogger As cLogger

Public Function GetLogger() As cLogger
    If gLogger Is Nothing Then
        Set gLogger = New cLogger
    End If
    Set GetLogger = gLogger
End Function
