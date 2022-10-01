Attribute VB_Name = "Module3"
Option Explicit

Function NewRate(x As Double, y As String) As Double
  If x < 50000 Then
        NewRate = x + 0.004 * x + 0.00015 * x + 25
    ElseIf x < 500000 Then
        NewRate = x + 0.0037 * x + 0.00015 * x + 25
    ElseIf x < 2000000 Then
        NewRate = x + 0.0034 * x + 0.00015 * x + 25
    ElseIf x < 10000000 Then
        NewRate = x + 0.003 * x + 0.00015 * x + 25
    Else
    NewRate = x + 0.0027 * x + 0.00015 * x + 25
    End If
If y = "IPO" Then NewRate = x

End Function
        
        
