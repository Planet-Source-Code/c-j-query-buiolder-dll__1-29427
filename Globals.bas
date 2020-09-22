Attribute VB_Name = "Globals"
'====================================================
'Module         Globals
'====================================================
Option Explicit

'====================================================
'Return the operator
'====================================================
Public Function GetOperator(sOperator As String) As String
    If sOperator = "" Then
        GetOperator = "="
    Else
        GetOperator = sOperator
    End If
End Function
