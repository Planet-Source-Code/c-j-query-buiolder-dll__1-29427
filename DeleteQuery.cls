VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeleteQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================
'Class          SelectQuery
'Description    SelectQuery declaration
'===========================================================
Option Explicit

'__________________________
'Private members
Private oCriterion As EasyQDll.Criterion
Private sDomain As String

'============================================================
'Initialise and terminate events
'============================================================
Private Sub Class_Initialize()
    Set oCriterion = New EasyQDll.Criterion
End Sub
Private Sub Class_Terminate()
    If Not oCriterion Is Nothing Then Set oCriterion = Nothing
End Sub

'============================================================
'Return a reference to the Domains object
'============================================================
Public Property Let pDomain(sNewValue As String)
    sDomain = sNewValue
End Property

'============================================================
'Return a reference to the Wheres object
'============================================================
Public Property Get Criterion() As EasyQDll.Criterion
    Set Criterion = oCriterion
End Property

'=============================================================
'Return the string
'=============================================================
Public Property Get pGenSQL() As String
Dim oCriteria As EasyQDll.Criteria
Dim sCriteria As String
    '_____________________________
    'If no domain was specified then exit
    If sDomain = "" Then Exit Property
    
    '_____________________________
    'Build the criteria string
    For Each oCriteria In oCriterion
        sCriteria = sCriteria & oCriteria.NameOne & GetOperator(oCriteria.Operator)
        If UCase(oCriteria.DataTypeTwo) = "STRING" Or UCase(oCriteria.DataTypeTwo) = "DATE" Then
            sCriteria = sCriteria & "'" & oCriteria.NameTwo & "' AND "
        Else
            sCriteria = sCriteria & oCriteria.NameTwo & " AND "
        End If
    Next oCriteria
    If sCriteria <> "" Then sCriteria = Left(sCriteria, Len(sCriteria) - 4)
    
    '________________________________
    'Build the delete sql string
    pGenSQL = "DELETE FROM " & sDomain
    If sCriteria <> "" Then pGenSQL = pGenSQL & " WHERE " & sCriteria
    oCriterion.Clear
    sDomain = ""
    
End Property


