VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpdateQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Wheres"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Domains"
Attribute VB_Ext_KEY = "Member2" ,"Fields"
'==============================================================
'Class          UpdateQuery
'Description    Update Query main
'==============================================================
Option Explicit

'____________________________
'Private members
Private oFields As EasyQDll.Fields
Private oDomains As EasyQDll.Domains
Private oCriterion As EasyQDll.Criterion
Private sDBToUpdate As String

'=============================================================
'Initialise and terminate events
'=============================================================
Private Sub Class_Initialize()
    Set oFields = New EasyQDll.Fields
    Set oDomains = New EasyQDll.Domains
    Set oCriterion = New EasyQDll.Criterion
End Sub
Private Sub Class_Terminate()
    If Not oFields Is Nothing Then Set oFields = Nothing
    If Not oDomains Is Nothing Then Set oDomains = Nothing
    If Not oCriterion Is Nothing Then Set oCriterion = Nothing
End Sub

'=============================================================
'Return a reference to the fields object
'=============================================================
Public Property Get Fields() As EasyQDll.Fields
    Set Fields = oFields
End Property

'=============================================================
'Return a reference to the domains object
'=============================================================
Public Property Get Domains() As EasyQDll.Domains
    Set Domains = oDomains
End Property

'=============================================================
'Return a reference to the wheres object
'=============================================================
Public Property Get Criterion() As EasyQDll.Criterion
    Set Criterion = oCriterion
End Property

'=============================================================
'Database to update
'=============================================================
Public Property Let pDBToUpdate(sNewValue As String)
    sDBToUpdate = sNewValue
End Property

'=============================================================
'Return the string
'=============================================================
Public Property Get pGenSQL() As String
Dim oField As EasyQDll.Field
Dim oDomain As EasyQDll.Domain
Dim oCriteria As EasyQDll.Criteria
Dim sValues As String
Dim sDomains As String
Dim sCriteria As String
    If sDBToUpdate = "" Then Exit Property
    sValues = ""
    sDomains = ""
    sCriteria = ""
    
    '________________________________
    'set up the where clause
    For Each oCriteria In oCriterion
        sCriteria = sCriteria & oCriteria.NameOne & GetOperator(oCriteria.Operator)
        If UCase(oCriteria.DataTypeTwo) = "STRING" Or UCase(oCriteria.DataTypeTwo) = "DATE" Then
            sCriteria = sCriteria & "'" & oCriteria.NameTwo & "' AND "
        Else
            sCriteria = sCriteria & oCriteria.NameTwo & " AND "
        End If
    Next oCriteria
    If sCriteria <> "" Then
        sCriteria = Left(sCriteria, Len(sCriteria) - 4)
    Else
        Exit Property
    End If
    
    '________________________________-
    'set up the fields for the update
    For Each oField In oFields
        If oField.DataType = "" Or UCase(oField.DataType) = "DATE" _
        Or UCase(oField.DataType) = "STRING" Then
            sValues = sValues & oField.Name & "='" & oField.Value & "',"
        Else
            sValues = sValues & oField.Name & "=" & oField.Value & ","
        End If
    Next oField
    sValues = Left(sValues, Len(sValues) - 1)
    
    '_______________________________
    'set up the domain names
    For Each oDomain In oDomains
        sDomains = sDomains & oDomain.Name & ","
    Next oDomain
    If sDomains <> "" Then sDomains = Left(sDomains, Len(sDomains) - 1)
    
    If sDomains <> "" Then
        pGenSQL = "UPDATE " & sDBToUpdate & " SET " & sValues & " FROM " & sDomains & _
        " WHERE " & sCriteria
    Else
        pGenSQL = "UPDATE " & sDBToUpdate & " SET " & sValues & " WHERE " & sCriteria
    End If
    
    oDomains.Clear
    oCriterion.Clear
    oFields.Clear
End Property


