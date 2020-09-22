VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Wheres"
Attribute VB_Ext_KEY = "Member1" ,"Domains"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'===========================================================
'Class          SelectQuery
'Description    SelectQuery declaration
'===========================================================
Option Explicit

'__________________________
'Private members
Private oDomains As EasyQDll.Domains
Private oCriterion As EasyQDll.Criterion
Private bSelectFields As Boolean

'============================================================
'Initialise and terminate events
'============================================================
Private Sub Class_Initialize()
    Set oDomains = New EasyQDll.Domains
    Set oCriterion = New EasyQDll.Criterion
End Sub
Private Sub Class_Terminate()
    If Not oDomains Is Nothing Then Set oDomains = Nothing
    If Not oCriterion Is Nothing Then Set oCriterion = Nothing
End Sub

'============================================================
'Return a reference to the Domains object
'============================================================
Public Property Get Domains() As EasyQDll.Domains
    Set Domains = oDomains
End Property

'============================================================
'Return a reference to the Wheres object
'============================================================
Public Property Get Criterion() As EasyQDll.Criterion
    Set Criterion = oCriterion
End Property

'============================================================
'Property for all fields
'============================================================
Public Property Let pSelectFields(bNewValue As Boolean)
    bSelectFields = bNewValue
End Property

'=============================================================
'Return the string
'=============================================================
Public Property Get pGenSQL() As String
Dim oField As EasyQDll.Field
Dim oDomain As EasyQDll.Domain
Dim oCriteria As EasyQDll.Criteria
Dim oOrder As EasyQDll.Order
Dim oGroup As EasyQDll.Group
Dim sFields As String
Dim sDomains As String
Dim sCriteria As String
Dim sGroups As String
Dim sOrders As String
    sFields = ""
    sDomains = ""
    sCriteria = ""
    sGroups = ""
    sOrders = ""
    
    '________________________________________________
    'If bSelectFields is false then use wildcard
    If bSelectFields = True Then
        For Each oDomain In oDomains
            For Each oField In oDomain.Fields
                sFields = sFields & oDomain.Name & "." & oField.Name & ","
            Next oField
        Next oDomain
        sFields = Left(sFields, Len(sFields) - 1)
    Else
        sFields = "*"
    End If
    
    '______________________________________________
    'Setup the domain string
    For Each oDomain In oDomains
        sDomains = sDomains & oDomain.Name & ","
        'Check the groupings
        For Each oGroup In oDomain.Groups
            sGroups = sGroups & oDomain.Name & "." & oGroup.Name & ","
        Next oGroup
        'Check the order by
        For Each oOrder In oDomain.Orders
            sOrders = sOrders & oDomain.Name & "." & oOrder.Name & ","
        Next oOrder
    Next oDomain
    If sDomains <> "" Then sDomains = Left(sDomains, Len(sDomains) - 1)
    If sGroups <> "" Then sGroups = Left(sGroups, Len(sGroups) - 1)
    If sOrders <> "" Then sOrders = Left(sOrders, Len(sOrders) - 1)
    
    '_____________________________________
    'Setup the where string
    For Each oCriteria In oCriterion
        sCriteria = sCriteria & oCriteria.NameOne & GetOperator(oCriteria.Operator)
        If UCase(oCriteria.DataTypeTwo) = "STRING" Or UCase(oCriteria.DataTypeTwo) = "DATE" Then
            sCriteria = sCriteria & "'" & oCriteria.NameTwo & "' AND "
        Else
            sCriteria = sCriteria & oCriteria.NameTwo & " AND "
        End If
    Next oCriteria
    If sCriteria <> "" Then sCriteria = Left(sCriteria, Len(sCriteria) - 4)
    
    '____________________________________
    'Build the sql string
    pGenSQL = "SELECT " & sFields & " FROM " & sDomains
    If sCriteria <> "" Then pGenSQL = pGenSQL & " WHERE " & sCriteria
    If sGroups <> "" Then pGenSQL = pGenSQL & " GROUP BY " & sGroups
    If sOrders <> "" Then pGenSQL = pGenSQL & " ORDER BY " & sOrders
    oDomains.Clear
    oCriterion.Clear
    bSelectFields = False
End Property

