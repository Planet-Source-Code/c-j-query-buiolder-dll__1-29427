VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InsertQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Fields"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'==============================================================
'Class          InsertQuery
'Description    Insert Query main
'==============================================================
Option Explicit

'____________________________
'Private members
Private oFields As EasyQDll.Fields
Private sDomain As String

'=============================================================
'Initialise and terminate events
'=============================================================
Private Sub Class_Initialize()
    Set oFields = New EasyQDll.Fields
End Sub
Private Sub Class_Terminate()
    If Not oFields Is Nothing Then Set oFields = Nothing
End Sub

'=============================================================
'Return a reference to the fields object
'=============================================================
Public Property Get Fields() As EasyQDll.Fields
    Set Fields = oFields
End Property

'=============================================================
'The domain for insert
'=============================================================
Public Property Let pDomain(sNewValue As String)
    sDomain = sNewValue
End Property

'=============================================================
'Return the string
'=============================================================
Public Property Get pGenSQL() As String
Dim oField As EasyQDll.Field
Dim sNames As String
Dim sValues As String

    '__________________________________
    'If not domain has been selected then exit
    If sDomain = "" Then Exit Property
    sNames = ""
    sValues = ""
    
    '___________________________________
    'Build the field string for the sql
    For Each oField In oFields
        sNames = sNames & oField.Name & ","
        If oField.DataType = "" Or UCase(oField.DataType) = "STRING" _
        Or UCase(oField.DataType) = "DATE" Then
            sValues = sValues & "'" & oField.Value & "',"
        Else
            sValues = sValues & oField.Value & ","
        End If
    Next oField
    
    '____________________________________
    'Build the sql string
    pGenSQL = "INSERT INTO " & sDomain & " (" & Left(sNames, Len(sNames) - 1) & ") VALUES (" & Left(sValues, Len(sValues) - 1) & ")"
    oFields.Clear
    
End Property
