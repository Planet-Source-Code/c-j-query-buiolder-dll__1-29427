VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=====================================================
'Class          Application
'Description    Dll entry point
'-----------------------------------------------------
'Written By C.Robb.
'=====================================================
Option Explicit

'______________________________
'Private members
Private oInsertQuery As EasyQDll.InsertQuery
Private oUpdateQuery As EasyQDll.UpdateQuery
Private oSelectQuery As EasyQDll.SelectQuery
Private oDeleteQuery As EasyQDll.DeleteQuery

'======================================================
'Initialise and terminate events
'======================================================
Private Sub Class_Initialize()
    Set oInsertQuery = New EasyQDll.InsertQuery
    Set oUpdateQuery = New EasyQDll.UpdateQuery
    Set oSelectQuery = New EasyQDll.SelectQuery
    Set oDeleteQuery = New EasyQDll.DeleteQuery
End Sub

Private Sub Class_Terminate()
    If Not oInsertQuery Is Nothing Then Set oInsertQuery = Nothing
    If Not oUpdateQuery Is Nothing Then Set oUpdateQuery = Nothing
    If Not oSelectQuery Is Nothing Then Set oSelectQuery = Nothing
    If Not oDeleteQuery Is Nothing Then Set oDeleteQuery = Nothing
End Sub

'======================================================
'Expose the insertquery object
'======================================================
Public Property Get InsertQuery() As EasyQDll.InsertQuery
    Set InsertQuery = oInsertQuery
End Property

'======================================================
'Expose the updatequery object
'======================================================
Public Property Get UpdateQuery() As EasyQDll.UpdateQuery
    Set UpdateQuery = oUpdateQuery
End Property

'======================================================
'Expose the selectquery object
'======================================================
Public Property Get SelectQuery() As EasyQDll.SelectQuery
    Set SelectQuery = oSelectQuery
End Property

'======================================================
'Expose the deletequery object
'======================================================
Public Property Get DeleteQuery() As EasyQDll.DeleteQuery
    Set DeleteQuery = oDeleteQuery
End Property


