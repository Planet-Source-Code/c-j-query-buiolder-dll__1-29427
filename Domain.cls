VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Domain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'=============================================================
'Class          Domain
'Description    Domain Declaration
'=============================================================
Option Explicit

'_________________________
'Public members
Public Name As String

'_________________________
'Private members
Private oFields As EasyQDll.Fields
Private oGroups As EasyQDll.Groups
Private oOrders As EasyQDll.Orders

'==============================================================
'Initialise and terminate events
'==============================================================
Private Sub Class_Initialize()
    Set oFields = New EasyQDll.Fields
    Set oGroups = New EasyQDll.Groups
    Set oOrders = New EasyQDll.Orders
End Sub
Private Sub Class_Terminate()
    If Not oFields Is Nothing Then Set oFields = Nothing
    If Not oGroups Is Nothing Then Set oGroups = Nothing
    If Not oOrders Is Nothing Then Set oOrders = Nothing
End Sub

'=============================================================
'Return a reference to the fields object
'=============================================================
Public Property Get Fields() As EasyQDll.Fields
    Set Fields = oFields
End Property

'=============================================================
'Return a reference to the groups object
'=============================================================
Public Property Get Groups() As EasyQDll.Groups
    Set Groups = oGroups
End Property

'============================================================
'Return a reference to the Orders object
'============================================================
Public Property Get Orders() As EasyQDll.Orders
    Set Orders = oOrders
End Property
