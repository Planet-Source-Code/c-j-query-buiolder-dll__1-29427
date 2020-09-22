VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Query Builder"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   Icon            =   "frmTestEQ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select Query"
      Height          =   495
      Left            =   6975
      TabIndex        =   0
      Top             =   2550
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Dele&te Query"
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Query"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   1275
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert Query"
      Height          =   495
      Left            =   6945
      TabIndex        =   1
      Top             =   630
      Width           =   1215
   End
   Begin VB.Label lblQuery 
      Alignment       =   2  'Center
      Caption         =   "Query Builder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   180
      TabIndex        =   4
      Top             =   270
      Width           =   6555
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'Form       Test control
'-------------------------------------------------------------------------
'Note : About the control
'1...
'In the criterion collection, the datatype and operator are optional.
'They will default to unquoted data for datatype and = for operator if
'not specified.
'This means that if you specify "Key1", "Key2" the sql will read Key1=Key2
'If you specify "Key1", "Jack", "String" the sql will read Key1='Jack'
'If you specify "Key1", "Key2",,"<>" the sql will read Key1<>Key2
'Please note that this default is for criterion only.
'2...
'When adding fields, the datatype will default to string if no datatype
'is specified.
'If you add "Name", "Chris", the value will appear as 'Chris'
'If you add "Key", "123", "Integer" the value will appear as 123
'=========================================================================
Option Explicit

'________________________________
'Private members
Private oSQL As EasyQDll.Application

'===========================================================
'Build the delete query
'===========================================================
Private Sub cmdDelete_Click()
    lblQuery.Caption = ""
    
    'Table to delete from
    oSQL.DeleteQuery.pDomain = "Clients"
    
    'Where criteria
    oSQL.DeleteQuery.Criterion.Add "Code", "123", "Integer", "<"
    
    'build the sql
    lblQuery.Caption = oSQL.DeleteQuery.pGenSQL
End Sub

'============================================================
'Build the insert query
'============================================================
Private Sub cmdInsert_Click()
    lblQuery.Caption = ""
    'Table to insert to
    oSQL.InsertQuery.pDomain = "Clients"
    
    'Fields and values to insert
    oSQL.InsertQuery.Fields.Add "Name", "Jack"
    oSQL.InsertQuery.Fields.Add "Surname", "Sprat"
    oSQL.InsertQuery.Fields.Add "AccountNo", "123456789", "Long"
    oSQL.InsertQuery.Fields.Add "IDNo", "8201010000000"
    oSQL.InsertQuery.Fields.Add "Code", "123", "Integer"
    oSQL.InsertQuery.Fields.Add "PaymentTerms", "Monthly"
    oSQL.InsertQuery.Fields.Add "AvailableCredit", "1000000.00", "Currency"
    
    'Initiate the sql build
    lblQuery.Caption = oSQL.InsertQuery.pGenSQL
End Sub

'============================================================
'Build a select string
'============================================================
Private Sub cmdSelect_Click()
    lblQuery.Caption = ""
    
    'To select only certain fields pSelectFields must be true
    oSQL.SelectQuery.pSelectFields = True
    
    'Tables to select from
    oSQL.SelectQuery.Domains.Add "Clients"
    oSQL.SelectQuery.Domains.Add "Status"
    
    'Fields to select (If pSelectFields = true)
    oSQL.SelectQuery.Domains("Clients").Fields.Add "Name"
    oSQL.SelectQuery.Domains("Clients").Fields.Add "Surname"
    oSQL.SelectQuery.Domains("Status").Fields.Add "ClientCode"
    
    'Where criteria
    oSQL.SelectQuery.Criterion.Add "Clients.Code", "Status.ClientCode"
    oSQL.SelectQuery.Criterion.Add "Clients.Code", "123", "Integer", "<>"
    
    'Order if any
    oSQL.SelectQuery.Domains("Clients").Orders.Add "IDNo"
    
    'Group if any
    oSQL.SelectQuery.Domains("Clients").Groups.Add "PaymentTerms"
    
    'Initiate the sql build
    lblQuery.Caption = oSQL.SelectQuery.pGenSQL
    
End Sub

'=============================================================
'Build an update query
'=============================================================
Private Sub cmdUpdate_Click()
    lblQuery.Caption = ""
    'Table to update
    oSQL.UpdateQuery.pDBToUpdate = "Clients"
    
    'Fields and values to update
    oSQL.UpdateQuery.Fields.Add "Surname", "Spells"
    oSQL.UpdateQuery.Fields.Add "AccountNo", "111111111", "Long"
    oSQL.UpdateQuery.Fields.Add "AvailableCredit", "500000", "Currency"
    
    'If the data to be updated has more the one table from which the criteria
    'is to be selected add the domains here. Please note that if you are
    'using more than one table you must specify the table name and field in the
    'Criteria
    oSQL.UpdateQuery.Domains.Add "Clients"
    oSQL.UpdateQuery.Domains.Add "Status"
    
    'Add the where criteria
    oSQL.UpdateQuery.Criterion.Add "Clients.Code", "Status.ClientCode"
    oSQL.UpdateQuery.Criterion.Add "Clients.PaymentsTerms", "Monthly", "String"
    
    'Initiate the sql build
    lblQuery.Caption = oSQL.UpdateQuery.pGenSQL
End Sub

'===========================================================
'Initialise the dll
'===========================================================
Private Sub Form_Load()
    Set oSQL = New EasyQDll.Application
End Sub

