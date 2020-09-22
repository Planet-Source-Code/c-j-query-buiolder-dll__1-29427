VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Criterion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Where"
Attribute VB_Ext_KEY = "Member0" ,"Where"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=======================================================
'Class          Criterion
'Description    Criterion collection methods
'=======================================================
Option Explicit

'______________________________
'Private members
Private mCol As Collection

'=======================================================
'Add the where names
'=======================================================
Public Function Add(sNameOne As String, sNameTwo As String, Optional sDataTypeTwo As String, Optional sOperator As String) As Criteria
    'create a new object
    Dim objNewMember As Criteria
    Set objNewMember = New Criteria


    'set the properties passed into the method
    objNewMember.NameOne = sNameOne
    objNewMember.NameTwo = sNameTwo
    objNewMember.DataTypeTwo = sDataTypeTwo
    objNewMember.Operator = sOperator
    mCol.Add objNewMember
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Criteria
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

'=========================================================
'Clear the collection
'=========================================================
Public Sub Clear()
    While mCol.Count
        mCol.Remove mCol.Count
    Wend
End Sub
