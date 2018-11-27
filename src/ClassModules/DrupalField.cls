VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrupalField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sType As String
Private iLength As Integer
Private vValue As Variant
Private sFieldName As String
Private oForeignEntity As DrupalEntity
Private bIdField As Boolean
'***************************CONSTRUCTOR****************************************
Private Sub class_initialize()
    bIdField = False
End Sub

'***************************LET PROPERTIES*************************************
Public Property Let DataType(sValue As String)
    sType = sValue
End Property

Public Property Let Length(iValue As Integer)
    iLength = iValue
End Property

Public Property Let Value(vInput)
    vValue = vInput
End Property

Public Property Let FieldName(sValue As String)
    sFieldName = sValue
End Property

Public Property Let IdField(bValue As Boolean)
    bIdField = bValue
End Property

Public Property Set ForeignEntity(oEntity)
    Set oForeignEntity = oEntity
End Property

'***************************Get PROPERTIES*************************************
Public Property Get Value()
    Value = vValue
End Property

Public Property Get Length() As Integer
    Length = iLength
End Property

Public Property Get DataType() As String
    DataType = sType
End Property

Public Property Get FieldName() As String
    FieldName = sFieldName
End Property

Public Property Get ForeignEntity() As DrupalEntity
    Set ForeignEntity = oForeignEntity
End Property

Public Property Get IdField() As Boolean
    IdField = bIdField
End Property

Public Sub QuickLoad(DataType, FieldName, Optional Length As Integer = 50)
    sFieldName = FieldName
    sType = DataType
    iLength = Length
End Sub

