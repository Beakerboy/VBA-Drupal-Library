VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrupalEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************DRUPAL ENTITY CLASS********************************
'Provides a parent class for Drupal Entities.
'Provides Shared Methods that can be used by all Entities
'
'*********************************VARIABLES************************************
Private sTable As String        'The table name
Private oId As DrupalField      'Primary Key
Private oLabel As DrupalField    'The entity name

'***************************CONSTRUCTOR****************************************
Private Sub Class_Initialize()
    InitializeFields
End Sub

Private Sub InitializeFields()
    Set oId = New DrupalField
    InitializeField oId, "integer", "id"

    Set oLabel = New DrupalField
    InitializeField oLabel, "string", "name", 50
End Sub

Public Sub InitializeField(ByRef oField As DrupalField, sType As String, sName As String, Optional iLength As Integer = 0)
    oField.DataType = sType
    oField.FieldName = sName
    If iLength <> 0 Then oField.Length = iLength
End Sub

'***************************LET PROPERTIES*************************************
Public Property Let Label(sValue As String)
    oLabel.Value = sValue
End Property

Public Property Let Table(sValue As String)
    sTable = sValue
End Property

Public Property Let ID(lValue As Long)
   oId.Value = lValue
End Property

Public Property Set LabelField(oValue As DrupalField)
    Set oLabel = oValue
End Property

Public Property Set IdField(oValue As DrupalField)
    Set oId = oValue
End Property

'***************************Get PROPERTIES*************************************
Public Property Get ID() As Long
    ID = oId.Value
End Property

Public Property Get Label() As String
    Label = oLabel.Value
End Property

Public Property Get Table() As String
    Table = sTable
End Property

Public Property Get LabelField() As DrupalField
    Set LabelField = oLabel
End Property

Public Property Get IdField() As DrupalField
    Set IdField = oId
End Property
'******************************METHODS*****************************************
Public Sub CreateFromArray(oEntity, ArrayOfValues)
    FieldList = oEntity.GetFields
    numfields = UBound(FieldList)
    Dim counter As Integer
    Dim Item As DrupalField
    For counter = 0 To numfields
        Set Item = FieldList(counter)
        Item.Value = ArrayOfValues(counter)
    Next counter
End Sub
