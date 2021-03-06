VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrupalEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements iDrupalEntity
'***************************DRUPAL ENTITY CLASS********************************
'Provides a parent class for Drupal Entities.
'Provides Shared Methods that can be used by all Entities
'
'*********************************VARIABLES************************************
Private sTable As String        'The table name
Private oId As DrupalField      'Primary Key
Private oLabel As DrupalField   'The entity name
Private bnoLabel As Boolean
Private dFields As Dictionary   'An array of entity fields. This currently only
                                'supports base fields.

'***************************CONSTRUCTOR****************************************
Private Sub Class_Initialize()
    Set dFields = New Dictionary
    bnoLabel = False
    InitializeFields
End Sub

Private Sub InitializeFields()
    Set oId = New DrupalField
    InitializeField oId, "integer", "id"
    oId.IdField = True
    
    Set oLabel = New DrupalField
    InitializeField oLabel, "string", "name", 50
End Sub

Private Sub InitializeField(ByRef oField As DrupalField, sType As String, sName As String, Optional iLength As Integer = 0)
    oField.DataType = sType
    oField.FieldName = sName
    If iLength <> 0 Then oField.Length = iLength
End Sub

Public Sub CreateField(sType As String, sName As String, Optional iLength As Integer = 0)
    Dim MyField As New DrupalField
    MyField.DataType = sType
    MyField.FieldName = sName
    If iLength <> 0 Then MyField.Length = iLength
    AddField MyField
End Sub

'***************************LET PROPERTIES*************************************
Public Property Let Label(sValue As String)
    oLabel.Value = Left(sValue, oLabel.Length - Drupal_CountCharsToEscape(sValue))
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
Public Sub noLabel()
    bnoLabel = True
End Sub
Public Sub SetValue(sFieldName As String, vValue)
    dFields(sFieldName).Value = vValue
End Sub

Public Function GetValue(sFieldName As String)
    GetValue = dFields(sFieldName).Value
End Function

Public Sub SetTargetValue(sFieldName As String, vValue As Variant)
    dFields(sFieldName).TargetEntity.Label = vValue
End Sub

'WIP
'Create a new Entity from an array of values.
'For example, from the results of a Select Query
'Public Sub CreateFromArray(oEntity, ArrayOfValues)
'    FieldList = oEntity.GetFields
'    numfields = UBound(FieldList)
'    Dim counter As Integer
'    Dim Item As DrupalField
'    For counter = 0 To numfields
'        Set Item = FieldList(counter)
'        Item.Value = ArrayOfValues(counter)
'    Next counter
'End Sub

Public Sub AddField(ByVal oField As DrupalField)
    Dim FieldName As String
    FieldName = oField.FieldName
    Set dFields(FieldName) = oField
End Sub

Public Function CreateEntityReference(FieldName As String, TargetEntity As iDrupalEntity)
    Dim NewField As New DrupalField
    NewField.Create "entity_reference", FieldName
    Set NewField.TargetEntity = TargetEntity
    'If ForeignEntity.ID is set, should NewField.Value = ForeignEntity.ID?
    AddField NewField
End Function

' Function: GetFields
'
' Get an array of all fields in the entity
Public Function GetFields() As Variant
    Dim Number_Fields As Long
    Number_Fields = dFields.count
    
    Dim FieldArray() As DrupalField
    ReDim FieldArray(Number_Fields)
    Dim i As Long
    i = 0
    If bnoLabel = False Then
        ReDim FieldArray(Number_Fields)
        Set FieldArray(i) = oLabel
        i = i + 1
    Else
        ReDim FieldArray(Number_Fields - 1)
    End If
    Dim key As Variant
    
    For Each key In dFields.Keys
        Set FieldArray(i) = dFields(key)
        i = i + 1
    Next key
    GetFields = FieldArray
End Function

'***************************Interface Methods**********************************
Public Property Get iDrupalEntity_Table()
    iDrupalEntity_Table = sTable
End Property

Public Property Get iDrupalEntity_ID() As Long
    iDrupalEntity_ID = oId.Value
End Property

Public Property Let iDrupalEntity_ID(lValue As Long)
    oId.Value = lValue
End Property

Public Property Let iDrupalEntity_Label(vValue As Variant)
    oLabel.Value = vValue
End Property

Public Property Get iDrupalEntity_IdField()
    Set iDrupalEntity_IdField = oId
End Property

Public Property Get iDrupalEntity_LabelField()
    Set iDrupalEntity_LabelField = oLabel
End Property

Public Property Get iDrupalEntity_Label()
    iDrupalEntity_Label = oLabel.Value
End Property

Public Function iDrupalEntity_GetFields()
    iDrupalEntity_GetFields = GetFields()
End Function

