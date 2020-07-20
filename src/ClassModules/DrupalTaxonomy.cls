VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrupalTaxonomy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements iDrupalEntity

Private oEntity As DrupalEntity
Private Sub Class_Initialize()
    Set oEntity = Create_DrupalEntity
    
    Dim Tid As DrupalField
    Set Tid = Create_DrupalField
    With Tid
        .FieldName = "tid"
        .DataType = "int"
        .IdField = True
    End With
    With oEntity
        .Table = "taxonomy_term_field_data"
        Set .IdField = Tid
    End With
End Sub

Public Property Get iDrupalEntity_Table()
    iDrupalEntity_Table = oEntity.Table
End Property

Public Property Get iDrupalEntity_ID() As Long
    iDrupalEntity_ID = oEntity.ID
End Property

Public Property Let iDrupalEntity_ID(lValue As Long)
    oEntity.ID = lValue
End Property

Public Property Let iDrupalEntity_Label(vValue As Variant)
    oEntity.Label = vValue
End Property

Public Property Get iDrupalEntity_IdField()
    Set iDrupalEntity_IdField = oEntity.IdField
End Property

Public Property Get iDrupalEntity_LabelField()
    Set iDrupalEntity_LabelField = oEntity.LabelField
End Property

Public Property Get iDrupalEntity_Label()
    iDrupalEntity_Label = oEntity.Label
End Property

Public Function iDrupalEntity_GetFields()
    iDrupalEntity_GetFields = oEntity.GetFields
End Function
