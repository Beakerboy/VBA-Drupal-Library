VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrupalField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private sType As String
Private iLength As Integer
Private vValue As Variant
Private sFieldName As String
Private oTarget As iDrupalEntity
Private bIdField As Boolean
'***************************CONSTRUCTOR****************************************
Private Sub Class_Initialize()
    bIdField = False
End Sub

'***************************LET PROPERTIES*************************************
Public Property Let DataType(sValue As String)
    'Allowed Types
        'boolean
        'string
        'timestamp
    sType = sValue
End Property

Public Property Let Length(iValue As Integer)
    iLength = iValue
End Property

Public Property Let Value(vInput)
    vValue = Validate(vInput)
End Property

Public Property Let FieldName(sValue As String)
    sFieldName = sValue
End Property

Public Property Let IdField(bValue As Boolean)
    bIdField = bValue
End Property

Public Property Set TargetEntity(oEntity As iDrupalEntity)
    Set oTarget = oEntity
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

Public Property Get TargetEntity() As iDrupalEntity
    Set TargetEntity = oTarget
End Property

Public Property Get IdField() As Boolean
    IdField = bIdField
End Property

Public Sub Create(DataType, FieldName, Optional Length As Integer = 50)
    sFieldName = FieldName
    sType = DataType
    iLength = Length
End Sub

Private Function Validate(vValue)
    Select Case sType
        Case Is = "boolean"
            If VarType(vValue) = vbBoolean Then
                Validate = vValue
            Else
                Dim slcase As String
                slcase = lcase(vValue)
                If slcase = "true" Or slcase = "t" Or slcase = "true" Or slcase = "y" Or slcase = "yes" Or slcase = "on" Or slcase = "1" Then
                    Validate = True
                Else
                    Validate = False
                End If
            End If
        Case Is = "decimal"
            Dim dData As Double
            dData = vValue
            Validate = dData
        Case Is = "string"
            Dim sData As String
            sData = vValue
            If iLength > 0 Then
                sData = Left(sData, iLength)
                Validate = Left(sData, iLength - Drupal_CountCharsToEscape(sData))
            Else
                Validate = sData
            End If
        Case Is = "timestamp"
            Dim Timestamp As Long
            Timestamp = vValue
            Validate = Timestamp
        Case Else
            Validate = vValue
    End Select
End Function
