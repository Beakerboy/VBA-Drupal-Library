Private sType As String
Private iLength As Integer
Private vValue As Variant
Private sFieldName As String
Private vExtraValue

Private Sub Class_Initialize()
    vExtraValue = 0
End Sub

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

Public Property Get ExtraValue()
    ExtraValue = vExtraValue
End Property
