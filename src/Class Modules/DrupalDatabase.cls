Private MySQLDatabase As SQLDatabase

Private Sub Class_Initialize()
    Set MySQLDatabase = New SQLDatabase
End Sub

Public Property Let DSN (sValue)
    MySQLDatabase.DSN = sValue
End Property

Public Property Let DbType (sValue)
    MySQLDatabase.DbType = sValue
End Property

Public Property Let UserName(sValue As String)
    MySQLDatabase.Username = sValue
End Property

Public Property Let Password(sValue As String)
    MySQLDatabase.Password = sValue
End Property

Public Sub Insert(ByRef oEntity)
    aFields = oEntity.GetFields()
    Set inSQL = New SQLInsert
    inSQL.Table = oEntity.Table
    inSQL.Fields = GetFieldNames(aFields)
    inSQL.Values = GetFieldValues(aFields)
    oEntity.ID = MySQLDatabase.InsertGetNewID(inSQL)
End Sub

Private Function GetFieldNames(aFields)
    Count = 0
    NumFields = UBound(aFields)
    Dim ReturnArray() As String
    ReDim ReturnArray(0 To NumFields)
    Dim Field As DrupalField
    
    For Count = 0 To NumFields
        Set Field = aFields(Count)
        ReturnArray(Count) = Field.FieldName
    Next Count
    GetFieldNames = ReturnArray
End Function

Private Function GetFieldValues(aFields)
    Dim Item As Variant
    Count = 0
    NumFields = UBound(aFields)
    Dim ReturnArray() As String
    ReDim ReturnArray(0 To NumFields)
    Dim Field As DrupalField
    For Count = 0 To NumFields
        Set Field = aFields(Count)
        If Field.DataType = "string" Then
            ReturnArray(Count) = str(Field.Value)
        ElseIf Field.DataType = "entity_reference" Then
            If Field.Value > 0 Then
                'id is set, use it instead of the value
                ReturnArray(Count) = Field.Value
            Else
                'Create a simple subselect
                InnerTable = Field.ExtraValue(0)
                innerValue = Field.ExtraValue(1)
                Set SeSQL = New SQLSelect
                SeSQL.Fields = Array("id")
                SeSQL.Table = InnerTable
                SeSQL.addWhere "name", "=", innerValue
                ReturnArray(Count) = "(" & SeSQL.toString & ") AS " & Field.FieldName
            End If
        Else
            ReturnArray(Count) = Field.Value
        End If
    Next Count
    GetFieldValues = ReturnArray
End Function
