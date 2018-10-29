VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrupalDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************DRUPAL DATABASE CLASS******************************
'Provides methods for manupulating Drupal Entities with VBA
'
Option Explicit
'*********************************VARIABLES************************************
Private MySQLDatabase As SQLDatabase

'***************************CONSTRUCTOR****************************************
Private Sub Class_Initialize()
    Set MySQLDatabase = New SQLDatabase
End Sub

'***************************LET PROPERTIES*************************************
Public Property Let DSN(sValue As String)
    MySQLDatabase.DSN = sValue
End Property

Public Property Let DBType(sValue As String)
    MySQLDatabase.DBType = sValue
End Property

Public Property Let Username(sValue As String)
    MySQLDatabase.Username = sValue
End Property

Public Property Let Password(sValue As String)
    MySQLDatabase.Password = sValue
End Property

Public Property Let TestMode(bValue As Boolean)
    MySQLDatabase.TestMode = bValue
End Property

'******************************METHODS*****************************************

'Create a new record in the database using values from the class object
Public Sub Insert(ByRef oEntity)
    Dim inSQL As SQLInsert
    Set inSQL = New SQLInsert
    inSQL.Table = oEntity.Table
    
    Dim aFields As Variant
    aFields = oEntity.GetFields()
    inSQL.Fields = GetFieldNames(aFields)
    inSQL.Values = getFieldValues(aFields)
    oEntity.ID = MySQLDatabase.InsertGetNewId(inSQL)
End Sub

'Create an array of field names from the array of DrupalField objects
Private Function GetFieldNames(aFields) As Variant
    Dim numfields As Integer
    numfields = UBound(aFields)
    Dim ReturnArray() As String
    ReDim ReturnArray(numfields)
    Dim Item As DrupalField
    
    Dim counter As Integer
    counter = 0
    For counter = 0 To numfields
        Set Item = aFields(counter)
        ReturnArray(counter) = Item.FieldName
    Next counter
    GetFieldNames = ReturnArray
End Function

'Create an array of field values from the array of DrupalField objects.
'Sanitize the data to prepare it for the database.
Private Function getFieldValues(aFields)
    Dim numfields As Integer
    numfields = UBound(aFields)
    Dim ReturnArray() As Variant
    ReDim ReturnArray(numfields)
    Dim Item As DrupalField
    Dim counter As Integer
    counter = 0
    For counter = 0 To numfields
        Set Item = aFields(counter)
        If Item.DataType = "string" Then
            ReturnArray(counter) = str(Item.Value)
        ElseIf Item.DataType = "entity_reference" Then
            If Item.Value > 0 Then
                'id is set, use it instead of the value
                'Should we use Item.ForeignEntity.Id instead of Item.Value?
                ReturnArray(counter) = Item.Value
            Else
                'Create a simple subselect
                Dim SeSQL As SQLSelect
                Set SeSQL = New SQLSelect
                SeSQL.Fields = Array(Item.ForeignEntity.IdField.FieldName)
                SeSQL.Table = Item.ForeignEntity.Table
                SeSQL.addWhere Item.ForeignEntity.LabelField.FieldName, "=", str(Item.ForeignEntity.Label)
                Dim MySubselect As SQLSubselect
                Set MySubselect = New SQLSubselect
                MySubselect.SelectAs = Item.FieldName
                Set MySubselect.SelectSQL = SeSQL
                MySubselect.SelectAs = Item.FieldName
                ReturnArray(counter) = MySubselect.ToString
            End If
        Else
            ReturnArray(counter) = Item.Value
        End If
    Next counter
    getFieldValues = ReturnArray
End Function

'Using the provided object's .Label property, look up the primary key ID value. For cases where
'The name is a unique identifier.
Public Sub GetIdFromName(ByRef oEntity)
    Dim SQL As SQLSelect
    Set SQL = New SQLSelect
    SQL.Table = oEntity.Table
    Dim IdField As DrupalField
    Set IdField = oEntity.IdField
    SQL.Fields = Array(IdField.FieldName)
    SQL.addWhere oEntity.LabelField.FieldName, "=", str(oEntity.Label)
    oEntity.ID = MySQLDatabase.Execute(SQL, IdField.FieldName)
End Sub

'Look up the largest delta value that currently exists in the database for a given entity_id
Private Function MaxDelta(sTable As String, lEntityId As Long) As Long
    Dim SQL As SQLSelect
    Set SQL = New SQLSelect
    SQL.Table = sTable
    SQL.Fields = Array("CASE WHEN max(delta) IS NULL THEN -1 ELSE max(delta) END AS maxdelta")
    SQL.addWhere "entity_id", "=", lEntityId
    MaxDelta = MySQLDatabase.Execute(SQL, "maxdelta")
End Function

'Add a link from an entity to another, through an entity_reference
Public Sub JoinEntities(ByVal oEntity, ByVal oField)
    Dim sBundle As String
    sBundle = oEntity.Table
    
    Dim sField As String
    sField = oField.Table
    
    Dim lFieldId As Long
    lFieldId = oField.ID
    
    Dim lEntityId As Long
    lEntityId = oEntity.ID
    
    Dim table_name As String
    Dim Delta As Integer
    Dim column_name As String
    table_name = sBundle & "__" & sField & "_id"
    Delta = MaxDelta(table_name, lEntityId) + 1
    column_name = sField & "_id_target_id"
    
    Dim inSQL As SQLInsert
    Set inSQL = New SQLInsert
    inSQL.Table = table_name
    inSQL.Fields = Array("bundle", "deleted", "entity_id", "revision_id", "langcode", "delta", column_name)
    inSQL.Values = Array(str(sBundle), 0, lEntityId, lEntityId, str("und"), Delta, lFieldId)
    MySQLDatabase.Execute inSQL
End Sub

'Update the VBA object with values from the database using the id field
Public Sub RefreshObject(ByRef oEntity)
    CreateFromId oEntity, oEntity.ID
End Sub

'Update values in the database using values in the VBA Object
Public Sub Update(oEntity)
    
End Sub

'Create a object with values from the database, using sa primary key
Public Sub CreateFromId(ByRef oEntity, lId As Long)
    Set sSQL = New SQLSelect
    sSQL.Fields = GetFieldNames(oEntity.GetFields)
    sSQL.Table = oEntity.Table
    sSQL.addWhere oEntity.IdField.FieldName, "=", lId
    ArrayOfValues = MyDatabase.Execute(sSQL)
    FieldList = oEntity.GetFields
    numfields = UBound(FieldList)
    Dim counter As Integer
    Dim Item As DrupalField
    For counter = 0 To numfields
        Set Item = FieldList(counter)
        Item.Value = ArrayOfValues(counter)
    Next counter
End Sub
