VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrupalDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Class: DrupalDatabase
'Provides methods for manupulating Drupal Entities with VBA
Option Explicit
'*********************************VARIABLES************************************
Private MySQLDatabase As SQLDatabase

'***************************CONSTRUCTOR****************************************
Private Sub Class_Initialize()
    Set MySQLDatabase = Create_SQLDatabase
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

Public Property Set Recordset(MyRecordset As iSQLRecordset)
    Set MySQLDatabase.Recordset = MyRecordset
End Property

Public Property Set Connection(MyConnection As iSQLConnection)
    Set MySQLDatabase.Connection = MyConnection
End Property

'******************************METHODS*****************************************
' Sub: Insert
' Create a new record in the database using values from the class object
Public Sub Insert(ByRef oEntity As iDrupalEntity)
    Dim inSQL As SQLInsert
    Set inSQL = Create_SQLInsert
    Dim aFields As Variant
    aFields = oEntity.GetFields()
    With inSQL
        .Table = oEntity.Table
        .Fields = GetFieldNames(aFields)
        .Values = getFieldValues(aFields)
    End With
    oEntity.ID = MySQLDatabase.InsertGetNewId(inSQL, oEntity.IdField.FieldName)
End Sub

' Function: GetFieldNames
' Create an array of field names from the array of DrupalField objects
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

' Function: GetFieldValues
' Create an array of field values from the array of DrupalField objects.
' Sanitize the data to prepare it for the database.
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
        Dim DataType As String
        DataType = Item.DataType
        If DataType = "string" Or DataType = "password" Then
            ReturnArray(counter) = Str(Item.Value)
        ElseIf DataType = "boolean" Then
            Dim FieldValue As String
            If Item.Value Then
                FieldValue = "TRUE"
            Else
                FieldValue = "FALSE"
            End If
            ReturnArray(counter) = FieldValue
        ElseIf Item.DataType = "entity_reference" Then
            If Item.Value > 0 Then
                'id is set, use it instead of the value
                'Should we use Item.ForeignEntity.Id instead of Item.Value?
                ReturnArray(counter) = Item.Value
            Else
                'Create a simple subselect
                Dim SeSQL As SQLSelect
                Set SeSQL = Create_SQLSelect
                With SeSQL
                    .Fields = Array(Item.TargetEntity.IdField.FieldName)
                    .Table = Item.TargetEntity.Table
                    .AddWhere Item.TargetEntity.LabelField.FieldName, Str(Item.TargetEntity.Label)
                End With

                Dim Interfaced As iSQLQuery
                Set Interfaced = SeSQL
                ReturnArray(counter) = "(" & Interfaced.toString & ")"
            End If
        Else
            ReturnArray(counter) = Item.Value
        End If
    Next counter
    getFieldValues = ReturnArray
End Function

' Sub: GetIdFromName
' Using the provided object's .Label property, look up the primary key ID value. For cases where
' the name is a unique identifier.
' Should this be renamed to SetIdFromName? We are not returning anything.
Public Sub GetIdFromName(ByRef oEntity As iDrupalEntity)
    Dim SQL As SQLSelect
    Set SQL = Create_SQLSelect
    Dim IdField As DrupalField
    Set IdField = oEntity.IdField
    With SQL
        .Table = oEntity.Table
        .Fields = Array(IdField.FieldName)
        .AddWhere oEntity.LabelField.FieldName, Str(oEntity.Label)
    End With
    oEntity.ID = MySQLDatabase.Execute(SQL, IdField.FieldName)
End Sub

' Sub: JoinEntities
' Add a link from an entity to another, through an entity_reference
Public Sub JoinEntities(ByVal oEntity As iDrupalEntity, ByVal oField As iDrupalEntity)
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
    Set inSQL = Create_SQLInsert
    With inSQL
        .Table = table_name
        .Fields = Array("bundle", "deleted", "entity_id", "revision_id", "langcode", "delta", column_name)
        .Values = Array(Str(sBundle), 0, lEntityId, lEntityId, Str("und"), Delta, lFieldId)
    End With
    MySQLDatabase.Execute inSQL
End Sub

' Sub: RefreshObject
' Update the VBA object with values from the database using the id field
Public Sub RefreshObject(ByRef oEntity As iDrupalEntity)
    CreateFromId oEntity, oEntity.ID
End Sub

' Sub: Update
' Update values in the database using values in the VBA Object
Public Sub Update(oEntity As iDrupalEntity)
    Dim MyUpdate As SQLUpdate
    Set MyUpdate = Create_SQLUpdate
    Dim aFields As Variant
    aFields = oEntity.GetFields()
    With MyUpdate
        .Table = oEntity.Table
        .Fields = GetFieldNames(aFields)
        .Values = getFieldValues(aFields)
        .AddWhere oEntity.IdField.FieldName, oEntity.ID
    End With
    MySQLDatabase.Execute MyUpdate
End Sub

'Sub: CreateFromId
' Create an object with values from the database, using a primary key
Public Sub CreateFromId(ByRef oEntity As iDrupalEntity, lId As Long)
    Dim sSql As SQLSelect
    Dim ArrayOfValues As Variant
    Dim FieldList As Variant
    Dim numfields As Integer
    Set sSQL = Create_SQLSelect
    sSQL.Fields = GetFieldNames(oEntity.GetFields)
    sSQL.Table = oEntity.Table
    sSQL.AddWhere oEntity.IdField.FieldName, lId
    ArrayOfValues = MyDatabase.Execute(sSQL)
    FieldList = oEntity.GetFields
    numfields = UBound(FieldList)
    Dim counter As Integer
    Dim Item As DrupalField
    For counter = 0 To numfields
        Set Item = FieldList(counter)
        If IsNull(ArrayOfValues(counter, 0)) Then
            ArrayOfValues(counter, 0) = ""
        End If
        Item.Value = ArrayOfValues(counter, 0)
    Next counter
End Sub

'**********************PRIVATE METHODS*****************************************
' Function: MaxDelta
' Look up the largest delta value that currently exists in the database for a given entity_id
Private Function MaxDelta(sTable As String, lEntityId As Long) As Long
    Dim SQL As SQLSelect
    Set SQL = Create_SQLSelect
    With SQL
        .Table = sTable
        .AddField "CASE WHEN max(delta) IS NULL THEN -1 ELSE max(delta) END", "maxdelta"
        .AddWhere "entity_id", lEntityId
    End With
    MaxDelta = MySQLDatabase.Execute(SQL, "maxdelta")
End Function
