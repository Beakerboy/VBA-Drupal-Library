'This Class represents an Entity Reference

'**************************CLASS VARIABLES*************************************
Private sBundle As String       'The table name
Private sField As String
Private lFieldId As Long
Private lEntityId As Long
Private MyCredentials As Credentials

'***************************CONSTRUCTOR****************************************
Private Sub Class_Initialize()
End Sub

'***************************LET PROPERTIES*************************************
Public Property Let Bundle(sValue As String)
    sBundle = sValue
End Property

Public Property Let Field(sValue As String)
    sField = sValue
End Property
Property Let Credentials(oMyCred As Credentials)
    Set MyCredentials = oMyCred
End Property

'***************************GET PROPERTIES*************************************

'******************************METHODS*****************************************
Public Function MaxDelta(lEntityId As Long, sTable As String) As Long
    deltaFieldName = sTable & ".delta"
    Set Sql = New SQLSelect
    Sql.Credentials = MyCredentials
    Sql.Table = sTable
    Sql.Fields = Array("CASE WHEN max(delta) IS NULL THEN -1 ELSE max(delta) END AS maxdelta")
    Sql.addWhere "entity_id", "=", lEntityId
    MaxDelta = Sql.Execute("maxdelta")
End Function

'Attach a entity to another
'A Field with id FieldID is attached to a Bundle with id EntityID
'lFieldId
'lEntityId

Public Sub AttachById(FieldId As Long, EntityId As Long)
    lFieldId = FieldId
    lEntityId = EntityId
    Dim table_name As String
    table_name = sBundle & "__" & sField & "_id"
    Delta = MaxDelta(lEntityId, table_name) + 1
    column_name = sField & "_id_target_id"
    
    Set inSQL = New SQLInsert
    inSQL.Credentials = MyCredentials
    inSQL.Table = table_name
    inSQL.Fields = Array("bundle", "deleted", "entity_id", "revision_id", "langcode", "delta", column_name)
    inSQL.Values = Array(str(sBundle), 0, lEntityId, lEntityId, str("und"), Delta, lFieldId)
    inSQL.Execute
End Sub
