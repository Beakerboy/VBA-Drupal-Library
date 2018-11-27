Attribute VB_Name = "DrupalUnitTests"
Public Function Drupal_RunTests()
    '*******************************Drupal Field*******************************
    Dim MyField As DrupalField
    Set MyField = Create_DrupalField
    With MyField
        'Configure the fields
        .DataType = "string"
        .Length = 10
        .FieldName = "favorite_color"
        .IdField = True
        
        'Lastly, set the value to ensure data validation occurs
        .Value = "redredredred"
    End With
    CheckValue MyField.DataType, "string"
    CheckValue MyField.Length, 10
    CheckValue MyField.IdField, True
    CheckValue MyField.FieldName, "favorite_color"
    CheckValue MyField.Value, "redredredr"
    
    Dim Field2 As DrupalField
    Set Field2 = Create_DrupalField
    Field2.QuickLoad "decimal", "quantity"
    Field2.Value = 19.2
    CheckValue Field2.DataType, "decimal"
    CheckValue Field2.Length, 50
    CheckValue Field2.IdField, False
    CheckValue Field2.FieldName, "quantity"
    CheckValue Field2.Value, 19.2
    
    Field2.Value = "19.2"
    CheckValue Field2.Value, 19.2
    
    On Error GoTo Err
        Field2.Value = "text"
        CheckError Field2.Value
Err:
    
    
    '*******************************Drupal Entity*******************************
    'Test a Custom DrupalEntity Child
    Dim MyUserEntity As DrupalUser
    Set MyUserEntity = New DrupalUser
    With MyUserEntity
        'Let Properties
        .Label = "myusername"
        .ID = 2
        .Timezone = "UTC"
        .Password = "secr3t"
    End With
        
    CheckValue MyUserEntity.ID, 2
    CheckValue MyUserEntity.Label, "myusername"
    CheckValue MyUserEntity.Table, "users"
    CheckValue MyUserEntity.Timezone, "UTC"
    
    'Check using this custom class in an entity reference
    'And test use of just the base class
    Dim MediaID As DrupalField
    Set MediaID = Create_DrupalField
    MediaID.QuickLoad "integer", "mid"
    MediaID.IdField = True
    
    Dim Filename As DrupalField
    Set Filename = Create_DrupalField
    Filename.QuickLoad "string", "filename"
    
    Dim MyMediaEntity As DrupalEntity
    Set MyMediaEntity = Create_DrupalEntity
    With MyMediaEntity
        'Set Properties
        .Table = "media"
        Set .IdField = MediaID
        Set .LabelField = Filename
        .Label = "foo.jpg"
        .CreateEntityReference "uid", MyUserEntity
        .SetTargetValue "uid", "user2"
    End With
    CheckValue MyMediaEntity.IdField.FieldName, "mid"
    CheckValue MyMediaEntity.LabelField.FieldName, "filename"
      
    '******************************Drupal Database*****************************
    Dim MyDatabase As DrupalDatabase
    Set MyDatabase = Create_DrupalDatabase
    Dim TestConnection As New DrupalTestConnection
    Dim TestRecordset As New DrupalTestRecordset
    TestRecordset.ReturnValue = 1
    With MyDatabase
        .DSN = "mydsn"
        .DBType = "psql"
        .Password = "Pa$$word"
        .Username = "user"
        Set .Connection = TestConnection
        Set .Recordset = TestRecordset
        .Insert MyMediaEntity
    End With
    Dim Expected As String
    
    TestRecordset.ReturnValue = 42
    MyDatabase.Insert MyUserEntity
    CheckValue TestRecordset.Query, "INSERT INTO users (name, pass, timezone) VALUES ('user2', 'secr3t', 'UTC') RETURNING uid"
    
    'Check that an insert query is formatted correctly when an entity reference is present.
    Expected = "INSERT INTO media (filename, uid) VALUES ('foo.jpg', (SELECT uid FROM users WHERE name='user2')) RETURNING mid"
    CheckValue TestRecordset.Query, Expected
    'Check that the returned value is in the data object
    CheckValue MyUserEntity.iDrupalEntity_ID, 42
    
    TestRecordset.ClearReturnValues
    TestRecordset.ClearQueries
    TestRecordset.ReturnValue = 2
    MyDatabase.GetIdFromName MyUserEntity
    'Check the format of the select query
    CheckValue TestRecordset.Query, "SELECT uid FROM users WHERE name='user2'"
    'Check that the new returned value is in the data object
    CheckValue MyUserEntity.iDrupalEntity_ID, 2
    
    'Check that an Entity reference with its ID will insert without the subselect
    TestRecordset.ReturnValue = 17
    MyMediaEntity.SetValue "uid", MyUserEntity.iDrupalEntity_ID
    MyDatabase.Insert MyMediaEntity
    
    Expected = "INSERT INTO media (filename, uid) VALUES ('foo.jpg', 2) RETURNING mid"
    CheckValue TestRecordset.Query, Expected
    
    Dim MyNode As DrupalEntity
    Set MyNode = Create_DrupalNode
    With MyNode
        .Label = "User2 Homepage"
        .SetValue "status", "oN"
        .SetValue "uid", 91
    End With
    
    TestRecordset.ClearReturnValues
    TestRecordset.ClearQueries
    TestRecordset.ReturnValue = 432
    MyDatabase.Insert MyNode
    Expected = "INSERT INTO node (title, uid, status) VALUES ('User2 Homepage', 91, TRUE) RETURNING nid"
    CheckValue TestRecordset.Query, Expected
    TestRecordset.ReturnValue = -1
    MyDatabase.JoinEntities MyNode, MyMediaEntity
    Expected = "INSERT INTO node__media_id (bundle, deleted, entity_id, revision_id, langcode, delta, media_id_target_id) VALUES ('node', 0, 432, 432, 'und', 0, 17)"
    CheckValue TestRecordset.Query, Expected
    Expected = "SELECT CASE WHEN max(delta) IS NULL THEN -1 ELSE max(delta) END AS maxdelta FROM node__media_id WHERE entity_id=432"
    CheckValue TestRecordset.Query, Expected
End Function

Function CheckValue(MyTest, ExpectedValue)
    If MyTest <> ExpectedValue Then
        MsgBox "Expected: " & ExpectedValue & vbNewLine & "Provided: " & MyTest
    End If
End Function

Function CheckError(MyTest)
    If MyTest <> ExpectedValue Then
        MsgBox "Expected: " & "ERROR" & vbNewLine & "Provided: " & MyTest
    End If
End Function
