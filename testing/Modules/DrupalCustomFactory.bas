Attribute VB_Name = "DrupalCustomFactory"
Public Function Create_DrupalNode()
    Dim MyObject As DrupalEntity
    Dim Nid As DrupalField
    Dim Title As DrupalField
    Dim UserEntity As New DrupalUser
    
    Set MyObject = Create_DrupalEntity
    Set Nid = Create_DrupalField
    Set Title = Create_DrupalField
    Set UserEntity = New DrupalUser
    
    With Nid
        .FieldName = "nid"
        .IdField = True
        .DataType = "int"
    End With
    
    With Title
        .FieldName = "title"
        .DataType = "string"
        .Length = 255
    End With
    
    With MyObject
        .Table = "node"
        Set .IdField = Nid
        Set .LabelField = Title
        .CreateEntityReference "uid", UserEntity
        .CreateField "boolean", "status"
    End With
    Set Create_DrupalNode = MyObject
End Function

