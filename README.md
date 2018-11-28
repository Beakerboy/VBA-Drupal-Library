VBA Drupal Library
=====================

### Interact with Drupal Entities in VBA
This Library allows a user to easily move data between excel and a database running the Drupal CMS

Features
--------
 * [DrupalDatabase](#database-class)
 * [DrupalField](#field-class)
 * [DrupalEntity](#entity-class)
 
  Setup
-----

Import the files into a spreadsheet using Microsoft Visual Basic for Applications. These scripts also require the [VBA-SQL-Library](https://github.com/Beakerboy/VBA-SQL-Library) for generalized SQL Query Objects. Also the Microsoft Scripting Runtime project must be enabled.
 
 Security
-----
A user will need database authentication credentials to access the database.

 Usage
-----
 
 ### Database Class
 The DrupalDatabase object is used to perform actions against the databse using Drupal Entities.
 
 ```vba
    Private MyDatabase As DrupalDatabase
    Set MyDatabase = New DrupalDatabase
    MyDatabase.DSN = "foodb"
    MyDatabase.DBType = "mssql"
    'Open UserForm
    Login.Show
    'After Button is pressed assign values
    MyDatabase.Username = Login.Username
    MyDatabase.Password = Login.Password
    Unload Login
```

### Field Class
The DrupaField represents a column in the database. It contains all the meta-information like the data type, length, column name, as well as the value. Inspiration comes from the field definition a module developer uses in a custom Entity file in Drupal.
 * .DataType = __type__
 * .Length = __number__
 * .Value = __value__
 * .FieldName = __name__
 * .IdField = __boolean__
 * .TargetEntity __iDrupalEntity__
 * .Create __type__, __name__, _length_

The currently supported types are boolean, decimal, integer, password, string, and timestamp.
The field class cannot be extended into a custom class at this time.

#### Example
```vba
'For a string with a length of 50:
Set oField = Create_DrupalField
oField.Create "string", "name", 50
```
After the Field is configured, a value can be added with ```oField.Value="Lorum Ipsum```. The value will be validated based on the chosen DataType.

### Entity Class
The DrupalEntity class is a parent class for any other Entities, but can be used as-is. Custom Entities can extend this class and add custom properties and methods.
* .Label = __name__
* .Table = __dbtable__
* .ID = __integer__
* .LabelField = __custom-field__
* .idField = __custom-field__
* .AddField __DrupalField__
* .CreateField __type__, __name__, _length_
* .CreateEntityReference __filedname__, __DrupalEntity__
* .SetValue __field__, __value__
* .GetValue __field__
* .SetTargetValue __field__, __value__
* .GetFields

An example of Extending the Base class is this Drupal User Entity. The id field is named 'uid' and the label field remains as the default 'name'. Two additional fields are added, 'pass' and 'timezone'. We add Properties for the fields, and ensure all the required interface methods are in place. 
```vba
Implements iDrupalEntity

Private oEntity As DrupalEntity

Private Sub Class_Initialize()
    Set oEntity = Create_DrupalEntity
    
    Dim Uid As DrupalField
    Set Uid = Create_DrupalField
    With Uid
        .FieldName = "uid"
        .DataType = "int"
        .IdField = True
    End With
    With oEntity
        .Table = "users"
        Set .IdField = Uid
        .CreateField "password", "pass"
        .CreateField "string", "timezone", 32
    End With
End Sub

Public Property Get Timezone() As String
    Timezone = oEntity.GetValue("timezone")
End Property

Public Property Let Timezone(sValue As String)
    oEntity.SetValue "timezone", sValue
End Property

Public Property Let Password(sValue As String)
    oEntity.SetValue "pass", sValue
End Property

Public Property Let ID(lValue As Long)
    oEntity.ID = lValue
End Property

Public Property Get ID() As Long
    ID = oEntity.ID
End Property

Public Property Let Label(sValue As String)
    oEntity.Label = sValue
End Property

Public Property Get Label() As String
    Label = oEntity.Label
End Property

Public Property Get Table() As String
    Table = oEntity.Table
End Property

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
```
Alternatively, the DrupalEntity class can be used as is. This is sufficient if you do not desire or require custom functions or properties. This Node object has an id field named 'nid', a label field named 'title', a status field, and an entity reference to a user.
```vba
    Dim DrupalNode As DrupalEntity
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
```

