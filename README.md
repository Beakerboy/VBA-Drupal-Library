VBA Drupal Library
=====================

### Interact with Drupal Entities in VBA
This Library allows a user to easily move data between excel and a database running the Drupal CMS

Features
--------
 * [DrupalDatabase](#database-class)
 * [DrupalField](#field-class)
 * [DrupalEntity](#entity-class)
 * [DrupalNode](#node-class)
 
  Setup
-----

Import the files into a spreadsheet using Microsoft Visual Basic for Applications. These scripts also require the [VBA-SQL-Library](https://github.com/Beakerboy/VBA-SQL-Library) for generalized SQL Query Objects.
 
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
```vba
'For a string with a length of 50:
Set oField = New DrupalField
oField.DataType = "string"
oField.FieldName = "name"
oField.Length = 50
```

### Entity Class
The DrupalEntity class is a parent class for any other Entities. Your Entities will extend this class to add additional fields.

### Node Class
The library contains a representation of Drupal's Node Entity as an example of how to extend the DrupalEntity

To add a new node to your database:
```vba
'Create and populate a node object
Set MyNode = New DrupalNode
MyNode.Label = "New Node"
MyNode.Status = False

'Look up the username for the user "mydrupalusername"
Set MyUser = New DrupalUser
MyUser.Name = "mydrupalusername"
MyDatabase.getIdFromName MyUser

MyNode.Uid = MyUser.Id
MyDatabase.Insert MyNode         'Push the data to the database
```

The library also is able to attach entities together if their entity_reference is confugured to allow multiple values

```vba
MyDatabase.JoinEntities MyNode SomeCustomEntity    ''creates an entry in the node__somecustomentity_id table
```
