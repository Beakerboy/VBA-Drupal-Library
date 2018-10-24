'This Class represents the Node Entity

'***********************REQUIRED VARIABLES*************************************
rivate sTable As String        'The table name
Private lId As Long             'Primary Key value

'**************************CLASS VARIABLES*************************************
Private sTitle As DrupalField           'The node title
Private lUid As DrupalField
Private bStatus As DrupalField
Private dCreated as DrupalField
Private cChanged as DrupalField
Private bPromote as DrupalField
Private bSticky as DrupalField
Private cRevisionTimestamp as DrupalField
Private lRevisionUid as DrupalField
Private sRevisionLog as DrupalField
Private bRevisionTranslationAffected as DrupalField

'***************************CONSTRUCTOR****************************************
Private Sub Class_Initialize()
    sTable = "node"
    Set oEntity = New Entity
    
    Set sTitle = New DrupalField
    sTitle.DataType = "string"
    Set lUid As DrupalField
    lUid.Datatype = "integer"
    Set bStatus As DrupalField
    bStatus.Datatype = "boolean"
    bStatus.Value = True
    Set dCreated as DrupalField
    Set cChanged as DrupalField
    Set bPromote as DrupalField
    bPromote.Datatype = "boolean"
    pPromote.Value = True
    Set bSticky as DrupalField
    bSticky.Datatype = "boolean"
    bSticky.Value = False
    Set cRevisionTimestamp as DrupalField
    Set lRevisionUid as DrupalField
    Set sRevisionLog as DrupalField
    Set bRevisionTranslationAffected as DrupalField

End Sub

'***************************LET PROPERTIES*************************************
Public Property Let ID(lValue As Long)
End Property

Public Property Let Title(sValue As String)
    sTitle.Value = sValue
End Property

'***************************GET PROPERTIES*************************************
Public Property Get ID() As Long
    ID = iId
End Property

Public Property Get Table() As String
    Table = sTable
End Property

Public Property Get Title() As String
    Title = sTitle.Value
End Property
                            
'******************************METHODS*****************************************
Public Function GetFields()
    GetFields = Array(sTitle, lUid, bStatus)
End Function
                                                                    
