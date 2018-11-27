VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrupalTestRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iSQLRecordset

Private aQuery() As String
Private aReturnValue() As Variant
Private iQueryCount As Integer
Private iReturnCount As Integer

Private Sub Class_Initialize()
    ClearReturnValues
    ClearQueries
End Sub

Public Property Get Query()
    Query = aQuery(iQueryCount - 1)
    iQueryCount = iQueryCount - 1
End Property

Public Property Let ReturnValue(vValue)
    aReturnValue(iReturnCount) = vValue
    iReturnCount = iReturnCount + 1
End Property

Public Sub ClearReturnValues()
    ReDim aReturnValue(5)
    iReturnCount = 0
    
End Sub

Public Sub ClearQueries()
    ReDim aQuery(5)
    iQueryCount = 0
End Sub

Public Function iSQLRecordset_State()

End Function

Public Sub iSQLRecordset_OpenRecordset(MyQuery As String, Connection)
    aQuery(iQueryCount) = MyQuery
    iQueryCount = iQueryCount + 1
End Sub

Public Sub iSQLRecordset_CloseRecordset()

End Sub

Public Function iSQLRecordset_GetValue(MyFieldname)
    iSQLRecordset_GetValue = aReturnValue(iReturnCount - 1)
    iReturnCount = iReturnCount - 1
End Function

