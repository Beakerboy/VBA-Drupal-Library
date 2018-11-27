Attribute VB_Name = "DrupalHelperFunctions"
Public Function Drupal_CountCharsToEscape(sInput As String) As Long
    Dim lResult As Long
    Dim sParts() As String
    
    sParts = Split(sInput, "'")
    Dim ArraySize As Long
    ArraySize = UBound(sParts, 1)
    If ArraySize < 0 Then
        lResult = 0
    Else
        lResult = ArraySize
    End If
    CountCharsToEscape = lResult
End Function
