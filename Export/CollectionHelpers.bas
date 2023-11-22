Attribute VB_Name = "CollectionHelpers"
'@Folder "MVVM.ColumnState.Helpers"
Option Explicit

Public Function ConcatCollection(ByVal Collection As Collection, ByVal Delimiter As String) As String
    Dim Result As String
    
    Dim Item As Variant
    For Each Item In Collection
        Result = Result & Item & Delimiter
    Next Item
    
    If Len(Result) > 1 Then
        Result = Left$(Result, Len(Result) - 1)
    End If
    
    ConcatCollection = Result
End Function

Public Function ExistsInCollection(ByVal Collection As Object, ByVal Value As Variant) As Boolean
    Debug.Assert Not Collection Is Nothing
    
    Dim ThisValue As Variant
    For Each ThisValue In Collection
        'If ThisValue = Value Then
        If CStr(ThisValue) = CStr(Value) Then
            'If StrComp(ThisValue, Value) Then ' Run-time error '458' Variable uses an Automation Type supported in Visual Basic
            ExistsInCollection = True
            Exit Function
        End If
    Next ThisValue
End Function

Public Sub CollectionClear(ByVal Collection As Collection)
    Debug.Assert Not Collection Is Nothing
    
    Dim i As Long
    For i = Collection.Count To 1 Step -1
        Collection.Remove i
    Next i
End Sub

