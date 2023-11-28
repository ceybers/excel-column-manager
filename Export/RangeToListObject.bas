Attribute VB_Name = "RangeToListObject"
'@Folder "Helpers"
Option Explicit

'@Description "Returns the name of a ListColumn if a Range intersects with a given ListObject."
Public Function TryRangeToListHeader(ByVal ListObject As ListObject, ByVal Range As Range, _
                                     ByRef OutHeader As String) As Boolean
Attribute TryRangeToListHeader.VB_Description = "Returns the name of a ListColumn if a Range intersects with a given ListObject."
    Dim Result As String
    Result = RangeToListHeader(ListObject, Range)
    If Result <> Empty Then
        OutHeader = Result
        TryRangeToListHeader = True
    End If
End Function

Private Function RangeToListHeader(ByVal ListObject As ListObject, ByVal Range As Range) As String
    ' If it fails, will evaluate = Empty
    ' Technically we could take first cell in the Range and offset it to 1 row above
    Debug.Assert Not ListObject Is Nothing
    Debug.Assert Not Range Is Nothing
    Dim Result As String
    
    Dim Intersection As Range
    Set Intersection = Application.Intersect(ListObject.HeaderRowRange, Range.EntireColumn)
    If Not Intersection Is Nothing Then
        If Intersection.Cells.Count = 1 Then
            Result = Intersection.Value2
        End If
    End If
    
    RangeToListHeader = Result
End Function
