Attribute VB_Name = "ListColumnHelpers"
'@Folder "Helpers"
Option Explicit

Public Function TryGetListColumn(ByVal ListObject As ListObject, ByVal ListColumnName As String, _
                                 ByRef OutListColumn As ListColumn) As Boolean
    Dim ListColumn As ListColumn
    For Each ListColumn In ListObject.ListColumns
        If ListColumn.Name = ListColumnName Then
            Set OutListColumn = ListColumn
            TryGetListColumn = True
            Exit Function
        End If
    Next ListColumn
End Function

Public Function Exists(ByVal ListObject As ListObject, ByVal ListColumnName As String) As Boolean
    Exists = TryGetListColumn(ListObject, ListColumnName, Nothing)
End Function


