Attribute VB_Name = "ListObjectHelpers"
'@Folder "MVVM.ColumnState.Helpers"
Option Explicit

'@Description "Returns a Collection containing all the ListObjects in the given Workbook."
Public Function GetAllListObjects(ByVal Workbook As Workbook) As Collection
    Dim Result As Collection
    Set Result = New Collection
    Set GetAllListObjects = Result
    
    If Workbook Is Nothing Then Exit Function
    
    Dim Worksheet As Worksheet
    For Each Worksheet In Workbook.Worksheets
        Dim ListObject As ListObject
        For Each ListObject In Worksheet.ListObjects
            GetAllListObjects.Add Item:=ListObject, Key:=ListObject.Name
        Next ListObject
    Next Worksheet
End Function

'@Description "Returns True if a ListObject with the given name exists in the given Workbook."
Public Function ListObjectExists(ByVal Workbook As Workbook, ByVal ListObjectName As String) As Boolean
    Dim AllListObjects As Collection
    Set AllListObjects = GetAllListObjects(Workbook)
    
    Dim ListObject As ListObject
    For Each ListObject In AllListObjects
        If ListObject.Name = ListObjectName Then
            ListObjectExists = True
            Exit Function
        End If
    Next ListObject
End Function

'@Description "Returns True if the given ListObject contains a ListColumn with the given name."
Public Function HasListColumn(ByVal ListObject As ListObject, ByVal ListColumnName As String) As Boolean
    Dim ListColumn As ListColumn
    For Each ListColumn In ListObject.ListColumns
        If ListColumn.Name = ListColumnName Then
            HasListColumn = True
            Exit Function
        End If
    Next ListColumn
End Function

