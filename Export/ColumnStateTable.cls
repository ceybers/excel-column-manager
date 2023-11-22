VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnStateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.ColumnState.Entities"
Option Explicit

Private Type TState
    ListObject As ListObject
    ListObjectName As String
    ColumnStates As Collection
End Type

Private This As TState

Public Property Get ColumnStates() As Collection
    Set ColumnStates = This.ColumnStates
End Property

Public Property Get ListObjectName() As String
    ListObjectName = This.ListObjectName
End Property

' Used when reassociating orphaned states before serializing and saving them.
Public Property Let ListObjectName(ByVal vNewValue As String)
    This.ListObjectName = vNewValue
End Property

Public Property Get ListObject() As ListObject
    Set ListObject = This.ListObject
End Property

Public Property Get IsValid() As Boolean
    IsValid = Not This.ColumnStates Is Nothing
End Property

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    This.ListObjectName = ListObject.Name
    LoadColumns
End Sub

Private Sub LoadColumns()
    CollectionHelpers.CollectionClear This.ColumnStates
    Dim ListColumn As ListColumn
    For Each ListColumn In This.ListObject.ListColumns
        LoadColumn ListColumn
    Next ListColumn
End Sub

Private Sub LoadColumn(ByVal ListColumn As ListColumn)
    Dim ColumnStateSingle As ColumnStateSingle
    Set ColumnStateSingle = New ColumnStateSingle
    ColumnStateSingle.Load ListColumn
    This.ColumnStates.Add Item:=ColumnStateSingle
End Sub

Public Function ToString() As String
    If Not IsValid Then Exit Function
    
    Dim Result As String
    Result = This.ListObject.Name & " has " & This.ColumnStates.Count & " column(s)"
    ToString = Result
End Function

Public Sub DebugPrint()
    Dim ColumnStateSingle As ColumnStateSingle
    For Each ColumnStateSingle In This.ColumnStates
        Debug.Print ColumnStateSingle.ToBase64
        'Debug.Print ColumnStateSingle.ToString
    Next ColumnStateSingle
End Sub

Private Sub Class_Initialize()
    Set This.ColumnStates = New Collection
End Sub

Public Function ToBase64() As String
    Dim Result As String
    'Result = Result & This.WorksheetName & ":"
    Result = Result & This.ListObjectName & ":"
    
    Dim ColumnStateSingle As ColumnStateSingle
    For Each ColumnStateSingle In This.ColumnStates
        Result = Result & ColumnStateSingle.ToBase64() & ";"
    Next ColumnStateSingle
    
    Result = Left$(Result, Len(Result) - 1)
    
    ToBase64 = Result
End Function

Public Sub LoadFromBase64(ByVal Base64Value As String, ByVal ListObject As ListObject)
    CollectionHelpers.CollectionClear This.ColumnStates
    
    This.ListObjectName = Split(Base64Value, ":")(0)
    
    Dim ColumnStateSingles() As String
    ColumnStateSingles = Split(Split(Base64Value, ":")(1), ";")
    
    Dim i As Long
    For i = 0 To UBound(ColumnStateSingles)
        Dim ColumnStateSingle As ColumnStateSingle
        Set ColumnStateSingle = New ColumnStateSingle
        ColumnStateSingle.LoadFromBase64 ColumnStateSingles(i)
        ColumnStateSingle.Index = (i + 1)
        ColumnStateSingle.UpdateExists ListObject
        This.ColumnStates.Add Item:=ColumnStateSingle
    Next i
End Sub

Public Function GetCaption() As String
    Dim Result As String
    
    Dim TotalColumns As Long
    TotalColumns = This.ColumnStates.Count
    
    Dim VisibleColumns As Long
    Dim ColumnStateSingle As ColumnStateSingle
    For Each ColumnStateSingle In This.ColumnStates
        If Not ColumnStateSingle.Hidden Then
            VisibleColumns = VisibleColumns + 1
        End If
    Next ColumnStateSingle
    
    Result = VisibleColumns & " of " & TotalColumns & " column(s) visible"
    GetCaption = Result
End Function

Public Function CanApply(ByVal ListObject As ListObject)
    Dim ColumnStateSingle As ColumnStateSingle
    
    For Each ColumnStateSingle In This.ColumnStates
        If ListObjectHelpers.HasListColumn(ListObject, ColumnStateSingle.Name) Then
            CanApply = True
            Exit Function
        End If
    Next ColumnStateSingle
End Function

Public Function IsPartialMatch(ByVal ListObject As ListObject) As Boolean
    Dim ColumnState As ColumnStateSingle
    
    Dim n As Long
    
    For Each ColumnState In This.ColumnStates
        If ListObjectHelpers.HasListColumn(ListObject, ColumnState.Name) Then
            n = n + 1
        End If
    Next ColumnState
    
    IsPartialMatch = (n > 0) And (n < This.ColumnStates.Count)
End Function

Public Function Apply(ByVal ListObject As ListObject) As Boolean
    Dim ColumnStateSingle As ColumnStateSingle
    For Each ColumnStateSingle In This.ColumnStates
        ApplyState ColumnStateSingle, ListObject
    Next ColumnStateSingle
    
    Apply = True
End Function

Private Sub ApplyState(ByVal ColumnStateSingle As ColumnStateSingle, ListObject As ListObject)
    If ColumnStateSingle.Exists = False Then Exit Sub
    
    Dim ListColumn As ListColumn
    Set ListColumn = ListObject.ListColumns(ColumnStateSingle.Name)
    
    With ListColumn.DataBodyRange
        .EntireColumn.Hidden = ColumnStateSingle.Hidden
        If ColumnStateSingle.Width = -1 Then
            .EntireColumn.AutoFit
        Else
            .ColumnWidth = ColumnStateSingle.Width
        End If
    End With
End Sub

Public Function Equals(ByVal RHS As ColumnStateTable)
    Debug.Assert Not RHS Is Nothing
    Equals = (Me.ToBase64 = RHS.ToBase64)
End Function
