Attribute VB_Name = "modMain"
'@IgnoreModule EmptyIfBlock
'@Folder "ColumnState"
Option Explicit

Private Const MSG_NO_TABLE_SELECTED As String = "Select a table before running Persistent Column State Tool."
Private Const MSG_TITLE As String = "Persistent Column State Tool"

Private Const DEBUG_MODE As Boolean = True

'@EntryPoint "Open UserForm for PersistentColumnStateTool"
Public Sub PersistentColumnStateTool()
    Dim Target As ListObject
    
    If TryGetDebugListObject(Target) Then
    ElseIf TryGetSelectedListObject(Target) Then
    ElseIf TryGetSingleListObject(Target) Then
    Else
        MsgBox MSG_NO_TABLE_SELECTED, vbExclamation, MSG_TITLE
        Exit Sub
    End If
    
    RunPersistentColumnStateTool Selection.ListObject
End Sub

Private Function TryGetDebugListObject(ByRef OutListObject As ListObject) As Boolean
    If DEBUG_MODE Then
        ThisWorkbook.Worksheets.Item(1).Range("A2").Activate
        Set OutListObject = Selection.ListObject
        TryGetDebugListObject = True
    End If
End Function

Private Function TryGetSelectedListObject(ByRef OutListObject As ListObject) As Boolean
    If Not Selection.ListObject Is Nothing Then
        Set OutListObject = Selection.ListObject
        TryGetSelectedListObject = True
    End If
End Function

Private Function TryGetSingleListObject(ByRef OutListObject As ListObject) As Boolean
    If Selection.Parent.ListObjects.Count = 1 Then
        Set OutListObject = Selection.Parent.ListObjects.Item(1)
        TryGetSingleListObject = True
    End If
End Function

Private Sub RunPersistentColumnStateTool(ByVal ListObject As ListObject)
    Application.ScreenUpdating = False
    
    Dim Model As StatesModel
    Set Model = New StatesModel
    Model.Load ThisWorkbook
    
    Dim ViewModel As StateManagerViewModel
    Set ViewModel = New StateManagerViewModel
    ViewModel.Load Model, ListObject
    
    Dim View As IView
    Set View = New ExplorerView
    
    View.ShowDialog ViewModel
    
    Application.ScreenUpdating = True
End Sub
