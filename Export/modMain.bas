Attribute VB_Name = "modMain"
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
    Dim Model As StatesModel
    Set Model = New StatesModel
    Model.Load ThisWorkbook
    
    Dim ViewModel As StateManagerViewModel
    Set ViewModel = New StateManagerViewModel
    ViewModel.Load Model, ListObject
    
    Dim View As IView
    Set View = New ExplorerView
    
    View.ShowDialog ViewModel
End Sub

Public Sub ResetModel()
    Dim Model As StatesModel
    Set Model = New StatesModel
    Model.Load ThisWorkbook
    Model.RemoveAll
    
    Dim State As ISerializable
    Set State = New ColumnsState
    State.Deserialize ("Table1:0.0.0:1.-4152:Q29sQQ==,8,0,1;Q29sQg==,8,0,1;Q29sQw==,8,0,1")
    Model.Add State
    
    Set State = New ColumnsState
    State.Deserialize ("Table1:0.0.0:1.-4152:Q29sRA==,8,0,1;Q29sQg==,16,0,1;Q29sQw==,32,0,1")
    Model.Add State
    
    Set State = New ColumnsState
    State.Deserialize ("Table1:0.0.0:1.-4152:Q29sQQ==,8,0,1;Q29sQg==,0,-1,1;Q29sQw==,3.43,0,1")
    Model.Add State
    
    Set State = New ColumnsState
    State.Deserialize ("Table2:0.0.0:1.-4152:QUFB,8,0,1;QkJC,0,-1,1;Q0ND,3.43,0,1")
    Model.Add State
    
    Set State = New ColumnsState
    State.Deserialize ("Orphan:0.0.0:1.-4152:Q29sQQ==,10,0,1;Q29sQg==,20,0,1;Q29sQw==,30,0,1")
    Model.Add State
    
    Model.Save
End Sub

