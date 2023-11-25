Attribute VB_Name = "modMain"
'@Folder "ColumnState"
Option Explicit

Private Const MSG_NO_TABLE_SELECTED As String = "Select a table before running Persistent Column State Tool."
Private Const MSG_TITLE As String = "Persistent Column State Tool"

Public Sub TestMVVM()
    Dim Model As StatesModel
    Set Model = New StatesModel
    Model.Load ThisWorkbook
    
    Dim ViewModel As StateManagerViewModel
    Set ViewModel = New StateManagerViewModel
    ViewModel.Load Model
    Set ViewModel.Target = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    
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
    Set State = New ColumnsState2
    State.Deserialize ("Table1:Q29sQQ==,8,0;Q29sQg==,8,0;Q29sQw==,8,0")
    Model.Add State
    
    Set State = New ColumnsState2
    State.Deserialize ("Table1:Q29sQQ==,8,0;Q29sQg==,16,0;Q29sQw==,32,0")
    Model.Add State
    
    Set State = New ColumnsState2
    State.Deserialize ("Table1:Q29sQQ==,8,0;Q29sQg==,0,-1;Q29sQw==,3.43,0")
    Model.Add State
    
    Set State = New ColumnsState2
    State.Deserialize ("Table2:Q29sQQ==,8,0;Q29sQg==,0,-1;Q29sQw==,3.43,0")
    Model.Add State
    
    Set State = New ColumnsState2
    State.Deserialize ("Orphan:Q29sQQ==,8,0;Q29sQg==,0,-1;Q29sQw==,3.43,0")
    Model.Add State
    
    Model.Save
End Sub

'@EntryPoint "Open UserForm for PersistentColumnStateTool"
Public Sub PersistentColumnStateTool()
    ' DEBUG
    ThisWorkbook.Worksheets.Item(1).Range("A2").Activate
    
    If Selection.ListObject Is Nothing Then
        MsgBox MSG_NO_TABLE_SELECTED, vbExclamation, MSG_TITLE
        Exit Sub
    End If
    
    Dim ViewModel As ColumnStateViewModel
    Set ViewModel = New ColumnStateViewModel
    ViewModel.Load Selection.ListObject
    
    Dim ViewAsInterface As IView
    Set ViewAsInterface = New frmColumnStateView
    
    ViewAsInterface.ShowDialog ViewModel
End Sub

