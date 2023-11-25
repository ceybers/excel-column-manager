Attribute VB_Name = "modMain"
'@Folder "ColumnState"
Option Explicit

Private Const MSG_NO_TABLE_SELECTED As String = "Select a table before running Persistent Column State Tool."
Private Const MSG_TITLE As String = "Persistent Column State Tool"

'@EntryPoint "Open UserForm for PersistentColumnStateTool"
Public Sub PersistentColumnStateTool()
    ' DEBUG
    ThisWorkbook.Worksheets.Item(1).Range("A2").Activate
    
    If Selection.ListObject Is Nothing Then
        MsgBox MSG_NO_TABLE_SELECTED, vbExclamation, MSG_TITLE
        Exit Sub
    End If

    Dim Model As StatesModel
    Set Model = New StatesModel
    Model.Load ThisWorkbook
    
    Dim ViewModel As StateManagerViewModel
    Set ViewModel = New StateManagerViewModel
    ViewModel.Load Model, Selection.ListObject
    
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
    State.Deserialize ("Table1:Q29sQQ==,8,0;Q29sQg==,8,0;Q29sQw==,8,0")
    Model.Add State
    
    Set State = New ColumnsState
    State.Deserialize ("Table1:Q29sRA==,8,0;Q29sQg==,16,0;Q29sQw==,32,0")
    Model.Add State
    
    Set State = New ColumnsState
    State.Deserialize ("Table1:Q29sQQ==,8,0;Q29sQg==,0,-1;Q29sQw==,3.43,0")
    Model.Add State
    
    Set State = New ColumnsState
    State.Deserialize ("Table2:QUFB,8,0;QkJC,0,-1;Q0ND,3.43,0")
    Model.Add State
    
    Set State = New ColumnsState
    State.Deserialize ("Orphan:Q29sQQ==,10,0;Q29sQg==,20,0;Q29sQw==,30,0")
    Model.Add State
    
    Model.Save
End Sub

