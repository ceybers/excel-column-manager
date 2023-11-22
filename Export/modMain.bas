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
    
    Dim ViewModel As ColumnStateViewModel
    Set ViewModel = New ColumnStateViewModel
    ViewModel.Load Selection.ListObject
    
    Dim ViewAsInterface As IView
    Set ViewAsInterface = New frmColumnStateView
    
    ViewAsInterface.ShowDialog ViewModel
End Sub

