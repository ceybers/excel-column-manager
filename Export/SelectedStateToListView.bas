Attribute VB_Name = "SelectedStateToListView"
'@Folder("MVVM.ColumnState.ValueConverters")
Option Explicit

Public Sub Initialize(ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Item Name", Width:=ListView.Width - 8
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
    End With
End Sub

Public Sub Load(ByVal ListView As ListView, ByVal ViewModel As SelectedStateViewModel)
    ListView.ListItems.Clear
    
    If ViewModel Is Nothing Then Exit Sub
    If ViewModel.State Is Nothing Then Exit Sub
    
    Dim Child As IListable
    For Each Child In ViewModel.Items
        If Child.Visible = True Then
            ListView.ListItems.Add Text:=Child.Caption
        End If
    Next Child
End Sub

