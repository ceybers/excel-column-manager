Attribute VB_Name = "SelectedStateToListView"
'@Folder("MVVM.ColumnState.ValueConverters")
Option Explicit

Public Sub Initialize(ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="#", Width:=24
        .ColumnHeaders.Add Text:="Column Name", Width:=80
        .ColumnHeaders.Add Text:="Width", Width:=40
        .ColumnHeaders.Add Text:="Visible", Width:=40
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
    
    Dim i As Long
    
    Dim Child As IListable
    For Each Child In ViewModel.Items
        i = i + 1
        
        Dim ListItem As ListItem
        If Child.Visible = True Then
            Set ListItem = ListView.ListItems.Add(Text:=CStr(i))
        End If
        
        Dim State As ColumnState2
        Set State = Child
        
        With ListItem.ListSubItems
            .Add Text:=Child.Caption
            .Add Text:=CStr(State.Width)
            .Add Text:=IIf(State.Width = 0, "Hidden", "Visible")
        End With
    Next Child
End Sub

