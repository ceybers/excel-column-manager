Attribute VB_Name = "SelectedStateToListView"
'@Folder("MVVM.ColumnState.ValueConverters")
Option Explicit

Public Sub Initialize(ByVal ListView As ListView)
    Dim il As ImageList
    Set il = New ImageList
    With il
        .ImageWidth = 16
        .ImageHeight = 16
        .ListImages.Add Key:="Visible", Picture:=Application.CommandBars.GetImageMso("VisibilityVisible", 16, 16)
        .ListImages.Add Key:="Hidden", Picture:=Application.CommandBars.GetImageMso("VisibilityHidden", 16, 16)
        .ListImages.Add Key:="Exists", Picture:=Application.CommandBars.GetImageMso("AcceptInvitation", 16, 16)
        .ListImages.Add Key:="NotExists", Picture:=Application.CommandBars.GetImageMso("DeclineInvitation", 16, 16)
    End With
        
    With ListView
        Set .Icons = il
        Set .SmallIcons = il
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="#", Width:=24
        .ColumnHeaders.Add Text:="Column Name", Width:=70
        .ColumnHeaders.Add Text:="Width", Width:=40
        .ColumnHeaders.Add Text:="Visible", Width:=64
        .ColumnHeaders.Item(3).Alignment = lvwColumnRight
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
        AddItem ListView, Child
    Next Child
End Sub

Private Sub AddItem(ByVal ListView As ListView, ByVal Child As IListable)
    If Child.Visible = False Then Exit Sub
    
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Text:="")
    ListItem.Text = ListItem.Index
    ListItem.SmallIcon = IIf(IsOrphan(Child), "NotExists", "Exists")
    
    With ListItem.ListSubItems
        .Add Text:=Child.Caption

        If IsColumnHidden(Child) Then
            .Add Text:="" 'width
            .Add Text:="Hidden"
        Else
            .Add Text:=ColumnWidth(Child)
            .Add Text:="Visible"
        End If
    
        .Item(3).ReportIcon = .Item(3).Text
    End With
End Sub

Private Function IsColumnHidden(ByVal State As ColumnState2)
    IsColumnHidden = (State.Width = 0)
End Function

Private Function ColumnWidth(ByVal State As ColumnState2)
    ColumnWidth = Format(State.Width, "0.00") & "u"
End Function

Private Function IsOrphan(ByVal State As ColumnState2)
    IsOrphan = State.Orphan
End Function
