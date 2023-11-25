Attribute VB_Name = "SelectedStateToListView"
'@Folder("MVVM.ColumnState.ValueConverters")
Option Explicit

Private Const WIDTH_TO_AVOID_SCROLLBAR As Long = 8

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
    
    FillColumnHeaderWidth ListView, 2
End Sub

Private Sub FillColumnHeaderWidth(ByVal ListView As ListView, ByVal ColumnIndex As Long)
    Dim TotalColumnWidth As Long
    Dim ColumnHeader As ColumnHeader
    For Each ColumnHeader In ListView.ColumnHeaders
        TotalColumnWidth = TotalColumnWidth + ColumnHeader.Width
    Next ColumnHeader
    
    Dim RemainingWidth As Long
    RemainingWidth = ListView.Width - TotalColumnWidth
    
    Dim TargetColumnHeader As ColumnHeader
    Set TargetColumnHeader = ListView.ColumnHeaders.Item(ColumnIndex)
    
    TargetColumnHeader.Width = TargetColumnHeader.Width + RemainingWidth - WIDTH_TO_AVOID_SCROLLBAR
End Sub

Public Sub Load(ByVal ListView As ListView, ByVal ViewModel As SelectedStateViewModel)
    ListView.ListItems.Clear
    
    If ViewModel Is Nothing Then Exit Sub
    If ViewModel.State Is Nothing Then Exit Sub
    
    Dim ShowNonMatchingCols As Boolean
    ShowNonMatchingCols = ViewModel.ShowNonMatchingCols
    
    Dim Child As IListable
    For Each Child In ViewModel.Items
        AddItem ListView, Child, ShowNonMatchingCols
    Next Child
    
    Dim BuiltinState As IListable
    Set BuiltinState = ViewModel.State
    If BuiltinState.ParentKey = modConstants.BUILTIN_KEY Then
        AddBuiltinItem ListView, ViewModel.State
    End If
End Sub

Private Sub AddItem(ByVal ListView As ListView, ByVal Child As IListable, ByVal ShowNonMatchingCols As Boolean)
    If Child.Visible = False Then Exit Sub
    
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Text:=vbNullString)
    ListItem.Text = ListItem.Index
    ListItem.SmallIcon = IIf(IsOrphan(Child), "NotExists", "Exists")
    
    With ListItem.ListSubItems
        .Add Text:=Child.Caption

        If IsColumnHidden(Child) Then
            .Add Text:=ColumnWidth(Child)
            .Add Text:="Hidden"
        Else
            .Add Text:=ColumnWidth(Child)
            .Add Text:="Visible"
        End If
    
        .Item(3).ReportIcon = .Item(3).Text
    End With
    
    If Not ShowNonMatchingCols Then
        If IsOrphan(Child) Then
            ListView.ListItems.Remove ListItem.Index
        End If
    End If
End Sub

Private Function IsColumnHidden(ByVal State As ColumnState) As Boolean
    IsColumnHidden = State.Hidden
End Function

Private Function ColumnWidth(ByVal State As ColumnState) As String
    ColumnWidth = Format$(State.Width, "0.00") & "u"
End Function

Private Function IsOrphan(ByVal State As ColumnState) As Boolean
    IsOrphan = State.Orphan
End Function

Private Sub AddBuiltinItem(ByVal ListView As ListView, ByVal State As IListable)
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Text:=vbNullString)
    ListItem.ListSubItems.Add Text:=State.Caption
End Sub

