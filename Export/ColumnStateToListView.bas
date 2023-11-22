Attribute VB_Name = "ColumnStateToListView"
'@Folder("MVVM.ColumnState.ValueConverters")
Option Explicit

Private Const MSO_COLUMN_EXISTS As String = "AcceptInvitation"
Private Const MSO_COLUMN_NOT_EXISTS As String = "DeclineInvitation"
Private Const MSO_VISIBLE_TRUE As String = "VisibilityVisible"
Private Const MSO_VISIBLE_FALSE As String = "VisibilityHidden"

Public Sub InitializeListView(ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add text:="#", Width:=24
        .ColumnHeaders.Add text:="Column Name", Width:=80
        .ColumnHeaders.Add text:="Width", Width:=40
        .ColumnHeaders.Add text:="Visible", Width:=64
        .Appearance = cc3D
        .BorderStyle = ccNone
        .Gridlines = True
        .View = lvwReport
        .FullRowSelect = True
        .HotTracking = False
        .LabelEdit = lvwManual
        Set .SmallIcons = ImageListHelpers.GetImageList
    End With
End Sub

Public Sub Load(ByVal ViewModel As ColumnStateViewModel, ByVal ListView As ListView)
    ListView.ListItems.Clear
    If ViewModel.SelectedState Is Nothing Then Exit Sub
    
    Dim FilterUnmatched As Boolean
    FilterUnmatched = ViewModel.DoFilterUnmatched
    
    Dim ColumnState As ColumnStateSingle
    For Each ColumnState In ViewModel.SelectedState.ColumnStates
        LoadColumnStateToListView ListView, ColumnState, FilterUnmatched
    Next ColumnState
End Sub

Private Sub LoadColumnStateToListView(ByVal ListView As ListView, ByVal ColumnState As ColumnStateSingle, ByVal FilterUnmatched As Boolean)
    If FilterUnmatched Then
        If ColumnState.Exists = False Then
            Exit Sub
        End If
    End If
        
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(text:=CStr(ColumnState.Index), SmallIcon:=MSO_COLUMN_EXISTS)
    
    ListItem.ListSubItems.Add text:=ColumnState.Name
    ListItem.ListSubItems.Add text:=CStr(ColumnState.Width)
    
    If ColumnState.Hidden = True Then
        ListItem.ListSubItems.Add text:="Hidden", ReportIcon:=MSO_VISIBLE_FALSE
    Else
        ListItem.ListSubItems.Add text:="Visible", ReportIcon:=MSO_VISIBLE_TRUE
    End If
    
    If Not ColumnState.Exists Then
        ListItem.SmallIcon = MSO_COLUMN_NOT_EXISTS
    End If
End Sub


