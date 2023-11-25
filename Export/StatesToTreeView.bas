Attribute VB_Name = "StatesToTreeView"
'@Folder("MVVM.ColumnState.ValueConverters")
Option Explicit

Private Const ROOT_CAPTION As String = "Column States"
Private Const CURRENT_SUFFIX_CAPTION As String = " (current)"
Private Const ORPHANS_CAPTION As String = "Orphans"

Private Const ROOT_KEY As String = "::ROOT"
Private Const ORPHAN_KEY As String = "::ORPHAN"
Private Const LO_KEY_PREFIX As String = "lo::"

Public Sub Initialize(ByVal TreeView As TreeView)
    Dim il As ImageList
    Set il = New ImageList
    With il
        .ImageWidth = 16
        .ImageHeight = 16
        .ListImages.Add Key:="DDD", Picture:=Application.CommandBars.GetImageMso("TableInsert", 16, 16)
        .ListImages.Add Key:="EEE", Picture:=Application.CommandBars.GetImageMso("Help", 16, 16)
        .ListImages.Add Key:="AAA", Picture:=Application.CommandBars.GetImageMso("TableRowSelect", 16, 16)
        .ListImages.Add Key:="BBB", Picture:=Application.CommandBars.GetImageMso("TableSelect", 16, 16)
        .ListImages.Add Key:="CCC", Picture:=Application.CommandBars.GetImageMso("TableAutoFitFixedColumnWidth", 16, 16)
    End With
    
    With TreeView
        Set .ImageList = il
        .Nodes.Clear
        .Nodes.Add Key:="ROOT", Text:="States", Image:="CCC"
        .Nodes.Item(1).Expanded = True
        .FullRowSelect = False
        .HideSelection = False
        .Indentation = 16
        .Style = tvwTreelinesPictureText
        .LabelEdit = tvwManual
    End With
End Sub

Public Sub Load(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    AddParentNode TreeView, ViewModel
    AddTables TreeView, ViewModel
    AddStates TreeView, ViewModel
End Sub

Private Sub AddParentNode(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    Dim ParentNode As Node
    With TreeView.Nodes
        .Remove (1)
        Set ParentNode = .Add(Key:=ROOT_KEY, Text:=ROOT_CAPTION, Image:="CCC")
        .Item(1).Expanded = True
    End With
End Sub

Private Sub AddTables(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    Dim Workbook As Workbook
    Set Workbook = ViewModel.Target.Parent.Parent
    Dim TableNames As Collection
    Set TableNames = New Collection
    
    Dim Current As ColumnsState2
    Set Current = ViewModel.Current.State
    TableNames.Add Current.Name
    
    Dim HasOrphans As Boolean
    
    Dim State As ColumnsState2
    For Each State In ViewModel.States.CollectionView
        Dim TableName As String
        TableName = State.Name
        If Not CollectionHelpers.ExistsInCollection(TableNames, TableName) Then
            If ListObjectHelpers.ListObjectExists(Workbook, TableName) Then
                TableNames.Add TableName
            Else
                SetOrphan State
                HasOrphans = True
            End If
        End If
    Next State
    
    Dim ParentNode As Node
    Set ParentNode = TreeView.Nodes.Item(1)
    
    Dim TableToCreate As Variant
    For Each TableToCreate In TableNames
        TreeView.Nodes.Add Relative:=ParentNode, Relationship:=tvwChild, _
                           Key:=LO_KEY_PREFIX & TableToCreate, Text:=TableToCreate, Image:="DDD"
    Next TableToCreate
        
    If HasOrphans Then
        TreeView.Nodes.Add Relative:=ParentNode, Relationship:=tvwChild, _
                           Key:=ORPHAN_KEY, Text:=ORPHANS_CAPTION, Image:="EEE"
    End If
End Sub

Private Sub SetOrphan(ByVal OrphanState As IListable)
    OrphanState.ParentKey = ORPHAN_KEY
End Sub

Private Sub AddStates(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    Dim State As IListable
    For Each State In ViewModel.States.CollectionView
        Dim TableNode As Node
        Set TableNode = TreeView.Nodes.Item(State.ParentKey)
        TableNode.Expanded = True
        TreeView.Nodes.Add Relative:=TableNode, Relationship:=tvwChild, _
                           Key:=State.Key, Text:=State.Caption, Image:="AAA", SelectedImage:="BBB"
    Next State
End Sub


