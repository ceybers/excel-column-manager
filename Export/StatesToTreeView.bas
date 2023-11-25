Attribute VB_Name = "StatesToTreeView"
'@Folder("MVVM.ColumnState.ValueConverters")
Option Explicit

Private Const ROOT_CAPTION As String = "Column States"
Private Const CURRENT_SUFFIX_CAPTION As String = " (current)"
Private Const UNSAVED_CAPTION As String = "(current)"
Private Const ORPHANS_CAPTION As String = "Orphans"
Private Const NO_STATES_CAPTION As String = "No saved Column States found."

Public Sub Initialize(ByVal TreeView As TreeView)
    Dim il As ImageList
    Set il = New ImageList
    With il
        .ImageWidth = 16
        .ImageHeight = 16
        .ListImages.Add Key:="msoUnsaved", Picture:=Application.CommandBars.GetImageMso("TableStyleBandedColumns", 16, 16)
        .ListImages.Add Key:="msoTable", Picture:=Application.CommandBars.GetImageMso("TableInsert", 16, 16)
        .ListImages.Add Key:="msoOrphan", Picture:=Application.CommandBars.GetImageMso("Help", 16, 16)
        .ListImages.Add Key:="msoItem", Picture:=Application.CommandBars.GetImageMso("TableRowSelect", 16, 16)
        .ListImages.Add Key:="msoSelected", Picture:=Application.CommandBars.GetImageMso("TableSelect", 16, 16)
        .ListImages.Add Key:="msoRoot", Picture:=Application.CommandBars.GetImageMso("TableAutoFitFixedColumnWidth", 16, 16)
    End With
    
    With TreeView
        Set .ImageList = il
        .Nodes.Clear
        .Nodes.Add Key:="ROOT", Text:="States", Image:="msoRoot"
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
    AddUnsavedState TreeView, ViewModel
    AddStates TreeView, ViewModel
    CheckNoResults TreeView, ViewModel
End Sub

Private Sub AddParentNode(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    Dim ParentNode As Node
    With TreeView.Nodes
        .Remove (1)
        Set ParentNode = .Add(Key:=ROOT_KEY, Text:=ROOT_CAPTION, Image:="msoRoot")
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
            If State.Orphan Then
                HasOrphans = True
            Else
                TableNames.Add TableName
            End If
        End If
    Next State
    
    Dim ParentNode As Node
    Set ParentNode = TreeView.Nodes.Item(1)
    
    Dim TableToCreate As Variant
    For Each TableToCreate In TableNames
        TreeView.Nodes.Add Relative:=ParentNode, Relationship:=tvwChild, _
                           Key:=LO_KEY_PREFIX & TableToCreate, Text:=TableToCreate, _
                           Image:="msoTable"
    Next TableToCreate
        
    If HasOrphans Then
        TreeView.Nodes.Add Relative:=ParentNode, Relationship:=tvwChild, _
                           Key:=ORPHAN_KEY, Text:=ORPHANS_CAPTION, _
                           Image:="msoOrphan"
    End If
End Sub

Private Sub SetOrphan(ByVal OrphanState As IListable)
    OrphanState.ParentKey = ORPHAN_KEY
End Sub

Private Sub AddUnsavedState(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    Dim Current As ColumnsState2
    Set Current = ViewModel.Current.State
    
    Dim TableNode As Node
    Set TableNode = TreeView.Nodes.Item(LO_KEY_PREFIX & Current.Name)
    
    Dim Node As Node
    Set Node = TreeView.Nodes.Add(Relative:=TableNode, Relationship:=tvwChild, _
                                  Key:=UNSAVED_KEY, Text:=UNSAVED_CAPTION, _
                                  Image:="msoItem", SelectedImage:="msoUnsaved")
    Node.Bold = True
    Node.Selected = True
End Sub

Private Sub AddStates(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    Dim RemoveUnsaved As Boolean
    
    Dim State As IListable
    For Each State In ViewModel.States.CollectionView
        'Debug.Assert State.ParentKey <> modConstants.ORPHAN_KEY
    
        Dim TableNode As Node
        Set TableNode = TreeView.Nodes.Item(State.ParentKey)
        TableNode.Expanded = True
        
        Dim Node As Node
        Set Node = TreeView.Nodes.Add(Relative:=TableNode, Relationship:=tvwChild, _
                                      Key:=State.Key, Text:=State.Caption, _
                                      Image:="msoItem", SelectedImage:="msoSelected")
                           
        If MatchesCurrent(ViewModel, State) Then
            RemoveUnsaved = True
        End If
    Next State
    
    If RemoveUnsaved Then
        TreeView.Nodes.Remove UNSAVED_KEY
        Node.Bold = True
        Node.Selected = True
    End If
End Sub

Private Function MatchesCurrent(ByVal ViewModel As StateManagerViewModel, ByVal State As IState) As Boolean
    MatchesCurrent = State.Equals(ViewModel.Current.State)
End Function

Private Sub CheckNoResults(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    If TreeView.Nodes.Count > 2 Then Exit Sub
    
    Dim Node As Node
    Set Node = TreeView.Nodes.Add(Relative:=TreeView.Nodes.Item(1), Relationship:=tvwChild, _
                                  Key:=NO_STATES_KEY, Text:=NO_STATES_CAPTION)
    Node.ForeColor = modConstants.GREY_TEXT_COLOR
End Sub


