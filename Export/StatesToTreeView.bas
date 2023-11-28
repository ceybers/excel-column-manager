Attribute VB_Name = "StatesToTreeView"
'@Folder("MVVM.ColumnState.ValueConverters")
Option Explicit

Private Const ROOT_CAPTION As String = "Column States"
Private Const CURRENT_SUFFIX_CAPTION As String = " (active)"
Private Const UNSAVED_CAPTION As String = "(current unsaved state)"
Private Const ORPHANS_CAPTION As String = "Orphans"
Private Const BUILTIN_CAPTION As String = "Built-in States"
Private Const SEARCH_CAPTION As String = "Search results"
Private Const NO_STATES_CAPTION As String = "No saved Column States found."
Private Const ICON_SIZE As Long = 16

Public Sub Initialize(ByVal TreeView As TreeView4)
    Dim TreeViewImageList As ImageList
    Set TreeViewImageList = GetImageList

    With TreeView
        .Nodes.Clear
        Set .ImageList = TreeViewImageList
        .FullRowSelect = False
        .HideSelection = False
        .Indentation = ICON_SIZE
        .Style = tvwTreelinesPictureText
        .LabelEdit = tvwAutomatic
    End With
End Sub

Private Function GetImageList() As ImageList
    Set GetImageList = New ImageList
    With GetImageList
        .ImageWidth = ICON_SIZE
        .ImageHeight = ICON_SIZE
        .ListImages.Add Key:=MSO_ROOT, Picture:=frmPictures.lblRoot.Picture
        .ListImages.Add Key:=MSO_BUILTIN, Picture:=frmPictures.lblTableBuiltin.Picture
        .ListImages.Add Key:=MSO_BUILTINITEM, Picture:=frmPictures.lblMethod.Picture
        .ListImages.Add Key:=MSO_TABLE, Picture:=frmPictures.lblTable.Picture
        .ListImages.Add Key:=MSO_ORPHAN, Picture:=frmPictures.lblTableOrphans.Picture
        .ListImages.Add Key:=MSO_UNSAVED, Picture:=frmPictures.lblFieldFriend.Picture
        .ListImages.Add Key:=MSO_SELECTEDSTATE, Picture:=frmPictures.lblField.Picture
        .ListImages.Add Key:=MSO_STATE, Picture:=frmPictures.lblField.Picture
        .ListImages.Add Key:=MSO_SEARCH, Picture:=frmPictures.lblSearch.Picture
    End With
End Function

Public Sub Load(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    AddParentNode TreeView
    AddBuiltin TreeView
    AddTables TreeView, ViewModel
    AddNoSearchResults TreeView, ViewModel
    AddUnsavedState TreeView, ViewModel
    AddStates TreeView, ViewModel
    CheckNoResults TreeView, ViewModel
    TrySelectSelectedItem TreeView, ViewModel
End Sub

Private Sub AddParentNode(ByVal TreeView As TreeView)
    With TreeView.Nodes
        If .Count > 0 Then .Remove (1)
        .Add Key:=ROOT_KEY, Text:=ROOT_CAPTION, Image:=MSO_ROOT
        .Item(1).Expanded = True
    End With
End Sub

Private Sub AddBuiltin(ByVal TreeView As TreeView)
    Dim ParentNode As Node
    Set ParentNode = TreeView.Nodes.Item(1)
    
    With TreeView.Nodes
        .Add Relative:=ParentNode, relationship:=tvwChild, _
             Key:=BUILTIN_KEY, Text:=BUILTIN_CAPTION, _
             Image:=MSO_BUILTIN
        .Item(1).Expanded = True
    End With
End Sub

Private Sub AddTables(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    Dim TableNames As Collection
    Set TableNames = New Collection
    
    Dim Current As ColumnsState
    Set Current = ViewModel.Current.State
    TableNames.Add Current.Name
    
    Dim HasOrphans As Boolean
    
    Dim Item As Object
    For Each Item In ViewModel.States.CollectionView
        If TypeOf Item Is ColumnsState Then
            Dim State As ColumnsState
            Set State = Item
            If IsOrphan(TableNames, State) Then
                If State.Orphan Then
                    HasOrphans = True
                Else
                    TableNames.Add State.Name
                End If
            End If
        End If
    Next Item
    
    Dim ParentNode As Node
    Set ParentNode = TreeView.Nodes.Item(1)
    
    Dim TableToCreate As Variant
    For Each TableToCreate In TableNames
        TreeView.Nodes.Add Relative:=ParentNode, relationship:=tvwChild, _
                           Key:=LO_KEY_PREFIX & TableToCreate, Text:=TableToCreate, _
                           Image:=MSO_TABLE
    Next TableToCreate
        
    ' Make first table (current) bold
    Dim FirstTable As Node
    Set FirstTable = TreeView.Nodes.Item(LO_KEY_PREFIX & TableNames.Item(1))
    FirstTable.Bold = True
    FirstTable.Text = FirstTable.Text & CURRENT_SUFFIX_CAPTION

    If HasOrphans Then
        TreeView.Nodes.Add Relative:=ParentNode, relationship:=tvwChild, _
                           Key:=ORPHAN_KEY, Text:=ORPHANS_CAPTION, _
                           Image:=MSO_ORPHAN
    End If
    
End Sub

Private Sub AddNoSearchResults(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    Dim ParentNode As Node
    Set ParentNode = TreeView.Nodes.Item(1)
    
    Dim FolderNode As Node
    Set FolderNode = TreeView.Nodes.Add(Relative:=ParentNode, relationship:=tvwChild, _
                                        Key:=SEARCH_FOLDER_KEY, Text:=SEARCH_CAPTION, _
                                        Image:=MSO_SEARCH)
    FolderNode.Expanded = True
                       
    Dim Node As Node
    Set Node = TreeView.Nodes.Add(Relative:=FolderNode, relationship:=tvwChild, _
                                  Key:=NO_SEARCH_RESULT_KEY, Text:=NO_STATES_CAPTION)
    Node.ForeColor = modConstants.GREY_TEXT_COLOR
End Sub

Private Function IsOrphan(ByVal TableNames As Collection, ByVal State As ColumnsState) As Boolean
    IsOrphan = Not CollectionHelpers.ExistsInCollection(TableNames, State.Name)
End Function

Private Sub SetOrphan(ByVal OrphanState As IListable)
    OrphanState.ParentKey = ORPHAN_KEY
End Sub

Private Sub AddUnsavedState(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    Dim Current As ColumnsState
    Set Current = ViewModel.Current.State
    
    Dim TableNode As Node
    Set TableNode = TreeView.Nodes.Item(LO_KEY_PREFIX & Current.Name)
    
    Dim Node As Node
    Set Node = TreeView.Nodes.Add(Relative:=TableNode, relationship:=tvwChild, _
                                  Key:=UNSAVED_KEY, Text:=UNSAVED_CAPTION, _
                                  Image:=MSO_UNSAVED, SelectedImage:=MSO_UNSAVED)
    Node.Bold = True
    Node.Selected = True
End Sub

Private Sub AddStates(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    Dim RemoveUnsaved As Boolean
    
    Dim State As IListable
    For Each State In ViewModel.States.CollectionView
        Dim TableNode As Node
        Set TableNode = TreeView.Nodes.Item(State.ParentKey)
        TableNode.Expanded = True
        
        Dim Node As Node
        Set Node = TreeView.Nodes.Add(Relative:=TableNode, relationship:=tvwChild, _
                                      Key:=State.Key, Text:=State.Caption, _
                                      Image:=MSO_STATE, SelectedImage:=MSO_SELECTEDSTATE)
                           
        If State.ParentKey = BUILTIN_KEY Then
            Node.Image = MSO_BUILTINITEM
            Node.SelectedImage = MSO_BUILTINITEM
        End If
        
        If RemoveUnsaved = False Then
            If MatchesCurrent(ViewModel, State) Then
                RemoveUnsaved = True
                Node.Bold = True
                Node.Selected = True
            End If
        End If
    Next State
    
    If RemoveUnsaved Then
        TreeView.Nodes.Remove UNSAVED_KEY
    End If
End Sub

Private Function MatchesCurrent(ByVal ViewModel As StateManagerViewModel, ByVal State As IState) As Boolean
    MatchesCurrent = State.Equals(ViewModel.Current.State)
End Function

Private Sub CheckNoResults(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    If ViewModel.States.HasNoSearchResults Then Exit Sub
    
    With TreeView.Nodes
        .Remove NO_SEARCH_RESULT_KEY
        .Remove SEARCH_FOLDER_KEY
    End With
End Sub

Private Sub TrySelectSelectedItem(ByVal TreeView As TreeView, ByVal ViewModel As StateManagerViewModel)
    If TreeView.SelectedItem Is Nothing Then Exit Sub
    ViewModel.TrySelect TreeView.SelectedItem.Key
End Sub


