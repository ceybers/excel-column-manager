Attribute VB_Name = "StatesToTreeView"
'@Folder("MVVM.ColumnState.ValueConverters")
Option Explicit

Public Sub Initialize(ByVal TreeView As TreeView)
    Dim il As ImageList
    Set il = New ImageList
    With il
        .ImageWidth = 16
        .ImageHeight = 16
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

Public Sub Load(ByVal TreeView As TreeView, ByVal ViewModel As StatesViewModel)
    Dim ParentNode As Node
    With TreeView.Nodes
        .Remove (1)
        Set ParentNode = .Add(Key:="ROOT", Text:="States", Image:="CCC")
        .Item(1).Expanded = True
    End With
    
    Dim State As IListable
    For Each State In ViewModel.CollectionView
        TreeView.Nodes.Add Relative:=ParentNode, Relationship:=tvwChild, Key:=State.Key, Text:=State.Caption, Image:="AAA", SelectedImage:="BBB"
    Next State
End Sub

