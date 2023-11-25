VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StatesViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.ColumnState.ViewModel")
Option Explicit

Private Type TState
    Items As Collection
    CollectionView As Collection
    HasOrphans As Boolean
    
    Filter As String
End Type

Private This As TState

Public Property Get Count() As Long
    Count = This.Items.Count
End Property

Public Property Get CollectionView() As Collection
    Set CollectionView = This.CollectionView
End Property

Public Property Get Filter() As String
    Filter = This.Filter
End Property

Public Property Let Filter(ByVal vNewValue As String)
    This.Filter = vNewValue
    UpdateCollectionView
End Property

Private Sub Class_Initialize()
    Set This.Items = New Collection
    Set This.CollectionView = New Collection
End Sub

Public Sub Load(ByVal Model As StatesModel)
    Dim i As Long
    i = 1
    
    CollectionHelpers.CollectionClear This.Items
    
    Dim State As IState
    For Each State In Model.Items
        Dim StateListable As IListable
        Set StateListable = State
        StateListable.Key = "K" & Format(i, "000") ' WARNING this means only one model can assign keys!
        This.Items.Add Item:=StateListable, Key:=StateListable.Key
        i = i + 1
    Next State
    
    UpdateCollectionView
End Sub

Private Sub UpdateCollectionView()
    ' Greedy filter. Will cause us to always delete all ListItems and recreate them.
    ' Won't preserve checked/selected status, or active ListItem in control, or caret focus.
    CollectionHelpers.CollectionClear This.CollectionView
    
    Dim FilterString As String
    If This.Filter = vbNullString Then
        FilterString = "*"
    Else
        FilterString = "*" & This.Filter & "*"
    End If
    
    Dim StateListable As IListable
    For Each StateListable In This.Items
        If StateListable.Caption Like FilterString Then
            This.CollectionView.Add Item:=StateListable, Key:=StateListable.Key
        End If
    Next StateListable
End Sub

Public Sub UpdateOrphans(ByVal Workbook As Workbook)
    This.HasOrphans = False
    
    Dim State As ColumnsState2
    For Each State In This.Items
        If Not ListObjectHelpers.ListObjectExists(Workbook, State.Name) Then
            State.Orphan = True
            SetOrphan State
            This.HasOrphans = True
        End If
    Next State
End Sub

Private Sub SetOrphan(ByVal State As IListable)
    State.ParentKey = ORPHAN_KEY
End Sub

Public Function HasOrphans() As Boolean
    HasOrphans = This.HasOrphans
End Function

Public Sub RemoveOrphans(ByVal Model As StatesModel)
    Dim State As ColumnsState2
    For Each State In This.Items
        If State.Orphan Then
            Model.Remove State
        End If
    Next State
End Sub

