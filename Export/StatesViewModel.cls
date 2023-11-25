VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StatesViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
'@Folder("MVVM.ColumnState.ViewModel")
Option Explicit

Private Type TState
    Target As ListObject
    
    Items As Collection
    CollectionView As Collection
    HasOrphans As Boolean
    
    Filter As String
    
    Model As StatesModel
    Options As OptionsViewModel
    Workbook As Workbook
End Type

Private This As TState

Public Property Get Target() As ListObject
    Set Target = This.Target
End Property

Public Property Set Target(ByVal vNewValue As ListObject)
    Set This.Target = vNewValue
    UpdateMatchTarget
    UpdateCollectionView
End Property

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

Public Sub Load(ByVal Model As StatesModel, ByVal Options As OptionsViewModel, ByVal Workbook As Workbook)
    Set This.Options = Options
    Set This.Model = Model
    Set This.Workbook = Workbook
    Reload
End Sub

Public Sub Reload()
    UpdateMatchTarget
    
    Dim i As Long
    i = 1
    
    CollectionHelpers.CollectionClear This.Items
    
    Dim State As IState
    For Each State In This.Model.Items
        Dim StateListable As IListable
        Set StateListable = State
        StateListable.Key = "K" & Format$(i, "000") ' WARNING this means only one model can assign keys!
        This.Items.Add Item:=StateListable, Key:=StateListable.Key
        i = i + 1
    Next State
    
    UpdateCollectionView
End Sub

Private Sub UpdateCollectionView()
    ' Greedy filter. Will cause us to always delete all ListItems and recreate them.
    ' Won't preserve checked/selected status, or active ListItem in control, or caret focus.
    CollectionHelpers.CollectionClear This.CollectionView
    
    UpdateOrphans This.Workbook
    
    Dim FilterString As String
    If This.Filter = vbNullString Then
        FilterString = "*"
    ElseIf DoMatchWhole Then
        FilterString = This.Filter
    Else
        FilterString = "*" & This.Filter & "*"
    End If
    
    Dim StateListable As IListable
    For Each StateListable In This.Items
        If DoFilter(StateListable, FilterString) Then
            This.CollectionView.Add Item:=StateListable, Key:=StateListable.Key
        End If
    Next StateListable
End Sub

Private Function DoFilter(ByVal StateListable As IListable, ByVal FilterString As String) As Boolean
    If DoMatchCase Then
        DoFilter = StateListable.Caption Like FilterString
    Else
        DoFilter = UCase$(StateListable.Caption) Like UCase$(FilterString)
    End If
    
    Dim State As ColumnsState
    Set State = StateListable
        
    If Not DoShowOrphanStates Then
        If State.Orphan Then
            DoFilter = False
        End If
    End If
    
    If Not DoShowPartialMatch Then
        If State.Match <> FULL_MATCH Then
            DoFilter = False
        End If
    End If
End Function

Public Sub UpdateOrphans(ByVal Workbook As Workbook)
    This.HasOrphans = False
    
    Dim State As ColumnsState
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
    Dim State As ColumnsState
    For Each State In This.Items
        If State.Orphan Then
            Model.Remove State
        End If
    Next State
End Sub

Private Sub UpdateMatchTarget()
    Dim MatchCount As Long
    Dim TotalCount As Long
        
    Dim State As ColumnsState
    For Each State In This.Items
        MatchCount = 0
        TotalCount = 0
        
        Dim Child As ColumnState
        For Each Child In State.Items
            Child.Orphan = Not ListObjectHelpers.HasListColumn(This.Target, Child.Name)
            If Not Child.Orphan Then MatchCount = MatchCount + 1
            TotalCount = TotalCount + 1
        Next Child
        
        If MatchCount = TotalCount Then
            State.Match = FULL_MATCH
        ElseIf MatchCount = 0 Then
            State.Match = NO_MATCH
        Else
            State.Match = PARTIAL_MATCH
        End If
    Next State
End Sub

' Options
Private Function DoMatchCase() As Boolean
    DoMatchCase = This.Options.GetFlag(modOptionsConst.DO_SEARCH_CASE)
End Function

Private Function DoMatchWhole() As Boolean
    DoMatchWhole = This.Options.GetFlag(modOptionsConst.DO_SEARCH_WHOLE)
End Function

Private Function DoShowOrphanStates() As Boolean
    DoShowOrphanStates = This.Options.GetFlag(modOptionsConst.DO_SHOW_ORPHAN_STATES)
End Function

Private Function DoShowPartialMatch() As Boolean
    DoShowPartialMatch = This.Options.GetFlag(modOptionsConst.DO_SHOW_PARTIAL_MATCH)
End Function

