VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StateManagerViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "ViewModel that contains VMs for Current, Select, and Browsing States."
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
'@ModuleDescription "ViewModel that contains VMs for Current, Select, and Browsing States."
'@Folder("MVVM.ColumnState.ViewModel")
Option Explicit

Private Type TState
    Target As ListObject
    
    Current As CurrentStateViewModel
    Selected As SelectedStateViewModel
    States As StatesViewModel
    Options As OptionsViewModel
    
    Model As StatesModel
End Type

Private This As TState

Public Property Get Target() As ListObject
    Set Target = This.Target
End Property

Public Property Get CloseOnApply() As Boolean
    CloseOnApply = This.Options.GetFlag(modOptionsConst.DO_CLOSE_ON_APPLY)
End Property

'Public Property Set Target(ByVal vNewValue As ListObject)
'    Set This.Target = vNewValue
'End Property

Public Property Get Current() As CurrentStateViewModel
    Set Current = This.Current
End Property

Public Property Get Selected() As SelectedStateViewModel
    Set Selected = This.Selected
End Property

Public Property Get States() As StatesViewModel
    Set States = This.States
End Property

Private Sub Class_Initialize()
    Set This.Current = New CurrentStateViewModel
    Set This.Selected = New SelectedStateViewModel
    Set This.States = New StatesViewModel
    Set This.Options = New OptionsViewModel
End Sub

Public Sub Load(ByVal Model As StatesModel, ByVal Target As ListObject)
    Set This.Target = Target
    UpdateCurrent
        
    Set This.Model = Model
    
    This.Options.Load Model
    
    This.States.Load Model, This.Options, This.Target.Parent.Parent
    Set This.States.Target = This.Target
    
    This.Selected.Load This.Options
End Sub

Public Sub TrySelect(ByVal Key As String)
    If Key = UNSAVED_KEY Then
        Set This.Selected.State = This.Current.State
    ElseIf Left$(Key, 3) = "B::" Then            ' TODO Const builtin prefix
        Set This.Selected.State = This.States.CollectionView.Item(Key)
    ElseIf Left$(Key, 1) <> "K" Then
        Set This.Selected.State = Nothing
        Exit Sub
    Else
        Set This.Selected.State = This.States.CollectionView.Item(Key)
        This.Selected.UpdateOrphans This.Target
    End If
End Sub

Private Sub UpdateCurrent()
    Set This.Current.State = ColumnsState.Create(This.Target)
    This.States.UpdateOrphans Target.Parent.Parent
End Sub

Public Function CanSave() As Boolean
    CanSave = Not This.Model.Exists(This.Current.State)
End Function

Public Sub Save()
    This.Model.Add This.Current.State
    SaveAndReload
End Sub

Public Function CanApply() As Boolean
    If This.Selected.State Is Nothing Then Exit Function
    If This.Current.State.Equals(This.Selected.State) Then Exit Function
    
    If Not ApplyAllowed(This.Selected.State) Then Exit Function

    CanApply = True
End Function

Private Function ApplyAllowed(ByVal State As IState) As Boolean
    If Not TypeOf State Is ColumnsState Then     ' TODO This is a bad idea
        ApplyAllowed = True
        Exit Function
    End If
    
    Dim StateCast As ColumnsState
    Set StateCast = State
    If StateCast.Match = FULL_MATCH Then
        ApplyAllowed = True
    ElseIf This.Options.GetFlag(modOptionsConst.DO_ALLOW_APPLY_PARTIAL_MATCH) Then
        ApplyAllowed = True
    End If
End Function

Public Sub Apply()
    Application.ScreenUpdating = False
    
    If This.Options.GetFlag(modOptionsConst.DO_HIDE_NONMATCH_COLS) Then
        Target.Range.EntireColumn.Hidden = True
    End If
    
    ApplyStateToTarget This.Selected.State
    UpdateCurrent
    
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    
    Application.ScreenUpdating = True
End Sub

Private Sub ApplyStateToTarget(ByVal State As IState)
    State.Apply This.Target
    
    If TypeOf State Is ColumnsState Then
        TryAssociateOrphan State
    End If

    UpdateCurrent
End Sub

Private Sub TryAssociateOrphan(ByVal State As ColumnsState)
    If Not State.Orphan Then Exit Sub
    If Not This.Options.GetFlag(modOptionsConst.DO_ASSOCIATE_ORPHAN_STATES) Then Exit Sub
    
    Dim NewState As ColumnsState
    Set NewState = CloneState(State)
    NewState.Name = This.Target.Name
    
    This.Model.Remove State
    This.Model.Add NewState
    
    SaveAndReload
End Sub

Private Function CloneState(ByVal OrphanState As ISerializable) As ColumnsState
    Dim Result As ISerializable
    Set Result = New ColumnsState
    Result.Deserialize OrphanState.Serialize
    Set CloneState = Result
End Function

Public Function CanPrune() As Boolean
    CanPrune = This.States.HasOrphans
End Function

Public Sub Prune()
    This.States.RemoveOrphans This.Model
    
    SaveAndReload
End Sub

Public Function CanRemove() As Boolean
    If This.Selected.State Is Nothing Then Exit Function
    CanRemove = True
End Function

Public Sub Remove()
    This.Model.Remove This.Selected.State
    Set This.Selected.State = Nothing
    
    SaveAndReload
End Sub

Public Function CanRemoveAll() As Boolean
    If This.States.Count = 0 Then Exit Function
    CanRemoveAll = True
End Function

Public Sub RemoveAll()
    This.Model.RemoveAll
    
    Set This.Selected.State = Nothing
    
    SaveAndReload
End Sub

Public Function CanExport() As Boolean
    CanExport = Not This.Selected.State Is Nothing
End Function

Public Function TryImport(ByVal SerialString As String, ByRef OutState As IState) As Boolean
    Dim State As ISerializable
    Set State = New ColumnsState
    If Not State.Deserialize(SerialString) Then
        Exit Function
    End If
    
    Set OutState = State
    
    If This.Model.Exists(State) Then
        Exit Function
    End If
    
    This.Model.Add State
    Set This.Selected.State = State
    SaveAndReload

    TryImport = True
End Function

Public Sub ShowOptions()
    Dim View As IView
    Set View = New OptionsView
    
    View.ShowDialog This.Options
    
    Set View = Nothing
End Sub

Private Sub SaveAndReload()
    This.Model.Save
    This.States.Reload
End Sub

Public Function CanRename() As Boolean
    If This.Selected.State Is Nothing Then Exit Function
    If This.Selected.State.ParentKey = modConstants.BUILTIN_KEY Then Exit Function
    
    CanRename = True
End Function

Public Function Rename() As Boolean
    If Not CanRename Then Exit Function
    
    If This.Selected.TryRename(vbNullString) Then
        This.Model.Save
        This.States.Reload
        Rename = True
    End If
End Function
