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
'@ModuleDescription "ViewModel that contains VMs for Current, Select, and Browsing States."
'@Folder("MVVM.ColumnState.ViewModel")
Option Explicit

Private Type TState
    Target As ListObject
    
    Current As CurrentStateViewModel
    Selected As SelectedStateViewModel
    States As StatesViewModel
    
    Model As StatesModel
End Type

Private This As TState

Public Property Get Target() As ListObject
    Set Target = This.Target
End Property

Public Property Set Target(ByVal vNewValue As ListObject)
    Set This.Target = vNewValue
    UpdateCurrent
End Property

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
End Sub

Public Sub Load(ByVal Model As StatesModel)
    Set This.Model = Model
    This.States.Load Model
End Sub

Public Sub TrySelect(ByVal Key As String)
    'If Key = "ROOT" Then
    '    Set This.Selected.State = Nothing
    '    Exit Sub
    'End If
    If Left$(Key, 1) <> "K" Then
        Set This.Selected.State = Nothing
        Exit Sub
    End If
    Set This.Selected.State = This.States.CollectionView.Item(Key)
End Sub

Private Sub UpdateCurrent()
    Set This.Current.State = ColumnsState2.Create(This.Target)
End Sub

Public Function CanSave() As Boolean
    CanSave = Not This.Model.Exists(This.Current.State)
End Function

Public Sub Save()
    This.Model.Add This.Current.State
    ' DEBUG This.Model.Save
    This.States.Load This.Model
End Sub

Public Function CanApply() As Boolean
    If This.Selected.State Is Nothing Then Exit Function
    If This.Current.State.Equals(This.Selected.State) Then Exit Function
    CanApply = True
End Function

Public Sub Apply()
    ApplyStateToTarget This.Selected.State
    UpdateCurrent
End Sub

Private Function ApplyStateToTarget(ByVal State As IState)
    State.Apply This.Target
    UpdateCurrent
End Function

Public Function CanPrune() As Boolean
    CanPrune = False
End Function

Public Sub Prune()
    ' NYI
End Sub

Public Function CanRemove() As Boolean
    If This.Selected.State Is Nothing Then Exit Function
    CanRemove = True
End Function

Public Function Remove() As Boolean
    This.Model.Remove This.Selected.State
    ' DEBUG This.Model.Save
    Set This.Selected.State = Nothing
    This.States.Load This.Model
End Function

Public Function RemoveAll() As Boolean
    This.Model.RemoveAll
    ' DEBUG This.Model.Save
    Set This.Selected.State = Nothing
    This.States.Load This.Model
End Function

Public Function CanExport() As Boolean
    CanExport = Not This.Selected.State Is Nothing
End Function

