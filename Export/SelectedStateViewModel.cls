VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectedStateViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.ColumnState.ViewModel")
Option Explicit

Private Const RENAME_MESSAGE As String = "Rename the selected Column State:"
Private Const RENAME_TITLE As String = "Rename Column State"
    
Private Type TState
    State As IState
    'HideSmallColumns As Boolean
    Items As Collection
    
    Options As OptionsViewModel
End Type

Private This As TState

Public Property Get ShowNonMatchingCols() As Boolean
    ShowNonMatchingCols = This.Options.GetFlag(modOptionsConst.DO_SHOW_NONMATCHING_COLS)
End Property

Public Property Get Items() As Collection
    Set Items = This.Items
End Property

Public Property Get State() As IListable
    Set State = This.State
End Property

Public Property Set State(ByVal vNewValue As IListable)
    Set This.State = vNewValue
    UpdateItems
End Property

'Public Property Get HideSmallColumns() As Boolean
'    HideSmallColumns = This.HideSmallColumns
'End Property

'Public Property Let HideSmallColumns(ByVal vNewValue As Boolean)
'    This.HideSmallColumns = vNewValue
'    UpdateHiddenItems
'End Property

Private Sub Class_Initialize()
    Set This.Items = New Collection
End Sub

Public Sub Load(ByVal Options As OptionsViewModel)
    Set This.Options = Options
End Sub

Private Sub UpdateItems()
    CollectionHelpers.CollectionClear This.Items
    
    If This.State Is Nothing Then Exit Sub
    
    If TypeOf This.State Is ColumnsState Then AddChildren This.State
    
    UpdateHiddenItems
End Sub

Private Sub AddChildren(ByVal State As ColumnsState)
    Dim Child As IListable
    For Each Child In State.Items
        This.Items.Add Item:=Child
    Next Child
End Sub

Private Sub UpdateHiddenItems()
    Dim Child As IListable
    For Each Child In This.Items
        Dim State As ColumnState
        Set State = Child
        Child.Visible = True                     'Not (This.HideSmallColumns And State.Width <= 8#)
    Next Child
End Sub

Public Sub UpdateOrphans(ByVal Target As ListObject)
    Dim Child As ColumnState
    For Each Child In This.Items
        Child.Orphan = Not ListObjectHelpers.HasListColumn(Target, Child.Name)
    Next Child
End Sub

Public Function TryRename() As Boolean
    If TypeOf This.State Is ColumnsState Then
        TryRename = TryRenameColumnsState(This.State)
    End If
End Function

Private Function TryRenameColumnsState(ByVal State As ColumnsState) As Boolean
    Dim CurrentName As String
    CurrentName = State.Caption
    
    Dim NewName As String
    NewName = InputBox(RENAME_MESSAGE, RENAME_TITLE, CurrentName)
    
    If NewName = vbNullString Then Exit Function
    If NewName = CurrentName Then Exit Function
    
    State.Caption = NewName
    TryRenameColumnsState = True
End Function

Public Function IsBuiltin() As Boolean
    If This.State Is Nothing Then Exit Function
    IsBuiltin = (This.State.StateType = BUILTIN_STATE)
    
    'Dim StateListable As IListable
    'Set StateListable = This.State
    'If StateListable.ParentKey = BUILTIN_KEY Then
    '    IsBuiltin = True
    'End If
End Function

