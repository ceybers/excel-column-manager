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

Private Type TState
    State As ColumnsState
    HideSmallColumns As Boolean
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

Public Property Get HideSmallColumns() As Boolean
    HideSmallColumns = This.HideSmallColumns
End Property

Public Property Let HideSmallColumns(ByVal vNewValue As Boolean)
    This.HideSmallColumns = vNewValue
    UpdateHiddenItems
End Property

Private Sub Class_Initialize()
    Set This.Items = New Collection
End Sub

Public Sub Load(ByVal Options As OptionsViewModel)
    Set This.Options = Options
End Sub

Private Sub UpdateItems()
    CollectionHelpers.CollectionClear This.Items
    
    If Not This.State Is Nothing Then
        Dim Child As IListable
        For Each Child In This.State.Items
            This.Items.Add Item:=Child           ', Key:=Child.Key           ' No keys needed
        Next Child
    End If
    
    UpdateHiddenItems
End Sub

Private Sub UpdateHiddenItems()
    Dim Child As IListable
    For Each Child In This.Items
        Dim State As ColumnState
        Set State = Child
        Child.Visible = Not (This.HideSmallColumns And State.Width <= 8#)
    Next Child
End Sub

Public Sub UpdateOrphans(ByVal Target As ListObject)
    Dim Child As ColumnState
    For Each Child In This.Items
        Child.Orphan = Not ListObjectHelpers.HasListColumn(Target, Child.Name)
    Next Child
End Sub

