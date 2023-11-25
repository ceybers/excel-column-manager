VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FirstOnlyColumnsState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
'@ModuleDescription "Built-in state that displays only the first ListColumns in a single ListObject."
'@Folder("State.Entities")
Option Explicit
Implements IState
Implements ISerializable
Implements IListable

Private Const BUILTIN_NAME As String = "(Show first column only)"
Private Const THIS_KEY As String = "B::FirstOnly"

Private Type TState
    'ListObject As ListObject
    Name As String
    ' IListable
    Key As String
    ParentKey As String
    Visible As Boolean
End Type

Private This As TState

'@DefaultMember
Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    This.Name = BUILTIN_NAME
    This.ParentKey = BUILTIN_KEY
End Property

Public Property Get Match() As ListObjectMatch
    Match = FULL_MATCH
End Property

'Public Property Let Match(ByVal vNewValue As ListObjectMatch)
'End Property

'@Description "Returns True if this ListObject of this state cannot be found in the VM's workbook."
Public Property Get Orphan() As Boolean
    Orphan = False
End Property

'Public Property Let Orphan(ByVal RHS As Boolean)
'End Property

Public Function ToString() As String
    ToString = BUILTIN_NAME
End Function

' Requires PredeclaredId annotation
Public Function Create() As IState
    Set Create = New FirstOnlyColumnsState
End Function

' IState
Private Function IState_From(ByVal Object As Object) As Boolean
    IState_From = False
End Function

Private Function IState_CanApply(ByVal Target As Object) As Boolean
    If Not TypeOf Target Is ListObject Then Exit Function
    IState_CanApply = True
End Function

Private Function IState_Apply(ByVal Target As Object) As Boolean
    If Not TypeOf Target Is ListObject Then Exit Function
    
    Dim ListObject As ListObject
    Set ListObject = Target
    
    Dim ListColumn As ListColumn
    For Each ListColumn In ListObject.ListColumns
        ListColumn.Range.EntireColumn.ColumnWidth = 0
    Next ListColumn
    
    ListObject.ListColumns.Item(1).Range.EntireColumn.ColumnWidth = 8.43
    
    IState_Apply = True
End Function

' ISerializable
Private Function ISerializable_Serialize() As String
End Function

Private Function ISerializable_Deserialize(ByVal SerialString As String) As Boolean
End Function

Private Function IState_Equals(ByVal RHS As IState) As Boolean
    If RHS Is Nothing Then Exit Function

    IState_Equals = (TypeOf RHS Is StandardColumnsState)
End Function

Private Property Get IListable_Caption() As String
    IListable_Caption = BUILTIN_NAME
End Property

Private Property Get IListable_Key() As String
    IListable_Key = THIS_KEY
End Property

Private Property Let IListable_Key(ByVal vNewValue As String)
End Property

Private Property Get IListable_ParentKey() As String
    IListable_ParentKey = BUILTIN_KEY
End Property

Private Property Let IListable_ParentKey(ByVal vNewValue As String)
End Property

Private Property Get IListable_Visible() As Boolean
    IListable_Visible = This.Visible
End Property

Private Property Let IListable_Visible(ByVal vNewValue As Boolean)
    This.Visible = vNewValue
End Property

