VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnsState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Stores the state of multiple ListColumns from a single ListObject."
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
'@ModuleDescription "Stores the state of multiple ListColumns from a single ListObject."
'@PredeclaredId
'@Folder("State.Entities")
Option Explicit
Implements IState
Implements ISerializable
Implements IListable

Private Const MAX_CHILDREN_IN_CAPTION As Long = 5
Private Const MAX_COLUMN_CAPTION As Long = 12

Private Type TState
    ListObject As ListObject
    Name As String
    Caption As String
    Freeze As FreezePaneState
    Outline As OutlineState
    ListColumns As Collection
    Orphan As Boolean
    Match As ListObjectMatch
    StateType As StateType
    ' IListable
    Key As String
    ParentKey As String
    Visible As Boolean
End Type

Private This As TState

Public Property Get Items() As Collection
    Set Items = This.ListColumns
End Property

'@DefaultMember
Public Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    Name = This.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    This.Name = vNewValue
    This.ParentKey = LO_KEY_PREFIX & vNewValue
End Property

Public Property Get Caption() As String
    Caption = This.Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    This.Caption = vNewValue
    If vNewValue = vbNullString Then
        SetAutoCaption
    End If
End Property

Public Property Get Match() As ListObjectMatch
    Match = This.Match
End Property

Public Property Let Match(ByVal vNewValue As ListObjectMatch)
    This.Match = vNewValue
End Property

'@Description "Returns True if this ListObject of this state cannot be found in the VM's workbook."
Public Property Get Orphan() As Boolean
Attribute Orphan.VB_Description = "Returns True if this ListObject of this state cannot be found in the VM's workbook."
    Orphan = This.Orphan
End Property

Public Property Let Orphan(ByVal RHS As Boolean)
    This.Orphan = RHS
End Property

Private Property Get IState_StateType() As StateType
    IState_StateType = This.StateType
End Property

Public Function ToString() As String
    Dim Result As String
    Result = This.Name & " has " & This.ListColumns.Count & " column(s). "
    
    Dim State As ColumnState
    For Each State In This.ListColumns
        Result = Result & State.ToString & ", "
    Next State
    
    Result = Left$(Result, Len(Result) - 2) & "."
    ToString = Result
End Function

Private Sub Class_Initialize()
    Set This.Freeze = New FreezePaneState
    Set This.Outline = New OutlineState
    Set This.ListColumns = New Collection
End Sub

Public Function Create(ByVal ListObject As ListObject) As IState
    Dim Result As IState
    Set Result = New ColumnsState
    Result.From ListObject
    Set Create = Result
End Function

Private Function IState_From(ByVal Object As Object) As Boolean
    If Not TypeOf Object Is ListObject Then Exit Function
    
    Set This.ListObject = Object
    Name = This.ListObject.Name
    
    Set This.Freeze = FreezePaneState.Create
    Set This.Outline = OutlineState.Create(This.ListObject.Parent)
    
    Dim ListColumn As ListColumn
    For Each ListColumn In This.ListObject.ListColumns
        This.ListColumns.Add ColumnState.Create(ListColumn)
    Next ListColumn
    
    SetAutoCaption
    
    This.StateType = UNSAVED_STATE
    
    IState_From = True
End Function

Private Function IState_CanApply(ByVal Target As Object) As Boolean
    If Not TypeOf Target Is ListObject Then Exit Function
    Dim TargetListObject As ListObject
    Set TargetListObject = Target

    Dim State As IState
    For Each State In This.ListColumns
        If State.CanApply(TargetListObject) Then
            IState_CanApply = True
            Exit Function
        End If
    Next State
End Function

Private Function IState_Apply(ByVal Target As Object) As Boolean
    If Not TypeOf Target Is ListObject Then Exit Function
    Dim TargetListObject As ListObject
    Set TargetListObject = Target
    
    Dim State As IState
    For Each State In This.ListColumns
        State.Apply TargetListObject
    Next State
    
    Dim Freeze As IState
    Set Freeze = This.Freeze
    Freeze.Apply Nothing
    
    Dim Outline As IState
    Set Outline = This.Outline
    Outline.Apply TargetListObject.Parent
    
    IState_Apply = True
End Function

Private Function ISerializable_Serialize() As String
    Dim Result As String
    Result = Result & This.Name & ":"

    Result = Result & StringtoBase64(This.Caption) & ":"
    Result = Result & ToSerializable(This.Freeze) & ":"
    Result = Result & ToSerializable(This.Outline) & ":"
    
    Dim Child As ISerializable
    For Each Child In This.ListColumns
        Result = Result & Child.Serialize & ";"
    Next Child
    
    Result = Left$(Result, Len(Result) - 1)
    
    ISerializable_Serialize = Result
End Function

Private Function ToSerializable(ByVal State As ISerializable) As String
    ToSerializable = State.Serialize
End Function

Private Function ISerializable_Deserialize(ByVal SerialString As String) As Boolean
    Dim SplitText() As String
    SplitText = Split(SerialString, ":")
    If UBound(SplitText) <> 4 Then Exit Function
    
    This.Name = SplitText(0)
    This.ParentKey = modConstants.LO_KEY_PREFIX & This.Name
    
    This.Caption = Base64toString(SplitText(1))
    
    Dim Freeze As ISerializable
    Set Freeze = This.Freeze
    Freeze.Deserialize SplitText(2)
    
    Dim Outline As ISerializable
    Set Outline = This.Outline
    Outline.Deserialize SplitText(3)
    
    If Not DeserializeChildren(SplitText(4)) Then
        Exit Function
    End If
    
    This.StateType = WORKBOOK_STATE
    
    ISerializable_Deserialize = True
End Function

Private Function DeserializeChildren(ByVal ChildSerials As String) As Boolean
    CollectionHelpers.CollectionClear This.ListColumns
    
    Dim SplitSerials() As String
    SplitSerials = Split(ChildSerials, ";")
    
    Dim i As Long
    For i = 0 To UBound(SplitSerials)
        Dim Child As ISerializable
        Set Child = New ColumnState
        If Not Child.Deserialize(SplitSerials(i)) Then
            Exit Function
        End If
        This.ListColumns.Add Item:=Child, Key:="K" & Format$(i, "000")
        
        Dim listable As IListable
        Set listable = Child
        listable.Key = "K" & Format$(i, "000")
    Next i
    
    DeserializeChildren = True
End Function

Private Function IState_Equals(ByVal RHS As IState) As Boolean
    If RHS Is Nothing Then Exit Function
    If Not TypeOf RHS Is ColumnsState Then Exit Function
    
    IState_Equals = SerializableEquals(Me, RHS)
End Function

Private Function SerializableEquals(ByVal LHS As ISerializable, ByVal RHS As ISerializable) As Boolean
    SerializableEquals = (LHS.Serialize = RHS.Serialize)
End Function

Private Property Get IListable_Caption() As String
    IListable_Caption = This.Caption
End Property

Private Property Get IListable_Key() As String
    IListable_Key = This.Key
End Property

Private Property Let IListable_Key(ByVal vNewValue As String)
    This.Key = vNewValue
End Property

Private Property Get IListable_ParentKey() As String
    IListable_ParentKey = This.ParentKey
End Property

Private Property Let IListable_ParentKey(ByVal vNewValue As String)
    This.ParentKey = vNewValue
End Property

Private Property Get IListable_Visible() As Boolean
    IListable_Visible = This.Visible
End Property

Private Property Let IListable_Visible(ByVal vNewValue As Boolean)
    This.Visible = vNewValue
End Property

' ---
Private Function VisibleCount() As Long
    Dim Child As ColumnState
    For Each Child In This.ListColumns
        If Child.Width > 0 Then VisibleCount = VisibleCount + 1
    Next Child
End Function

Public Sub SetAutoCaption()
    Dim Result As String
    
    Dim MaxChildren As Long
    MaxChildren = This.ListColumns.Count
    If MaxChildren > MAX_CHILDREN_IN_CAPTION Then MaxChildren = MAX_CHILDREN_IN_CAPTION
    
    Dim i As Long
    For i = 1 To MaxChildren
        Dim Child As ColumnState
        Set Child = This.ListColumns.Item(i)
        Dim ChildName As String
        ChildName = Child.Name
        If Len(ChildName) > MAX_COLUMN_CAPTION Then
            ChildName = Left$(ChildName, MAX_COLUMN_CAPTION - 3) & "..."
        End If
        Result = Result & ChildName & ","
    Next i
    If This.ListColumns.Count > MAX_CHILDREN_IN_CAPTION Then
        ' Adds a second comma if there are more elements
        ' e.g. foo..., bar..., xyz...,,
        Result = Result & ","
    End If
    
    If Len(Result) > 2 Then
        Result = Left$(Result, Len(Result) - 1)
    End If
    
    Dim Suffix As String
    Suffix = " (" & VisibleCount() & "/" & This.ListColumns.Count & ")"
    
    Result = Result & Suffix
    
    If This.Match = PARTIAL_MATCH Then
        Result = Result & " (partial)"
    End If
    
    If This.Match = NO_MATCH Then
        Result = Result & " (no match)"
    End If
    
    Result = Result & " (" & Now() & ")"
    
    This.Caption = Result
End Sub
