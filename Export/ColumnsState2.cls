VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnsState2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Stores the state of multiple ListColumns from a single ListObject."
'@ModuleDescription "Stores the state of multiple ListColumns from a single ListObject."
'@PredeclaredId
'@Folder("State.Entities")
Option Explicit
Implements IState
Implements ISerializable
Implements IListable

Private Type TState
    ListObject As ListObject
    Name As String
    ListColumns As Collection
    Key As String
    Visible As Boolean
End Type

Private This As TState

Public Property Get Items() As Collection
    Set Items = This.ListColumns
End Property

'@DefaultMember
Public Property Get Name() As String
    Name = This.Name
End Property

Public Function ToString() As String
    Dim Result As String
    Result = This.Name & " has " & This.ListColumns.Count & " column(s). "
    
    Dim State As ColumnState2
    For Each State In This.ListColumns
        Result = Result & State.ToString & ", "
    Next State
    
    Result = Left(Result, Len(Result) - 2) & "."
    ToString = Result
End Function

Private Sub Class_Initialize()
    Set This.ListColumns = New Collection
End Sub

' Requires PredeclaredId annotation
Public Function Create(ByVal ListObject As ListObject) As IState
    Dim Result As IState
    Set Result = New ColumnsState2
    Result.From ListObject
    Set Create = Result
End Function

' IState
Private Function IState_From(ByVal Object As Object) As Boolean
    If Not TypeOf Object Is ListObject Then Exit Function
    
    Set This.ListObject = Object
    This.Name = This.ListObject.Name
    
    Dim ListColumn As ListColumn
    For Each ListColumn In This.ListObject.ListColumns
        This.ListColumns.Add ColumnState2.Create(ListColumn)
    Next ListColumn
    
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
    
    IState_Apply = True
End Function

' ISerializable
Private Function ISerializable_Serialize() As String
    Dim Result As String
    Result = Result & This.Name & ":"
    
    Dim Child As ISerializable
    For Each Child In This.ListColumns
        Result = Result & Child.Serialize & ";"
    Next Child
    
    Result = Left$(Result, Len(Result) - 1)
    
    ISerializable_Serialize = Result
End Function

Private Function ISerializable_Deserialize(ByVal SerialString As String) As Boolean
    Dim SplitText() As String
    SplitText = Split(SerialString, ":")
    If UBound(SplitText) <> 1 Then Exit Function
    
    This.Name = SplitText(0)
    
    DeserializeChildren SplitText(1)
    
    ISerializable_Deserialize = True
End Function

Private Sub DeserializeChildren(ByVal ChildSerials As String)
    CollectionHelpers.CollectionClear This.ListColumns
    
    Dim SplitSerials() As String
    SplitSerials = Split(ChildSerials, ";")
    
    Dim i As Long
    For i = 0 To UBound(SplitSerials)
        Dim Child As ISerializable
        Set Child = New ColumnState2
        Child.Deserialize SplitSerials(i)
        This.ListColumns.Add Item:=Child, Key:="K" & Format(i, "000")
    Next i
End Sub

Private Function IState_Equals(ByVal RHS As IState) As Boolean
    If RHS Is Nothing Then Exit Function
    If Not TypeOf RHS Is ColumnsState2 Then Exit Function
    Dim State As ColumnsState2
    Set State = RHS
    IState_Equals = (Me.ToString = State.ToString)
End Function

Private Property Get IListable_Caption() As String
    IListable_Caption = This.Name & " (" & VisibleCount() & "/" & This.ListColumns.Count & " visible)"
End Property

Private Property Get IListable_Key() As String
    IListable_Key = This.Key
End Property

Private Property Let IListable_Key(ByVal vNewValue As String)
    This.Key = vNewValue
End Property

Private Property Get IListable_Visible() As Boolean
    IListable_Visible = This.Visible
End Property

Private Property Let IListable_Visible(ByVal vNewValue As Boolean)
    This.Visible = vNewValue
End Property

' ---
Private Function VisibleCount() As Long
    Dim Child As ColumnState2
    For Each Child In This.ListColumns
        If Child.Width > 0 Then VisibleCount = VisibleCount + 1
    Next Child
End Function

