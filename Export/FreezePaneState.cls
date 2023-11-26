VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FreezePaneState"
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

Private Type TState
    Enabled As Boolean
    Row As Long
    Column As Long
End Type

Private This As TState

Public Function ToString() As String
    Dim Result As String
    Result = This.Enabled & ", R" & This.Row & "C" & This.Column
    ToString = Result
End Function

' Requires PredeclaredId annotation
Public Function Create() As IState
    Dim Result As IState
    Set Result = New FreezePaneState
    Result.From Nothing
    Set Create = Result
End Function

Public Sub ApplyTopLeft()
    Dim SerialState As ISerializable
    Set SerialState = New FreezePaneState
    SerialState.Deserialize "-1.1.1"
    
    Dim State As IState
    Set State = SerialState
    State.Apply Nothing
End Sub

Public Sub ApplyNone()
    Dim SerialState As ISerializable
    Set SerialState = New FreezePaneState
    SerialState.Deserialize "0.0.0"
    
    Dim State As IState
    Set State = SerialState
    State.Apply Nothing
End Sub

' IState
Private Function IState_From(ByVal Object As Object) As Boolean
    With This
        .Enabled = ActiveWindow.Split
        .Row = ActiveWindow.SplitRow
        .Column = ActiveWindow.SplitColumn
    End With
    
    IState_From = True
End Function

Private Function IState_CanApply(ByVal Target As Object) As Boolean
    IState_CanApply = True
End Function

Private Function IState_Apply(ByVal Target As Object) As Boolean
    Application.ScreenUpdating = False
    With ActiveWindow
        .SplitRow = This.Row
        .SplitColumn = This.Column
        .FreezePanes = This.Enabled
    End With
    Application.ScreenUpdating = True            'Doesn't seem to help much
    
    IState_Apply = True
End Function

' ISerializable
Private Function ISerializable_Serialize() As String
    Dim Result As String
    
    Result = CStr(CLng(This.Enabled)) & "." & CLng(This.Row) & "." & CLng(This.Column)
    
    ISerializable_Serialize = Result
End Function

Private Function ISerializable_Deserialize(ByVal SerialString As String) As Boolean
    Dim SplitText() As String
    SplitText = Split(SerialString, ".")
    If UBound(SplitText) <> 2 Then Exit Function
    
    With This
        .Enabled = CBool(CLng(SplitText(0)))
        .Row = CLng(SplitText(1))
        .Column = CLng(SplitText(2))
    End With
    
    ISerializable_Deserialize = True
End Function

Private Function IState_Equals(ByVal RHS As IState) As Boolean
    If RHS Is Nothing Then Exit Function
    If Not TypeOf RHS Is ColumnsState Then Exit Function

    IState_Equals = SerializableEquals(Me, RHS)
End Function

Private Function SerializableEquals(ByVal LHS As ISerializable, ByVal RHS As ISerializable) As Boolean
    SerializableEquals = (LHS.Serialize = RHS.Serialize)
End Function

