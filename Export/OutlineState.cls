VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutlineState"
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
    Row As Long
    Column As Long
End Type

Private This As TState

Private Property Get IState_StateType() As StateType
    IState_StateType = StateType.UNDEFINED
End Property

Public Function ToString() As String
    Dim Result As String
    Result = "Outline: " & This.Row & "C" & This.Column
    ToString = Result
End Function

Public Function Create(ByVal Worksheet As Worksheet) As IState
    Dim Result As IState
    Set Result = New OutlineState
    Result.From Worksheet
    Set Create = Result
End Function

Private Function IState_From(ByVal Object As Object) As Boolean
    Debug.Assert TypeOf Object Is Worksheet
    Dim Worksheet As Worksheet
    Set Worksheet = Object
    With This
        .Row = Worksheet.Outline.SummaryRow
        .Column = Worksheet.Outline.SummaryColumn
    End With
    
    IState_From = True
End Function

Private Function IState_CanApply(ByVal Target As Object) As Boolean
    If Not (This.Column = -4131 Or This.Column = -4152) Then Exit Function
    If Not (This.Row = 0 Or This.Row = 1) Then Exit Function
    IState_CanApply = TypeOf Target Is Worksheet
End Function

Private Function IState_Apply(ByVal Target As Object) As Boolean
    Debug.Assert TypeOf Target Is Worksheet
    
    Dim Worksheet As Worksheet
    Set Worksheet = Target
    
    If Not IState_CanApply(Target) Then Exit Function
    
    Dim Before As Range
    Set Before = Selection
    
    Dim UsedRange As Range
    Set UsedRange = Target.UsedRange
    Target.Cells(1, UsedRange.Columns.Count + 1).Select
    
    With Worksheet.Outline
        .SummaryRow = This.Row
        .SummaryColumn = This.Column
    End With
    
    Before.Select
    
    IState_Apply = True
End Function

Private Function ISerializable_Serialize() As String
    Dim Result As String
    
    Result = CLng(This.Row) & "." & CLng(This.Column)
    
    ISerializable_Serialize = Result
End Function

Private Function ISerializable_Deserialize(ByVal SerialString As String) As Boolean
    Dim SplitText() As String
    SplitText = Split(SerialString, ".")
    If UBound(SplitText) <> 1 Then Exit Function
    
    With This
        .Row = CLng(SplitText(0))
        .Column = CLng(SplitText(1))
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

