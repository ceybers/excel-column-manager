VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Stores the state of a ListColumn's width and visibility."
'@ModuleDescription "Stores the state of a ListColumn's width and visibility."
'@PredeclaredId
'@Folder "State.Entities"
Option Explicit
Implements IState
Implements ISerializable
Implements IListable

Private Type TState
    ListColumn As ListColumn
    Name As String
    Width As Double
    Hidden As Boolean
    OutlineLevel As Long
    Orphan As Boolean
    ' IListable
    Key As String
    Visible As Boolean
End Type

Private This As TState

'@DefaultMember
Public Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    Name = This.Name
End Property

Public Property Let Name(ByVal RHS As String)
    This.Name = RHS
End Property

Public Property Get Width() As Double
    Width = This.Width
End Property

Public Property Let Width(ByVal RHS As Double)
    This.Width = RHS
End Property

Public Property Get Hidden() As Boolean
    Hidden = This.Hidden
End Property

Public Property Let Hidden(ByVal RHS As Boolean)
    This.Hidden = RHS
End Property

Public Property Get Orphan() As Boolean
    Orphan = This.Orphan
End Property

Public Property Let Orphan(ByVal RHS As Boolean)
    This.Orphan = RHS
End Property

Public Property Get IsSmallColumn() As Boolean
    IsSmallColumn = False
End Property

Private Property Get IState_StateType() As StateType
    IState_StateType = StateType.UNDEFINED
End Property

Public Function ToString() As String
    ToString = This.Name & ".Width = " & This.Width
End Function

' Requires PredeclaredId
Public Function Create(ByVal ListColumn As ListColumn) As IState
    Dim Result As IState
    Set Result = New ColumnState
    Result.From ListColumn
    Set Create = Result
End Function

' IState
Private Function IState_From(ByVal Object As Object) As Boolean
    If Not TypeOf Object Is ListColumn Then Exit Function
    
    Set This.ListColumn = Object
    
    This.Name = This.ListColumn.Name
    This.Width = This.ListColumn.Range.ColumnWidth
    
    This.Hidden = This.ListColumn.Range.EntireColumn.Hidden
    If This.Hidden Then
        Application.ScreenUpdating = False
        This.ListColumn.Range.EntireColumn.Hidden = False
        This.Width = This.ListColumn.Range.ColumnWidth
        This.ListColumn.Range.EntireColumn.Hidden = True
        Application.ScreenUpdating = True
    End If
    
    This.Orphan = False
    This.OutlineLevel = This.ListColumn.Range.EntireColumn.OutlineLevel
    
    IState_From = True
End Function

Private Function IState_CanApply(ByVal Target As Object) As Boolean
    If Not TypeOf Target Is ListObject Then Exit Function
    
    IState_CanApply = ListColumnHelpers.Exists(Target, This.Name)
End Function

Private Function IState_Apply(ByVal Target As Object) As Boolean
    If Not TypeOf Target Is ListObject Then Exit Function
    
    Dim ListColumn As ListColumn
    If ListColumnHelpers.TryGetListColumn(Target, This.Name, ListColumn) Then
        ListColumn.Range.ColumnWidth = This.Width
        ListColumn.Range.EntireColumn.Hidden = This.Hidden
        ListColumn.Range.EntireColumn.OutlineLevel = This.OutlineLevel
        IState_Apply = True
    End If
End Function

' ISerializable
Private Function ISerializable_Serialize() As String
    Dim Result As String
    Result = StringtoBase64(Name) & "," & CDbl(Width) & "," & CStr(CLng(This.Hidden)) & "," & CLng(This.OutlineLevel)
    ISerializable_Serialize = Result
End Function

Private Function ISerializable_Deserialize(ByVal SerialString As String) As Boolean
    Dim SplitText() As String
    SplitText = Split(SerialString, ",")
    If UBound(SplitText) <> 3 Then Exit Function
    If Not IsNumeric(SplitText(2)) Then Exit Function ' Test for .Hidden
    
    With This
        Set .ListColumn = Nothing
        .Name = Base64toString(SplitText(0))
        .Width = CDbl(SplitText(1))
        .Hidden = CBool(CLng(SplitText(2)))
        .OutlineLevel = CLng(SplitText(3))
    End With
    
    ISerializable_Deserialize = True
End Function

Private Function IState_Equals(ByVal RHS As IState) As Boolean
    If Not TypeOf RHS Is ColumnState Then Exit Function
    Dim State As ColumnState
    Set State = RHS
    IState_Equals = (Me.ToString = State.ToString)
End Function

Private Property Get IListable_Caption() As String
    IListable_Caption = This.Name
End Property

Private Property Get IListable_Key() As String
    IListable_Key = This.Key
End Property

Private Property Let IListable_Key(ByVal vNewValue As String)
    This.Key = vNewValue
End Property

Private Property Get IListable_ParentKey() As String
    IListable_ParentKey = vbNullString
End Property

'@Ignore EmptyMethod
Private Property Let IListable_ParentKey(ByVal vNewValue As String)
End Property

Private Property Get IListable_Visible() As Boolean
    IListable_Visible = This.Visible
End Property

Private Property Let IListable_Visible(ByVal vNewValue As Boolean)
    This.Visible = vNewValue
End Property
