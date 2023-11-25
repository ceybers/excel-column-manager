VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnStateSingle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.ColumnState.Entities"
Option Explicit

Private Type TState
    ListColumn As ListColumn
    Name As String
    Width As Double
    Hidden As Boolean
    Index As Long
    Exists As Boolean
End Type

Private This As TState

Public Property Get ListColumn() As ListColumn
    Set ListColumn = This.ListColumn
End Property

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Get Width() As Double
    Width = This.Width
End Property

Public Property Get Hidden() As Boolean
    Hidden = This.Hidden
End Property

Public Property Get Index() As Long
    Index = This.Index
End Property

Public Property Let Index(ByVal vNewValue As Long)
    This.Index = vNewValue
End Property

Public Property Get Exists() As Boolean
    Exists = This.Exists
End Property

Public Property Get IsValid() As Boolean
    ' This.ListColumn can be Nothing if we haven't (or can't) re-reference it
    IsValid = This.Name <> vbNullString
End Property

Public Sub Load(ByVal ListColumn As ListColumn)
    Set This.ListColumn = ListColumn
    This.Name = ListColumn.Name
    If ListColumn.DataBodyRange.EntireColumn.Hidden = True Then
        This.Hidden = True
        'This.Width = ListColumn.DataBodyRange.EntireColumn.Previous.Width
    Else
        This.Hidden = False
        This.Width = ListColumn.DataBodyRange.ColumnWidth
    End If
    This.Index = ListColumn.Index
    This.Exists = True                           ' Loaded from actual ListColumn
End Sub

Public Function ToString() As String
    If Not IsValid Then Exit Function
    
    Dim Result As String
    Result = Name & " is " & Width & " units wide and hidden = " & Hidden
    ToString = Result
End Function

Public Function ToBase64() As String
    Debug.Assert IsValid
    
    Dim Result As String
    Result = StringtoBase64(Name) & "," & CDbl(Width) & "," & CStr(CLng(Hidden))
    ToBase64 = Result
End Function

Public Sub LoadFromBase64(ByVal Base64Value As String)
    Dim SplitText() As String
    SplitText = Split(Base64Value, ",")
    Debug.Assert UBound(SplitText) = 2
    With This
        Set .ListColumn = Nothing
        .Name = Base64toString(SplitText(0))
        .Width = CDbl(SplitText(1))
        .Hidden = CBool(CLng(SplitText(2)))
    End With
End Sub

Public Sub UpdateExists(ByVal ListObject As ListObject)
    This.Exists = ListObjectHelpers.HasListColumn(ListObject, This.Name)
End Sub

