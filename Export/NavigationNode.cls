VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NavigationNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("State.Entities")
Option Explicit
Implements IListable

Private Type TState
    Caption As String
    Key As String
    ParentKey As String
    Visible As Boolean
End Type

Private This As TState

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

Public Sub Load(ByVal Object As Object)
    Debug.Assert Not Object Is Nothing
    If TypeOf Object Is ListObject Then LoadListObject Object
    If TypeOf Object Is Worksheet Then LoadWorksheet Object
    If TypeOf Object Is Workbook Then LoadWorkbook Object
End Sub

Private Sub LoadListObject(ByVal ListObject As ListObject)
    This.Caption = ListObject.Name
    This.ParentKey = WB_KEY_PREFIX & ListObject.Parent.Name
End Sub

Private Sub LoadWorksheet(ByVal Worksheet As Worksheet)
    This.Caption = Worksheet.Name
    This.ParentKey = WS_KEY_PREFIX & Worksheet.Parent.Name
End Sub

Private Sub LoadWorkbook(ByVal Workbook As Workbook)
    This.Caption = Workbook.Name
End Sub

Public Function Create(ByVal Object As Object) As IListable
    Dim Result As NavigationNode
    Set Result = New NavigationNode
    Result.Load Object
    Set Create = Result
End Function

