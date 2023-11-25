VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IListable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder("State.Abstract")
Option Explicit

Public Property Get Caption() As String
End Property

Public Property Get Key() As String
End Property

Public Property Let Key(ByVal vNewValue As String)
End Property

Public Property Get ParentKey() As String
End Property

Public Property Let ParentKey(ByVal vNewValue As String)
End Property

Public Property Get Visible() As Boolean
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
End Property

' Selected
' Checked
' Enabled
' Visible
