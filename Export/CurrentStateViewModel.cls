VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CurrentStateViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "ViewModel that holds the current active State of the target."
'@ModuleDescription "ViewModel that holds the current active State of the target."
'@Folder("MVVM.ColumnState.ViewModel")
Option Explicit

Private Type TState
    State As IState
End Type

Private This As TState

Public Property Get State() As IState
    Set State = This.State
End Property

Public Property Set State(ByVal vNewValue As IState)
    Set This.State = vNewValue
End Property

Public Function IsProtected() as Boolean
    If This.State Is Nothing Then Exit Function
    
    IsProtected = IsColumnsStateProtected(This.State)
End Function

Private Function IsColumnsStateProtected(ByVal State as ColumnsState) as Boolean
    If Not TypeOf State is ColumnsState Then Exit Function
    If State.ListObject is Nothing Then Exit Function

    Dim Worksheeet as Worksheet
    Set Worksheet = State.ListObject.Parent

    IsColumnsStateProtected = Worksheet.ProtectContents