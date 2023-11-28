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

