VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder "State.Abstract"
Option Explicit

Public Function From(ByVal Target As Object) As Boolean
End Function

Public Function CanApply(ByVal Target As Object) As Boolean
End Function

Public Function Apply(ByVal Target As Object) As Boolean
End Function

Public Function Equals(ByVal RHS As IState) As Boolean
End Function

