VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptionsViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.ColumnState.ViewModel")
Option Explicit

Private Type TState
    Model As StatesModel
End Type

Private This As TState

Public Sub Load(ByVal Model As StatesModel)
    Set This.Model = Model
End Sub

Public Sub Save()
    This.Model.Save
End Sub

Public Function GetFlag(ByVal FlagName As String) As Boolean
    GetFlag = This.Model.Settings.Workbook.GetFlag(FlagName)
End Function

Public Sub SetFlag(ByVal FlagName As String, ByVal Value As Boolean)
    This.Model.Settings.Workbook.SetFlag FlagName, Value
End Sub

