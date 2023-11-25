VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISerializable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder "State.Abstract"
Option Explicit

Public Function Serialize() As String
End Function

Public Function Deserialize(ByVal SerialString As String) As Boolean
End Function

