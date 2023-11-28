Attribute VB_Name = "modDebugResetModel"
'@Folder("ColumnState")
Option Explicit

Public Sub ResetModel()
    Dim Model As StatesModel
    Set Model = New StatesModel
    Model.Load ThisWorkbook
    Model.RemoveAll
    
    Dim State As ISerializable
    Set State = New ColumnsState
    State.Deserialize ("Table1:Q2FwdGlvbg==:0.0.0:1.-4152:Q29sQQ==,8,0,1;Q29sQg==,8,0,1;Q29sQw==,8,0,1")
    Model.Add State
    
    Set State = New ColumnsState
    State.Deserialize ("Table1:Q2FwdGlvbg==:0.0.0:1.-4152:Q29sRA==,8,0,1;Q29sQg==,16,0,1;Q29sQw==,32,0,1")
    Model.Add State
    
    Set State = New ColumnsState
    State.Deserialize ("Table1:Q2FwdGlvbg==:0.0.0:1.-4152:Q29sQQ==,8,0,1;Q29sQg==,0,-1,1;Q29sQw==,3.43,0,1")
    Model.Add State
    
    Set State = New ColumnsState
    State.Deserialize ("Table2:Q2FwdGlvbg==:0.0.0:1.-4152:QUFB,8,0,1;QkJC,0,-1,1;Q0ND,3.43,0,1")
    Model.Add State
    
    Set State = New ColumnsState
    State.Deserialize ("Orphan:Q2FwdGlvbg==:0.0.0:1.-4152:Q29sQQ==,10,0,1;Q29sQg==,20,0,1;Q29sQw==,30,0,1")
    Model.Add State
    
    Model.Save
End Sub
