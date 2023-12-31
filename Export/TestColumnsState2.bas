Attribute VB_Name = "TestColumnsState2"
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    
    With lo.ListColumns
        .Item(1).Name = "ColA"
        .Item(1).Range.ColumnWidth = 8
        .Item(1).Name = "ColB"
        .Item(2).Range.ColumnWidth = 8
        .Item(1).Name = "ColC"
        .Item(3).Range.ColumnWidth = 8
    End With
    
    ActiveWindow.FreezePanes = False
    ActiveWindow.Split = False
    ActiveWindow.SplitRow = 0
    ActiveWindow.SplitColumn = 0
    
    lo.HeaderRowRange.Cells.Item(1, 1).Offset(0, lo.ListColumns.Count).Select
    lo.Parent.Outline.SummaryRow = 1
    lo.Parent.Outline.SummaryColumn = -4152
    lo.DataBodyRange.Cells.Item(1, 1).Select
    
    Debug.Assert lo.ListColumns.Count = 3
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    
    With lo.ListColumns
        .Item(1).Name = "ColA"
        .Item(1).Range.ColumnWidth = 8
        .Item(1).Name = "ColB"
        .Item(2).Range.ColumnWidth = 8
        .Item(1).Name = "ColC"
        .Item(3).Range.ColumnWidth = 8
    End With
End Sub

'@TestMethod("Uncategorized")
Private Sub TestColumnsStateToString()
    On Error GoTo TestFail
    
    'Arrange:
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    
    Dim lc As ListColumn
    Set lc = lo.ListColumns.Item(1)
    
    'Act:
    Dim State As ColumnsState
    Dim StateString As String
    
    lc.Range.ColumnWidth = 4
    Set State = ColumnsState.Create(lo)
    StateString = "Table1 has 3 column(s). ColA.Width = 4, ColB.Width = 8, ColC.Width = 8."
    If StateString <> State.ToString Then
        Err.Description = "Col.width = 4"
        GoTo TestFail
    End If
    
    lc.Range.EntireColumn.Hidden = True
    Set State = ColumnsState.Create(lo)
    StateString = "Table1 has 3 column(s). ColA.Width = 4, ColB.Width = 8, ColC.Width = 8."
    If StateString <> State.ToString Then
        Err.Description = "Col.width = 0"
        GoTo TestFail
    End If
    
    lc.Range.ColumnWidth = 8
    Set State = ColumnsState.Create(lo)
    StateString = "Table1 has 3 column(s). ColA.Width = 8, ColB.Width = 8, ColC.Width = 8."
    If StateString <> State.ToString Then
        Err.Description = "Col.width = 8"
        GoTo TestFail
    End If
    
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestColumnsStateSerialize()
    On Error GoTo TestFail
    
    'Arrange:
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    
    Dim lc As ListColumn
    Set lc = lo.ListColumns.Item(1)
    
    'Act:
    Dim State As ISerializable
    
    Dim SerialString As String
    
    SerialString = "Table1:Q29sQSxDb2xCLENvbEMgKDMvMyk=:0.0.0:1.-4152:Q29sQQ==,4,0,1;Q29sQg==,8,0,1;Q29sQw==,8,0,1"
    lc.Range.ColumnWidth = 4
    Set State = ColumnsState.Create(lo)
    If SerialString <> State.Serialize Then
        Err.Description = "Serialize 4,0"
        GoTo TestFail
    End If
    
    SerialString = "Table1:Q2FwdGlvbg==:0.0.0:1.-4152:Q29sQQ==,4,-1,1;Q29sQg==,8,0,1;Q29sQw==,8,0,1"
    lc.Range.EntireColumn.Hidden = True
    Set State = ColumnsState.Create(lo)
    If SerialString <> State.Serialize Then
        Err.Description = "Serialize 4,-1"
        GoTo TestFail
    End If
    
    SerialString = "Table1:Q2FwdGlvbg==:0.0.0:1.-4152:Q29sQQ==,8,0,1;Q29sQg==,8,0,1;Q29sQw==,8,0,1"
    lc.Range.ColumnWidth = 8
    Set State = ColumnsState.Create(lo)
    If SerialString <> State.Serialize Then
        Err.Description = "Serialize 8,0"
        GoTo TestFail
    End If
    
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestColumnsStateDeserialize()
    On Error GoTo TestFail
    
    'Arrange:
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    
    Dim lc As ListColumn
    Set lc = lo.ListColumns.Item(1)
    
    'Act:
    Dim State As ISerializable
    Dim SerialString As String
    
    SerialString = "Table1:Q2FwdGlvbg==:0.0.0:1.-4152:Q29sQQ==,4,-1,1;Q29sQg==,8,0,1;Q29sQw==,8,0,1"
    lc.Range.EntireColumn.Hidden = True
    Set State = ColumnsState.Create(lo)
    
    If Not State.Deserialize(SerialString) Then
        Err.Description = "Deserialize a hidden column - routine failed"
        GoTo TestFail
    End If
    If lc.Range.EntireColumn.Hidden = False Then
        Err.Description = "Deserialize a hidden column -  hidden state failed"
        GoTo TestFail
    End If
    If lc.Range.ColumnWidth <> 0 Then
        Err.Description = "Deserialize a hidden column -  column width failed"
        GoTo TestFail
    End If
    
    SerialString = "Table1:Q2FwdGlvbg==:0.0.0:1.-4152:Q29sQQ==,8,0,1;Q29sQg==,8,0,1;Q29sQw==,8,0,1"
    lc.Range.ColumnWidth = 8
    Set State = ColumnsState.Create(lo)
    
    If Not State.Deserialize(SerialString) Then
        Err.Description = "Deserialize a visible column - routine failed"
        GoTo TestFail
    End If
    If lc.Range.EntireColumn.Hidden = True Then
        Err.Description = "Deserialize a visible column - hidden state failed"
        GoTo TestFail
    End If
    If lc.Range.ColumnWidth <> 8 Then
        Err.Description = "Deserialize a visible column - column width failed"
        GoTo TestFail
    End If
    
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestColumnsStateApply()
    On Error GoTo TestFail
    
    'Arrange:
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    
    Dim lc As ListColumn
    Set lc = lo.ListColumns.Item(1)
    
    'Act:
    Dim SerialState As ISerializable
    Dim SerialString As String
    Dim State As IState
    
    SerialString = "Table1:Q2FwdGlvbg==:0.0.0:1.-4152:Q29sQQ==,4,-1,1;Q29sQg==,8,0,1;Q29sQw==,8,0,1"
    Set SerialState = New ColumnsState
    SerialState.Deserialize SerialString
    Set State = SerialState
    
    If Not State.Apply(lo) Then
        Err.Description = "Apply a hidden column - routine failed"
    End If
    If lc.Range.EntireColumn.Hidden = False Then
        Err.Description = "Apply a hidden column - hidden state failed"
        GoTo TestFail
    End If
    If lc.Range.ColumnWidth <> 0 Then
        Err.Description = "Apply a hidden column - column width failed"
        GoTo TestFail
    End If
    
    SerialString = "Table1:Q2FwdGlvbg==:0.0.0:1.-4152:Q29sQQ==,8,0,1;Q29sQg==,8,0,1;Q29sQw==,8,0,1"
    Set SerialState = New ColumnsState
    SerialState.Deserialize SerialString
    Set State = SerialState
    
    If Not State.Apply(lo) Then
        Err.Description = "Apply a visible column - routine failed"
    End If
    If lc.Range.EntireColumn.Hidden = True Then
        Err.Description = "Apply a visible column - hidden state failed"
        GoTo TestFail
    End If
    If lc.Range.ColumnWidth <> 8 Then
        Err.Description = "Apply a visible column - column width failed"
        GoTo TestFail
    End If
    
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

