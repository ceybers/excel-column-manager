Attribute VB_Name = "TestColumnsState2"
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
    Dim State As ColumnsState2
    Dim StateString As String
    
    lc.Range.ColumnWidth = 4
    Set State = ColumnsState2.Create(lo)
    StateString = "Table1 has 3 column(s). ColA.Width = 4, ColB.Width = 8, ColC.Width = 8."
    If StateString <> State.ToString Then
        Err.Description = "Col.width = 4"
        GoTo TestFail
    End If
    
    lc.Range.EntireColumn.Hidden = True
    Set State = ColumnsState2.Create(lo)
    StateString = "Table1 has 3 column(s). ColA.Width = 0, ColB.Width = 8, ColC.Width = 8."
    If StateString <> State.ToString Then
        Err.Description = "Col.width = 0"
        GoTo TestFail
    End If
    
    lc.Range.ColumnWidth = 8
    Set State = ColumnsState2.Create(lo)
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
    
    SerialString = "Table1:Q29sQQ==,4,0;Q29sQg==,8,0;Q29sQw==,8,0"
    lc.Range.ColumnWidth = 4
    Set State = ColumnsState2.Create(lo)
    If SerialString <> State.Serialize Then
        Err.Description = "Serialize 4,0"
        GoTo TestFail
    End If
    
    SerialString = "Table1:Q29sQQ==,0,-1;Q29sQg==,8,0;Q29sQw==,8,0"
    lc.Range.EntireColumn.Hidden = True
    Set State = ColumnsState2.Create(lo)
    If SerialString <> State.Serialize Then
        Err.Description = "Serialize 0,-1"
        GoTo TestFail
    End If
    
    SerialString = "Table1:Q29sQQ==,8,0;Q29sQg==,8,0;Q29sQw==,8,0"
    lc.Range.ColumnWidth = 8
    Set State = ColumnsState2.Create(lo)
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
    
    SerialString = "Table1:Q29sQQ==,0,-1;Q29sQg==,8,0;Q29sQw==,8,0"
    lc.Range.EntireColumn.Hidden = True
    Set State = ColumnsState2.Create(lo)
    
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
    
    SerialString = "Table1:Q29sQQ==,8,0;Q29sQg==,8,0;Q29sQw==,8,0"
    lc.Range.ColumnWidth = 8
    Set State = ColumnsState2.Create(lo)
    
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
    
    SerialString = "Table1:Q29sQQ==,0,-1;Q29sQg==,8,0;Q29sQw==,8,0"
    Set SerialState = New ColumnsState2
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
    
    SerialString = "Table1:Q29sQQ==,8,0;Q29sQg==,8,0;Q29sQw==,8,0"
    Set SerialState = New ColumnsState2
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

