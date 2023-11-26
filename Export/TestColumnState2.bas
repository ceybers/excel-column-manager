Attribute VB_Name = "TestColumnState2"
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
        .Item(1).Range.ColumnWidth = 8
        .Item(2).Range.ColumnWidth = 8
        .Item(3).Range.ColumnWidth = 8
    End With
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    
    With lo.ListColumns
        .Item(1).Range.ColumnWidth = 8
        .Item(2).Range.ColumnWidth = 8
        .Item(3).Range.ColumnWidth = 8
    End With
End Sub

'@TestMethod("Uncategorized")
Private Sub TestColumnStateToString()
    On Error GoTo TestFail
    
    'Arrange:
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    
    Dim lc As ListColumn
    Set lc = lo.ListColumns.Item(1)
    
    'Act:
    Dim State As ColumnState
    Dim StateString As String
    
    lc.Range.ColumnWidth = 4
    Set State = ColumnState.Create(lc)
    StateString = "ColA.Width = 4"
    If StateString <> State.ToString Then
        Err.Description = "Col.width = 4"
        GoTo TestFail
    End If
    
    lc.Range.EntireColumn.Hidden = True
    Set State = ColumnState.Create(lc)
    StateString = "ColA.Width = 4"
    If StateString <> State.ToString Then
        Err.Description = "Col.Hidden = True"
        GoTo TestFail
    End If
    
    lc.Range.ColumnWidth = 8
    Set State = ColumnState.Create(lc)
    StateString = "ColA.Width = 8"
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
Private Sub TestColumnStateSerialize()
    On Error GoTo TestFail
    
    'Arrange:
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    
    Dim lc As ListColumn
    Set lc = lo.ListColumns.Item(1)
    
    'Act:
    Dim State As ISerializable
    
    Dim SerialString As String
    
    SerialString = "Q29sQQ==,4,0,1"
    lc.Range.ColumnWidth = 4
    Set State = ColumnState.Create(lc)
    If SerialString <> State.Serialize Then
        Err.Description = "Serialize 4,0"
        GoTo TestFail
    End If
    
    SerialString = "Q29sQQ==,4,-1,1"
    lc.Range.EntireColumn.Hidden = True
    Set State = ColumnState.Create(lc)
    If SerialString <> State.Serialize Then
        Err.Description = "Serialize 4,-1"
        GoTo TestFail
    End If
    
    SerialString = "Q29sQQ==,8,0,1"
    lc.Range.ColumnWidth = 8
    Set State = ColumnState.Create(lc)
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
Private Sub TestColumnStateDeserialize()
    On Error GoTo TestFail
    
    'Arrange:
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    
    Dim lc As ListColumn
    Set lc = lo.ListColumns.Item(1)
    
    'Act:
    Dim State As ISerializable
    Dim SerialString As String
    
    SerialString = "Q29sQQ==,4,-1,1"
    lc.Range.EntireColumn.Hidden = True
    Set State = ColumnState.Create(lc)
    
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
    
    SerialString = "Q29sQQ==,8,0,1"
    lc.Range.ColumnWidth = 8
    Set State = ColumnState.Create(lc)
    
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
Private Sub TestColumnStateApply()
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
    
    Set SerialState = New ColumnState
    SerialString = "Q29sQQ==,0,-1,1"
    SerialState.Deserialize SerialString
    Set State = SerialState
    
    If Not State.Apply(lo) Then                  ' ColumnState.Apply must be applied to ListObject!
        Err.Description = "Apply a hidden column - routine failed"
    End If
    If lc.Range.EntireColumn.Hidden = False Then
        Err.Description = "Apply a hidden column -  hidden state failed"
        GoTo TestFail
    End If
    If lc.Range.ColumnWidth <> 0 Then
        Err.Description = "Apply a hidden column -  column width failed"
        GoTo TestFail
    End If
    
    Set SerialState = New ColumnState
    SerialString = "Q29sQQ==,8,0,1"
    SerialState.Deserialize SerialString
    Set State = SerialState
    
    If Not State.Apply(lo) Then                  ' ColumnState.Apply must be applied to ListObject!
        Err.Description = "Apply a visible column - routine failed"
    End If
    If lc.Range.EntireColumn.Hidden = True Then
        Err.Description = "Apply a visible column -  hidden state failed"
        GoTo TestFail
    End If
    If lc.Range.ColumnWidth <> 8 Then
        Err.Description = "Apply a visible column -  column width failed"
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

