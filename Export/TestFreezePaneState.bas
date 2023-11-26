Attribute VB_Name = "TestFreezePaneState"
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
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 0
    ActiveWindow.SplitColumn = 0
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Uncategorized")
Private Sub TestFreezePaneStateToString()
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    Dim State As FreezePaneState
    Set State = FreezePaneState.Create
    If State.ToString <> "False, R0C0" Then
        Err.Description = "ToString no freeze failed"
        GoTo TestFail
    End If
    
    ActiveWindow.SplitRow = 1
    ActiveWindow.SplitColumn = 2
    ActiveWindow.FreezePanes = True
    Set State = FreezePaneState.Create
    If State.ToString <> "True, R1C2" Then
        Err.Description = "ToString with freeze failed"
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
Private Sub TestFreezePaneStateSerialize()
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    Dim State As ISerializable
    Set State = FreezePaneState.Create
    If State.Serialize <> "0.0.0" Then
        Err.Description = "Serialize no freeze failed"
        GoTo TestFail
    End If
    
    ActiveWindow.SplitRow = 1
    ActiveWindow.SplitColumn = 2
    ActiveWindow.FreezePanes = True
    Set State = FreezePaneState.Create
    
    If State.Serialize <> "-1.1.2" Then
        Err.Description = "Serialize with freeze failed"
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
Private Sub TestFreezePaneStateDeserialize()
    On Error GoTo TestFail
    
    'Arrange:
    Dim State As ISerializable
    Dim SerialString As String
    
    'Act:
    Set State = FreezePaneState.Create
    SerialString = "0.0.0"
    If Not State.Deserialize(SerialString) Then
        Err.Description = "Deserialize no freeze failed"
        GoTo TestFail
    End If
    
    ActiveWindow.SplitRow = 1
    ActiveWindow.SplitColumn = 2
    ActiveWindow.FreezePanes = True
    Set State = FreezePaneState.Create
    SerialString = "-1.1.2"
    If Not State.Deserialize(SerialString) Then
        Err.Description = "Deserialize with freeze failed"
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
Private Sub TestFreezePaneStateApply()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SerialState As ISerializable
    Dim SerialString As String
    Dim State As IState
    
    'Act:
    Set SerialState = FreezePaneState.Create
    SerialString = "0.0.0"
    SerialState.Deserialize SerialString
    Set State = SerialState
    If Not State.Apply(Nothing) Then
        Err.Description = "Apply no freeze failed on routine"
        GoTo TestFail
    End If
    If Not ActiveWindow.FreezePanes = False Then
        Err.Description = "Apply no freeze failed on freezepanes property"
        GoTo TestFail
    End If
    
    Set SerialState = FreezePaneState.Create
    SerialString = "-1.1.2"
    SerialState.Deserialize SerialString
    Set State = SerialState
    If Not State.Apply(Nothing) Then
        Err.Description = "Apply with freeze failed on routine"
        GoTo TestFail
    End If
    If Not ActiveWindow.FreezePanes = True Then
        Err.Description = "Apply with freeze failed on freezepanes property"
        GoTo TestFail
    End If
    If Not ActiveWindow.SplitRow = 1 Then
        Err.Description = "Apply with freeze failed on SplitRow property"
        GoTo TestFail
    End If
    If Not ActiveWindow.SplitColumn = 2 Then
        Err.Description = "Apply with freeze failed on SplitColumn property"
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

