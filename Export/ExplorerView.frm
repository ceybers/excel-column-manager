VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExplorerView 
   Caption         =   "Column State Management Explorer"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7395
   OleObjectBlob   =   "ExplorerView.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "ExplorerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
'@Folder "MVVM.ColumnState.Views"
Option Explicit
Implements IView

Private Const MSG_TITLE As String = "Column State Manager"
Private Const RESIZE_WIDTH As Long = 480 '380
Private Const RESIZE_HEIGHT As Long = 320 '260
Private Const SEARCH_WATERMARK As String = "Search..."

Private Type TState
    ViewModel As StateManagerViewModel
    IsCancelled As Boolean
End Type

Private This As TState

' Buttons for Actions
Private Sub cmbSave_Click()
    TrySave
End Sub

Private Sub cmbApply_Click()
    TryApply
End Sub

Private Sub cmbExport_Click()
    TryExport
End Sub

Private Sub cmbImport_Click()
    TryImport
End Sub

Private Sub cmbPrune_Click()
    TryPrune
End Sub

Private Sub cmbRemove_Click()
    TryRemove
End Sub

Private Sub cmbRemoveAll_Click()
    TryRemoveAll
End Sub

Private Sub cmbOptions_Click()
    TryShowOptions
End Sub

Private Sub cmbAbout_Click()
    frmAbout.Show
End Sub

' ---
Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub tvStates_NodeClick(ByVal Node As MSComctlLib.Node)
    This.ViewModel.TrySelect Node.Key
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub txtFilterStates_Change()
    If Me.txtFilterStates.Value = SEARCH_WATERMARK Then Exit Sub
    
    This.ViewModel.States.Filter = Me.txtFilterStates.Value
    UpdateListViewLHS
End Sub

Private Sub txtFilterStates_Enter()
    If Me.txtFilterStates.Text = SEARCH_WATERMARK Then
        Me.txtFilterStates.Text = vbNullString
        Me.txtFilterStates.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub txtFilterStates_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtFilterStates.Text = vbNullString Then
        Me.txtFilterStates.Text = SEARCH_WATERMARK
        Me.txtFilterStates.ForeColor = modConstants.GREY_TEXT_COLOR
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub UserForm_Activate()
    Set Me.lblTargetIcon.Picture = Application.CommandBars.GetImageMso("CreateTableInDesignView", 16, 16)
    Set Me.cmbSave.Picture = Application.CommandBars.GetImageMso("DataFormAddRecord", 16, 16)
    Set Me.cmbApply.Picture = Application.CommandBars.GetImageMso("QueryBuilder", 16, 16)
    Set Me.cmbExport.Picture = Application.CommandBars.GetImageMso("Copy", 16, 16)
    Set Me.cmbImport.Picture = Application.CommandBars.GetImageMso("Paste", 16, 16)
    Set Me.cmbPrune.Picture = Application.CommandBars.GetImageMso("TextBoxLinkBreak", 16, 16) 'Prune
    Set Me.cmbRemove.Picture = Application.CommandBars.GetImageMso("TableRowsDelete", 16, 16) ' Remove
    Set Me.cmbRemoveAll.Picture = Application.CommandBars.GetImageMso("TableDelete", 16, 16) ' Remove All
    Set Me.cmbOptions.Picture = Application.CommandBars.GetImageMso("OmsViewAccountSetting", 16, 16)
    Set Me.cmbAbout.Picture = Application.CommandBars.GetImageMso("Help", 16, 16)
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As Boolean
    ResizeWindow
    
    Set This.ViewModel = ViewModel
    This.IsCancelled = False
    
    InitializeControls

    UpdateCurrentState
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub ResizeWindow()
    If Me.Width > RESIZE_WIDTH Then Exit Sub
    If Me.Height > RESIZE_HEIGHT Then Exit Sub
    
    Dim DeltaX As Long
    Dim DeltaX2 As Long
    DeltaX = RESIZE_WIDTH - Me.Width
    DeltaX2 = DeltaX / 2
    Dim DeltaY As Long
    Dim DeltaY2 As Long
    DeltaY = RESIZE_HEIGHT - Me.Height
    DeltaY2 = DeltaY / 2
    
    Me.Width = RESIZE_WIDTH
    Me.Height = RESIZE_HEIGHT
    
    Me.lblStatusBar.Top = Me.lblStatusBar.Top + DeltaY
    Me.lblStatusBar.Width = Me.lblStatusBar.Width + DeltaX
    
    Me.tvStates.Width = Me.tvStates.Width + DeltaX2
    Me.txtFilterStates.Width = Me.txtFilterStates.Width + DeltaX2
    Me.tvStates.Height = Me.tvStates.Height + DeltaY
    
    Me.lvState.Left = Me.lvState.Left + DeltaX2
    Me.lvState.Width = Me.lvState.Width + DeltaX2
    Me.lblSelectedState.Left = Me.lblSelectedState.Left + DeltaX2
    Me.lblSelectedState.Width = Me.lblSelectedState.Width + DeltaX2
    Me.lvState.Height = Me.lvState.Height + DeltaY
End Sub
Private Sub InitializeControls()
    Me.txtFilterStates.ForeColor = modConstants.GREY_TEXT_COLOR
    
    '@Ignore ArgumentWithIncompatibleObjectType
    StatesToTreeView.Initialize Me.tvStates
    '@Ignore ArgumentWithIncompatibleObjectType
    SelectedStateToListView.Initialize Me.lvState
End Sub

Private Sub UpdateButtons()
    Me.cmbSave.Enabled = This.ViewModel.CanSave
    Me.cmbApply.Enabled = This.ViewModel.CanApply
    Me.cmbExport.Enabled = This.ViewModel.CanExport
    Me.cmbPrune.Enabled = This.ViewModel.CanPrune
    Me.cmbRemove.Enabled = This.ViewModel.CanRemove
    Me.cmbRemoveAll.Enabled = This.ViewModel.CanRemoveAll
End Sub

Private Sub UpdateCurrentState()
    Dim ListableState As ColumnsState
    Set ListableState = This.ViewModel.Current.State
    
    Me.cbbTarget.Text = ListableState.Name
End Sub

Private Sub UpdateListViewLHS()
    '@Ignore ArgumentWithIncompatibleObjectType
    StatesToTreeView.Load Me.tvStates, This.ViewModel
    If Me.tvStates.SelectedItem Is Nothing Then Exit Sub
    This.ViewModel.TrySelect Me.tvStates.SelectedItem.Key
End Sub

Private Sub UpdateListViewRHS()
    If This.ViewModel.Selected.State Is Nothing Then
        If Me.tvStates.SelectedItem Is Nothing Then
            Me.lblSelectedState.Caption = vbNullString
        Else
            Me.lblSelectedState.Caption = "Contents of '" & Me.tvStates.SelectedItem.Text & "'"
        End If
    Else
        Me.lblSelectedState.Caption = "Contents of '" & This.ViewModel.Selected.State.Caption & "'"
    End If
    '@Ignore ArgumentWithIncompatibleObjectType
    SelectedStateToListView.Load Me.lvState, This.ViewModel.Selected
End Sub

Private Sub TrySave()
    This.ViewModel.Save
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub TryApply()
    This.ViewModel.Apply
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
    If This.ViewModel.CloseOnApply Then
        Me.Hide
    End If
End Sub

Private Sub TryPrune()
    If vbNo = MsgBox(MSG_PRUNE_STATES, vbExclamation + vbYesNo + vbDefaultButton2, MSG_TITLE) Then
        Exit Sub
    End If
    
    This.ViewModel.Prune
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub TryRemove()
    If vbNo = MsgBox(MSG_REMOVE_STATE, vbExclamation + vbYesNo + vbDefaultButton2, MSG_TITLE) Then
        Exit Sub
    End If
    
    This.ViewModel.Remove
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub TryRemoveAll()
    If vbNo = MsgBox(MSG_REMOVE_STATES, vbExclamation + vbYesNo + vbDefaultButton2, MSG_TITLE) Then
        Exit Sub
    End If
    
    This.ViewModel.RemoveAll
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub TryExport()
    Dim State As ISerializable
    Set State = This.ViewModel.Selected.State
    If State Is Nothing Then Exit Sub
    InputBox "Serial string for selected state", "Export State to Serial String", State.Serialize
End Sub

Private Sub TryImport()
    Dim SerialString As String
    SerialString = InputBox("Serial string for selected state", "Export State to Serial String", _
                            "Table1:Q29sRA==,8,0;Q29sQg==,16,0;Q29sQw==,102,0")
    
    Dim State As IListable
    If This.ViewModel.TryImport(SerialString, State) Then
        MsgBox "Import OK!", vbInformation + vbOKOnly, MSG_TITLE
        UpdateListViewLHS
        
        ' TODO This is a bad idea
        Me.tvStates.Nodes.Item(State.Key).Selected = True ' Simulates the click to update control state
        This.ViewModel.TrySelect State.Key       ' Simulates the click event
        
        UpdateListViewRHS
        UpdateButtons
    Else
        If State Is Nothing Then
            MsgBox "Import FAIL! serial malformed", vbCritical + vbOKOnly, MSG_TITLE
        Else
            MsgBox "Import FAIL! Already exists", vbCritical + vbOKOnly, MSG_TITLE
        End If
    End If
End Sub

Private Sub TryShowOptions()
    'Me.Hide
    This.ViewModel.ShowOptions
    'Me.Show ' Doesn't help
    'UpdateListViewLHS
    'UpdateListViewRHS
    'UpdateButtons
End Sub


