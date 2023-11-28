VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExplorerView 
   Caption         =   "Column State Management Explorer"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7395
   OleObjectBlob   =   "ExplorerView.frx":0000
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
Private Const RESIZE_WIDTH As Long = 640         '380
Private Const RESIZE_HEIGHT As Long = 480        '260
Private Const SEARCH_WATERMARK As String = "Search..."

Private Const MSG_TITLE_EXPORT As String = "Export Column State to Serial String"
Private Const MSG_TITLE_IMPORT As String = "Import Column State from Serial String"
Private Const MSG_EXPORT As String = "This is a Base64 serialized string that represents a Column State."
Private Const MSG_IMPORT As String = "Please input a Base64 serialized string that represents a Column State."
Private Const MSG_IMPORT_SUCCEEDED As String = "Column State imported successfully from Base64 serial string."
Private Const MSG_IMPORT_FAILED_MALFORMED As String = "Column State could not be imported!" & vbCrLf & "Serialized string could be not be deserialized."
Private Const MSG_IMPORT_FAILED_DUPLICATE As String = "Column State was not imported as it already exists!"

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
Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub tvStates_DblClick()
    TryRename
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
    Set Me.lblTargetIcon.Picture = frmPictures.lblTarget.Picture
    Set Me.cmbSave.Picture = frmPictures.lblSave.Picture
    Set Me.cmbApply.Picture = frmPictures.lblApply.Picture
    Set Me.cmbExport.Picture = frmPictures.lblCopy.Picture
    Set Me.cmbImport.Picture = frmPictures.lblPaste.Picture
    Set Me.cmbPrune.Picture = frmPictures.lblPrune.Picture
    Set Me.cmbRemove.Picture = frmPictures.lblRemove.Picture
    Set Me.cmbRemoveAll.Picture = frmPictures.lblRemoveAll.Picture
    Set Me.cmbOptions.Picture = frmPictures.lblOptions.Picture
    Set Me.cmbAbout.Picture = frmPictures.lblHelp.Picture
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

    If This.ViewModel.IsTargetProtected Then
        Me.cbbTarget.Text = Me.cbbTarget.Text & TARGET_LOCKED_SUFFIX
    End If
End Sub

Private Sub UpdateListViewLHS()
    '@Ignore ArgumentWithIncompatibleObjectType
    StatesToTreeView.Load Me.tvStates, This.ViewModel
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
    If This.ViewModel.IsTargetProtected Then
        MsgBox "The table you are trying to update is protected!", vbCritical + vbOKOnly, MSG_TITLE
        Exit Sub
    End If
    
    This.ViewModel.Apply
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
    If This.ViewModel.CloseOnApply Then
        Me.Hide
    End If
End Sub

Private Sub TryPrune()
    If vbNo = MsgBox(MSG_PRUNE_STATES, vbQuestion + vbYesNo + vbDefaultButton2, MSG_TITLE) Then
        Exit Sub
    End If
    
    This.ViewModel.Prune
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub TryRemove()
    If vbNo = MsgBox(MSG_REMOVE_STATE, vbQuestion + vbYesNo + vbDefaultButton2, MSG_TITLE) Then
        Exit Sub
    End If
    
    This.ViewModel.Remove
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub TryRemoveAll()
    If vbNo = MsgBox(MSG_REMOVE_STATES, vbQuestion + vbYesNo + vbDefaultButton2, MSG_TITLE) Then
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
    InputBox MSG_EXPORT, MSG_TITLE_EXPORT, State.Serialize
End Sub

Private Sub TryImport()
    Dim SerialString As String
    SerialString = InputBox(MSG_IMPORT, MSG_TITLE_IMPORT, _
                            vbNullString)        ' TODO Implement watermark text and or example
    
    If SerialString = vbNullString Then Exit Sub

    Dim State As IListable
    If This.ViewModel.TryImport(SerialString, State) Then
        MsgBox MSG_IMPORT_SUCCEEDED, vbInformation + vbOKOnly, MSG_TITLE_IMPORT
        UpdateListViewLHS
        
        ' TODO This is a bad idea
        Me.tvStates.Nodes.Item(State.Key).Selected = True ' Simulates the click to update control state
        This.ViewModel.TrySelect State.Key       ' Simulates the click event
        
        UpdateListViewRHS
        UpdateButtons
    Else
        If State Is Nothing Then
            MsgBox MSG_IMPORT_FAILED_MALFORMED, vbCritical + vbOKOnly, MSG_TITLE_IMPORT
        Else
            MsgBox MSG_IMPORT_FAILED_DUPLICATE, vbCritical + vbOKOnly, MSG_TITLE_IMPORT
        End If
    End If
End Sub

Private Sub TryShowOptions()
    This.ViewModel.ShowOptions
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub TryRename()
    If This.ViewModel.Selected.IsBuiltin Then
        This.ViewModel.Apply
        Exit Sub
    End If
    
    Dim CurrentKey As String
    If Not This.ViewModel.Selected Is Nothing Then
        If Not This.ViewModel.Selected.State Is Nothing Then
            CurrentKey = This.ViewModel.Selected.State.Key
        End If
    End If
    If This.ViewModel.Rename() Then
        UpdateListViewLHS
        UpdateListViewRHS
        If CurrentKey <> vbNullString Then
            ' TODO FIX Bad idea
            Me.tvStates.Nodes.Item(CurrentKey).Selected = True
            This.ViewModel.TrySelect CurrentKey
        End If
    End If
End Sub


