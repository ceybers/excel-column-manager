VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExplorerView 
   Caption         =   "Sample State Management Explorer"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7470
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
'@Folder "MVVM.ColumnState.Views"
Option Explicit
Implements IView

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
    This.ViewModel.Prune
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub cmbRemove_Click()
    TryRemove
End Sub

Private Sub cmbRemoveAll_Click()
    TryRemoveAll
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
    This.ViewModel.States.Filter = Me.txtFilterStates.Value
    UpdateListViewLHS
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub UserForm_Activate()
    Set Me.Label4.Picture = Application.CommandBars.GetImageMso("CreateTableInDesignView", 16, 16)
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
    Set This.ViewModel = ViewModel
    This.IsCancelled = False
    
    InitializeControls

    UpdateCurrentState
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
    
    Me.txtFilterStates.SetFocus
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub InitializeControls()
    StatesToTreeView.Initialize Me.tvStates
    SelectedStateToListView.Initialize Me.lvState
End Sub

Private Sub UpdateButtons()
    Me.cmbSave.Enabled = This.ViewModel.CanSave
    Me.cmbApply.Enabled = This.ViewModel.CanApply
    Me.cmbExport.Enabled = This.ViewModel.CanExport
    Me.cmbPrune.Enabled = This.ViewModel.CanPrune
    Me.cmbRemove.Enabled = This.ViewModel.CanRemove
End Sub

Private Sub UpdateCurrentState()
    Dim ListableState As ColumnsState2
    Set ListableState = This.ViewModel.Current.State
    
    Me.cbbTarget.Text = ListableState.Name
End Sub

Private Sub UpdateListViewLHS()
    StatesToTreeView.Load Me.tvStates, This.ViewModel
    If Me.tvStates.SelectedItem Is Nothing Then Exit Sub
    This.ViewModel.TrySelect Me.tvStates.SelectedItem.Key
End Sub

Private Sub UpdateListViewRHS()
    If This.ViewModel.Selected.State Is Nothing Then
        If Me.tvStates.SelectedItem Is Nothing Then
            Me.lblSelectedState.Caption = ""
        Else
            Me.lblSelectedState.Caption = "Contents of '" & Me.tvStates.SelectedItem.Text & "'"
        End If
    Else
        Me.lblSelectedState.Caption = "Contents of '" & This.ViewModel.Selected.State.Caption & "'"
    End If
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
End Sub

Private Sub TryRemove()
    If vbNo = MsgBox("Are you sure?", vbYesNo + vbDefaultButton2) Then
        Exit Sub
    End If
    
    This.ViewModel.Remove
    UpdateListViewLHS
    UpdateListViewRHS
    UpdateButtons
End Sub

Private Sub TryRemoveAll()
    If vbNo = MsgBox("Are you sure?", vbYesNo + vbDefaultButton2) Then
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

End Sub

