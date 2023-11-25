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
    Set Me.CommandButton1.Picture = Application.CommandBars.GetImageMso("DataFormAddRecord", 16, 16)
    Set Me.CommandButton2.Picture = Application.CommandBars.GetImageMso("QueryBuilder", 16, 16)
    Set Me.CommandButton3.Picture = Application.CommandBars.GetImageMso("Copy", 16, 16)
    Set Me.CommandButton4.Picture = Application.CommandBars.GetImageMso("Paste", 16, 16)
    Set Me.CommandButton5.Picture = Application.CommandBars.GetImageMso("TextBoxLinkBreak", 16, 16) 'Prune
    Set Me.CommandButton6.Picture = Application.CommandBars.GetImageMso("TableRowsDelete", 16, 16) ' Remove
    Set Me.CommandButton7.Picture = Application.CommandBars.GetImageMso("TableDelete", 16, 16) ' Remove All
    Set Me.CommandButton8.Picture = Application.CommandBars.GetImageMso("OmsViewAccountSetting", 16, 16)
    Set Me.CommandButton9.Picture = Application.CommandBars.GetImageMso("Help", 16, 16)
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
    'Me.cmbApply.Enabled = 'This.ViewModel.CanApply
    'Me.cmbRemove.Enabled = 'This.ViewModel.CanRemove
End Sub

Private Sub UpdateCurrentState()
    Dim ListableState As IListable
    Set ListableState = This.ViewModel.Current.State
    
    Me.cbbTarget.Text = ListableState.Caption
End Sub

Private Sub UpdateListViewLHS()
    StatesToTreeView.Load Me.tvStates, This.ViewModel.States
    If Me.tvStates.SelectedItem Is Nothing Then Exit Sub
    This.ViewModel.TrySelect Me.tvStates.SelectedItem.Key
End Sub

Private Sub UpdateListViewRHS()
    SelectedStateToListView.Load Me.lvState, This.ViewModel.Selected
End Sub

