VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ZZZColumnStateView 
   Caption         =   "Persistent Column State Tool"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "ZZZColumnStateView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ZZZColumnStateView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder "MVVM.ColumnState.Views"
Option Explicit
Implements IView

Private Const MSG_TITLE As String = "Column State Tool"
Private Const MSG_REMOVE_STATE As String = "Remove this Column State?"
Private Const MSG_REMOVE_ALL_STATES As String = "Remove ALL Column States?"
Private Const MSG_EXPORT_STATE As String = "Column State in Base64 format:"

Private Type TState
    IsCancelled As Boolean
    ViewModel As ZZZColumnStateViewModel
End Type

Private This As TState

Private Sub cboFilterUnmatched_Change()
    This.ViewModel.DoFilterUnmatched = Me.cboFilterUnmatched.Value
    UpdateListView
End Sub

Private Sub cboCloseOnApply_Change()
    This.ViewModel.DoCloseOnApply = Me.cboCloseOnApply.Value
End Sub

Private Sub cboHideUnmatched_Click()
    This.ViewModel.DoHideUnmatched = Me.cboHideUnmatched.Value
End Sub

Private Sub cboPartialApply_Click()
    This.ViewModel.DoPartialApply = Me.cboPartialApply.Value
    UpdateControls
End Sub

Private Sub cboPartialMatch_Click()
    This.ViewModel.DoPartialMatch = Me.cboPartialMatch.Value
    UpdateControls
End Sub

Private Sub cboReassociate_Click()
    This.ViewModel.DoAssociateOnApply = Me.cboReassociate.Value
End Sub

Private Sub cmbImport_Click()
    Dim StateString As String
    StateString = InputBox(MSG_EXPORT_STATE, MSG_TITLE)
    This.ViewModel.TryImport StateString
    UpdateControls
End Sub

Private Sub cmbApply_Click()
    This.ViewModel.Apply
    
    If This.ViewModel.DoCloseOnApply Then
        Me.Hide
    Else
        UpdateControls
    End If
End Sub

Private Sub cmbClose_Click()
    OnCancel
End Sub

Private Sub cmbExport_Click()
    InputBox MSG_EXPORT_STATE, MSG_TITLE, This.ViewModel.SelectedState.ToBase64
End Sub

Private Sub cmbPrune_Click()
    This.ViewModel.Prune
    
    UpdateControls
End Sub

Private Sub cmbRemove_Click()
    If vbNo = MsgBox(MSG_REMOVE_STATE, _
                     vbExclamation + vbYesNo + vbDefaultButton2, _
                     MSG_TITLE) Then
        Exit Sub
    End If
    
    This.ViewModel.RemoveSelected
    
    UpdateControls
End Sub

Private Sub cmbRemoveAll_Click()
    If vbNo = MsgBox(MSG_REMOVE_ALL_STATES, _
                     vbExclamation + vbYesNo + vbDefaultButton2, _
                     MSG_TITLE) Then
        Exit Sub
    End If
    
    This.ViewModel.RemoveAll
    
    UpdateControls
End Sub

Private Sub cmbSave_Click()
    This.ViewModel.Save
    
    UpdateControls
End Sub

Private Sub lblOptionsPicture_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    frmAbout.Show
End Sub

Private Sub tvStates_NodeClick(ByVal Node As MSComctlLib.Node)
    This.ViewModel.TrySelect Node.Key
    
    UpdateListView
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As Boolean
    Set This.ViewModel = ViewModel
    
    InitalizeLabelPictures
    InitalizeFromViewModel
    This.IsCancelled = False
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub InitalizeFromViewModel()
    ZZZColumnStateToTreeView.InitializeTreeView Me.tvStates
    ZZZColumnStateToListView.InitializeListView Me.lvPreview
    
    UpdateControls
    
    InitializeOptions
End Sub

Private Sub UpdateControls()
    UpdateSelectedTable
    UpdateTreeView
    UpdateListView
    Me.cmbClose.SetFocus
    Me.cboPartialApply.Enabled = Me.cboPartialMatch.Value
End Sub

Private Sub UpdateSelectedTable()
    Me.txtTableName.Value = This.ViewModel.CurrentState.ListObjectName
    Me.txtSortOrder.Value = This.ViewModel.CurrentState.GetCaption
    
    Me.cmbSave.Enabled = This.ViewModel.CanSave
    Me.cmbSave.Caption = IIf(This.ViewModel.CanSave, "Save", "Saved")
End Sub

Private Sub UpdateTreeView()
    ZZZColumnStateToTreeView.Load This.ViewModel, Me.tvStates
    
    Me.cmbPrune.Enabled = This.ViewModel.CanPrune
    Me.cmbRemove.Enabled = False
    Me.cmbRemoveAll.Enabled = (This.ViewModel.States.Count > 0)
End Sub

Private Sub UpdateListView()
    ZZZColumnStateToListView.Load This.ViewModel, Me.lvPreview
    
    Me.cmbApply.Caption = "Apply"
    Me.cmbApply.Enabled = False
    Me.cmbExport.Enabled = False
    Me.cmbRemove.Enabled = False
    
    If This.ViewModel.SelectedState Is Nothing Then Exit Sub
    Me.cmbExport.Enabled = True
    
    If This.ViewModel.SelectedState.CanApply(This.ViewModel.ListObject) Then
        Me.cmbApply.Caption = "Apply"
        Me.cmbApply.Enabled = True
    End If
    
    If Not This.ViewModel.DoPartialApply Then
        If This.ViewModel.SelectedState.IsPartialMatch(This.ViewModel.ListObject) Then
            'Me.cmbApply.Caption = "Partial match"
            Me.cmbApply.Enabled = False
        End If
    End If
    
    ' Check if SelectedSortState is already applied as CurrentSortState
    If Not This.ViewModel.CurrentState Is Nothing Then
        If This.ViewModel.SelectedState.Equals(This.ViewModel.CurrentState) Then
            Me.cmbApply.Caption = "Applied"
            Me.cmbApply.Enabled = False
        End If
    End If
    
    Me.cmbRemove.Enabled = True
    
    If Me.tvStates.SelectedItem.Key = "UNSAVED" Then
        Me.cmbRemove.Enabled = False
    End If
End Sub

Private Sub InitializeOptions()
    Me.cboCloseOnApply.Value = This.ViewModel.DoCloseOnApply
    Me.cboPartialApply.Value = This.ViewModel.DoPartialApply
    Me.cboPartialMatch.Value = This.ViewModel.DoPartialMatch
    Me.cboReassociate.Value = This.ViewModel.DoAssociateOnApply
    Me.cboFilterUnmatched.Value = This.ViewModel.DoFilterUnmatched
    Me.cboHideUnmatched = This.ViewModel.DoHideUnmatched
End Sub

Private Sub InitalizeLabelPictures()
    InitalizeLabelPicture Me.lblOptionsPicture, "AdvancedFileProperties"
    InitalizeLabelPicture Me.lblPreviewSortOrderPicture, "FindTaggedNotes"
    InitalizeLabelPicture Me.lblSavedSortOrdersPicture, "SaveSelectionToQuickTablesGallery"
    InitalizeLabelPicture Me.lblSelectedTablePicture, "TableAutoFormat"
End Sub

Private Sub InitalizeLabelPicture(ByVal Label As MSForms.Label, ByVal ImageMsoName As String)
    Set Label.Picture = Application.CommandBars.GetImageMso(ImageMsoName, 24, 24)
End Sub


