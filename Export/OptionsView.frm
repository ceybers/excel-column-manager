VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OptionsView 
   Caption         =   "Options"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4830
   OleObjectBlob   =   "OptionsView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OptionsView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("MVVM.ColumnState.Views")
Option Explicit
Implements IView

Private Type TState
    ViewModel As OptionsViewModel
    IsCancelled As Boolean
End Type

Private This As TState

' ---
Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub cboShowOrphanStates_Click()
    TryUpdate DO_SHOW_ORPHAN_STATES, Me.cboShowOrphanStates.Value
End Sub

Private Sub cboAssociateOrphanStates_Click()
    TryUpdate DO_ASSOCIATE_ORPHAN_STATES, Me.cboAssociateOrphanStates.Value
End Sub

Private Sub cboShowPartialMatch_Click()
    TryUpdate DO_SHOW_PARTIAL_MATCH, Me.cboShowPartialMatch.Value
End Sub

Private Sub cboAllowApplyPartialMatch_Click()
    TryUpdate DO_ALLOW_APPLY_PARTIAL_MATCH, Me.cboAllowApplyPartialMatch.Value
End Sub

Private Sub cboShowNonmatchingCols_Click()
    TryUpdate DO_SHOW_NONMATCHING_COLS, Me.cboShowNonmatchingCols.Value
End Sub

Private Sub cboSearchCase_Click()
    TryUpdate DO_SEARCH_CASE, Me.cboSearchCase.Value
End Sub

Private Sub cboSearchWhole_Click()
    TryUpdate DO_SEARCH_WHOLE, Me.cboSearchWhole.Value
End Sub

Private Sub cboCloseOnApply_Click()
    TryUpdate DO_CLOSE_ON_APPLY, Me.cboCloseOnApply.Value
End Sub

Private Sub cmbApply_Click()
    ' NYI
End Sub

Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbOK_Click()
    This.ViewModel.Save
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As Boolean
    Set This.ViewModel = ViewModel
    This.IsCancelled = False
    
    UpdateControls
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub UpdateControls()
    Me.cboShowOrphanStates.Value = This.ViewModel.GetFlag(DO_SHOW_ORPHAN_STATES)
    Me.cboAssociateOrphanStates.Value = This.ViewModel.GetFlag(DO_ASSOCIATE_ORPHAN_STATES)
    Me.cboShowPartialMatch.Value = This.ViewModel.GetFlag(DO_SHOW_PARTIAL_MATCH)
    Me.cboAllowApplyPartialMatch.Value = This.ViewModel.GetFlag(DO_ALLOW_APPLY_PARTIAL_MATCH)
    Me.cboShowNonmatchingCols.Value = This.ViewModel.GetFlag(DO_SHOW_NONMATCHING_COLS)
    Me.cboSearchCase.Value = This.ViewModel.GetFlag(DO_SEARCH_CASE)
    Me.cboSearchWhole.Value = This.ViewModel.GetFlag(DO_SEARCH_WHOLE)
    Me.cboCloseOnApply.Value = This.ViewModel.GetFlag(DO_CLOSE_ON_APPLY)
End Sub

Public Sub TryUpdate(ByVal FlagName As String, ByVal Value As Boolean)
    This.ViewModel.SetFlag FlagName, Value
End Sub

