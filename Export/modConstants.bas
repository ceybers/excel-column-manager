Attribute VB_Name = "modConstants"
'@Folder("ColumnState")
Option Explicit

Public Const GREY_TEXT_COLOR As Long = 12632256  'RGB(192,192,192)
Public Const LO_KEY_PREFIX As String = "lo::"

Public Const MSG_PRUNE_STATES As String = "Prune all orphaned Column States?"
Public Const MSG_REMOVE_STATE As String = "Remove this Column State?"
Public Const MSG_REMOVE_STATES As String = "Remove ALL Column States?"

Public Const ROOT_KEY As String = "::ROOT"
Public Const ORPHAN_KEY As String = "::ORPHAN"
Public Const UNSAVED_KEY As String = "::UNSAVED"
Public Const NO_STATES_KEY As String = "::NOSTATES"
