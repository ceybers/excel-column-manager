Attribute VB_Name = "StateTypes"
'@Folder("State.Constants")
Option Explicit

Public Enum StateType
    UNDEFINED = 0
    BUILTIN_STATE
    UNSAVED_STATE
    WORKBOOK_STATE
    USER_STATE
    RECENT_SEARCH_STATE
End Enum

