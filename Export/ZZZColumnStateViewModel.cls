VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZZZColumnStateViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.ColumnState.ViewModel"
Option Explicit

Private Const XML_SETTINGS_NAME As String = "PersistentColumnState"
Private Const SETTING_COLLECTION_NAME As String = "ColumnStates"
Private Const SETTING_LAST_UPDATED As String = "LastUpdated"

Private Const DO_ASSOCIATE_ON_APPLY As String = "DO_ASSOCIATE_ON_APPLY"
Private Const DO_PARTIAL_MATCH As String = "DO_PARTIAL_MATCH"
Private Const DO_PARTIAL_APPLY As String = "DO_PARTIAL_APPLY"
Private Const DO_CLOSE_ON_APPLY As String = "DO_CLOSE_ON_APPLY"
Private Const DO_HIDE_UNMATCHED As String = "DO_HIDE_UNMATCHED"
Private Const DO_FILTER_UNMATCHED As String = "DO_FILTER_UNMATCHED"

Private Type TState
    Workbook As Workbook
    ListObject As ListObject
    SelectedState As ZZZColumnStateTable
    States As Collection
    Settings As ISettingsModel
End Type

Private This As TState

Public Property Get DoFilterUnmatched() As Boolean
    DoFilterUnmatched = This.Settings.Workbook.GetFlag(DO_FILTER_UNMATCHED)
End Property

Public Property Let DoFilterUnmatched(ByVal vNewValue As Boolean)
    This.Settings.Workbook.SetFlag DO_FILTER_UNMATCHED, vNewValue
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
End Property

Public Property Get DoHideUnmatched() As Boolean
    DoHideUnmatched = This.Settings.Workbook.GetFlag(DO_HIDE_UNMATCHED)
End Property

Public Property Let DoHideUnmatched(ByVal vNewValue As Boolean)
    This.Settings.Workbook.SetFlag DO_HIDE_UNMATCHED, vNewValue
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
End Property

Public Property Get DoAssociateOnApply() As Boolean
    DoAssociateOnApply = This.Settings.Workbook.GetFlag(DO_ASSOCIATE_ON_APPLY)
End Property

Public Property Let DoAssociateOnApply(ByVal vNewValue As Boolean)
    This.Settings.Workbook.SetFlag DO_ASSOCIATE_ON_APPLY, vNewValue
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
End Property

Public Property Get DoPartialMatch() As Boolean
    DoPartialMatch = This.Settings.Workbook.GetFlag(DO_PARTIAL_MATCH)
End Property

Public Property Let DoPartialMatch(ByVal vNewValue As Boolean)
    This.Settings.Workbook.SetFlag DO_PARTIAL_MATCH, vNewValue
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
End Property

Public Property Get DoPartialApply() As Boolean
    DoPartialApply = This.Settings.Workbook.GetFlag(DO_PARTIAL_APPLY)
End Property

Public Property Let DoPartialApply(ByVal vNewValue As Boolean)
    This.Settings.Workbook.SetFlag DO_PARTIAL_APPLY, vNewValue
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
End Property

Public Property Get DoCloseOnApply() As Boolean
    DoCloseOnApply = This.Settings.Workbook.GetFlag(DO_CLOSE_ON_APPLY)
End Property

Public Property Let DoCloseOnApply(ByVal vNewValue As Boolean)
    This.Settings.Workbook.SetFlag DO_CLOSE_ON_APPLY, vNewValue
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
End Property

Public Property Get CurrentState() As ZZZColumnStateTable
    Dim Result As ZZZColumnStateTable
    Set Result = New ZZZColumnStateTable
    Result.Load This.ListObject
    Set CurrentState = Result
End Property

Public Property Get SelectedState() As ZZZColumnStateTable
    Set SelectedState = This.SelectedState
End Property

Public Property Get Workbook() As Workbook
    Set Workbook = This.Workbook
End Property

Public Property Get ListObject() As ListObject
    Set ListObject = This.ListObject
End Property

Public Property Get States() As Collection
    Set States = This.States
End Property

Public Property Get IsValid() As Boolean
    IsValid = (Not This.ListObject Is Nothing)
End Property

Private Sub Class_Initialize()
    Set This.States = New Collection
End Sub

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    Set This.Workbook = This.ListObject.Parent.Parent
    
    LoadSettings
    
    Reload
End Sub

Private Sub LoadSettings()
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = XMLSettingsFactory.CreateWorkbookSettings( _
                           Workbook:=This.Workbook, _
                           RootNode:=XML_SETTINGS_NAME)
    
    Dim ASettingsModel As ISettingsModel
    Set ASettingsModel = SettingsModel.Create() _
        .AddWorkbookSettings(WorkbookSettings)
    
    Set This.Settings = ASettingsModel
    
    If This.Settings.Workbook.GetSetting(SETTING_LAST_UPDATED) = Empty Then
        ApplyDefaultSettings
    End If
End Sub

Private Sub ApplyDefaultSettings()
    With This.Settings.Workbook
        .SetFlag DO_ASSOCIATE_ON_APPLY, True
        .SetFlag DO_PARTIAL_MATCH, True
        .SetFlag DO_PARTIAL_APPLY, True
        .SetFlag DO_CLOSE_ON_APPLY, True
        .SetFlag DO_HIDE_UNMATCHED, True
        .SetFlag DO_FILTER_UNMATCHED, False
        .SetSetting SETTING_LAST_UPDATED, CStr(Now())
    End With
End Sub

Private Sub Reload()
    Dim SerializedStates As Collection
    Set SerializedStates = This.Settings.Workbook.GetCollection(SETTING_COLLECTION_NAME)

    CollectionHelpers.CollectionClear This.States
    
    Dim SerializedState As Variant
    For Each SerializedState In SerializedStates
        Dim ColumnState As ZZZColumnStateTable
        Set ColumnState = New ZZZColumnStateTable
        ColumnState.LoadFromBase64 SerializedState, This.ListObject
        This.States.Add ColumnState
    Next SerializedState
End Sub

Public Function Apply() As Boolean
    If This.SelectedState Is Nothing Then Exit Function
    If Not This.SelectedState.CanApply(This.ListObject) Then Exit Function
    
    If Me.DoHideUnmatched Then
        Dim ListColumn As ListColumn
        For Each ListColumn In This.ListObject.ListColumns
            ListColumn.DataBodyRange.EntireColumn.Hidden = True
        Next ListColumn
    End If
    
    This.SelectedState.Apply This.ListObject
    
    If Me.DoHideUnmatched Then
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
    End If
    
    If Me.DoAssociateOnApply And IsStateOrphaned(This.SelectedState) Then
        AssociateSortOrderState This.SelectedState
    End If

    Apply = True
End Function

Public Sub RemoveSelected()
    If This.SelectedState Is Nothing Then Exit Sub
    
    Dim i As Long
    For i = This.States.Count To 1 Step -1
        Dim State As ZZZColumnStateTable
        Set State = This.States.Item(i)
        If State Is This.SelectedState Then
            This.States.Remove i
        End If
    Next i

    Set This.SelectedState = Nothing
    
    SaveStates
End Sub

Public Sub RemoveAll()
    Set This.SelectedState = Nothing
    
    CollectionHelpers.CollectionClear This.States
    
    SaveStates
End Sub

Public Function CanSave() As Boolean
    CanSave = True                               'Not Exists(This.CurrentState)
End Function

Private Function Exists(ByVal ColumnState As ZZZColumnStateTable) As Boolean
    Dim RHSState As ZZZColumnStateTable
    For Each RHSState In This.States
        If ColumnState.Equals(RHSState) Then
            Exists = True
            Exit Function
        End If
    Next RHSState
End Function

Public Function TrySelect(ByVal Key As String) As Boolean
    If Key = "UNSAVED" Then
        Set This.SelectedState = CurrentState
        TrySelect = True
        Exit Function
    End If
    
    Dim ColumnState As ZZZColumnStateTable
    For Each ColumnState In This.States
        If ColumnState.ToBase64 = Key Then
            Set This.SelectedState = ColumnState
            TrySelect = True
            Exit Function
        End If
    Next ColumnState
    
    Set This.SelectedState = Nothing
End Function

Public Sub Save()
    Debug.Assert Not CurrentState Is Nothing
    
    ' TODO Check if it exists first!
    This.States.Add Item:=CurrentState
    
    SaveStates
End Sub

Private Function IsStateOrphaned(ByVal ColumnStateTable As ZZZColumnStateTable) As Boolean
    Dim AllListObjects As Collection
    Set AllListObjects = ListObjectHelpers.GetAllListObjects(This.Workbook)
    
    Dim ListObject As ListObject
    For Each ListObject In AllListObjects
        If ListObject.Name = ColumnStateTable.ListObjectName Then
            Exit Function
        End If
    Next ListObject
    
    IsStateOrphaned = True
End Function

Private Sub AssociateSortOrderState(ByVal State As ZZZColumnStateTable)
    Dim OldKey As String
    OldKey = State.ToBase64
    
    'State.WorksheetName = This.ListObject.Parent.Name
    State.ListObjectName = This.ListObject.Name

    Dim SavedCollection As Collection
    Set SavedCollection = This.Settings.Workbook.GetCollection(SETTING_COLLECTION_NAME)
    
    Dim NewCollection As Collection
    Set NewCollection = New Collection
    
    Dim SerializedState As Variant
    For Each SerializedState In SavedCollection
        If SerializedState = OldKey Then
            NewCollection.Add State.ToBase64
        Else
            NewCollection.Add SerializedState
        End If
    Next SerializedState
    
    This.Settings.Workbook.SetCollection SETTING_COLLECTION_NAME, NewCollection
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
End Sub

Public Sub TryImport(ByVal StateString As String)
    Dim ColumnStateTable As ZZZColumnStateTable
    Set ColumnStateTable = New ZZZColumnStateTable
    ColumnStateTable.LoadFromBase64 StateString, This.ListObject
    
    If Exists(ColumnStateTable) Then
        Exit Sub
    End If
    
    This.States.Add Item:=ColumnStateTable
    
    SaveStates
End Sub

Public Function CanPrune() As Boolean
    ' TODO Why is this a For Next loop?
    Dim i As Long
    For i = This.States.Count To 1 Step -1
        If Not ListObjectHelpers.ListObjectExists(This.Workbook, This.States.Item(i).ListObjectName) Then
            CanPrune = True
            Exit Function
        End If
    Next i
End Function

Public Sub Prune()
    Dim ItemsToDelete As Collection
    Set ItemsToDelete = New Collection
    
    Dim i As Long
    For i = This.States.Count To 1 Step -1
        If Not ListObjectHelpers.ListObjectExists(This.Workbook, This.States.Item(i).ListObjectName) Then
            ItemsToDelete.Add i
        End If
    Next i
    
    Dim Index As Variant
    For Each Index In ItemsToDelete
        This.States.Remove Index
    Next Index
    
    SaveStates
End Sub

Private Sub SaveStates()
    Dim NewCollection As Collection
    Set NewCollection = New Collection
    
    Dim State As ZZZColumnStateTable
    For Each State In This.States
        NewCollection.Add State.ToBase64
    Next State
    
    This.Settings.Workbook.SetCollection SETTING_COLLECTION_NAME, NewCollection
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
End Sub


