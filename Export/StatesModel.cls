VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StatesModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "MVVM Model that stores serialized State objects in a Workbook's CustomXML"
'@NoIndent
'@ModuleDescription "MVVM Model that stores serialized State objects in a Workbook's CustomXML"
'@Folder "MVVM.ColumnState.Models"
Option Explicit

Private Const XML_SETTINGS_NAME As String = "PersistentColumnState"
Private Const SETTING_COLLECTION_NAME As String = "ColumnStates"
Private Const SETTING_LAST_UPDATED As String = "LastUpdated"

Private Type TState
    Items As Collection
    Settings As ISettingsModel
End Type

Private This As TState

Public Property Get Items()
    Set Items = This.Items
End Property

'@Description "Loads State objects from persistent serial store."
Public Sub Load(ByVal Workbook As Workbook)
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = XMLSettingsFactory.CreateWorkbookSettings( _
                           Workbook:=Workbook, _
                           RootNode:=XML_SETTINGS_NAME)
    
    Dim ASettingsModel As ISettingsModel
    Set ASettingsModel = SettingsModel.Create() _
        .AddWorkbookSettings(WorkbookSettings)
    
    Set This.Settings = ASettingsModel
    
    'If This.Settings.Workbook.GetSetting(SETTING_LAST_UPDATED) = Empty Then
    '    ApplyDefaultSettings
    'End If
    
    CollectionHelpers.CollectionClear This.Items
    
    Dim SerialStrings As Collection
    Set SerialStrings = This.Settings.Workbook.GetCollection(SETTING_COLLECTION_NAME)
    
    Dim SerialString As Variant
    For Each SerialString In SerialStrings
        LoadOne SerialString
    Next SerialString
End Sub

Private Sub LoadOne(ByVal SerialString As String)
    Dim State As ISerializable
    Set State = New ColumnsState2
    State.Deserialize SerialString
    This.Items.Add State
End Sub

'@Description "Saves State objects into persistent serial store."
Public Sub Save()
    Dim SerialStrings As Collection
    Set SerialStrings = New Collection
    
    Dim Item As ISerializable
    For Each Item In This.Items
        SerialStrings.Add Item.Serialize
    Next Item
    
    This.Settings.Workbook.SetCollection SETTING_COLLECTION_NAME, SerialStrings
End Sub

Private Sub Class_Initialize()
    Set This.Items = New Collection
End Sub

'@Description "Returns true if State object exists in model's collection."
Public Function Exists(ByVal State As IState) As Boolean
    Dim RHSState As IState
    For Each RHSState In This.Items
        If State.Equals(RHSState) Then
            Exists = True
        End If
    Next RHSState
End Function

'@Description "Adds a State object to the model's collection if it is not already there."
Public Sub Add(ByVal State As IState)
    If Not Exists(ColumnsState2) Then
        This.Items.Add State
    End If
End Sub

'@Description "Removes a State object from the model's collection."
Public Sub Remove(ByVal State As IState)
    Dim i As Long
    For i = This.Items.Count To 1 Step -1
        Dim RHSState As IState
        Set RHSState = This.Items(i)
        If State.Equals(RHSState) Then
            This.Items.Remove (i)
            Exit Sub
        End If
    Next i
End Sub

'@Description "Removes all State objects from the model's collection."
Public Sub RemoveAll()
    Dim i
    For i = This.Items.Count To 1 Step -1
        This.Items.Remove (i)
    Next i
End Sub

