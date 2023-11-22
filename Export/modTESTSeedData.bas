Attribute VB_Name = "modTESTSeedData"
'@Folder "ColumnState"
Option Explicit

Private Const XML_SETTINGS_NAME As String = "PersistentColumnState"

Public Sub TESTSeedTestData()
    ThisWorkbook.Worksheets.Item(1).UsedRange.EntireColumn.Hidden = False
    ThisWorkbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit
    
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = XMLSettingsFactory.CreateWorkbookSettings( _
                           Workbook:=ThisWorkbook, _
                           RootNode:=XML_SETTINGS_NAME)
    
    Dim ASettingsModel As ISettingsModel
    Set ASettingsModel = SettingsModel.Create() _
        .AddWorkbookSettings(WorkbookSettings)
      
    WorkbookSettings.Load
    ASettingsModel.Workbook.Reset

    Dim ColumnStates As Collection
    Set ColumnStates = New Collection
    
    With ColumnStates
        .Add Item:="Table1:Zm9vYmFy,42,0;Q29sRQ==,16,0"
        .Add Item:="Table1:Zm9vYmFy,8.43,0;Q29sQg==,8.43,0;Q29sQw==,8.43,0;Q29sRA==,8.43,0;Q29sRQ==,-1,0"
        .Add Item:="Table1:Zm9vYmFy,8.43,-1;Q29sQg==,16,0;Q29sQw==,0,-1;Q29sRA==,0,-1;Q29sRQ==,0,-1"
        .Add Item:="Table2:Zm9vYmFy,8.43,0;Q29sQg==,8.43,0;Q29sQw==,8.43,0;Q29sRA==,8.43,0;Q29sRQ==,8.43,0"
        .Add Item:="Orphan:Zm9vYmFy,8.43,0;Q29sQg==,8.43,0;Q29sQw==,8.43,0;Q29sRA==,8.43,0;Q29sRQ==,8.43,0"
    End With
    
    ASettingsModel.Workbook.SetCollection "ColumnStates", ColumnStates
    
    Set ASettingsModel = Nothing
    Set WorkbookSettings = Nothing
    
    MsgBox "Reset column states to hard-coded test values.", vbInformation + vbOKOnly
End Sub

