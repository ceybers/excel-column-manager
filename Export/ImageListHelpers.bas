Attribute VB_Name = "ImageListHelpers"
'@Folder "MVVM.ColumnState.Helpers"
Option Explicit

Private Const IMAGEMSO_SIZE As Long = 16
Private Const IMAGEMSO_NAMES As String = "FileSaveAsExcelXlsx,CreateTable,AcceptInvitation,DeclineInvitation,SortUp,SortDown,SortDialog,TableInsert,HeaderFooterSheetNameInsert,CancelRequest,SendCopyFlag,TableStyleColumnHeaders,TableStyleRowHeaders,GroupTableDrawBorders,InlineEditMenu,WorkflowPending,FontFillBackColorPicker,TextFillColorPicker,GroupFont,ConditionalFormattingIconSetsGallery,ContentControlDropDownList,ChangeCaseDialogClassic,TableColumnSelect,TableStyleBandedColumns,VisibilityVisible,VisibilityHidden"

Public Function GetImageList() As ImageList
    Dim Result As ImageList
    Set Result = New ImageList
    
    Result.ImageWidth = IMAGEMSO_SIZE
    Result.ImageHeight = IMAGEMSO_SIZE
    
    Dim ImageNameArray() As String
    ImageNameArray = Split(IMAGEMSO_NAMES, ",")
    
    Dim ImageName As Variant
    For Each ImageName In ImageNameArray
        AddImageToImageList Result, ImageName
    Next ImageName
    
    Set GetImageList = Result
End Function

Private Sub AddImageToImageList(ByVal ImageList As ImageList, ByVal ImageMso As String)
    Dim Picture As IPictureDisp
    Set Picture = Application.CommandBars.GetImageMso(ImageMso, IMAGEMSO_SIZE, IMAGEMSO_SIZE)
    ImageList.ListImages.Add Key:=ImageMso, Picture:=Picture
End Sub

