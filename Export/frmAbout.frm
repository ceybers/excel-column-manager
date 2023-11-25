VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "About Excel Column State Tool"
   ClientHeight    =   1845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3990
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
'@Folder "MVVM.ColumnState.Views"
Option Explicit

Private Sub cmbClose_Click()
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    Dim Picture As IPictureDisp
    Set Picture = Application.CommandBars.GetImageMso("CreateTableInDesignView", 32, 32)
    Set Me.lblPicHeader.Picture = Picture
End Sub

