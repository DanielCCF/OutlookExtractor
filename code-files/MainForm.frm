VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12930
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Activate()
    
    Windows(ThisWorkbook.Name).Visible = False
    Application.Visible = False
    
End Sub

Private Sub UserForm_Deactivate()



End Sub

Private Sub UserForm_Terminate()

    Windows(ThisWorkbook.Name).Visible = True
    Application.Visible = True
    
End Sub
