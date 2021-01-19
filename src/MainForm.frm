VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Outlook Extractor"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5280
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MainController As CController
'Private Mailboxes As Object


Private Sub UserForm_Initialize()

    Set MainController = New CController
'    Set Mailboxes = MainController.GetMailboxes
                    LoadAvailableMailboxes
'    Windows(ThisWorkbook.Name).Visible = False
'    Application.Visible = False
   
End Sub


Private Sub UserForm_Terminate()

'    Windows(ThisWorkbook.Name).Visible = True
'    Application.Visible = True

End Sub


Private Sub LoadAvailableMailboxes()

    With SSupport
        .EraseCurrentMailBoxes
        
    End With
End Sub


'Private Sub LoadAvailableMailboxes()
'
'    Dim box As Object
'
'    For Each box In Mailboxes
'        MailboxExtractComboBox.AddItem box.Name
'    Next
'
'End Sub


'========================
'Home Page
'========================


Private Sub PreconfiguredExtractionsComboBox_Enter()
    
    LoadPreconfiguredExtractions
    
End Sub


Private Sub LoadPreconfiguredExtractions()
    
    Dim extractions() As Variant
    Dim i
    
    extractions = SMainToolOptions.GetExtractionsNames
    
    For i = LBound(extractions) To UBound(extractions)
        PreconfiguredExtractionsComboBox.AddItem extractions(i)
    Next
    
    
End Sub


'========================
'Mailbox Page
'========================

Private Sub AddMailboxButton_Click()

End Sub


'========================
'Filters Page
'========================


Private Sub FilterTypeComboBox_Enter()
    
    Dim cell As Range
    
    For Each cell In SSupport.Range("FilterTypesTable")
        FilterTypeComboBox.AddItem cell
    Next
    
End Sub
