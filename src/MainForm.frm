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
Private Mailboxes As Object


Private Sub UserForm_Initialize()

    Set MainController = New CController
    Set Mailboxes = MainController.GetMailboxes
                    LoadAvailableMailboxes
    LoadMailPropertiesForFiltering
    LoadFilterTypes
'    Windows(ThisWorkbook.Name).Visible = False
'    Application.Visible = False
   
End Sub


Private Sub UserForm_Terminate()

'    Windows(ThisWorkbook.Name).Visible = True
'    Application.Visible = True

End Sub



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

Private Sub LoadAvailableMailboxes()

    Dim box As Object

    With SSupport
        .EraseCurrentMailBoxes
        For Each box In Mailboxes
            MailboxExtractComboBox.AddItem box.FolderPath
            .AddMailbox box.FolderPath, box.entryId, box.storeId
        Next
    End With
    
End Sub


Private Sub AddMailboxButton_Click()
    
    If IncludeSubfoldersYes + IncludeSubfoldersNo = 0 Or MailboxExtractComboBox.Value = "" Then
        MsgBox "No folder or option for subfolder selected. Please, fill this information", vbExclamation
        Exit Sub
    End If
    
    With MailboxList
        .AddItem
        .List(.ListCount - 1, 0) = MailboxExtractComboBox.Value
        If IncludeSubfoldersYes Then
            .List(.ListCount - 1, 1) = IncludeSubfoldersYes.Caption
        Else
            .List(.ListCount - 1, 1) = IncludeSubfoldersNo.Caption
        End If
    End With
    
End Sub


Private Sub RemoveMailboxButton_Click()
    
    On Error Resume Next
    MailboxList.RemoveItem MailboxList.ListIndex

End Sub


Private Sub MailboxList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    With MailboxList
        If .ListIndex = -1 Then Exit Sub
        MailboxExtractComboBox.Value = .List(.ListIndex, 0)
        If .List(.ListIndex, 1) = IncludeSubfoldersYes.Caption Then
            IncludeSubfoldersYes.Value = True
        Else
            IncludeSubfoldersNo.Value = True
        End If
        .RemoveItem MailboxList.ListIndex
    End With
    
End Sub


'========================
' Download Page
'========================

Private Sub AttachFolderButton_Click()

    Dim currentFolder As String
    
    currentFolder = MainController.GetDownloadFolder(FolderStoreFilesTextBox.Value)
    
    If currentFolder <> "" Then FolderStoreFilesTextBox = currentFolder
    
End Sub


Private Sub DownloadAttachmentsCheckBox_Click()

    If DownloadAttachmentsCheckBox Then
        NewestOptionButton.Visible = True
        OldestOptionButton.Visible = True
        NewestOptionButton.Value = True
    Else
        NewestOptionButton.Visible = False
        OldestOptionButton.Visible = False
        NewestOptionButton.Value = False
        OldestOptionButton.Value = False
    End If

End Sub


'========================
'Filters Page
'========================


Private Sub LoadFilterTypes()
    
    Dim item
    
    For Each item In SSupport.GetFilterTypes
        MailPropertyComboBox.AddItem item
    Next

End Sub


Private Sub LoadMailPropertiesForFiltering()
    
    Dim item
    
    For Each item In SSupport.GetMailProperties
        MailPropertyComboBox.AddItem item
    Next

End Sub

