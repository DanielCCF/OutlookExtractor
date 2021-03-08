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
Private ListBoxes As Variant
Private CheckBoxesList As Variant
Private FlagCheckBoxesList As Variant
Const INVALID_FIELD_COLOR As Variant = &H6464FF
Const NORMAL_FIELD_COLOR As Variant = &HF0F0FF
Const BACKGROUND_COLOR As Variant = &H80000004


Private Sub UserForm_Initialize()

    Set MainController = New CController
    Set Mailboxes = MainController.GetMailboxes
                    LoadAvailableMailboxes
    ListBoxes = Array(MailboxList, FiltersListBox)
    CheckBoxesList = Array(DownloadAttachmentsCheckBox, GetMailAsFileCheckBox, GetMailPropertiesCheckBox)
    FlagCheckBoxesList = Array(FlagDownloadAttachLabel, FlagGetMailAsFileLabel, FlagGetMailPropertiesLabel)
    LoadMailPropertiesForFiltering
    LoadFilterTypes
    LoadPreconfiguredExtractions
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


Private Sub LoadPreconfiguredExtractions()
    
    Dim extractions() As CExtraction
    Dim i As Long
    
    extractions = MainController.GetExtractionsNames
    
    PreconfiguredExtractionsComboBox.AddItem ""
    For i = LBound(extractions) To UBound(extractions)
        PreconfiguredExtractionsComboBox.AddItem extractions(i).ExtractionName
    Next
    
    
End Sub

Private Sub DeleteExtractionButton_Click()

    If MsgBox("Are you sure deleting this extraction? It will be impossible to revert.") = vbNo Then _
        Exit Sub
        
    
        
End Sub

Private Sub PreconfiguredExtractionsComboBox_Change()
    
    DeleteExtractionButton.Visible = (PreconfiguredExtractionsComboBox.value <> "")
    DeleteExtractionButton.Enabled = (PreconfiguredExtractionsComboBox.value <> "")

End Sub

Private Sub ExecuteButton_Click()
    
    Dim chosenExtraction As CExtraction
    Dim chosenmailboxes As CMailbox
    Dim chosenFilters As CFilters
    Dim chosenDownloadOptions As CDownloadOptions
    
    If HasEmptyFields Then
        MsgBox "Some fields were not filled, please check the tabs.", vbExclamation
        Exit Sub
    End If
    
    RemoveInvalidFieldIndicator
    
    GetCurrentUserInput chosenmailboxes, _
                               chosenFilters, _
                               chosenDownloadOptions
                               
'    RecordAsNewExtraction chosenExtraction, _
'                          chosenmailboxes, _
'                          chosenFilters, _
'                          chosenDownloadOptions
'

End Sub


Private Sub SaveButton_Click()

    If HasEmptyFields Then
        MsgBox "Some fields were not filled, please check the tabs.", vbExclamation
        Exit Sub
    End If
    
    RemoveInvalidFieldIndicator
    
    
End Sub


Private Function HasEmptyFields() As Boolean

    Dim i As Integer
    Dim downloadOptionsChecked As Integer

    With FolderStoreFilesTextBox
        If .value = "" Then
            .BackColor = INVALID_FIELD_COLOR
            HasEmptyFields = True
        End If
    End With
    
    For i = LBound(ListBoxes) To UBound(ListBoxes)
        If ListBoxes(i).ListCount = 0 Then
            ListBoxes(i).BackColor = INVALID_FIELD_COLOR
            HasEmptyFields = True
        End If
    Next
    
    For i = LBound(CheckBoxesList) To UBound(CheckBoxesList)
        downloadOptionsChecked = downloadOptionsChecked + CInt(CheckBoxesList(i).value)
    Next
    If downloadOptionsChecked = 0 Then
        HasEmptyFields = True
        For i = LBound(FlagCheckBoxesList) To UBound(FlagCheckBoxesList)
            FlagCheckBoxesList(i).ForeColor = INVALID_FIELD_COLOR
        Next
    End If
    
End Function


Private Sub RemoveInvalidFieldIndicator()

    Dim i As Integer
    
    FolderStoreFilesTextBox.BackColor = NORMAL_FIELD_COLOR
    
    For i = LBound(ListBoxes) To UBound(ListBoxes)
        ListBoxes.BackColor = NORMAL_FIELD_COLOR
    Next
    
    For i = LBound(FlagCheckBoxesList) To UBound(FlagCheckBoxesList)
        FlagCheckBoxesList.ForeColor = BACKGROUND_COLOR
    Next
    
End Sub


Private Sub RecordAsNewExtraction(ByRef chosenExtraction As CExtraction, _
                                  ByRef chosenmailboxes() As CMailbox, _
                                  ByRef chosenFilters() As CFilters, _
                                  ByRef chosenDownloadOptions As CDownloadOptions)
    
    Dim userChoosedAnOption As Boolean
    Dim ExtractionName As String
    Dim tempError As Object
    
    If MsgBox("Do you want to record the current configuration for later use?", vbYesNo) = vbNo Then _
        Exit Sub
    
    Do Until userChooseAnOption
        ExtractionName = InputBox("Type a name for your extraction")
        If ExtractionName = "" Then
            If MsgBox("The name is empty, do you still want to choose a name?", vbYesNo) = vbNo Then _
                userChoosedAnOption = True
        Else
            MainController.SaveConfiguration chosenExtraction, _
                                             chosenmailboxes, _
                                             chosenFilters, _
                                             chosenDownloadOptions
            userChooseAnOption = True
            RecordAsNewExtraction = True
        End If
    Loop
    
End Sub


Private Sub GetCurrentUserInput(ByRef chosenExtraction As CExtraction, _
                                ByRef chosenmailboxes() As CMailbox, _
                                ByRef chosenFilters() As CFilters, _
                                ByRef chosenDownloadOptions As CDownloadOptions)
    
    Dim i As Integer
    
    chosenExtraction.ExtractionName = PreconfiguredExtractionsComboBox.value
    
    With MailboxList
        For i = 1 To .ListCount
            chosenmailboxes.ExtractionName = chosenExtraction.ExtractionName
            chosenmailboxes.MailboxItemId = SSupport.GetFolderId(.list(i, 0))
            chosenmailboxes.IncludeSubfolders = .list(i, 1)
        Next
    End With
    
    With FiltersListBox
        For i = 1 To .ListCount
            chosenFilters.ExtractionName
            chosenFilters.FilterType
            chosenFilters.FilterValue
            chosenFilters.MailProperty
        Next
    End With

End Sub


'========================
'Mailbox Page
'========================


Private Sub MailboxList_Change()
    
    MailboxExtractComboBox.BackColor = NORMAL_FIELD_COLOR
    
End Sub


Private Sub LoadAvailableMailboxes()

    Dim box As Object

    With SSupport
        .EraseCurrentMailBoxes
        For Each box In Mailboxes
            MailboxExtractComboBox.AddItem box.FolderPath
            .AddMailbox box.FolderPath, box.entryId ', box.storeId
        Next
    End With
    
End Sub


Private Sub EditMailboxButton_Click()

    With MailboxList
        If .ListIndex = -1 Then Exit Sub
        MailboxExtractComboBox.value = .list(.ListIndex, 0)
        If .list(.ListIndex, 1) = IncludeSubfoldersYes.Caption Then
            IncludeSubfoldersYes.value = True
        Else
            IncludeSubfoldersNo.value = True
        End If
        .RemoveItem MailboxList.ListIndex
    End With
    
End Sub


Private Sub AddMailboxButton_Click()
    
    If IncludeSubfoldersYes + IncludeSubfoldersNo = 0 Or MailboxExtractComboBox.value = "" Then
        MsgBox "No folder or option for subfolder selected. Please, fill this information", vbExclamation
        Exit Sub
    End If
    
    With MailboxList
        .AddItem
        .list(.ListCount - 1, 0) = MailboxExtractComboBox.value
        If IncludeSubfoldersYes Then
            .list(.ListCount - 1, 1) = IncludeSubfoldersYes.Caption
        Else
            .list(.ListCount - 1, 1) = IncludeSubfoldersNo.Caption
        End If
    End With
    
End Sub


Private Sub RemoveMailboxButton_Click()
    
    On Error Resume Next
    MailboxList.RemoveItem MailboxList.ListIndex

End Sub


'========================
' Download Page
'========================


Private Sub FolderStoreFilesTextBox_Change()
    
    FolderStoreFilesTextBox.BackColor = NORMAL_FIELD_COLOR

End Sub


Private Sub FiltersListBox_Change()
    
    FiltersListBox.BackColor = NORMAL_FIELD_COLOR
    
End Sub


Private Sub AttachFolderButton_Click()

    Dim currentFolder As String
    
    currentFolder = MainController.GetDownloadFolder(FolderStoreFilesTextBox.value)
    
    If currentFolder <> "" Then FolderStoreFilesTextBox = currentFolder
    
End Sub


Private Sub DownloadAttachmentsCheckBox_Click()
    
    ResetFlagColors
    If DownloadAttachmentsCheckBox Then
        NewestOptionButton.Visible = True
        OldestOptionButton.Visible = True
        NewestOptionButton.value = True
    Else
        NewestOptionButton.Visible = False
        OldestOptionButton.Visible = False
        NewestOptionButton.value = False
        OldestOptionButton.value = False
    End If

End Sub


Private Sub GetMailAsFileCheckBox_Change()

    ResetFlagColors
    
End Sub


Private Sub GetMailPropertiesCheckBox_Change()

    ResetFlagColors
    
End Sub


Private Sub ResetFlagColors()

    Dim i As Integer
    
    For i = LBound(FlagCheckBoxesList) To UBound(FlagCheckBoxesList)
        FlagCheckBoxesList(i).ForeColor = BACKGROUND_COLOR
    Next

End Sub

Private Sub AfterDateTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    TreatDateField AfterDateTextBox.Object
    
End Sub


Private Sub BeforeDateTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    TreatDateField BeforeDateTextBox.Object
    
    If AfterDateTextBox <> "" And AfterDateTextBox > BeforeDateTextBox And BeforeDateTextBox <> "" Then
        MsgBox "The final date is lower than start date. Please, fix the dates.", vbExclamation
        BeforeDateTextBox = ""
    End If
    
End Sub


Private Sub TreatDateField(ByRef field As Object)

    If Not MainController.IsDate(field) And field <> "" Then
        MsgBox "The given value is not a date. Please, insert a valid date.", vbExclamation
        field = ""
    Else
        field = Format(field, "DD/MM/YYYY")
    End If

End Sub


'========================
'Filters Page
'========================


Private Sub LoadFilterTypes()
    
    Dim item
    
    For Each item In SSupport.GetFilterTypes
        FilterTypeComboBox.AddItem item
    Next

End Sub


Private Sub LoadMailPropertiesForFiltering()
    
    Dim item
    
    For Each item In SSupport.GetMailProperties
        MailPropertyComboBox.AddItem item
    Next

End Sub


Private Sub AddFilterButton_Click()

    If MailPropertyComboBox = 0 Or FilterTypeComboBox = 0 Then
        MsgBox "No mail property or filter type was selected. Please, fill this information", vbExclamation
        Exit Sub
    End If
    
    With FiltersListBox
        .AddItem
        .list(.ListCount - 1, 0) = MailPropertyComboBox.value
        .list(.ListCount - 1, 1) = FilterTypeComboBox.value
        .list(.ListCount - 1, 2) = FilterValueTextBox.value
    End With
    
    
End Sub


Private Sub RemoveFilterButton_Click()
    
    On Error Resume Next
    FiltersListBox.RemoveItem FiltersListBox.ListIndex

End Sub


Private Sub EditButton_Click()

    With FiltersListBox
        If .ListIndex = -1 Then Exit Sub
        MailPropertyComboBox = .list(.ListIndex, 0)
        FilterTypeComboBox = .list(.ListIndex, 1)
        FilterValueTextBox = .list(.ListIndex, 2)
        .RemoveItem FiltersListBox.ListIndex
    End With
    
End Sub


Private Sub HomeButton_Click()

    MultiPage.value = 0
    
End Sub
