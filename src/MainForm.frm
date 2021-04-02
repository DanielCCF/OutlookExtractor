VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Outlook Extractor"
   ClientHeight    =   5715
   ClientLeft      =   180
   ClientTop       =   705
   ClientWidth     =   5220
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
Private mailboxes As Object

Private ChosenExtraction As CExtraction
Private ChosenMailboxes() As CMailbox
Private ChosenFilters() As CFilters
Private ChosenDownloadOptions As CDownloadOptions

Private ListBoxes As Variant
Private CheckBoxesList As Variant
Private FlagCheckBoxesList As Variant

Const INVALID_FIELD_COLOR As Variant = &H6464FF
Const NORMAL_FIELD_COLOR As Variant = &HF0F0FF
Const BACKGROUND_COLOR As Variant = &H80000004


Private Sub UserForm_Initialize()

    Set MainController = New CController
    Set mailboxes = MainController.GetMailboxes
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
    SSupport.EraseCurrentMailBoxes
    
End Sub


Private Sub GetCurrentUserInput()
    
    Dim i As Integer
    
    Set ChosenExtraction = New CExtraction
    ChosenExtraction.ExtractionName = PreconfiguredExtractionsComboBox.value
    
    With MailboxList
        ReDim ChosenMailboxes(.ListCount - 1)
        For i = 0 To .ListCount - 1
            Set ChosenMailboxes(i) = New CMailbox
            ChosenMailboxes(i).ExtractionName = ChosenExtraction.ExtractionName
            ChosenMailboxes(i).MailboxItemId = SSupport.GetFolderId(.list(i, 0))
            ChosenMailboxes(i).IncludeSubfolders = ("Yes" = .list(i, 1))
        Next
    End With
    
    With FiltersListBox
        ReDim ChosenFilters(.ListCount - 1)
        For i = 0 To .ListCount - 1
            Set ChosenFilters(i) = New CFilters
            ChosenFilters(i).ExtractionName = ChosenExtraction.ExtractionName
            ChosenFilters(i).MailProperty = .list(i, 0)
            ChosenFilters(i).FilterType = .list(i, 1)
            ChosenFilters(i).FilterValue = .list(i, 2)
        Next
    End With
    
    Set ChosenDownloadOptions = New CDownloadOptions
    With ChosenDownloadOptions
        .ExtractionName = ChosenExtraction.ExtractionName
        .DownloadFolder = FolderStoreFilesTextBox.value
        .DownloadAttachments = DownloadAttachmentsCheckBox.value
        .GetMailAsFile = GetMailAsFileCheckBox.value
        .GetMailProperties = GetMailPropertiesCheckBox.value
        .AfterDate = ConvertTextToDate(AfterDateTextBox.value)
        .BeforeDate = ConvertTextToDate(BeforeDateTextBox.value)
    End With
    
End Sub


Private Function ConvertTextToDate(ByVal text As String) As Date

    If text = "" Then
        ConvertTextToDate = CDate(0)
    Else
        ConvertTextToDate = CDate(text)
    End If

End Function


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

    If MsgBox("Are you sure deleting this extraction? It is impossible to revert.") = vbNo Then _
        Exit Sub
        
    MainController.DeleteDataFrom ChosenExtraction
    
    MsgBox "Data deleted successfully!"
    
End Sub

Private Sub PreconfiguredExtractionsComboBox_Change()
    
    DeleteExtractionButton.Visible = (PreconfiguredExtractionsComboBox.value <> "")
    DeleteExtractionButton.Enabled = (PreconfiguredExtractionsComboBox.value <> "")

End Sub

Private Sub ExecuteButton_Click()
    
    Dim ChosenExtraction As CExtraction
    Dim ChosenMailboxes() As CMailbox
    Dim ChosenFilters() As CFilters
    Dim ChosenDownloadOptions As CDownloadOptions
    
    If HasEmptyFields Then
        MsgBox "Some fields were not filled, please check the tabs.", vbExclamation
        Exit Sub
    End If
    
    RemoveInvalidFieldIndicator
    
    GetCurrentUserInput
    
    MainController.Execute ChosenMailboxes, ChosenFilters, ChosenDownloadOptions

End Sub


Private Sub SaveButton_Click()

    If HasEmptyFields Then
        MsgBox "Some fields were not filled, please check the tabs.", vbExclamation
        Exit Sub
    End If
    
    RemoveInvalidFieldIndicator
    
    GetCurrentUserInput
    
    RecordAsNewExtraction
    
    MsgBox "Saved successfully!", vbInformation

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
        ListBoxes(i).BackColor = NORMAL_FIELD_COLOR
    Next
    
    For i = LBound(FlagCheckBoxesList) To UBound(FlagCheckBoxesList)
        FlagCheckBoxesList(i).ForeColor = BACKGROUND_COLOR
    Next
    
End Sub


Private Sub RecordAsNewExtraction()
    
    Dim userChoosedAnOption As Boolean
    '@TODO Remove
    Dim ExtractionName As String
    Dim tempError As Object
    
    If MsgBox("Do you want to record the current configuration for later use?", vbYesNo) = vbNo Then _
        Exit Sub
    
    Set ChosenExtraction = New CExtraction
    Do Until userChoosedAnOption
        ChosenExtraction.ExtractionName = InputBox("Type a name for your extraction")
        If ChosenExtraction.ExtractionName = "" Then
            userChoosedAnOption = UserGaveUpSaving
        ElseIf MainController.IsAlreadyInUse(ChosenExtraction) Then
            userChoosedAnOption = CanOverwrite
        Else
            '@TODO Check whata heck is going on
            FillInputDataWithExtractionName
            MainController.SaveConfiguration ChosenExtraction, _
                                             ChosenMailboxes, _
                                             ChosenFilters, _
                                             ChosenDownloadOptions
            userChoosedAnOption = True
        End If
    Loop
    
End Sub


Private Function UserGaveUpSaving() As Boolean
    
    UserGaveUpSaving = True
    If MsgBox("The name is empty, do you still want to save this configuration?", vbYesNo) = vbYes Then _
        UserGaveUpSaving = False
        
End Function


Private Function CanOverwrite() As Boolean

    If MsgBox("This name was already choosen, " & _
              "do you want to overwrite?", vbYesNo) = vbYes Then
        MainController.DeleteDataFrom ChosenExtraction
        MainController.SaveConfiguration ChosenExtraction, _
                                     ChosenMailboxes, _
                                     ChosenFilters, _
                                     ChosenDownloadOptions
        CanOverwrite = True
    End If

End Function


Private Function FillInputDataWithExtractionName()
    
    Dim i As Integer
    
    With ChosenExtraction
        For i = LBound(ChosenFilters) To UBound(ChosenFilters)
            ChosenFilters(i).ExtractionName = .ExtractionName
        Next
        For i = LBound(ChosenMailboxes) To UBound(ChosenMailboxes)
            ChosenMailboxes(i).ExtractionName = .ExtractionName
        Next
        ChosenDownloadOptions.ExtractionName = .ExtractionName
    End With

End Function


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
        For Each box In mailboxes
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
'Download Page
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


Private Sub LoadMailPropertiesForFiltering()
    
    Dim i As Integer
    Dim MailProperties() As CMailProperties
    
    MailProperties = MainController.GetMailProperties
    For i = LBound(MailProperties) To UBound(MailProperties)
        MailPropertyComboBox.AddItem MailProperties(i).Property
    Next

End Sub


Private Sub LoadFilterTypes()
    
    Dim i As Integer
    Dim FilterTypes() As CFilterTypes
    
    FilterTypes = MainController.GetFiltersTypes
    For i = LBound(FilterTypes) To UBound(FilterTypes)
        FilterTypeComboBox.AddItem FilterTypes(i).TypeName
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
