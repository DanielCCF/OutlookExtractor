VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Outlook Extractor"
   ClientHeight    =   5730
   ClientLeft      =   75
   ClientTop       =   285
   ClientWidth     =   5175
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private MainController                 As CController
Private mailboxes                      As Object

Private ChosenExtraction               As CExtraction
Private ChosenMailboxes()              As CMailbox
Private ChosenFilters()                As CFilters
Private ChosenDownloadOptions          As CDownloadOptions

Private ListBoxes                      As Variant
Private CheckBoxesList                 As Variant
Private FlagCheckBoxesList             As Variant
Private ListCommboBoxes                As Variant
Private ListLabels                     As Variant
Private RadioButtonsList               As Variant


Const INVALID_FIELD_COLOR              As Variant = &H6464FF
Const NORMAL_FIELD_COLOR               As Variant = &H80000005
Const BACKGROUND_COLOR                 As Variant = &H80000004


Private Sub UserForm_Initialize()

    
    On Error GoTo ErrorHandling
    
    Set MainController = New CController
    Set mailboxes = MainController.GetMailboxes
    
    Windows(ThisWorkbook.name).Visible = False
    Application.Visible = False
    
    ChangeToLoadingStatus
    
    LoadAvailableMailboxes
        
    ListBoxes = Array(MailboxList, FiltersListBox)
    CheckBoxesList = Array(DownloadAttachmentsCheckBox, GetMailAsFileCheckBox, GetMailPropertiesCheckBox)
    FlagCheckBoxesList = Array(FlagDownloadAttachLabel, FlagGetMailAsFileLabel, FlagGetMailPropertiesLabel)
    ListCommboBoxes = Array(MailboxExtractComboBox, MailPropertyComboBox, FilterTypeComboBox)
    ListLabels = Array(FilterValueTextBox, FolderStoreFilesTextBox, AfterDateTextBox, BeforeDateTextBox)
    RadioButtonsList = Array(IncludeSubfoldersYes, IncludeSubfoldersNo)
    
    LoadMailPropertiesForFiltering
    LoadFilterTypes
    LoadPreconfiguredExtractions
       
    ChangeToReadyStatus
    
    Exit Sub
    
ErrorHandling:

    MsgBox "An error happened during the tool initialization stage" & _
           ", please check if Outlook is available, and if has an " & _
           "account configured. Here is the full error description: " & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Error Description: " & Err.Description & vbNewLine & _
           "Error Source: " & Err.Source & vbNewLine, vbCritical

End Sub


Private Sub UserForm_Terminate()
    
    On Error Resume Next
    Windows(ThisWorkbook.name).Visible = True
    Application.Visible = True
    SSupport.EraseCurrentMailBoxes
    'ThisWorkbook.Close False

End Sub


Private Sub ChangeToLoadingStatus()

    PreconfiguredExtractionsLabel.Visible = False
    PreconfiguredExtractionsComboBox.Visible = False
    ExecuteButton.Visible = False
    DeleteExtractionButton.Visible = False
    SaveButton.Visible = False

    With MultiPage
        .Pages("MailboxPage").Enabled = False
        .Pages("FiltersPage").Enabled = False
        .Pages("DownloadPage").Enabled = False
    End With

End Sub


Private Sub ChangeToReadyStatus()

    PreconfiguredExtractionsLabel.Visible = True
    PreconfiguredExtractionsComboBox.Visible = True
    ExecuteButton.Visible = True
    DeleteExtractionButton.Visible = True
    SaveButton.Visible = True

    With MultiPage
        .Pages("MailboxPage").Enabled = True
        .Pages("FiltersPage").Enabled = True
        .Pages("DownloadPage").Enabled = True
    End With

    LoadingLabel.Visible = False
End Sub


Private Sub GetCurrentUserInput()
    
    Dim i                              As Integer
    
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
        .afterDate = ConvertTextToDate(AfterDateTextBox.value)
        .beforeDate = ConvertTextToDate(BeforeDateTextBox.value)
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
    
    Dim extractions()                  As CExtraction
    Dim i                              As Long
    
    extractions = MainController.GetExtractionsNames
    PreconfiguredExtractionsComboBox.Clear
    PreconfiguredExtractionsComboBox.AddItem ""
    For i = LBound(extractions) To UBound(extractions)
        PreconfiguredExtractionsComboBox.AddItem extractions(i).ExtractionName
    Next
    
    
End Sub


Private Sub DeleteExtractionButton_Click()

    On Error GoTo ErrorHandling
    
    If MsgBox("Are you sure deleting this extraction? It is impossible to revert.", _
              vbYesNo) = vbNo Then Exit Sub
            
    GetCurrentUserInput
    
    MainController.DeleteDataFrom ChosenExtraction
    LoadPreconfiguredExtractions
    
    MsgBox "Data deleted successfully!", vbInformation
    
    Exit Sub
    
ErrorHandling:
    
    MsgBox "An error happened during the deleting" & _
           ". Here is the full error description: " & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Error Description: " & Err.Description & vbNewLine & _
           "Error Source: " & Err.Source & vbNewLine, vbCritical
           
End Sub


Private Sub PreconfiguredExtractionsComboBox_Change()
    
    Dim currentExtractionIsFilled      As Boolean
    
    currentExtractionIsFilled = (PreconfiguredExtractionsComboBox.value <> "")
    
    EraseUIInformation
    
    DeleteExtractionButton.Visible = currentExtractionIsFilled
    DeleteExtractionButton.Enabled = currentExtractionIsFilled
    
    If currentExtractionIsFilled Then BringExtractionInformation
    
End Sub


Sub EraseUIInformation()
    
    Dim i                              As Byte
    
    For i = LBound(ListBoxes) To UBound(ListBoxes)
        ListBoxes(i).Clear
    Next
    
    EraseValueFromUIObjectList ListCommboBoxes
    EraseValueFromUIObjectList ListLabels
    EraseValueFromUIObjectList RadioButtonsList
    EraseValueFromUIObjectList CheckBoxesList

End Sub


Private Sub EraseValueFromUIObjectList(ByVal objList As Variant)

    Dim i                              As Byte
    
    For i = LBound(objList) To UBound(objList)
        On Error Resume Next
        objList(i).value = False
        If Err.Number <> 0 Or objList(i).value = "FALSE" Then _
                                                 objList(i).value = ""
        On Error GoTo 0
    Next

End Sub


Private Sub BringExtractionInformation()
    
    Dim objCurrentExtraction           As CExtraction
    
    Set objCurrentExtraction = New CExtraction
    
    objCurrentExtraction.ExtractionName = PreconfiguredExtractionsComboBox.value
    
    FillMailboxPage objCurrentExtraction
    FillFiltersPage objCurrentExtraction
    FillDownloadPage objCurrentExtraction
    
End Sub


Private Sub FillMailboxPage(ByRef objCurrentExtraction As CExtraction)

    Dim i                              As Integer
    Dim objMailboxes()                 As CMailbox

    objMailboxes = MainController.GetMailboxesFrom(objCurrentExtraction)
    On Error Resume Next
    If UBound(objMailboxes) = -1 Then Exit Sub
    On Error GoTo 0
    
    For i = LBound(objMailboxes) To UBound(objMailboxes)
        With MailboxList
            .AddItem
            .list(.ListCount - 1, 0) = MainController.GetFullFolderNameFromId(objMailboxes(i).MailboxItemId)
            If CBool(objMailboxes(i).IncludeSubfolders) Then
                .list(.ListCount - 1, 1) = "Yes"
            Else
                .list(.ListCount - 1, 1) = "No"
            End If
        End With
    Next

End Sub


Private Sub FillFiltersPage(ByRef objCurrentExtraction As CExtraction)

    Dim i                              As Integer
    Dim objFilters()                   As CFilters
        
    objFilters = MainController.GetFiltersFrom(objCurrentExtraction)
    On Error Resume Next
    If UBound(objFilters) = -1 Then Exit Sub
    On Error GoTo 0
    
    For i = LBound(objFilters) To UBound(objFilters)
        With FiltersListBox
            .AddItem
            .list(.ListCount - 1, 0) = objFilters(i).MailProperty
            .list(.ListCount - 1, 1) = objFilters(i).FilterType
            .list(.ListCount - 1, 2) = objFilters(i).FilterValue
        End With
    Next
    
End Sub


Private Sub FillDownloadPage(ByRef objCurrentExtraction As CExtraction)

    Dim i                              As Integer
    Dim objDownloadOptions             As CDownloadOptions
        
    Set objDownloadOptions = MainController.GetDownloadOptionsFrom(objCurrentExtraction)
    If objDownloadOptions Is Nothing Then Exit Sub
    
    With objDownloadOptions
        If .afterDate <> CDate(0) Then AfterDateTextBox = .afterDate
        If .afterDate <> CDate(0) Then BeforeDateTextBox = .beforeDate
        FolderStoreFilesTextBox = .DownloadFolder
        DownloadAttachmentsCheckBox = CBool(.DownloadAttachments)
        GetMailAsFileCheckBox = CBool(.GetMailAsFile)
        GetMailPropertiesCheckBox = CBool(.GetMailProperties)
    End With
    
End Sub


Private Sub ClearArrayListBoxes(ByVal arrListBoxes As Variant)
    
    Dim i                              As Byte
    
    For i = LBound(arrListBoxes) To UBound(arrListBoxes)
        arrListBoxes(i).Clear
    Next
    
End Sub


Private Sub ExecuteButton_Click()


    On Error GoTo ErrorHandling
    
    If HasEmptyFields Then
        MsgBox "Some fields were not filled, please check the tabs.", vbExclamation
        Exit Sub
    End If
    
    RemoveInvalidFieldIndicator
    
    GetCurrentUserInput
    
    MainController.Execute ChosenMailboxes, ChosenFilters, ChosenDownloadOptions
    
    MsgBox "Execution completed!", vbInformation
    
    Exit Sub
    
ErrorHandling:
    MsgBox "An error happened during the tool execution" & _
           ", please check if all the information was provided" & _
           ". Here is the full error description: " & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Error Description: " & Err.Description & vbNewLine & _
           "Error Source: " & Err.Source & vbNewLine, vbCritical

End Sub


Private Sub SaveButton_Click()

    On Error GoTo ErrorHandling
    
    If HasEmptyFields Then
        MsgBox "Some fields were not filled, please check the tabs.", vbExclamation
        Exit Sub
    End If
    
    RemoveInvalidFieldIndicator
    
    GetCurrentUserInput
    
    RecordAsNewExtraction

    Exit Sub
    
ErrorHandling:
    MsgBox "An error happened during the tool saving stage" & _
           ", please check if all the information was provided" & _
           ". Here is the full error description: " & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Error Description: " & Err.Description & vbNewLine & _
           "Error Source: " & Err.Source & vbNewLine, vbCritical

End Sub


Private Function HasEmptyFields() As Boolean

    Dim i                              As Integer
    Dim downloadOptionsChecked         As Integer

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

    Dim i                              As Integer
    
    FolderStoreFilesTextBox.BackColor = NORMAL_FIELD_COLOR
    
    For i = LBound(ListBoxes) To UBound(ListBoxes)
        ListBoxes(i).BackColor = NORMAL_FIELD_COLOR
    Next
    
    For i = LBound(FlagCheckBoxesList) To UBound(FlagCheckBoxesList)
        FlagCheckBoxesList(i).ForeColor = BACKGROUND_COLOR
    Next
    
End Sub


Private Sub RecordAsNewExtraction()
    
    Dim userChoosedAnOption            As Boolean
    Dim tempError                      As Object
    
    If MsgBox("Do you want to record the current configuration for later use?", vbYesNo) = vbNo Then _
                                                                                           Exit Sub
    
    Set ChosenExtraction = New CExtraction
    Do Until userChoosedAnOption
        ChosenExtraction.ExtractionName = InputBox("Type a name for your extraction")
        If ChosenExtraction.ExtractionName = "" Then
            userChoosedAnOption = UserGaveUpSaving
        ElseIf MainController.IsAlreadyInUse(ChosenExtraction) Then
            FillInputDataWithExtractionName
            userChoosedAnOption = CanOverwrite
        Else
            FillInputDataWithExtractionName
            MainController.SaveConfiguration ChosenExtraction, _
                                             ChosenMailboxes, _
                                             ChosenFilters, _
                                             ChosenDownloadOptions
            userChoosedAnOption = True
            MsgBox "Saved successfully!", vbInformation
        End If
    Loop
    
    LoadPreconfiguredExtractions
    
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
    
    Dim i                              As Integer
    
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

    Dim box                            As Object

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

    Dim currentFolder                  As String
    
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

    Dim i                              As Integer
    
    For i = LBound(FlagCheckBoxesList) To UBound(FlagCheckBoxesList)
        FlagCheckBoxesList(i).ForeColor = BACKGROUND_COLOR
    Next

End Sub


Private Sub AfterDateTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim exactTimeIsSpecified           As Boolean
    
    If AfterDateTextBox.value = "" Then Exit Sub
    
    TreatDateField AfterDateTextBox.Object
    
    exactTimeIsSpecified = (InStr(1, AfterDateTextBox.value, ":") > 1)
    If Not exactTimeIsSpecified Then
        AfterDateTextBox.value = AfterDateTextBox.value & " 00:00:00"
    End If
End Sub


Private Sub BeforeDateTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim userSpecifiedExactTime         As Boolean
    Const A_MINUTE_BEFORE_COMPLETING_HOUR = 0.99999
    
    If BeforeDateTextBox.value = "" Then Exit Sub
    
    userSpecifiedExactTime = (InStr(1, BeforeDateTextBox.value, ":") > 1)
    
    TreatDateField BeforeDateTextBox.Object
    
    If Not userSpecifiedExactTime And BeforeDateTextBox.value <> "" Then _
       BeforeDateTextBox.value = CDate(CDate(BeforeDateTextBox.value) + A_MINUTE_BEFORE_COMPLETING_HOUR)
    
    If AfterDateTextBox <> "" And AfterDateTextBox > BeforeDateTextBox And BeforeDateTextBox <> "" Then
        MsgBox "The final date is lower than start date. Please, fix the dates.", vbExclamation
        BeforeDateTextBox = ""
    End If
    
End Sub


Private Sub TreatDateField(ByRef field As Object)

    If Not MainController.IsDate(field) Or field = "" Or InStr(1, field.value, "/") = 0 Then
        MsgBox "The given value is not a date. Please, insert a valid date.", vbExclamation
        field = ""
    Else
        field = CDate(field)
    End If

End Sub


'========================
'Filters Page
'========================


Private Sub LoadMailPropertiesForFiltering()
    
    Dim i                              As Integer
    Dim MailProperties()               As CMailProperties
    
    MailProperties = MainController.GetMailProperties
    For i = LBound(MailProperties) To UBound(MailProperties)
        MailPropertyComboBox.AddItem MailProperties(i).Property
    Next

End Sub


Private Sub LoadFilterTypes()
    
    Dim i                              As Integer
    Dim FilterTypes()                  As CFilterTypes
    
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



