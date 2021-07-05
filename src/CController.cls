VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private UtilsOutlook                   As CUtils_Outlook
Private UtilsFolder                    As CUtils_Folder
Private UtilsRegex                     As CUtils_Regex
Private UtilsArray                     As CUtils_Array
Private ObjOutlook                     As Object
Private ObjModel                       As CModel


Private Sub Class_Initialize()

    Set UtilsOutlook = New CUtils_Outlook
    Set UtilsFolder = New CUtils_Folder
    Set UtilsRegex = New CUtils_Regex
    UtilsRegex.IgnoreCaseSensity = True
    UtilsRegex.HasMultipleLines = True
    Set UtilsArray = New CUtils_Array
    Set ObjOutlook = UtilsOutlook.GetOutlookInstance(True)
    Set ObjModel = New CModel
    
End Sub


Sub DeleteDataFrom(ByVal currentExtraction As CExtraction)
    
    ObjModel.DeleteFromAllTables (currentExtraction.ExtractionName)

End Sub


Function GetMailProperties() As CMailProperties()

    GetMailProperties = ObjModel.GetFromMailPropertiesTable

End Function


Function GetFiltersTypes() As CFilterTypes()

    GetFiltersTypes = ObjModel.GetFromFiltersTypeTable

End Function


Function GetExtractionsNames() As CExtraction()
    
    GetExtractionsNames = ObjModel.GetFromExtractionTable("")

End Function


Function GetMailboxes() As Collection
    
    Set GetMailboxes = New Collection
    
    UtilsOutlook.FillCollectionWithAllSubFolders GetMailboxes, ObjOutlook.Session
    
End Function


Function GetDownloadFolder(ByVal initialPath As String) As String

    GetDownloadFolder = UtilsFolder.GetChosenFolderName(initialPath)

End Function


Function IsDate(ByVal target As String) As Boolean
    
    On Error Resume Next
    target = CDate(target)
    If Err.Number = 0 Then IsDate = True
    On Error GoTo 0

End Function


Function GetMailboxesFrom(ByVal currentExtraction As CExtraction) As CMailbox()

    GetMailboxesFrom = ObjModel.GetFromMailboxTable(currentExtraction.ExtractionName)

End Function


Function GetFiltersFrom(ByVal currentExtraction As CExtraction) As CFilters()

    GetFiltersFrom = ObjModel.GetFromFiltersTable(currentExtraction.ExtractionName)

End Function


Function GetDownloadOptionsFrom(ByVal currentExtraction As CExtraction) As CDownloadOptions
    
    Dim currentDownloadOptions
    
    currentDownloadOptions = ObjModel.GetFromDownloadOptionsTable(currentExtraction.ExtractionName)
    If UBound(currentDownloadOptions) = -1 Then Exit Function
    Set GetDownloadOptionsFrom = currentDownloadOptions(1)

End Function


Sub SaveConfiguration(ByRef ChosenExtraction As CExtraction, _
                      ByRef ChosenMailboxes() As CMailbox, _
                      ByRef ChosenFilters() As CFilters, _
                      ByRef ChosenDownloadOptions As CDownloadOptions)
    
    Dim i                              As Integer
    
    If IsAlreadyInUse(ChosenExtraction) Then
        Err.Raise 513, _
                  "\\OutlookExtractor\CController\SaveConfiguration", _
                  "The given Extraction Name is already in use, please choose other one."
    End If
    
    With ObjModel
        .WriteOnExtractionTable ChosenExtraction
        For i = LBound(ChosenMailboxes) To UBound(ChosenMailboxes)
            .WriteOnMailboxTable ChosenMailboxes(i)
        Next
        For i = LBound(ChosenFilters) To UBound(ChosenFilters)
            .WriteOnFiltersTable ChosenFilters(i)
        Next
       
        .WriteOnDownloadOptionsTable ChosenDownloadOptions
    End With
    
End Sub


Function IsAlreadyInUse(ByRef ChosenExtraction As CExtraction) As Boolean
    
    Dim extractionData()               As CExtraction
    
    extractionData = ObjModel.GetFromExtractionTable(ChosenExtraction.ExtractionName)
    IsAlreadyInUse = UtilsArray.IsInitializedArray(extractionData)
            
End Function


Function Execute(ByRef ChosenMailboxes() As CMailbox, _
                 ByRef ChosenFilters() As CFilters, _
                 ByRef ChosenDownloadOptions As CDownloadOptions)
                 
    Dim mailboxes                      As Collection
    Dim mailbox                        As Object
    Dim mailsIds                       As Collection
    Dim filteredMails                  As Collection
    
    Set mailboxes = GetChosenMailboxes(ChosenMailboxes)
       
    For Each mailbox In mailboxes
        Set mailsIds = UtilsOutlook.GetMailsEntryIdCollectionFromFolder(mailbox)
        FilterMailsIdsWithChosenFilters mailsIds, ChosenFilters
        ApplyDownloadOptionsToFilteredMails ChosenDownloadOptions, mailsIds
    Next

End Function


Private Function GetChosenMailboxes(ByRef ChosenMailboxes() As CMailbox) As Collection

    Dim i                              As Long
    Dim currentFolder                  As Object
    
    Set GetChosenMailboxes = New Collection
    
    For i = LBound(ChosenMailboxes) To UBound(ChosenMailboxes)
        Set currentFolder = ObjOutlook.Session.GetFolderFromID(ChosenMailboxes(i).MailboxItemId)
        GetChosenMailboxes.Add currentFolder
        If CBool(ChosenMailboxes(i).IncludeSubfolders) Then
            UtilsOutlook.FillCollectionWithAllSubFolders GetChosenMailboxes, currentFolder
        End If
    Next
    
End Function


Private Sub FilterMailsIdsWithChosenFilters(ByRef mailsIds As Collection, _
                                            ByRef ChosenFilters() As CFilters)
    
    Dim i                              As Long
    Dim mail                           As Object
    Dim outInstance                    As Object
    
    For i = mailsIds.Count To 1 Step -1
        Set outInstance = UtilsOutlook.GetOutlookInstance(True)
        Set mail = outInstance.Session.GetItemFromID(mailsIds.item(i))
        If DoesNotFitInUserCriteria(mail, ChosenFilters) Then
            mailsIds.Remove i
        End If
    Next
    
End Sub


Private Function DoesNotFitInUserCriteria(ByRef mail As Object, _
                                          ByRef ChosenFilters() As CFilters) As Boolean

    Dim i                              As Integer
    Dim currentMailPropertyText        As String
    Dim currentFilterValue             As String
    
    For i = LBound(ChosenFilters) To UBound(ChosenFilters)
        currentMailPropertyText = GetMailPropertyText(mail, ChosenFilters(i))
        currentFilterValue = CreateSearchString(ChosenFilters(i))
        DoesNotFitInUserCriteria = Not UtilsRegex.HasPattern(currentFilterValue, currentMailPropertyText)
        If DoesNotFitInUserCriteria Then Exit Function
    Next

End Function


Private Function GetMailPropertyText(ByRef mail As Object, _
                                     ByRef currentFilter As CFilters) As String
    
    Dim Property                       As Variant
    Dim joinMultipleProperties         As String
    Dim item                           As Variant
    Dim isMultiProperty                As Boolean
    Dim rightPropertyName              As String
    
    rightPropertyName = GetRightPropertyName(currentFilter)
    
    For Each Property In mail.ItemProperties
        
        If Property.name <> rightPropertyName Then _
           GoTo Continue
        
        On Error Resume Next
        isMultiProperty = (Property.Count > 1)
        On Error GoTo 0
        
        If isMultiProperty Then
            For Each item In Property
                joinMultipleProperties = item.value & joinMultipleProperties & ";"
            Next
            GetMailPropertyText = joinMultipleProperties
        Else
            GetMailPropertyText = Property.value
        End If
        
        Exit Function
Continue:
    Next

End Function


Private Function GetRightPropertyName(ByRef currentFilter As CFilters) As String
                                 
    Dim i                              As Long
    Dim mailPropertiesData()           As CMailProperties
    
    mailPropertiesData = ObjModel.GetFromMailPropertiesTable
    
    For i = LBound(mailPropertiesData) To UBound(mailPropertiesData)
        If currentFilter.MailProperty = mailPropertiesData(i).Property Then
            GetRightPropertyName = mailPropertiesData(i).RespectiveObjectPropertyName
            Exit Function
        End If
    Next

End Function


Private Function CreateSearchString(ByRef ChosenFilters As CFilters) As String

    Dim filterTypesData()              As CFilterTypes
    Dim regexSearchString              As String
    Dim i                              As Integer
    
    filterTypesData = ObjModel.GetFromFiltersTypeTable
    
    For i = LBound(filterTypesData) To UBound(filterTypesData)
        If ChosenFilters.FilterType = filterTypesData(i).TypeName Then
            regexSearchString = filterTypesData(i).RespectiveRegex
            CreateSearchString = Replace(regexSearchString, "dummy", ChosenFilters.FilterValue)
            Exit Function
        End If
    Next
    
End Function


Private Sub ApplyDownloadOptionsToFilteredMails(ByRef ChosenDownloadOptions As CDownloadOptions, _
                                                ByRef mailsIds As Collection)

    Dim mail                           As Variant
    Dim outInstance                    As Object
    Dim currentMailObject              As Object
    Dim targetFolder                   As String
    Dim currentFolder                  As String
    Dim mailPopertiesCollection        As Collection
    Dim mailObj                        As Object
    
    targetFolder = ChosenDownloadOptions.DownloadFolder
    Set outInstance = UtilsOutlook.GetOutlookInstance(True)
    Set mailPopertiesCollection = New Collection
    For Each mail In mailsIds
        Set currentMailObject = outInstance.Session.GetItemFromID(mail)
        If Not IsIncludedOnDates(currentMailObject, ChosenDownloadOptions) Then GoTo Continue
        currentFolder = UtilsFolder.CreateFolder(targetFolder)
        If CBool(ChosenDownloadOptions.GetMailAsFile) Then SaveMailAsFile currentMailObject, currentFolder
        If CBool(ChosenDownloadOptions.GetMailProperties) Then StoreMailPropertiesInCollection currentMailObject, mailPopertiesCollection
        If CBool(ChosenDownloadOptions.DownloadAttachments) Then SaveMailAttachments currentMailObject, currentFolder
Continue:
    Next
    
    If CBool(ChosenDownloadOptions.GetMailProperties) Then
        BuildMailPropertiesReport targetFolder, mailPopertiesCollection
    End If
    
End Sub


Private Function IsIncludedOnDates(ByRef currentMailObject As Object, _
                                   ByRef ChosenDownloadOptions As CDownloadOptions) As Boolean

    Dim before                         As Date
    Dim after                          As Date
    Dim receivedIn                     As Date
    
    before = ChosenDownloadOptions.beforeDate
    after = ChosenDownloadOptions.afterDate
    receivedIn = currentMailObject.ReceivedTime
    
    If before <> CDate(0) And after <> CDate(0) Then
        IsIncludedOnDates = (receivedIn <= before) And (receivedIn >= after)
    ElseIf before = CDate(0) And after <> CDate(0) Then
        IsIncludedOnDates = (receivedIn >= after)
    ElseIf before <> CDate(0) And after = CDate(0) Then
        IsIncludedOnDates = (receivedIn <= before)
    ElseIf before = CDate(0) And after = CDate(0) Then
        IsIncludedOnDates = True
    End If
    
End Function


Private Sub SaveMailAsFile(ByRef mail As Variant, _
                           ByVal targetFolder As String)
    
    Dim fileName                       As String
    
    fileName = targetFolder & "\" & _
               Format(mail.ReceivedTime, "YYYYMMDD_HHMMSS") & _
               Left(mail.Subject, 25) & ".msg"
    mail.SaveAs fileName

End Sub


Private Sub StoreMailPropertiesInCollection(ByRef mail As Variant, _
                                            ByRef targetCollection As Collection)

    Dim allAttachmentsNames            As String
    Dim item                           As Variant
    Dim isHeaderRow                    As Boolean
    
    For Each item In mail.Attachments
        allAttachmentsNames = allAttachmentsNames & item & ";"
    Next
    
    If allAttachmentsNames <> "" Then allAttachmentsNames = Left(allAttachmentsNames, Len(allAttachmentsNames) - 1)
    
    isHeaderRow = (targetCollection.Count = 0)
    With mail
        If isHeaderRow Then
            targetCollection.Add _
        Array("Received by", _
              "Sended by", _
              "Subject", _
              "CC", _
              "BCC", _
              "Body", _
              "Attachments")
        End If
        
        targetCollection.Add _
        Array(.ReceivedByName, _
              .SenderEmailAddress, _
              .Subject, _
              .CC, _
              .BCC, _
              .Body, _
              allAttachmentsNames)
    End With
    
End Sub


Private Sub SaveMailAttachments(ByRef mail As Variant, _
                                ByVal targetFolder As String)

    Dim item                           As Variant
    Dim currentFileName                As String
    
    For Each item In mail.Attachments
        currentFileName = Format(mail.ReceivedTime, "YYYYMMDD_HHMMSS_") & item
        item.SaveAsFile targetFolder & "\" & currentFileName
    Next
    
End Sub


Private Sub BuildMailPropertiesReport(ByVal targetFolder As String, _
                                      ByRef mailPopertiesCollection As Collection)

    Dim i                              As Integer
    Dim item                           As Variant
    Dim lastFilledRow                  As Long
    Dim wkbReport                      As Workbook
    Dim isEmptyReport                  As Boolean
    Dim timestampedFilename            As String
    
    Set wkbReport = Workbooks.Add
    
    For Each item In mailPopertiesCollection
        isEmptyReport = (wkbReport.Sheets(1).Cells(1, 1) = "")
        If isEmptyReport Then
            lastFilledRow = 0
        Else
            lastFilledRow = wkbReport.Sheets(1).Cells(1, 1).CurrentRegion.Rows.Count
        End If
        For i = LBound(item) To UBound(item)
            wkbReport.Sheets(1).Cells(lastFilledRow + 1, i + 1) = item(i)
        Next
    Next
    
    timestampedFilename = Format(Now, "YYYYMMDD_HHMMSS") & "_report.xlsx"
    With wkbReport
        .Sheets(1).Cells.WrapText = False
        .Sheets(1).Cells.ColumnWidth = 50
        .SaveAs targetFolder & "\" & timestampedFilename
        .Close False
    End With
    
End Sub


Function GetFullFolderNameFromId(ByVal folderId As String) As String

    GetFullFolderNameFromId = ObjOutlook.Session.GetFolderFromID(folderId).FolderPath

End Function



