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

Private UtilsOutlook As CUtils_Outlook
Private UtilsFolder As CUtils_Folder
Private UtilsRegex As CUtils_Regex
Private UtilsArray As CUtils_Array
Private ObjOutlook As Object
Private ObjModel As CModel

Private Sub Class_Initialize()

    Set UtilsOutlook = New CUtils_Outlook
    Set UtilsFolder = New CUtils_Folder
    Set UtilsRegex = New CUtils_Regex
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
    
    Dim dummy
    
    On Error Resume Next
        dummy = CDate(target)
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
    
    Dim i As Integer
    
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
    
    Dim extractionData() As CExtraction
    
    extractionData = ObjModel.GetFromExtractionTable(ChosenExtraction.ExtractionName)
    IsAlreadyInUse = UtilsArray.IsInitializedArray(extractionData)
            
End Function


Function Execute(ByRef ChosenMailboxes() As CMailbox, _
                 ByRef ChosenFilters() As CFilters, _
                 ByRef ChosenDownloadOptions As CDownloadOptions)
                 
    Dim mailboxes As Collection
    Dim mailbox As Object
    Dim mailsIds As Collection
    Dim filteredMails As Collection
    
    Set mailboxes = GetChosenMailboxes(ChosenMailboxes)
       
    For Each mailbox In mailboxes
        Set mailsIds = UtilsOutlook.GetMailsEntryIdCollectionFromFolder(mailbox)
'        Set filteredMails = FilterMailsIdsWithChosenFilters(mailsIds, ChosenFilters)
'        ApplyDownloadOptionsToFilteredMails ChosenDownloadOptions, filteredMails
    Next

End Function


Private Function GetChosenMailboxes(ByRef ChosenMailboxes() As CMailbox) As Collection

    Dim i As Long
    Dim currentFolder As Object
    
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
    
    Dim item As Variant
    Dim mail As Object
    
    
    For Each item In mailsIds
        Set mail = UtilsOutlook.GetOutlookInstance.Session.GetItemFromId(item)
        If DoesNotFitInUserCriteria(mail) Then
            mailsIds.Remove item
        End If
    Next
    
End Sub


Private Function DoesNotFitInUserCriteria(ByRef mail As Object, _
                                          ByRef ChosenFilters() As CFilters) As Boolean

    Dim i As Integer
    Dim currentMailPropertyText As String
    Dim currentFilterValue As String
    
    For i = LBound(ChosenFilters) To UBound(ChosenFilters)
        currentMailPropertyText = GetMailPropertyText(mail, ChosenFilters(i))
        currentFilterValue = CreateSearchString(ChosenFilters(i))
        DoesNotFitInUserCriteria = Not UtilsRegex.HasPattern(currentFilerValue, currentMailPropertyText)
        If DoesNotFitInUserCriteria Then Exit Function
    Next

End Function


Private Function GetMailPropertyText(ByRef mail As Object, _
                                     ByRef currentFilter As CFilters) As String
    
    Dim Property As Variant
    Dim joinMultipleProperties As String
    Dim item As Variant
    
    For Each Property In mail.ItemProperties
        If Property.name <> currentFilter.MailProperty Then _
            GoTo Continue
        
        If Property.Count > 1 Then
            For Each item In Property
                joinMultipleProperties = item.value & joinMultipleProperties & ";"
            Next
            GetMailPropertyText = GetMailPropertyText
        Else
            GetMailPropertyText = Property.value
        End If
   
Continue:
    Next

End Function


Private Function CreateSearchString(ByRef ChosenFilters As CFilters) As String

    Dim filterTypesData() As CFilterTypes
    Dim regexSearchString As String
    Dim i As Integer
    
    Set filterstypesdata = ObjModel.GetFromFiltersTypeTable
    
    For i = LBound(filterstypesdata) To UBound(filterstypesdata)
        If ChosenFilters.FilterType = filterstypesdata.TypeName Then
            regexSearchString = filterTypesData.RespectiveRegex
            CreateSearchString = Replace(regexSearchString, "dummy", ChosenFilters.FilterValue)
        End If
    Next
    
End Function


Function GetFullFolderNameFromId(ByVal folderId As String) As String

     GetFullFolderNameFromId = ObjOutlook.Session.GetFolderFromID(folderId).FolderPath

End Function
