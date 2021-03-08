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
Private ObjOutlook As Object
Private ObjModel As CModel

Private Sub Class_Initialize()

    Set UtilsOutlook = New CUtils_Outlook
    Set UtilsFolder = New CUtils_Folder
    Set ObjOutlook = UtilsOutlook.GetOutlookInstance(True)
    Set ObjModel = New CModel
    
End Sub


Sub DeleteDataFrom(ByVal currentExtraction As CExtraction)
    
    ObjModel.DeleteFromAllTables (currentExtraction.ExtractionName)

End Sub

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

    GetDownloadOptionsFrom = ObjModel.GetFromDownloadOptionsTable(currentExtraction.ExtractionName)

End Function


Sub SaveConfiguration(ByRef chosenExtraction As CExtraction, _
                      ByRef chosenmailboxes() As CMailbox, _
                      ByRef chosenFilters() As CFilters, _
                      ByRef chosenDownloadOptions As CDownloadOptions)
    
    If IsAlreadyInUse(chosenExtraction) Then
        Err.Raise 513, _
                  "\\OutlookExtractor\CController\SaveConfiguration", _
                  "The given Extraction Name is already in use, please choose other one."
    End If
    
    With ObjModel
       .WriteOnExtractionTable chosenExtraction
       .WriteOnMailboxTable chosenmailboxes
       .WriteOnFiltersTable chosenFilters
       .WriteOnDownloadOptionsTable chosenDownloadOptions
    End With
    
End Sub


Function IsAlreadyInUse(ByRef chosenExtraction As CExtraction) As Boolean
    
    IsAlreadyInUse = (ObjModel.GetFromExtractionTable(chosenExtraction.ExtractionName) = "")
            
End Function


Function Execute(ByRef chosenmailboxes() As CMailbox, _
                 ByRef chosenFilters() As CFilters, _
                 ByRef chosenDownloadOptions As CDownloadOptions)

    
    
End Function

