VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "OutlookExtractor"
Option Explicit

Private ExtractionTable As ListObject
Private MailboxTable As ListObject
Private DownloadOptionsTable As ListObject
Private FiltersTable As ListObject


Private Sub Class_Initialize()
    
    With SDatabase
        Set ExtractionTable = .ListObjects("ExtractionTable")
        Set MailboxTable = .ListObjects("MailboxTable")
        Set DownloadOptionsTable = .ListObjects("DownloadOptionsTable")
        Set FiltersTable = .ListObjects("FiltersTable")
    End With
    
End Sub


Sub WriteOnExtractionTable(ByRef info As Extraction)

    If Not (ExtractionTable.DataBodyRange.Find(info.extractionName, , , xlWhole, , , True) Is Nothing) Then
        Err.Raise 513, _
                  "\\OutlookExtractor\CModel\WriteOnExtractionTable", _
                  "The given extraction name already exists on the database. Please choose other name."

    End If
    
    WriteOnTable Array(info.extractionName), ExtractionTable.DataBodyRange
    
End Sub

Sub WriteOnMailboxTable(ByRef info As Mailbox)
    
    With info
        If .extractionName = "" Or .MailboxItemId = "" Then
            Err.Raise 513, _
                      "\\OutlookExtractor\CModel\WriteOnMailboxTable", _
                      "Mailbox Id and(or) Extraction Name are(is) empty"
        End If
        WriteOnTable Array(.extractionName, .MailboxItemId, .IncludeSubfolders), _
                    MailboxTable.DataBodyRange
    End With
    
End Sub


Sub WriteOnDownloadOptionsTable(ByRef info As DownloadOptions)

    With info
        If .extractionName = "" Or .DownloadFolder = "" Then
            Err.Raise 513, _
                      "\\OutlookExtractor\CModel\WriteOnDownloadOptionsTable", _
                      "Download folder and(or) Extraction Name are(is) empty"
        ElseIf .DownloadAttachments + .GetMailAsFile + .GetMailProperties = 0 Then
            Err.Raise 513, _
               "\\OutlookExtractor\CModel\WriteOnDownloadOptionsTable", _
               "None of the options was selected. Please choose at least one"
        End If
        
        WriteOnTable Array(.extractionName, .DownloadFolder, .DownloadAttachments, _
                           .GetMailProperties, .GetMailAsFile, .AfterDate, _
                           .BeforeDate), _
                    DownloadOptionsTable.DataBodyRange
    End With


End Sub


Sub WriteOnFiltersTable(ByRef info As Filters)

    
     With info
        If .extractionName = "" Or .MailProperty = "" Or .FilterType = "" Then
            Err.Raise 513, _
                      "\\OutlookExtractor\CModel\WriteOnFiltersTable", _
                      "Mail property and(or) Extraction Name and(or) Filter Type are(is) empty"
        End If
        WriteOnTable Array(.extractionName, .MailProperty, .FilterType, .FilterValue), _
                    FiltersTable.DataBodyRange
    End With


End Sub


Private Sub WriteOnTable(ByVal rowInformation As Variant, _
                         ByRef targetDataRange As Range)

    Dim rowsAmount As Long
    Dim referenceCell As Range
    Dim firstRowIsEmpty As Boolean
    Dim i As Integer
    
    If targetDataRange.Columns.Count <> UBound(rowInformation) + 1 Then
       Err.Raise 503, _
                  "\\OutlookExtractor\CModel\WriteOnTable", _
                  "Data mismatch between the given information and the Database. " & _
                  "Please, check if the Database was changed by comparing with a " & _
                  "new program copy."
    End If
    
    With targetDataRange
        rowsAmount = .rows.Count
        firstRowIsEmpty = Application.CountA(.rows(1).EntireRow) = 0
        If firstRowIsEmpty Then
            Set referenceCell = .Cells(1, 1)
        Else
            Set referenceCell = .Cells(rowsAmount, 1).Offset(1)
        End If
    End With
    
    For i = LBound(rowInformation) To UBound(rowInformation)
        referenceCell.Offset(, i) = rowInformation(i)
    Next
    
End Sub


Function GetFromExtractionTable(ByVal extracionName As String) As Collection

    Dim row As Range
    Dim dataRange As Range
    Dim currentObject As Extraction
    
    Set dataRange = GetFilteredTable(extractionName, "ExtractionTable")
    Set GetFromExtractionTable = New Collection
    
    
    For Each row In dataRange.rows
        currentObject.extractionName = row.Cells(1, 1)
        GetFromExtractionTable.Add currentObject
    Next
    
End Function


Function GetFromMailboxTable(ByVal extracionName As String) As Collection

    Dim row As Range
    Dim dataRange As Range
    Dim currentObject As Mailbox
    
    Set dataRange = GetFilteredTable(extractionName, "MailboxTable")
    Set GetFromExtractionTable = New Collection
    
    
    For Each row In dataRange.rows
        currentObject.extractionName = row.Cells(1, 1)
        currentObject.MailboxItemId = row.cels(1, 2)
        currentObject.IncludeSubfolders = row.Cells(1, 3)
        GetFromExtractionTable.Add currentObject
    Next

End Function


Function GetFromDownloadOptionsTable(ByVal extracionName As String) As Collection



End Function


Function GetFromFiltersTable(ByVal extracionName As String) As Collection



End Function


Private Function GetFilteredTable(ByVal extractionName As String, _
                                  ByRef targetTable As ListObject) As Range
                                  
    With targetTable
        .Range.AutoFilter
        .Range.AutoFilter 1, extractionName
        Set GetFilteredTable = .DataBodyRange.SpecialCells(xlCellTypeVisible)
    End With
    
End Function
