VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Const FIRST_ELEMENT_ARRAY = 1
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


Sub DeleteFromAllTables(ByVal ExtractionName As String)
    
    Dim tables As Variant
    Dim i As Integer
    Dim currentData As Range
    
    If Not (ExtractionTable.DataBodyRange.Find(info.ExtractionName, , , xlWhole, , , True) Is Nothing) Then
        Err.Raise 513, _
                  "\\OutlookExtractor\CModel\WriteOnExtractionTable", _
                  "The given extraction name already exists on the database. Please choose other name."
    ElseIf extracitoname = "" Then
        End Sub
    End If
    
    tables = Array(ExtractionTable, MailboxTable, DownloadOptionsTable, FiltersTable)
    
    For i = LBound(tables) To UBound(tables)
        Set currentData = GetFilteredTable(ExtractionName, tables(i))
        currentData.SpecialCells(xlCellTypeVisible).EntireRow.Delete
    Next

End Sub


Sub WriteOnExtractionTable(ByRef info As CExtraction)

    If Not (ExtractionTable.DataBodyRange.Find(info.ExtractionName, , , xlWhole, , , True) Is Nothing) Then
        Err.Raise 513, _
                  "\\OutlookExtractor\CModel\WriteOnExtractionTable", _
                  "The given extraction name already exists on the database. Please choose other name."

    End If
    
    WriteOnTable Array(info.ExtractionName), ExtractionTable.DataBodyRange
    
End Sub


Sub WriteOnMailboxTable(ByRef info As CMailbox)
    
    With info
        If .ExtractionName = "" Or .MailboxItemId = "" Then
            Err.Raise 513, _
                      "\\OutlookExtractor\CModel\WriteOnMailboxTable", _
                      "Mailbox Id and(or) Extraction Name are(is) empty"
        End If
        WriteOnTable Array(.ExtractionName, .MailboxItemId, .IncludeSubfolders), _
                    MailboxTable.DataBodyRange
    End With
    
End Sub


Sub WriteOnDownloadOptionsTable(ByRef info As CDownloadOptions)

    With info
        If .ExtractionName = "" Or .DownloadFolder = "" Then
            Err.Raise 513, _
                      "\\OutlookExtractor\CModel\WriteOnDownloadOptionsTable", _
                      "Download folder and(or) Extraction Name are(is) empty"
        ElseIf .DownloadAttachments + .GetMailAsFile + .GetMailProperties = 0 Then
            Err.Raise 513, _
               "\\OutlookExtractor\CModel\WriteOnDownloadOptionsTable", _
               "None of the options was selected. Please choose at least one"
        End If
        
        WriteOnTable Array(.ExtractionName, .DownloadFolder, .DownloadAttachments, _
                           .GetMailProperties, .GetMailAsFile, .AfterDate, _
                           .BeforeDate), _
                    DownloadOptionsTable.DataBodyRange
    End With


End Sub


Sub WriteOnFiltersTable(ByRef info As Filters)

    
     With info
        If .ExtractionName = "" Or .MailProperty = "" Or .FilterType = "" Then
            Err.Raise 513, _
                      "\\OutlookExtractor\CModel\WriteOnFiltersTable", _
                      "Mail property and(or) Extraction Name and(or) Filter Type are(is) empty"
        End If
        WriteOnTable Array(.ExtractionName, .MailProperty, .FilterType, .FilterValue), _
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
        rowsAmount = .Rows.Count
        firstRowIsEmpty = Application.CountA(.Rows(1).EntireRow) = 0
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


Function GetFromExtractionTable(ByVal ExtractionName As String) As CExtraction()

    Dim row As Range
    Dim dataRange As Range
    Dim list() As CExtraction
    Dim i As Long
    
    Set dataRange = GetFilteredTable(ExtractionName, ExtractionTable)
    
    ReDim list(dataRange.Rows.Count)
    For i = FIRST_ELEMENT_ARRAY To dataRange.Rows.Count
        Set list(i) = New CExtraction
        list(i).ExtractionName = dataRange.Rows(i).Cells(1, 1)
    Next
    
    GetFromExtractionTable = list
    
End Function


Function GetFromMailboxTable(ByVal ExtractionName As String) As CMailbox()

    Dim i As Long
    Dim dataRange As Range
    Dim list() As CMailbox
    
    Set dataRange = GetFilteredTable(ExtractionName, MailboxTable)
    
    ReDim list(dataRange.Rows.Count)
     For i = FIRST_ELEMENT_ARRAY To dataRange.Rows.Count
        Set list(i) = New CMailbox
        list(i).ExtractionName = dataRange.Rows(i).Cells(1, 1)
        list(i).MailboxItemId = dataRange.Rows(i).cels(1, 2)
        list(i).IncludeSubfolders = dataRange.Rows(i).Cells(1, 3)
    Next
    
    GetFromMailboxTable = list
    
End Function


Function GetFromDownloadOptionsTable(ByVal ExtractionName As String) As CDownloadOptions()

    Dim i As Long
    Dim dataRange As Range
    Dim list() As CDownloadOptions
    
    Set dataRange = GetFilteredTable(ExtractionName, DownloadOptionsTable)
    
    ReDim list(dataRange.Rows.Count)
    For i = FIRST_ELEMENT_ARRAY To dataRange.Rows.Count
        Set list(i) = New CDownloadOptions
        list(i).ExtractionName = dataRange.Rows(i).Cells(1, 1)
        list(i).DownloadFolder = dataRange.Rows(i).Cells(1, 2)
        list(i).DownloadAttachments = dataRange.Rows(i).Cells(1, 3)
        list(i).GetMailAsFile = dataRange.Rows(i).Cells(1, 4)
        list(i).GetMailProperties = dataRange.Rows(i).Cells(1, 5)
        list(i).AfterDate = dataRange.Rows(i).Cells(1, 6)
        list(i).BeforeDate = dataRange.Rows(i).Cells(1, 7)
    Next
    
    GetFromDownloadOptionsTable = list

End Function


Function GetFromFiltersTable(ByVal ExtractionName As String) As CFilters()

    Dim i As Long
    Dim dataRange As Range
    Dim list() As CFilters
    
    Set dataRange = GetFilteredTable(ExtractionName, FiltersTable)
    
    ReDim list(dataRange.Rows.Count)
    For i = FIRST_ELEMENT_ARRAY To dataRange.Rows.Count
        Set list(i) = New CFilters
        list(i).ExtractionName = dataRange.Rows(i).Cells(1, 1)
        list(i).MailProperty = dataRange.Rows(i).Cells(1, 2)
        list(i).FilterType = dataRange.Rows(i).Cells(1, 3)
        list(i).FilterValue = dataRange.Rows(i).Cells(1, 4)
    Next
    
    GetFromFiltersTable = list

End Function


Private Function GetFilteredTable(ByVal ExtractionName As String, _
                                  ByRef targetTable As ListObject) As Range
                                  
    With targetTable
        .Range.AutoFilter
        If ExtractionName <> "" Then .Range.AutoFilter 1, ExtractionName
        Set GetFilteredTable = .DataBodyRange.SpecialCells(xlCellTypeVisible)
    End With
    
End Function


