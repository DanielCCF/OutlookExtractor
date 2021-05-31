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
Private MailPropertiesTable As ListObject
Private FilterTypesTable As ListObject

Private Sub Class_Initialize()
    
    With SDatabase
        Set ExtractionTable = .ListObjects("ExtractionTable")
        Set MailboxTable = .ListObjects("MailboxTable")
        Set DownloadOptionsTable = .ListObjects("DownloadOptionsTable")
        Set FiltersTable = .ListObjects("FiltersTable")
        Set MailPropertiesTable = .ListObjects("MailPropertiesTable")
        Set FilterTypesTable = .ListObjects("FilterTypesTable")
    End With
    
End Sub


Sub DeleteFromAllTables(ByVal ExtractionName As String)
    
    Dim tables As Variant
    Dim i As Integer
    Dim currentData As Range
    
    If (ExtractionTable.DataBodyRange.Find(ExtractionName, , , xlWhole, , , False) Is Nothing) Then
        Err.Raise 513, _
                  "\\OutlookExtractor\CModel\DeleteFromAllTables", _
                  "The given extraction name does not exist on the database."
    ElseIf ExtractionName = "" Then
        Exit Sub
    End If
    
    tables = Array(ExtractionTable, MailboxTable, DownloadOptionsTable, FiltersTable)
    
    For i = LBound(tables) To UBound(tables)
        'The tables(i) it's not recognized as a ListObject, so it
        'was need to convert using the Parent object and tables(i) name
        Set currentData = GetFilteredTable(ExtractionName, tables(i).Parent.ListObjects(tables(i).name))
        If Not (currentData Is Nothing) Then
            currentData.SpecialCells(xlCellTypeVisible).Clear
            currentData.AutoFilter
            ClearEmptyRows tables(i)
        End If
    Next
    
End Sub


Private Sub ClearEmptyRows(ByRef targetTable As Variant)
    
    Dim row As Range
    Dim hasOneRow As Boolean
    Dim isEmptyRow As Boolean
    Dim lastRow As Long
    Dim i As Long
    Dim currentRow As Range
    
    lastRow = targetTable.DataBodyRange.Rows.Count
    
    For i = lastRow To 1 Step -1
        Set currentRow = targetTable.DataBodyRange.Rows(i)
        hasOneRow = (lastRow = 1)
        isEmptyRow = (Application.CountA(currentRow) = 0)
        If Not hasOneRow And isEmptyRow Then
            currentRow.Delete xlUp
        End If
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


Sub WriteOnFiltersTable(ByRef info As CFilters)

    
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
    
    If targetDataRange.Columns.Count <> UBound(rowInformation) Then
       Err.Raise 503, _
                  "\\OutlookExtractor\CModel\WriteOnTable", _
                  "Data mismatch between the given information and the Database. " & _
                  "Please, check if the Database was changed by comparing with a " & _
                  "new program copy."
    End If
    
    With targetDataRange
        rowsAmount = .Rows.Count
        firstRowIsEmpty = Application.CountA(.Rows(1)) = 0
        If firstRowIsEmpty Then
            Set referenceCell = .Cells(1, 1)
        Else
            Set referenceCell = .Cells(rowsAmount, 1).Offset(1)
        End If
    End With
    
    For i = LBound(rowInformation) To UBound(rowInformation)
        referenceCell.Offset(, i - 1) = rowInformation(i)
    Next
    
End Sub


Function GetFromExtractionTable(ByVal ExtractionName As String) As CExtraction()

    Dim row As Range
    Dim dataRange As Range
    Dim list() As CExtraction
    Dim i As Long
    
    Set dataRange = GetFilteredTable(ExtractionName, ExtractionTable)
    
    If dataRange Is Nothing Then Exit Function
    
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
    
    If dataRange Is Nothing Then Exit Function
    
    ReDim list(dataRange.Rows.Count)
     For i = FIRST_ELEMENT_ARRAY To dataRange.Rows.Count
        Set list(i) = New CMailbox
        list(i).ExtractionName = dataRange.Rows(i).Cells(1, 1)
        list(i).MailboxItemId = dataRange.Rows(i).Cells(1, 2)
        list(i).IncludeSubfolders = dataRange.Rows(i).Cells(1, 3)
    Next
    
    GetFromMailboxTable = list
    
End Function


Function GetFromDownloadOptionsTable(ByVal ExtractionName As String) As CDownloadOptions()

    Dim i As Long
    Dim dataRange As Range
    Dim list() As CDownloadOptions
    
    Set dataRange = GetFilteredTable(ExtractionName, DownloadOptionsTable)
    
    If dataRange Is Nothing Then Exit Function
    
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
    
    If dataRange Is Nothing Then Exit Function
    
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
        On Error Resume Next
            'For some reason, sometimes .DataBodyRange.SpecialCells(xlCellTypeVisible)
            'returns all the visible cells from SDatabase instead the targetTable.
            'This situation occurred on the ExtractionTable when it's empty.
            'To solve this, comparing the Address whas needed
            If .DataBodyRange.SpecialCells(xlCellTypeVisible).Address <> SDatabase.Cells.SpecialCells(xlCellTypeVisible).Address Then
                Set GetFilteredTable = .DataBodyRange.SpecialCells(xlCellTypeVisible)
            End If
        On Error GoTo 0
        .Range.AutoFilter
    End With
    
End Function


Function GetFromMailPropertiesTable() As CMailProperties()

    Dim i As Long
    Dim dataRange As Range
    Dim list() As CMailProperties
    
    Set dataRange = GetPropertiesOnDatabase(MailPropertiesTable)
    
    If dataRange Is Nothing Then Exit Function
    
    ReDim list(dataRange.Rows.Count)
    For i = FIRST_ELEMENT_ARRAY To dataRange.Rows.Count
        Set list(i) = New CMailProperties
        list(i).Property = dataRange.Rows(i).Cells(1, 1)
    Next
    
    GetFromMailPropertiesTable = list

End Function


Function GetFromFiltersTypeTable() As CFilterTypes()

    Dim i As Long
    Dim dataRange As Range
    Dim list() As CFilterTypes
    
    Set dataRange = GetPropertiesOnDatabase(FilterTypesTable)
    
    If dataRange Is Nothing Then Exit Function
    
    ReDim list(dataRange.Rows.Count)
    For i = FIRST_ELEMENT_ARRAY To dataRange.Rows.Count
        Set list(i) = New CFilterTypes
        list(i).TypeName = dataRange.Rows(i).Cells(1, 1)
        list(i).RespectiveRegex = dataRange.Rows(i).Cells(1, 2)
    Next
    
    GetFromFiltersTypeTable = list

End Function


Private Function GetPropertiesOnDatabase(ByRef targetTable As ListObject) As Range

    Set GetPropertiesOnDatabase = targetTable.DataBodyRange.SpecialCells(xlCellTypeVisible)

End Function
