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

Private ExtractionTable As ListObjects
Private MailboxTable As ListObjects
Private DownloadOptionsTable As ListObjects
Private FiltersTable As ListObjects

Type Extraction
    ExtractionName As String
End Type


Type Mailbox
    ExtractionName As String
    MailboxItemId As String
    IncludeSubfolders As Long
End Type


Type DownloadOptions
    ExtractionName As String
    DownloadFolder As String
    DownloadAttachments As Boolean
    GetMailProperties As Boolean
    GetMailAsFile As Boolean
    AfterDate As Date
    BeforeDate As Date
End Type


Type Filters
    ExtractionName As String
    MailProperty As String
    FilterType As String
    FilterValue As String
End Type


Private Sub Class_Initialize()
    
    With SDatabase
        Set ExtractionTable = .ListObjects("ExtractionTable")
        Set MailboxTable = .ListObjects("MailboxTable")
        Set DownloadOptionsTable = .ListObjects("DownloadOptionsTable")
        Set FiltersTable = .ListObjects("FiltersTable")
    End With
    
End Sub



Sub WriteOnExtractionTable(ByVal info As Extraction)

    If ExtractionTable.DataBodyRange.Find(info.Name) Then
        Err.Raise 513, & _
                  "The given extraction name already exists on the database. " & _
                  "Please choose other name." , & _
                  "\\OutlookExtractor\CModel\WriteOnExtractionTable"
    End If
    
    WriteOnTable Array(info.Name), ExtractionTable
    
End Sub


Sub WriteOnMailboxTable(ByVal info As Mailbox)



End Sub


Sub WriteOnDownloadOptionsTable(ByVal info As DownloadOptions)



End Sub


Sub WriteOnFiltersTable(ByVal info As Filters)



End Sub


Sub WriteOnTable(ByVal rowInformation As Variant, ByRef targetDataRange As Range)

    Dim colsAmount As Integer
    Dim rowsAmount As Long
    Dim referenceCell As Range
    Dim i As Integer
    
    If targetDataRange.Columns.Count <> UBound(rowInformation) + 1 Then
       Err.Raise 503, _
                  "\\OutlookExtractor\CModel\WriteOnTable", _
                  "Data mismatch between the given information and the Database. " & _
                  "Please, check if the Database was changed by comparing with a " & _
                  "new program copy."
    End If
    
    colsAmount = targetDataRange.Columns.Count
    rowsAmount = targetDataRange.rows.Count
    Set referenceCell = targetDataRange.Cells(rowsAmount, 1).Offset(1)
    
    For i = LBound(rowInformation) To UBound(rowInformation)
        referenceCell.Offset(, i) = rowInformation(i)
    Next
    
End Sub


Function ReadExtractionTable() As Collection



End Function


Function ReadMailboxTable() As Collection



End Function


Function ReadDownloadOptionsTable() As Collection



End Function


Function ReadFiltersTable() As Collection



End Function


