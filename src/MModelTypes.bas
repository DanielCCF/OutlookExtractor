Attribute VB_Name = "MModelTypes"
Option Explicit

Public Type Extraction
    extractionName As String
End Type


Public Type Mailbox
    extractionName As String
    MailboxItemId As String
    IncludeSubfolders As Boolean
End Type


Public Type DownloadOptions
    extractionName As String
    DownloadFolder As String
    DownloadAttachments As Boolean
    GetMailProperties As Boolean
    GetMailAsFile As Boolean
    AfterDate As Date
    BeforeDate As Date
End Type


Public Type Filters
    extractionName As String
    MailProperty As String
    FilterType As String
    FilterValue As String
End Type
