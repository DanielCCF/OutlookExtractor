VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit


Function GetMailboxInformation(ByVal dataRow As Long) As Range
    
    If dataRow > GetAmountMailboxes Then
        Err.Raise 513, "It was requested the row " & dataRow & " but the " & _
                      "mailboxes table only has" & GetAmountMailboxes
    End If
    
    With Me.ListObjects("MailboxesTable")
        Set GetMailboxInformation = .DataBodyRange.Rows(dataRow)
    End With

End Function


Sub AddMailbox(ByVal name As String, _
               ByVal entryId As String)
    'ByVal storeId As String)
    
    With GetReferenceCell
        .value = name
        .Offset(, 1).value = entryId
        '.Offset(, 2).value = storeId
    End With
    
End Sub


Private Function GetReferenceCell() As Range
    
    With Me.ListObjects("MailboxesTable")
        If Application.CountA(.DataBodyRange) = 0 Then
            Set GetReferenceCell = .DataBodyRange.Cells(1, 1)
        Else
            Set GetReferenceCell = .DataBodyRange.Cells(.DataBodyRange.Rows.Count, 1).Offset(1)
        End If
    End With
    
End Function


Sub EraseCurrentMailBoxes()
    
    Dim firstCellTable                 As Range
    
    With Me.Range("MailboxesTable")
        Set firstCellTable = .Cells(1, 1)
        .Delete
        firstCellTable = "Any value to bring table back"
        firstCellTable = ""
    End With

End Sub


Function GetAmountMailboxes()

    GetAmountMailboxes = Me.ListObjects("MailboxesTable").DataBodyRange.Rows.Count

End Function


Function GetMailProperties() As Collection
    
    Add1ColumnTableDataInCollection GetMailProperties, "MailPropertiesTable"

End Function


Function GetFilterTypes() As Collection

    Add1ColumnTableDataInCollection GetFilterTypes, "FilterTypesTable"

End Function


Private Sub Add1ColumnTableDataInCollection(ByRef targetCollection As Collection, _
                                            ByVal tableName As String)

    Dim rng                            As Range
    
    Set targetCollection = New Collection
    
    For Each rng In Me.ListObjects(tableName).DataBodyRange
        targetCollection.Add rng.value
    Next
End Sub


Function GetFolderId(ByVal folderName As String) As String


    On Error Resume Next
    GetFolderId = SSupport.ListObjects("MailboxesTable"). _
        DataBodyRange.Columns(1). _
        Find(folderName).Offset(, 1)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise 513, _
                  "\\SSupport\GetFolderId", _
                  "The folder: " & folderName & " was not found."
    End If
    On Error GoTo 0


End Function



