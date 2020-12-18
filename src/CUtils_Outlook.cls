VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUtils_Outlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================
'Description: This class is responsable for dealing with
'             general Outlook actions. Those actions are not
'             related to a specific process or program.
'
'==================================================================

Option Explicit


Function GetOutlookInstance() As Object

    On Error Resume Next
        Set GetOutlookInstance = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
End Function


Function CreateOutlookInstance() As Object
    
    Set GetOutlookInstance = CreateObject("Outlook.Application")
    
End Function


Sub SaveAttachments(ByVal mail As Object, _
                    ByVal targetFolder As String, _
                    ByVal saveWithMailTimestamp As Boolean, _
                    Optional textFilter As String)
                     
    Dim objAttachment As Object
    Dim hasTextFilter As Boolean
    Dim mailTimestamp As String
    
    For Each objAttachment In mail.Attachments
        hasTextFilter = InStr(1, objAttachment.fileName, textFilter) > 0
        If hasTextFilter Then
            If saveWithMailTimestamp Then _
                mailTimestamp = Format(mail.ReceivedTime, "yyyy_mm_dd__hh_mm_ss_")
            
            objAttachment.SaveAsFile targetFolder & "\" & mailTimestamp & objAttachment.fileName
        End If
    Next
    
End Sub


Sub FillCollectionWithSubFolders(ByRef foldersList As Collection, _
                                 ByVal parentFolder As Object)
    Dim f As Object
    
    For Each f In parentFolder.Folders
        foldersList.Add f
        FillCollectionWithSubFolders foldersList, f
    Next
    
End Sub


Function GetAllEmailsFromFolder(ByVal targetFolder As Object) As Collection

    Dim obj As Object
    
    Set GetAllEmailsFromFolder = New Collection

    For Each obj In targetFolder.Items
        If TypeOf obj Is Outlook.mailItem Then
            GetAllEmailsFromFolder.Add obj
        End If
    Next

End Function


Function GetMailsEntryIdCollectionFromFolder(ByVal targetFolder As Object) As Collection

    Dim item As Outlook.mailItem
    Dim obj As Object
    
    Set GetMailsEntryIdCollectionFromFolder = New Collection
    For Each obj In targetFolder.Items
        If TypeOf obj Is Outlook.mailItem Then
            GetMailsEntryIdCollectionFromFolder.Add obj.EntryID
        End If
    Next

End Function


Function GetExclusiveAttachmentsCollection(ByRef mailsCollection As Collection) As Collection
                                            
    Dim obj As Object
    Dim objItem As Variant
    Dim att As Object
    Dim curAtt As Variant
    Dim tempDictionary As Object
    Dim i As Long
    
    Set tempDictionary = CreateObject("Scripting.Dictionary")
    For i = 1 To mailsCollection.Count
        FillAttachmentsDictionary mailsCollection.item(i), tempDictionary
    Next

    For Each objItem In tempDictionary
        GetExclusiveAttachmentsCollection.Add tempDictionary.item(objItem)
    Next
    
End Function


Private Sub FillAttachmentsDictionary(ByVal currentMailId As String, _
                                      ByRef tempDictionary As Object)
    
    Dim objOutlook As Object
    Dim att As Outlook.attachment
    Dim attReceivedTime As String
    Dim fullKeyName
    Dim currentMail As Object
    
    Set objOutlook = GetOutlookInstance
    Set currentMail = objOutlook.Session.GetItemFromID(currentMailId)
    For Each att In currentMail.Attachments
        If tempDictionary.Exists(att.fileName) Then
            If currentMail.ReceivedTime > objOutlook.Session.GetItemFromID(tempDictionary.item(att.fileName)).ReceivedTime Then
                 tempDictionary.item(att.fileName) = currentMailId
            End If
        Else
            tempDictionary.Add att.fileName, currentMailId
        End If
    Next

End Sub



