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

Private Const MAIL_ITEM = 43

Function GetOutlookInstance(Optional createIfNotFound As Boolean) As Object

    On Error Resume Next
        Set GetOutlookInstance = GetObject(, "Outlook.Application")
        If GetOutlookInstance Is Nothing And createIfNotFound Then _
            Set GetOutlookInstance = CreateObject("Outlook.Application")
    On Error GoTo 0
    
End Function


Sub SaveAttachments(ByVal mail As Object, _
                    ByVal TargetFolder As String, _
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
            
            objAttachment.SaveAsFile TargetFolder & "\" & mailTimestamp & objAttachment.fileName
        End If
    Next
    
End Sub


Sub FillCollectionWithAllSubFolders(ByRef foldersList As Collection, _
                                    ByRef parentFolder As Object)
    Dim f As Variant
    Dim currentSession As Object
    
    Set currentSession = GetOutlookInstance(True).Session
    For Each f In parentFolder.Folders
        If f.DefaultItemType = 0 Then
            foldersList.Add currentSession.GetFolderFromID(f.entryId, f.storeId)
            FillCollectionWithAllSubFolders foldersList, currentSession.GetFolderFromID(f.entryId, f.storeId)
        End If
    Next

End Sub


Sub FillCollectionWithSubFolders(ByRef foldersList As Collection, _
                                 ByRef parentFolder As Object)
    Dim f As Object
    Dim currentSession As Object
    
    Set currentSession = GetOutlookInstance(True).Session
    For Each f In parentFolder.Folders
        If f.DefaultItemType = 0 Then
            foldersList.Add currentSession.GetFolderFromID(f.entryId, f.storeId)
        End If
    Next

End Sub


Function GetAllEmailsFromFolder(ByVal TargetFolder As Object) As Collection

    Dim obj As Object
    
    Set GetAllEmailsFromFolder = New Collection

    For Each obj In TargetFolder.Items
        If TypeOf obj Is Outlook.mailItem Then
            GetAllEmailsFromFolder.Add obj
        End If
    Next

End Function


Function GetMailsEntryIdCollectionFromFolder(ByVal TargetFolder As Object) As Collection

    Dim item As Object
    Dim obj As Object
    
    Set GetMailsEntryIdCollectionFromFolder = New Collection
    For Each obj In TargetFolder.Items
        If obj.Class = MAIL_ITEM Then
            GetMailsEntryIdCollectionFromFolder.Add obj.entryId
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
    
    Dim ObjOutlook As Object
    Dim att As Outlook.attachment
    Dim attReceivedTime As String
    Dim fullKeyName
    Dim currentMail As Object
    
    Set ObjOutlook = GetOutlookInstance
    Set currentMail = ObjOutlook.Session.GetItemFromID(currentMailId)
    For Each att In currentMail.Attachments
        If tempDictionary.Exists(att.fileName) Then
            If currentMail.ReceivedTime > ObjOutlook.Session.GetItemFromID(tempDictionary.item(att.fileName)).ReceivedTime Then
                 tempDictionary.item(att.fileName) = currentMailId
            End If
        Else
            tempDictionary.Add att.fileName, currentMailId
        End If
    Next

End Sub



Function GetSubFoldersFrom(ByRef folder As Object) As Collection


    Dim f As Object
    
    Set GetSubFoldersFrom = New Collection
    
    For Each f In folder.Folders
        GetSubFoldersFrom.Add folder
    Next
End Function


