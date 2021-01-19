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
Private objOutlook As Object

Private Sub Class_Initialize()

    Set UtilsOutlook = New CUtils_Outlook
    Set objOutlook = UtilsOutlook.GetOutlookInstance(True)
    
End Sub


Function GetMailboxes() As Collection
    
    Dim box As Object
    
    Set GetMailboxes = New Collection
    
    UtilsOutlook.FillCollectionWithSubFolders GetMailboxes, objOutlook.Session
    
End Function


Function GetMainfolders() As Collection
    
    Dim f As Object
    
    Set GetMainfolders = New Collection
    
    For Each f In objOutlook.Session.Folders
        GetMainfolders.Add f
    Next
    
End Function


Function GetSubFoldersFrom(ByVal folderEntryID As String, _
                           ByVal storeId As String) As Collection

    'Set GetSubFoldersFrom = New Collection
    
    Set GetSubFoldersFrom = UtilsOutlook.GetSubFoldersFrom( _
                                            objOutlook.Session.GetFolderFromID(folderEntryID, _
                                                                               storeId))
    
End Function

