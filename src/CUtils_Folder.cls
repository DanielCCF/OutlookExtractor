VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUtils_Folder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================
'Description: This class is responsable for dealing with
'             general folder actions. Those actions are not
'             related to a specific process or program.
'
'==================================================================
Option Explicit


Function GetChosenFolderName(Optional initialPath As String) As String

    Dim result As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = initialPath
        If .Show Then result = .SelectedItems(1)
    End With
    
    GetChosenFolderName = result
    
End Function


Function PickAllFilesNamesInAFolder(ByVal folder As String, _
                                    Optional stringFilter As String, _
                                    Optional delimiter As String) As String

    Dim result As String
    Dim file As Object
    Dim filesInChosenFolder As Object
    
    If delimiter = "" Then delimiter = ";"
    
    Set filesInChosenFolder = CreateObject("Scripting.FileSystemObject") _
                              .GetFolder(folder) _
                              .files

    For Each file In filesInChosenFolder
        If InStr(1, file.name, stringFilter) > 0 Then
            result = result & folder & "\" & file.name & delimiter
        End If
    Next
    
    If result <> "" Then _
        PickAllFilesNamesInAFolder = Left(result, Len(result) - 1)

End Function
