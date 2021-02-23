VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "OutlookExtractor"
Option Explicit

Function GetExtractionsNames() As Collection

    Dim cell As Variant
    
    Set GetExtractionsNames = New Collection
    For Each cell In Me.ListObjects("ExtractionTable").DataBodyRange.Columns(1)
        If cell.Value <> "" Then
            GetExtractionsNames.Add cell.Value
        End If
    Next

    
End Function



Sub a()

    GetExtractionsNames

End Sub
