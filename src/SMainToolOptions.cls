VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SMainToolOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function GetExtractionsNames() As Variant

    Dim tempArray() As Variant
    Dim cell As Variant
    
    For Each cell In SMainToolOptions.Range("MAIN_TOOL_OPTIONS_TABLE[EXTRACTION_NAME]")
        If (Not Not tempArray) = 0 Then
            ReDim tempArray(0)
        Else
            ReDim Preserve tempArray(UBound(tempArray) + 1)
        End If
        
        tempArray(UBound(tempArray)) = cell.Value
    Next
    
    GetExtractionsNames = tempArray
    
End Function


