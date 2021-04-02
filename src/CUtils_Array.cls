VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUtils_Array"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================
'Description: This class is responsable for dealing with
'             general file actions. Those actions are not
'             related to a specific process or program.
'
'==================================================================


Private Function RemoveDuplicatesArray1D(ByVal MyArray As Variant) As Variant
    
    Dim firstElement As Long, lastElement As Long
    Dim i As Long
    Dim item As String
    
    Dim tempArray() As String
    Dim tempCollection As New Collection
 
    firstElement = LBound(MyArray)
    lastElement = UBound(MyArray)
    ReDim tempArray(firstElement To lastElement)
 
    For i = firstElement To lastElement
        tempArray(i) = CStr(MyArray(i))
    Next i
    
    On Error Resume Next
    For i = firstElement To lastElement
        tempCollection.Add tempArray(i), tempArray(i)
    Next i
    Err.Clear
    On Error GoTo 0
 
    lastElement = tempCollection.Count + firstElement - 1
    ReDim tempArray(firstElement To lastElement)
    
    For i = firstElement To lastElement
        tempArray(i) = tempCollection(i - firstElement + 1)
    Next i
    
    ArrayRemoveDups = tempArray
 
End Function


Function IsInitializedArray(ByRef targetArray As Variant) As Boolean
    
    On Error Resume Next
        IsInitializedArray = (UBound(targetArray) > -1)
        IsInitializedArray = ((Not Not targetArray) = 0)
    On Error GoTo 0
    
End Function
