VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUtils_Regex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================
'Description: This class is responsable for dealing with
'             general RegEx actions. Those actions are not
'             related to a specific process or program.
'
'==================================================================
Option Explicit

Private RegEx As Object


Private Sub Class_Initialize()
    
    Set RegEx = CreateObject("vbscript.regexp")
    
End Sub


Property Let Pattern(ByVal textPattern As String)

    RegEx.Pattern = textPattern
    
End Property


Function HasPattern(ByVal Pattern As String, _
                    ByVal textToSearch As String) As Boolean
                     
    With RegEx
        .Pattern = Pattern
        HasPattern = .Test(textToSearch)
        .Pattern = ""
    End With
    
End Function


Function Extract(ByVal Pattern As String, _
                 ByVal textToSearch As String) As String
            
    With RegEx
        .Pattern = Pattern
    
        If .Test(textToSearch) Then
            Extract = .Execute(textToSearch)(0)
        Else
            Extract = vbNullString
        End If
    
        .Pattern = ""
    End With
    
End Function

