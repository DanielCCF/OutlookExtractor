VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExtraction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type extraction
    ExtractionName As String
End Type


Private this                           As extraction


Public Property Get ExtractionName() As String

    ExtractionName = this.ExtractionName

End Property


Public Property Let ExtractionName(ByVal value As String)

    this.ExtractionName = value

End Property


