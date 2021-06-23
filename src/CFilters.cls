VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Filters
    ExtractionName As String
    MailProperty As String
    FilterType As String
    FilterValue As String
End Type


Private this                           As Filters


Public Property Get ExtractionName() As String

    ExtractionName = this.ExtractionName

End Property


Public Property Get MailProperty() As String

    MailProperty = this.MailProperty

End Property


Public Property Get FilterType() As String

    FilterType = this.FilterType

End Property


Public Property Get FilterValue() As String

    FilterValue = this.FilterValue

End Property


Public Property Let ExtractionName(ByVal value As String)

    this.ExtractionName = value

End Property


Public Property Let MailProperty(ByVal value As String)

    this.MailProperty = value

End Property


Public Property Let FilterType(ByVal value As String)

    this.FilterType = value

End Property


Public Property Let FilterValue(ByVal value As String)

    this.FilterValue = value

End Property


