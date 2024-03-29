VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDownloadOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type DownloadOptions
    ExtractionName As String
    DownloadFolder As String
    DownloadAttachments As Boolean
    GetMailProperties As Boolean
    GetMailAsFile As Boolean
    afterDate As Date
    beforeDate As Date
End Type


Private this                           As DownloadOptions


Public Property Get ExtractionName() As String

    ExtractionName = this.ExtractionName

End Property


Public Property Get DownloadFolder() As String

    DownloadFolder = this.DownloadFolder

End Property


Public Property Get DownloadAttachments() As Boolean

    DownloadAttachments = this.DownloadAttachments

End Property


Public Property Get GetMailProperties() As Boolean

    GetMailProperties = this.GetMailProperties

End Property


Public Property Get GetMailAsFile() As Boolean

    GetMailAsFile = this.GetMailAsFile

End Property


Public Property Get afterDate() As Date

    afterDate = this.afterDate

End Property


Public Property Get beforeDate() As Date

    beforeDate = this.beforeDate

End Property


Public Property Let ExtractionName(ByVal value As String)

    this.ExtractionName = value

End Property


Public Property Let DownloadFolder(ByVal value As String)

    this.DownloadFolder = value

End Property


Public Property Let DownloadAttachments(ByVal value As Boolean)

    this.DownloadAttachments = value
    
End Property


Public Property Let GetMailProperties(ByVal value As Boolean)

    this.GetMailProperties = value

End Property


Public Property Let GetMailAsFile(ByVal value As Boolean)

    this.GetMailAsFile = value

End Property


Public Property Let afterDate(ByVal value As Date)

    this.afterDate = value

End Property


Public Property Let beforeDate(ByVal value As Date)

    this.beforeDate = value

End Property


