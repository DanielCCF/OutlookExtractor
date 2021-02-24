VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMailbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type Mailbox
    ExtractionName As String
    MailboxItemId As String
    IncludeSubfolders As Boolean
End Type


Private this As Mailbox


Public Property Get ExtractionName() As String

    ExtractionName = this.ExtractionName

End Property


Public Property Get MailboxItemId() As String

    MailboxItemId = this.MailboxItemId

End Property


Public Property Get IncludeSubfolders() As Boolean

    IncludeSubfolders = this.IncludeSubfolders

End Property


Public Property Let ExtractionName(ByVal value As String)

    this.ExtractionName = value

End Property


Public Property Let MailboxItemId(ByVal value As String)

    this.MailboxItemId = value

End Property


Public Property Let IncludeSubfolders(ByVal value As Boolean)

    this.IncludeSubfolders = value

End Property
