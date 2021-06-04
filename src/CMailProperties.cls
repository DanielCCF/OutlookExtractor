VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMailProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type MailProperties
    Property As String
    RespectiveObjectPropertyName As String
End Type

Private this As MailProperties


Public Property Get Property() As String

    Property = this.Property

End Property


Public Property Get RespectiveObjectPropertyName() As String

    RespectiveObjectPropertyName = this.RespectiveObjectPropertyName
    
End Property


Public Property Let Property(ByVal value As String)

    this.Property = value

End Property


Public Property Let RespectiveObjectPropertyName(ByVal value As String)

    this.RespectiveObjectPropertyName = value

End Property
