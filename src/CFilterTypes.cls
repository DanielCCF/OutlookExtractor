VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFilterTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type FilterTypes
    TypeName As String
    RespectiveRegex As String
End Type

Private this As FilterTypes

Public Property Get TypeName() As String

    TypeName = this.TypeName

End Property


Public Property Let TypeName(ByVal value As String)

    this.TypeName = value

End Property


Public Property Get RespectiveRegex() As String

    RespectiveRegex = this.RespectiveRegex

End Property


Public Property Let RespectiveRegex(ByVal value As String)

    this.RespectiveRegex = value

End Property
