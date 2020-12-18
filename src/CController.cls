VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private UtilsOutlook As Object


Private Sub Class_Initialize()

    Set UtilsOutlook = New CUtils_Outlook
    
End Sub



