VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CountDownGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Implementation.CountDown"
Option Explicit

Public startFrom As Long

Private Sub Class_Initialize()
    startFrom = 10
End Sub

'@Enumerator
Public Property Get NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = RawEnumWithAssem.NewCountdownEnum(startFrom)
End Property
