VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReturnLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder "Generators.Enumerate"
Option Explicit

Private internalValue As Long

'@DefaultMember
Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
    Value = internalValue
End Property

'@Ignore ProcedureNotUsed
Friend Property Let Value(ByVal newValue As Long)
    internalValue = newValue
End Property
