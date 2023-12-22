VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TQDMobj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder "Generators.TQDM"
Option Explicit

Public iterable As Object
Public index As ReturnLong
Public Length As Long

'@Ignore ProcedureNotUsed
'@Enumerator
Public Property Get NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = TQDMIterator(Me.iterable, Me.index, Me.Length)
End Property
