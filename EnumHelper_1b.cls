VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumHelper_1b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "CBuse - Array Iterators.Method1b"
Option Explicit


Private m_enum As IEnumVARIANT
Private m_arrByRef As Variant
Private m_currIndex As Long

'@Enumerator
Public Property Get EnumVariant() As IEnumVARIANT
Attribute EnumVariant.VB_UserMemId = -4
    Set EnumVariant = m_enum
End Property

Public Property Get Self() As EnumHelper_1b
    Set Self = Me
End Property

Public Sub Init(ByRef newEnum_ As IEnumVARIANT _
                , ByRef arr As Variant)
    Set m_enum = newEnum_
    m_arrByRef = GetArrayByRef(arr)
    m_currIndex = LBound(arr)
    MemLongPtr(ObjPtr(newEnum_) + PTR_SIZE * 2) = ObjPtr(Me)
End Sub

Public Function GetNext(ByRef v As Variant) As Boolean
    If m_currIndex < LBound(m_arrByRef) Or _
                                        m_currIndex > UBound(m_arrByRef) Then Exit Function
    If IsObject(m_arrByRef(m_currIndex)) Then
        Set v = m_arrByRef(m_currIndex)
    Else
        v = m_arrByRef(m_currIndex)
    End If
    m_currIndex = m_currIndex + 1
    GetNext = True
End Function
