VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoClass_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CBuse.Method2")
Option Explicit

Private Type VariantInfo
    IsObject As Boolean
    DummyOffset1 As Integer
    DummyOffset2 As Long
    #If Win64 Then
    DummyOffset3 As LongLong
    #End If
End Type

Private Type EnumerableVariant
    Value As Variant
    Info As VariantInfo
    nextPtr As LongPtr
End Type

Private m_coll As New collection
Private m_items() As EnumerableVariant
Private m_nextIndex As Long
Private m_bufferUBound As Long

Public Sub Add(ByRef Item As Variant)
    If m_nextIndex >= m_bufferUBound Then
        m_bufferUBound = m_bufferUBound * 2 + 1
        ReDim Preserve m_items(0 To m_bufferUBound)
    End If
    With m_items(m_nextIndex)
        .Info.IsObject = IsObject(Item)          'This can be used later to speed up item retrieval
        If .Info.IsObject Then Set .Value = Item Else .Value = Item
    End With
    m_nextIndex = m_nextIndex + 1
End Sub

Public Property Get Count() As Long
    Count = m_nextIndex
End Property

'@Enumerator
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = m_coll.[_NewEnum]
    Dim ptr As LongPtr: ptr = VarPtr(m_items(0))
    MemLongPtr(ObjPtr(NewEnum) + PTR_SIZE * 2) = ptr
    '
    If m_nextIndex > 1 Then
        #If Win64 Then
            Const size As Long = 48
        #Else
            Const size As Long = 28
        #End If
        Dim i As Long
        For i = 0 To m_nextIndex - 2
            ptr = ptr + size
            m_items(i).nextPtr = ptr
        Next i
    End If
End Function
