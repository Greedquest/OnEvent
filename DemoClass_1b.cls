VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoClass_1b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "CBuse - Array Iterators.Method1b"
Option Explicit


Private m_coll As New collection
Private m_items() As Variant
Private m_nextIndex As Long
Private m_bufferUBound As Long

Public Sub Add(ByRef Item As Variant)
    If m_nextIndex >= m_bufferUBound Then
        m_bufferUBound = m_bufferUBound * 2 + 1
        ReDim Preserve m_items(0 To m_bufferUBound)
    End If
    If IsObject(Item) Then
        Set m_items(m_nextIndex) = Item
    Else
        m_items(m_nextIndex) = Item
    End If
    m_nextIndex = m_nextIndex + 1
End Sub

Public Property Get Count() As Long
    Count = m_nextIndex
End Property

Public Function NewEnum() As EnumHelper_1b
    If m_bufferUBound > m_nextIndex - 1 Then
        m_bufferUBound = m_nextIndex - 1
        ReDim Preserve m_items(0 To m_bufferUBound)
    End If
    Set NewEnum = NewEnumHelper(m_coll, m_items)
End Function

