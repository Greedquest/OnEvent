VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoClass_1a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "CBuse.Method1a"
Option Explicit

Implements IEnumerable_1a

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

'@Enumerator
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    If m_bufferUBound > m_nextIndex - 1 Then
        m_bufferUBound = m_nextIndex - 1
        ReDim Preserve m_items(0 To m_bufferUBound)
    End If
    Set NewEnum = CreateNewEnumerable(m_coll, Me)
End Function

Private Function IEnumerable_1a_ItemByIndex(ByVal iIndex As Long, ByRef outVar As Variant) As Boolean
    If m_nextIndex = 0 Then Exit Function
    If iIndex < LBound(m_items) Or iIndex > UBound(m_items) Then Exit Function
    If IsObject(m_items(iIndex)) Then
        Set outVar = m_items(iIndex)
    Else
        outVar = m_items(iIndex)
    End If
    IEnumerable_1a_ItemByIndex = True
End Function
