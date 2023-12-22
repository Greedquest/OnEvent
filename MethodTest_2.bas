Attribute VB_Name = "MethodTest_2"
'@Folder "CBuse - Array Iterators"
Option Explicit

Private Type TagVariant
    vt As Integer
    wReserved1 As Integer
    wReserved2 As Integer
    wReserved3 As Integer
    ptr1 As LongPtr
    ptr2 As LongPtr
End Type

Private Type nextItemAddr
    itemByRef As TagVariant
    dummy1 As LongPtr
    dummy2 As LongPtr
    nextPtr As LongPtr
End Type

#If Win64 Then
    Const sizeOfNext = 48
#Else
    Const sizeOfNext = 28
#End If

Sub TestByRef()
    Const size As Long = 10
    '
    Dim c As collection:        Set c = New collection
    Dim e As IEnumVARIANT:      Set e = c.[_NewEnum]
    Dim arr() As Variant:       ReDim arr(0 To size - 1)
    Dim arr2() As nextItemAddr: ReDim arr2(0 To size - 1)
    Dim h As New EnumHelper:    Set h.EnumVariant = e
    Dim i As Long
    Dim v As Variant
    Dim ptr1 As LongPtr
    Dim ptr2 As LongPtr
    Dim vt As Integer
    '
    ptr1 = VarPtr(arr(0))
    ptr2 = VarPtr(arr2(1))
    vt = vbVariant + VT_BYREF
    For i = 0 To UBound(arr2)
        arr(i) = i                               'doesn't matter what values
        arr2(i).itemByRef.ptr1 = ptr1
        arr2(i).itemByRef.vt = vt                'could use the actual var type here but then the ptr needs updated with +8
        arr2(i).nextPtr = ptr2
        ptr1 = ptr1 + VARIANT_SIZE
        ptr2 = ptr2 + sizeOfNext
    Next i
    arr2(UBound(arr2)).nextPtr = 0
    '
    MemLongPtr(ObjPtr(e) + PTR_SIZE * 2) = VarPtr(arr2(0))
    '
    For Each v In h
        Debug.Print v
        v = Empty                                'Required to avoid 'Type Mismatch' error that gets raised on the 'Next v' line
    Next v
End Sub

