Attribute VB_Name = "Tests"
'@Folder "CBuse"
Option Explicit

Const LONG_SIZE As Long = 4

Private Enum CollEnumVariantOffsets
    vTblPtrOffset = 0
    refCountOffset = PTR_SIZE
    #If Win64 Then
    nextItemOffset = refCountOffset + PTR_SIZE
    #Else
    nextItemOffset = refCountOffset + LONG_SIZE
    #End If
    unkPtrOffset = nextItemOffset + PTR_SIZE
    collPtrOffset = unkPtrOffset + PTR_SIZE
End Enum

Sub TestCollEnumVariantStructure()
    Const LONG_SIZE As Long = 4
    Dim c As New collection
    Dim arr() As Variant
    c.Add 111
    c.Add 222
    '
    Dim i As stdole.IEnumVARIANT: Set i = c.[_NewEnum]
    Dim e As New EnumHelper:      Set e.EnumVariant = i
    Dim iPtr As LongPtr:          iPtr = ObjPtr(i)
    Dim v As Variant
    Dim addr As LongPtr
    Dim j As Long
    '
    Debug.Print "Virtual table: " & MemLongPtr(iPtr + vTblPtrOffset)
    Debug.Print "Reference count: " & MemLong(iPtr + refCountOffset)
    Debug.Print "Next item ptr: " & MemLongPtr(iPtr + nextItemOffset)
    Debug.Print "Unknown ptr: " & MemLongPtr(iPtr + unkPtrOffset)
    Debug.Print "Collection ptr: " & MemLongPtr(iPtr + collPtrOffset) & " (same as: " & ObjPtr(c) & ")"
    Debug.Print
    '
    j = 0
    For Each v In e
        j = j + 1
        Debug.Print "Item value for index " & j & ": " & v
        If j = 1 Then
            addr = MemLongPtr(iPtr + nextItemOffset)
            Debug.Print "Address of second item is: " & addr
            MemLong(addr + 8) = 555              '+8 as offset within the Variant
        End If
    Next v
    Debug.Print
    '
    Debug.Print "Next item ptr (after loop): " & MemLongPtr(iPtr + nextItemOffset)
    i.Reset
    Debug.Print "Next item ptr (after Reset): " & MemLongPtr(iPtr + nextItemOffset)
    Set c = Nothing
    Debug.Print "Next item ptr (after Collection destroyed): " & MemLongPtr(iPtr + nextItemOffset)
    Debug.Print "Collection ptr (after Collection destroyed): " & MemLongPtr(iPtr + collPtrOffset)
End Sub

Sub TestArrayEnumerator_1a()
    Dim d As New DemoClass_1a
    Dim i As Long
    Dim v As Variant
    '
    For i = 1 To 3
        d.Add i
    Next i
    '
    For Each v In d
        Debug.Print v
    Next v
End Sub

Sub TestEnumeratorSpeed_1a()
    Dim d As New DemoClass_1a
    Dim i As Long
    Dim v As Variant
    Const size As Long = 1000000
    Dim arr() As Variant: ReDim arr(1 To size)
    Dim t As Double
    '
    For i = 1 To size
        arr(i) = i
        d.Add i
    Next i
    '
    Debug.Print "Running 'For Each' on " & Format$(size, "#,##0") & " elements"
    Debug.Print String$(40, "-")
    '
    t = Timer
    For Each v In arr
    Next v
    Debug.Print "Array: " & Round(Timer - t, 3) & " (seconds)"
    '
    t = Timer
    For Each v In d
    Next v
    Debug.Print "Class: " & Round(Timer - t, 3) & " (seconds)"
End Sub

Sub TestArrayEnumerator_1b()
    Dim d As New DemoClass_1b
    Dim i As Long
    Dim v As Variant
    '
    For i = 1 To 3
        d.Add i
    Next i
    '
    For Each v In d.NewEnum
        Debug.Print v
    Next v
End Sub

Sub TestEnumeratorSpeed_1b()
    Dim d As New DemoClass_1b
    Dim i As Long
    Dim v As Variant
    Const size As Long = 1000000
    Dim arr() As Variant: ReDim arr(1 To size)
    Dim t As Double
    '
    For i = 1 To size
        arr(i) = i
        d.Add i
    Next i
    
    '
    Debug.Print "Running 'For Each' on " & Format$(size, "#,##0") & " elements"
    Debug.Print String$(40, "-")
    '
    t = Timer
    For Each v In arr
    Next v
    Debug.Print "Array: " & Round(Timer - t, 3) & " (seconds)"
    '
    t = Timer
    For Each v In d.NewEnum
    Next v
    Debug.Print "Class: " & Round(Timer - t, 3) & " (seconds)"
End Sub

Sub TestArrayEnumerator_2()
    Dim d As New DemoClass_2
    Dim i As Long
    Dim v As Variant
    '
    For i = 1 To 3
        d.Add i
    Next i
    '
    For Each v In d
        Debug.Print v
    Next v
End Sub

Sub TestEnumeratorSpeed_2()
    Dim d As New DemoClass_2
    Dim i As Long
    Dim v As Variant
    Const size As Long = 1000000
    Dim arr() As Variant: ReDim arr(1 To size)
    Dim t As Double
    '
    For i = 1 To size
        arr(i) = i
        d.Add i
    Next i
    '
    Debug.Print "Running 'For Each' on " & Format$(size, "#,##0") & " elements"
    Debug.Print String$(40, "-")
    '
    t = Timer
    For Each v In arr
    Next v
    Debug.Print "Array: " & Round(Timer - t, 3) & " (seconds)"
    '
    t = Timer
    For Each v In d
    Next v
    Debug.Print "Class: " & Round(Timer - t, 3) & " (seconds)"
End Sub

Sub TestEnumeratorSpeed_3()
    Dim d As New NumberRange
    Const size As Long = 1000000
    Dim arr() As Variant: ReDim arr(1 To size)
    Dim t As Double
    '
    Dim i As Long
    For i = 1 To size
        arr(i) = i
    Next i
    d.generatorTo size
    
    Dim v As Variant
    
    '
    Debug.Print "Running 'For Each' on " & Format$(size, "#,##0") & " elements"
    Debug.Print String$(40, "-")
    '
    t = Timer
    For Each v In arr
    Next v
    Debug.Print "Array: " & Round(Timer - t, 3) & " (seconds)"
    '
    t = Timer
    For Each v In d
    Next v
    Debug.Print "BASIC for each: " & Round(Timer - t, 3) & " (seconds)"
End Sub
