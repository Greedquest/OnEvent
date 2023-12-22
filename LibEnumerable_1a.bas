Attribute VB_Name = "LibEnumerable_1a"
'@Folder "CBuse.Method1a"
Option Explicit

Private virtualTablePointers(0 To 6) As LongPtr

Const LONG_SIZE As Long = 4

Private Type CollectionEnumeratorVariant
    vTablePointer As LongPtr
    referenceCount As Long
    #If Win64 Then
    alignmentBytes As Long
    #End If
    enumerationInterface As IEnumerable_1a
    unknownPointer As LongPtr
    collectionInstance As collection
End Type

Private Enum CollectionEnumeratorVariantOffsets
    vTablePointerOffset = 0
    referenceCountOffset = PTR_SIZE
    #If Win64 Then
    enumerationInterfaceOffset = referenceCountOffset + PTR_SIZE
    #Else
    enumerationInterfaceOffset = referenceCountOffset + LONG_SIZE
    #End If
    unknownPointerOffset = enumerationInterfaceOffset + PTR_SIZE
    collectionInstanceOffset = unknownPointerOffset + PTR_SIZE
End Enum

Private pointerToEnumeratorIndex As New collection
Private enumeratorIndices() As Long

' Creates a new enumerable collection by setting up a custom virtual table for IEnumVARIANT.
' This allows us to intercept and handle the enumeration of the collection.
Public Function CreateNewEnumerable(ByRef collection As collection, ByRef enumerationInterface As IEnumerable_1a) As IEnumVARIANT
    Set CreateNewEnumerable = collection.[_NewEnum]
    Dim pointer As LongPtr: pointer = ObjPtr(CreateNewEnumerable)
    If virtualTablePointers(0) = 0 Then
        ' Copy existing virtual table and replace Next and Reset methods.
        MemCopy VarPtr(virtualTablePointers(0)), MemLongPtr(pointer), PTR_SIZE * 7
        virtualTablePointers(3) = VBA.Int(AddressOf CustomNextMethod)
        virtualTablePointers(5) = VBA.Int(AddressOf CustomResetMethod)
    End If
    MemLongPtr(pointer) = VarPtr(virtualTablePointers(0))
    MemLongPtr(pointer + enumerationInterfaceOffset) = ObjPtr(enumerationInterface)
    
    ' Map the pointer to its index for tracking current enumeration position.
    Dim identifier As String: identifier = CStr(pointer)
    Dim index As Long
    On Error Resume Next
    index = UBound(enumeratorIndices) + 1
    ReDim Preserve enumeratorIndices(0 To index)
    pointerToEnumeratorIndex.Remove identifier
    pointerToEnumeratorIndex.Add index, identifier
    On Error GoTo 0
End Function

' Custom implementation of the Next method for IEnumVARIANT.
' It advances the enumeration and fetches the next item.
Private Function CustomNextMethod(ByRef enumerator As CollectionEnumeratorVariant _
                                 , ByVal elementCount As Long _
                                 , ByRef returnedVariant As Variant _
                                 , ByVal fetchedElementCount As LongPtr) As Long
    If enumerator.enumerationInterface Is Nothing Then GoTo EndOfEnumeration
    Dim index As Long: index = pointerToEnumeratorIndex(CStr(VarPtr(enumerator)))
    If Not enumerator.enumerationInterface.ItemByIndex(enumeratorIndices(index), returnedVariant) Then GoTo EndOfEnumeration
    enumeratorIndices(index) = enumeratorIndices(index) + 1
    Exit Function
EndOfEnumeration:
    Const S_FALSE = 1
    CustomNextMethod = S_FALSE
End Function

' Custom implementation of the Reset method for IEnumVARIANT, not implemented in this context.
Private Function CustomResetMethod(ByVal pointer As Long) As Long
    Const E_NOTIMPL = &H80004001
    CustomResetMethod = E_NOTIMPL
End Function


