Attribute VB_Name = "LibEnumerable_1b"
'@Folder "CBuse - Array Iterators.Method1b"
Option Explicit

Private vtbl(0 To 6) As LongPtr

Const LONG_SIZE As Long = 4

Private Type CollEnumVariant
    vTblPtr As LongPtr
    refCount As Long
    #If Win64 Then
    alignBytes As Long
    #End If
    enumH As EnumHelper_1b
    unkPtr As LongPtr
    coll As collection
End Type

Private Enum CollEnumVariantOffsets
    vTblPtrOffset = 0
    refCountOffset = PTR_SIZE
    #If Win64 Then
    enumHelperOffset = refCountOffset + PTR_SIZE
    #Else
    enumHelperOffset = refCountOffset + LONG_SIZE
    #End If
    unkPtrOffset = enumHelperOffset + PTR_SIZE
    collOffset = unkPtrOffset + PTR_SIZE
End Enum

Public Function NewEnumHelper(ByRef c As collection, ByRef arr As Variant) As EnumHelper_1b
    With New EnumHelper_1b
        .Init c.[_NewEnum], arr
        Set NewEnumHelper = .Self
        Dim ptr As LongPtr: ptr = ObjPtr(.EnumVariant)
    End With
    If vtbl(0) = 0 Then
        'We keep the 3 IUnknown methods
        'We replace Next
        'Skip cannot be called from VBA anyway due to the unsupported type so can be left as is
        'We replace Reset just in case, to avoid crashes when coercing to 'EnumHelper_1b'
        'Clone can be left as is
        '
        MemCopy VarPtr(vtbl(0)), MemLongPtr(ptr), PTR_SIZE * 7
        vtbl(3) = VBA.Int(AddressOf IEnumVARIANT_Next)
        vtbl(5) = VBA.Int(AddressOf IEnumVARIANT_Reset)
    End If
    MemLongPtr(ptr) = VarPtr(vtbl(0))
    MemLongPtr(ptr + enumHelperOffset) = ObjPtr(NewEnumHelper)
End Function

Private Function IEnumVARIANT_Next(ByRef thisEnum As CollEnumVariant _
                                   , ByVal celt As Long _
                                    , ByRef rgVar As Variant _
                                     , ByVal pceltFetched As LongPtr) As Long
    Const S_FALSE = 1
    If thisEnum.enumH Is Nothing Then
        IEnumVARIANT_Next = S_FALSE
    ElseIf Not thisEnum.enumH.GetNext(rgVar) Then
        IEnumVARIANT_Next = S_FALSE
    End If
End Function

Private Function IEnumVARIANT_Reset(ByVal ptr As Long) As Long
    Const E_NOTIMPL = &H80004001
    IEnumVARIANT_Reset = E_NOTIMPL
End Function


