Attribute VB_Name = "Enumeration"
'@Folder "Generators.Enumerate"
Option Explicit

Private Type TENUMERATOR
    VTablePtr   As LongPtr
    refCount  As Long
    baseEnum As IEnumVARIANT
    index As ReturnLong
End Type

Private Type IEnumVariantVTable
    IUnknown As IUnknownVTable
    Next As LongPtr
    Skip As LongPtr
    Reset As LongPtr
    Clone As LongPtr
End Type: Private IEnumVariantVTable As IEnumVariantVTable



'Private Instances(1 To 100) As TENUMERATOR

Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)

Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByVal lpiid As LongPtr) As Long
Private Declare PtrSafe Function SysAllocStringByteLen Lib "oleaut32" (ByVal psz As LongPtr, ByVal cblen As Long) As LongPtr

Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)


Private Property Get IEnumVariantVTableOffset(ByRef member As LongPtr) As LongPtr
    IEnumVariantVTableOffset = VarPtr(member) - VarPtr(IEnumVariantVTable)
End Property

'Private Static Function nextIndex() As Long
'    Dim i As Long
'    i = i + 1
'    If i > UBound(Instances) Then Err.Raise 5, , "Thats quite enough"
'    nextIndex = i
'End Function

Public Function Enumerate(ByVal iterable As Object, ByRef outIndex As ReturnLong) As Object
    Set outIndex = New ReturnLong
    Dim result As EnumerateObj
    Set result = New EnumerateObj
    Set result.iterable = iterable
    Set result.index = outIndex

    Set Enumerate = result

End Function


Private Function CallByDispid(ByVal accessor As Object, ByVal DISPID As DISPID) As Variant 'orig as object

    Dim localeID As Long 'Not really needed. Could pass 0 instead
    localeID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)

    Dim outExcepInfo As EXCEPINFOt

    Dim guidIID_NULL As GUIDt
    guidIID_NULL = GUIDFromString(IID_NULL)
    '@Ignore IntegerDataType
    Dim flags As Integer
    flags = VbMethod Or VbGet

    Dim params As DISPPARAMSt 'this empty should be sufficient if no params

    Dim outFirstBadArgIndex As Long

    'HRESULT Invoke(
    '  [in]      DISPID     dispIdMember,
    '  [in]      REFIID     riid,
    '  [in]      LCID       lcid,
    '  [in]      WORD       wFlags,
    '  [in, out] DISPPARAMS *pDispParams,
    '  [out]     VARIANT    *pVarResult,
    '  [out]     EXCEPINFO  *pExcepInfo,
    '  [out] UINT * puArgErr
    ');
    Debug.Print "INVOKED="; ObjPtr(accessor)
    Dim hresult As hResultCode
    On Error Resume Next
    hresult = CallFunction( _
        ObjPtr(accessor), IDispatchVTableOffset(IDispatchVTable.Invoke), _
        CR_HRESULT, CC_STDCALL, _
        DISPID, _
        VarPtr(guidIID_NULL), localeID, flags, _
        VarPtr(params), _
        VarPtr(CallByDispid), VarPtr(outExcepInfo), VarPtr(outFirstBadArgIndex) _
        )

    If hresult <> S_OK Then Stop
    On Error GoTo 0
End Function

Public Function NewEnumerator(ByVal iterable As Object, ByVal index As ReturnLong) As IEnumVARIANT
    Static VTable As IEnumVariantVTable
    If VTable.IUnknown.QueryInterface = NULL_PTR Then
        ' Setup the COM object's virtual table
        VTable.IUnknown.QueryInterface = VBA.CLngPtr(AddressOf IUnknown_QueryInterface)
        VTable.IUnknown.AddRef = VBA.CLngPtr(AddressOf IUnknown_AddRef)
        VTable.IUnknown.ReleaseRef = VBA.CLngPtr(AddressOf IUnknown_Release)
        VTable.Next = VBA.CLngPtr(AddressOf IEnumVARIANT_Next)
        VTable.Skip = VBA.CLngPtr(AddressOf IEnumVARIANT_Skip)
        VTable.Reset = VBA.CLngPtr(AddressOf IEnumVARIANT_Reset)
        VTable.Clone = VBA.CLngPtr(AddressOf IEnumVARIANT_Clone)
    End If

'    Dim i As Long
'    i = nextIndex
'    With Instances(i)
'        .VTablePtr = VarPtr(VTable.IUnknown.QueryInterface)
'        .refCount = 1
'        Set .baseEnum = CallByDispid(iterable, DISPID_NEWENUM)
'        Set .index = index
'        .index = 0
'    End With

    Dim NewEnum As TENUMERATOR
    With NewEnum
            .VTablePtr = VarPtr(VTable.IUnknown.QueryInterface)
            .refCount = 1
            Set .baseEnum = CallByDispid(iterable, DISPID_NEWENUM)
            Set .index = index
            .index = 0
    End With

    Dim pMemoryBlock As LongPtr
    pMemoryBlock = CoTaskAllocator.MemAlloc(LenB(NewEnum))
    If pMemoryBlock = NULL_PTR Then Exit Function

    CopyMemory ByVal pMemoryBlock, NewEnum, LenB(NewEnum)
    ZeroMemory NewEnum, LenB(NewEnum)


    MemLongPtr(VarPtr(NewEnumerator)) = pMemoryBlock 'VarPtr(Instances(i))
End Function

Private Function RefToIID(ByVal riid As LongPtr) As String
    ' copies an IID referenced into a binary string
    Const IID_CB As Long = 16&  ' GUID/IID size in bytes
    MemLongPtr(VarPtr(RefToIID)) = SysAllocStringByteLen(riid, IID_CB)
End Function

Private Function StrToIID(ByRef iid As String) As String
    ' converts a string to an IID
    StrToIID = RefToIID(NULL_PTR)
    IIDFromString StrPtr(iid), StrPtr(StrToIID)
End Function

Private Static Function IID_IUnknown() As String
    Dim iid As String
    If StrPtr(iid) = NULL_PTR Then iid = StrToIID("{00000000-0000-0000-C000-000000000046}")
    IID_IUnknown = iid
End Function

Private Static Function IID_IEnumVARIANT() As String
    Dim iid As String
    If StrPtr(iid) = NULL_PTR Then iid = StrToIID("{00020404-0000-0000-C000-000000000046}")
    IID_IEnumVARIANT = iid
End Function

Private Function InterlockedIncrement(ByRef Addend As Long) As Long
    Addend = Addend + 1
    InterlockedIncrement = Addend
End Function

Private Function InterlockedDecrement(ByRef Addend As Long) As Long
    Addend = Addend - 1
    InterlockedDecrement = Addend
End Function


Private Function IUnknown_QueryInterface(ByRef this As TENUMERATOR, _
                                         ByVal riid As LongPtr, _
                                         ByVal ppvObject As LongPtr _
                                         ) As Long
    If ppvObject = NULL_PTR Then
        IUnknown_QueryInterface = E_POINTER
        Exit Function
    End If
    Select Case RefToIID(riid)
        Case IID_IUnknown, IID_IEnumVARIANT
            MemLongPtr(ppvObject) = VarPtr(this)
            IUnknown_AddRef this
            IUnknown_QueryInterface = S_OK
        Case Else
            IUnknown_QueryInterface = E_NOINTERFACE
    End Select
End Function

Private Function IUnknown_AddRef(ByRef this As TENUMERATOR) As Long
    IUnknown_AddRef = InterlockedIncrement(this.refCount)
End Function

Private Function IUnknown_Release(ByRef this As TENUMERATOR) As Long
    Dim Count As Long
    Count = InterlockedDecrement(this.refCount)
    If Count = 0 Then
        Set this.baseEnum = Nothing
        Set this.index = Nothing
        CoTaskMemFree VarPtr(this)
        Debug.Print "Enumerator was released"
    End If
    IUnknown_Release = Count
End Function

Private Function IEnumVARIANT_Next(ByRef this As TENUMERATOR, _
                                   ByVal celt As Long, _
                                   ByVal rgVar As LongPtr, _
                                   ByVal pceltFetched As LongPtr _
                                   ) As Long

    'Const VARIANT_CB As Long = 16 ' VARIANT size in bytes

    'forward call to base IEnumVariant::Next
    Dim result As hResultCode
    result = CallFunction( _
        ObjPtr(this.baseEnum), IEnumVariantVTableOffset(IEnumVariantVTable.Next), CR_HRESULT, CC_STDCALL, _
        celt, rgVar, pceltFetched)

    'increment every time we get a new item
    If result = S_OK Then this.index = this.index + celt
    IEnumVARIANT_Next = result

End Function

'@Ignore ParameterNotUsed
'@Ignore ParameterNotUsed
Private Function IEnumVARIANT_Skip(ByRef this As TENUMERATOR, ByVal celt As Long) As Long
    IEnumVARIANT_Skip = E_NOTIMPL
End Function

'@Ignore ParameterNotUsed
Private Function IEnumVARIANT_Reset(ByRef this As TENUMERATOR) As Long
    IEnumVARIANT_Reset = E_NOTIMPL
End Function

'@Ignore ParameterNotUsed
'@Ignore ParameterNotUsed
Private Function IEnumVARIANT_Clone(ByRef this As TENUMERATOR, ByVal ppEnum As LongPtr) As Long
    IEnumVARIANT_Clone = E_NOTIMPL
End Function


