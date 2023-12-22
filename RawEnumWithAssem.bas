Attribute VB_Name = "RawEnumWithAssem"
'@Folder "Generators.Assem.CountDown"
Option Explicit

Private Type TENUMERATOR
    VTablePtr   As LongPtr
    refCount  As Long
    index       As Long
End Type

Private Type IEnumVariantVTable
    IUnknown As IUnknownVTable
    Next As LongPtr
    Skip As LongPtr
    Reset As LongPtr
    Clone As LongPtr
End Type

Private Enum API
    S_OK = 0
    S_FALSE = 1
    E_NOTIMPL = &H80004001
    E_NOINTERFACE = &H80004002
    E_POINTER = &H80004003
End Enum

Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal cb As LongPtr) As LongPtr
Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByVal lpiid As LongPtr) As Long
Private Declare PtrSafe Function SysAllocStringByteLen Lib "oleaut32" (ByVal psz As LongPtr, ByVal cblen As Long) As LongPtr
Private Declare PtrSafe Function VariantCopyToPtr Lib "oleaut32" Alias "VariantCopy" (ByVal pvargDest As LongPtr, ByRef pvargSrc As Variant) As Long
Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)

'@Ignore NonReturningFunction
Public Function NewCountdownEnum(ByVal startAt As Long) As IEnumVARIANT
    Static VTable As IEnumVariantVTable
    If VTable.IUnknown.QueryInterface = NULL_PTR Then
        ' Setup the COM object's virtual table
        VTable.IUnknown.QueryInterface = VBA.CLngPtr(AddressOf IUnknown_QueryInterface)
        VTable.IUnknown.AddRef = VBA.CLngPtr(AddressOf IUnknown_AddRef)
        VTable.IUnknown.ReleaseRef = X64.NoisyReleaseRefCode
        VTable.Next = VBA.CLngPtr(AddressOf IEnumVARIANT_Next)
        VTable.Skip = VBA.CLngPtr(AddressOf IEnumVARIANT_Skip)
        VTable.Reset = VBA.CLngPtr(AddressOf IEnumVARIANT_Reset)
        VTable.Clone = VBA.CLngPtr(AddressOf IEnumVARIANT_Clone)
    End If


    Dim instance As TENUMERATOR
    With instance
        .VTablePtr = VarPtr(VTable.IUnknown.QueryInterface)
        .refCount = 1
        .index = startAt
    End With

    Dim someMemory As LongPtr
    someMemory = CoTaskMemAlloc(LenB(instance))

    CopyMemory ByVal someMemory, instance, LenB(instance)
    ZeroMemory instance, LenB(instance)
    '@Ignore ValueRequired
    MemLongPtr(VarPtr(NewCountdownEnum)) = someMemory
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

Private Function inPlaceIncrement(ByRef x As Long) As Long
    x = x + 1
    inPlaceIncrement = x
End Function

Private Function inPlaceDecrement(ByRef x As Long) As Long
    x = x - 1
    inPlaceDecrement = x
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
    IUnknown_AddRef = inPlaceIncrement(this.refCount)
End Function

Private Function IEnumVARIANT_Next(ByRef this As TENUMERATOR, _
                                   ByVal celt As Long, _
                                   ByVal rgVar As LongPtr, _
                                   ByVal pceltFetched As LongPtr _
                                   ) As Long

    'Const VARIANT_CB As Long = 16 ' VARIANT size in bytes

    If rgVar = NULL_PTR Then
        IEnumVARIANT_Next = E_POINTER
        Exit Function
    End If
    Dim Fetched As Long
    Dim element As Variant

    Do While this.index > 0
        element = inPlaceDecrement(this.index)
        VariantCopyToPtr rgVar, element
        Fetched = Fetched + 1
        If Fetched = celt Then Exit Do
        rgVar = UnsignedAdd(rgVar, VARIANT_SIZE)
    Loop

    If pceltFetched Then MemLong(pceltFetched) = Fetched
    If Fetched < celt Then
        IEnumVARIANT_Next = S_FALSE
    Else
        IEnumVARIANT_Next = S_OK
    End If
End Function

'@Ignore ParameterNotUsed
Private Function IEnumVARIANT_Skip(ByRef this As TENUMERATOR, ByVal celt As Long) As Long
    IEnumVARIANT_Skip = E_NOTIMPL
End Function

'@Ignore ParameterNotUsed
Private Function IEnumVARIANT_Reset(ByRef this As TENUMERATOR) As Long
    IEnumVARIANT_Reset = E_NOTIMPL
End Function

'@Ignore ParameterNotUsed
Private Function IEnumVARIANT_Clone(ByRef this As TENUMERATOR, ByVal ppEnum As LongPtr) As Long
    IEnumVARIANT_Clone = E_NOTIMPL
End Function


