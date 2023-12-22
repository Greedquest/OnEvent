Attribute VB_Name = "X64"
'@Folder Implementation.Assem
Option Explicit

Type decrementableThing
    vTablePointer As LongPtr
    refCount As Long
End Type

Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const PAGE_READWRITE As Long = &H4
Private Const PAGE_EXECUTE_READWRITE  As Long = &H40


Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal byteCount As LongPtr) As LongPtr

Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
        ByVal lpAddress As LongPtr, _
        ByVal dwSize As Long, _
        ByVal flAllocationType As Long, _
        ByVal flProtect As Long) As LongPtr

Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
        ByVal lpLibFileName As String) As LongPtr

Declare PtrSafe Function GetProcAddress Lib "kernel32" ( _
        ByVal hModule As LongPtr, _
        ByVal lpProcName As String) As LongPtr


Private Function GetShellCode(ByRef caption As String, ByRef text As String, _
    ByVal addrmessageBoxW As LongPtr, _
    ByVal addrCoTaskMemFree As LongPtr) As Byte()
    Dim captionAddress As LongLong
    Dim textAddress As LongLong

    captionAddress = StrPtr(caption)
    textAddress = StrPtr(text)


    ' Incorporating the new code which first decrements the refCount,
    ' checks if it's zero, frees memory using CoTaskMemFree,
    ' and then shows the MessageBox if refCount is zero.
    With New ShellCodeBuilder
    'VERY STRANGELY only needed in twinBasic
        '.Append "48", "8B", "09"               ' MOV RCX, [RCX]         ; Get the pointer from [RCX] into RCX
        .Append "48", "8B", "41", "08"         ' MOV RAX, [RCX + 0x08]
        .Append "48", "FF", "C8"               ' DEC RAX
        .Append "48", "89", "41", "08"         ' MOV [RCX + 0x08], RAX
        .Append "48", "83", "F8", "00"         ' CMP RAX, 0
        .Append "74", "01"                     ' JE +1 (short jump if RAX is zero)
        '.Append "74", "01"                     ' JE [1 instruction]     ; (skips early return if RAX is zero)
        .Append "C3"                           ' RET (return from function)
        .Append "48", "B8", addrCoTaskMemFree  ' MOV RAX, addrCoTaskMemFree
        .Append "48", "83", "EC", "28"         ' SUB RSP, 0x28
        .Append "FF", "D0"                     ' CALL RAX
        .Append "48", "83", "C4", "28"         ' ADD RSP, 0x28
        .Append "41", "B9", 0&                 ' MOV R9D, 0
        .Append "49", "B8", captionAddress     ' MOV R8, captionAddress
        .Append "48", "BA", textAddress        ' MOV RDX, textAddress
        .Append "48", "31", "C9"               ' XOR RCX, RCX
        .Append "48", "B8", addrmessageBoxW    ' MOV RAX, addrmessageBoxW
        .Append "48", "83", "EC", "28"         ' SUB RSP, 0x28
        .Append "FF", "D0"                     ' CALL RAX
        .Append "48", "83", "C4", "28"         ' ADD RSP, 0x28
        .Append "48", "31", "C0"               ' XOR RAX, RAX
        .Append "C3"                           ' RET
        GetShellCode = .ToBytes
    End With

End Function





Function NoisyReleaseRefCode() As LongPtr
Static buffer As LongPtr
If buffer <> 0 Then
    NoisyReleaseRefCode = buffer
    Exit Function
End If
    Static text As String
    Static caption As String
    text = "foo"
    caption = "hello"

    Dim MessageBoxW As LongPtr
    MessageBoxW = pmessageBoxW()

    Dim memFree As LongPtr
    memFree = pmemFree()

    Dim code() As Byte
    code = GetShellCode(caption, text, MessageBoxW, pmemFree)


    buffer = VirtualAlloc(0&, UBound(code) - LBound(code) + 1, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)

    If buffer = 0 Then
        Debug.Print "VirtualAlloc() failed (Err:"; Err.LastDllError; ")."
        Exit Function
    End If

    CopyMemory ByVal buffer, code(1), UBound(code) - LBound(code) + 1

    Dim i As Long
    For i = LBound(code) To UBound(code)
        Debug.Print Hex$(code(i)); " ";
    Next i

    NoisyReleaseRefCode = buffer

'
'    Dim x As decrementableThing
'
'    Dim typeBuffer As LongPtr
'    typeBuffer = CoTaskMemAlloc(LenB(x))
'
'    If typeBuffer = 0 Then
'        Debug.Print "CoTaskMemAlloc() failed (Err:"; Err.LastDllError; ")."
'        Exit Function
'    Else
'        Debug.Print vbNewLine, "BUFFFFAAA", , Hex$(typeBuffer)
'    End If
'
'    x.refcount = 1
'    x.vtablepointer = &H1234567890ABCDEF^
'
'    CopyMemory ByVal typeBuffer, x, LenB(x)
'
'
'    Dim res As Long
'    res = CallFunction(0, buffer, CR_LONG, CC_STDCALL, typeBuffer)
'    CopyMemory x, ByVal typeBuffer, LenB(x) 'unsafe
'    'Debug.Print vbNewLine ; res, VarPtr(x) ; MemLongPtr(VarPtr(x) + 8), typeBuffer, MemLongLong(typeBuffer + 8)
'
'    Debug.Print vbNewLine; res, typeBuffer, x.refcount, "Done!"
End Function

Function pmessageBoxW() As LongPtr
    Dim hLib As LongPtr
    hLib = LoadLibrary("User32.dll")
    pmessageBoxW = GetProcAddress(hLib, "MessageBoxW")
End Function


Function pmemFree() As LongPtr
    Dim hLib As LongPtr
    hLib = LoadLibrary("Ole32.dll")
    pmemFree = GetProcAddress(hLib, "CoTaskMemFree")
End Function


