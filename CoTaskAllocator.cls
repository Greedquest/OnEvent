VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoTaskAllocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Archive.Implementation"
'@PredeclaredID
Option Explicit

Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal byteCount As LongPtr) As LongPtr


Private localCacheInstance As collection
Private Const name As String = "d5167d32-602c-4375-8eed-6ed642cad409" 'use ps [guid]::NewGuid() to avoid name clashes

Private Property Get defaultAppDomain() As AppDomain
    Static host As New mscoree.CorRuntimeHost
    Static result As mscorlib.AppDomain
    If result Is Nothing Then
        host.Start
        host.GetDefaultDomain result
    End If
    Set defaultAppDomain = result
End Property

Private Property Get openMemoryAddresses() As collection
    ' References:
    '  mscorlib.dll
    '  Common Language Runtime Execution Engine
    If localCacheInstance Is Nothing Then
        With defaultAppDomain
            'if collection not in cache then regenerate it
            If IsObject(.GetData(name)) Then
                'save it to a local copy for faster access (so we don't keep going through appDomain)
                Set localCacheInstance = .GetData(name)
            Else
                Set localCacheInstance = New collection
                .SetData name, localCacheInstance
            End If
        End With
    End If

    Set openMemoryAddresses = localCacheInstance
End Property

Public Function MemAlloc(ByVal cb As LongPtr) As LongPtr
    MemAlloc = CoTaskMemAlloc(cb)
    Debug.Print "Alloc "; MemAlloc
    openMemoryAddresses.Add MemAlloc
End Function

Public Sub FreeAll()
    'This is idempotent so can be called twice in a row without breaking anything
    Dim addr As Variant
    For Each addr In openMemoryAddresses
        Debug.Print "OOPS, leaking "; addr
'        Debug.Print "Free "; addr
'        CoTaskMemFree addr
    Next addr

    'to avoid double releasing memory next time we're called, we must clear the reference
    resetCache
End Sub

Private Sub resetCache()
    defaultAppDomain.SetData name, Empty
    Set localCacheInstance = Nothing
End Sub

Private Sub Class_Initialize()
    If Not Me Is CoTaskAllocator Then Err.Raise vbObjectError + 1, , "You cannot instantiate a new " & TypeName(Me) & ", use the predeclared instance"
    FreeAll
End Sub

Private Sub Class_Terminate()
    FreeAll
End Sub
