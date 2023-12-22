Attribute VB_Name = "LibOle"
Attribute VB_Description = "Credit: Ion Crisitian Buse https://gist.github.com/cristianbuse/b651a3cd740e27a78ea90bca9f7af4d1#file-libole-bas"
'@Folder "COMTools.OLE"
'@ModuleDescription("Credit: Ion Crisitian Buse https://gist.github.com/cristianbuse/b651a3cd740e27a78ea90bca9f7af4d1#file-libole-bas")
Option Explicit

#If Mac = False Then
    Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As LongPtr, ByRef pclsid As Any) As Long





#End If


'OLE Automation Protocol GUIDs
Public Const IID_NULL As String = "{00000000-0000-0000-0000-000000000000}"

'*******************************************************************************
'Converts a string to a GUID struct
'Note that 'CLSIDFromString' win API is only slightly faster (<10%) compared
'   to the pure VB approach (used for MAc only) but it has the advantage of
'   raising other types of errors (like class is not in registry)
'*******************************************************************************
#If Mac Then
Public Function GUIDFromString(ByVal sGUID As String) As GUIDt
    Const methodName As String = "GUIDFromString"
    Const hexPrefix As String = "&H"
    Static pattern As String
    '
    If LenB(pattern) = 0 Then pattern = Replace(IID_NULL, "0", "[0-9A-F]")
    If Not sGUID Like pattern Then Err.Raise 5, methodName, "Invalid string"
    '
    Dim parts() As String: parts = Split(Mid$(sGUID, 2, Len(sGUID) - 2), "-")
    Dim i As Long
    '
    With GUIDFromString
        .Data1 = CLng(hexPrefix & parts(0))
        .Data2 = CInt(hexPrefix & parts(1))
        .Data3 = CInt(hexPrefix & parts(2))
        For i = 0 To 1
            .Data4(i) = CByte(hexPrefix & Mid$(parts(3), i * 2 + 1, 2))
        Next i
        For i = 2 To 7
            .Data4(i) = CByte(hexPrefix & Mid$(parts(4), (i - 1) * 2 - 1, 2))
        Next i
    End With
End Function
#Else
'https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-clsidfromstring

'@Ignore NonReturningFunction
Public Function GUIDFromString(ByVal sGUID As String) As GUIDt
    Const methodName As String = "GUIDFromString"
    Dim hresult As Long: hresult = CLSIDFromString(StrPtr(sGUID), GUIDFromString)
    If hresult <> S_OK Then Err.Raise hresult, methodName, "Invalid string"
End Function
#End If

'*******************************************************************************
'Converts a GUID struct to a string
'Note that this approach is 4 times faster than running a combination of the
'   following 3 Windows APIs: StringFromCLSID, SysReAllocString, CoTaskMemFree
'*******************************************************************************




'*******************************************************************************
'Converts a CLSID string to a progid string. Windows only
'Returns an empty string if not successful
'*******************************************************************************
#If Mac = False Then

'@Ignore NonReturningFunction

#End If
