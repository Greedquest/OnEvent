VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellCodeBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Generators.Assem"
Option Explicit

Private result() As Byte
Private pos As Long

Private Sub Class_Initialize()
    ' Initial allocation with some buffer
    ReDim result(1 To 1000)
    pos = 1
End Sub

Private Sub EnsureSpace(ByVal RequiredArgs As Long)
    Dim requiredSpace As Long
    Dim newSize As Long

    requiredSpace = 8 * RequiredArgs

    If UBound(result) - pos + 1 < requiredSpace Then
        ' Calculate the new size based on power of 2
        newSize = 2 * UBound(result)
        Do While newSize - pos + 1 < requiredSpace
            newSize = newSize * 2
        Loop
        ReDim Preserve result(1 To newSize)
    End If
End Sub

Public Sub Append(ParamArray args() As Variant)
    Dim i As Long
    
    ' Ensure there's enough space
    EnsureSpace UBound(args) + 1

    For i = LBound(args) To UBound(args)
        Select Case VarType(args(i))
            Case vbByte                          ' Byte
                result(pos) = args(i)
                pos = pos + 1

            Case vbInteger                       ' Integer
                MemInt(VarPtr(result(pos))) = args(i)
                pos = pos + 2

            Case vbLong                          ' Long
                MemLong(VarPtr(result(pos))) = args(i)
                pos = pos + 4

            Case vbLongLong                      ' LongLong (for 64-bit VBA only)
                MemLongLong(VarPtr(result(pos))) = args(i)
                pos = pos + 8

            Case vbString
                If Len(args(i)) > 2 Then Err.Raise 5, , args(i) & " cannot be a byte literal"
                result(pos) = CByte("&h" & args(i))
                pos = pos + 1

            Case Else
                Err.Raise 5, , TypeName(args(i)) & " not supported in arg " & i + 1
        End Select
    Next i
End Sub

Public Function ToBytes() As Byte()
    ' Trim the array to its actual size
    ReDim Preserve result(1 To pos - 1)
    ToBytes = result
End Function

