Attribute VB_Name = "QuickProfiler"
'@Folder("Libs.Profiler")
Option Explicit

Private Enum BOOL
    S_FALSE = 0
End Enum

Private Declare PtrSafe Function QPF Lib "kernel32" _
Alias "QueryPerformanceFrequency" (ByRef outFrequency As Currency) As BOOL

Private Declare PtrSafe Function QPC Lib "kernel32" _
Alias "QueryPerformanceCounter" (ByRef outTickCount As Currency) As BOOL

Private results As collection
Private counterOUT As Currency

Public Sub ResetMeasurements()
    Set results = New collection
    counterOUT = 0
End Sub

Public Sub Measure()
    Dim counterIN As Currency
    Debug.Assert QPC(counterIN) <> S_FALSE
    
   
    
    results.Add counterOUT
    results.Add counterIN
    
    
    Debug.Assert QPC(counterOUT) <> S_FALSE
End Sub

Public Function ToRange(Optional ByVal topLeft As Range) As Range

    If topLeft Is Nothing Then
        With ThisWorkbook.Worksheets.Add
            .Range("A1") = "Data"
            Set topLeft = .ListObjects.Add(xlSrcRange, Range("A1"), , xlYes).HeaderRowRange.Offset(1, 0)
        End With
    End If
            
    'closing measurement
    Measure
        
    Dim freq As Currency
    Debug.Assert QPF(freq) <> S_FALSE
    
    Dim resultsArray As Variant
    ReDim resultsArray(1 To results.Count, 1 To 1)
    Dim i As Long, result As Variant
    For Each result In results
        i = i + 1
        resultsArray(i, 1) = result / freq
    Next result
    
    topLeft.Resize(results.Count) = resultsArray
    Set ToRange = topLeft
    
End Function
