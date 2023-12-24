Attribute VB_Name = "ProfilerTests"
'@Folder("Libs.Profiler")
Option Explicit

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub simpletest()
ResetMeasurements
Dim i As Long
For i = 5000 To 1 Step -1
Measure
Measure
Measure
Measure
Measure
Measure
Measure
Measure
Measure
Measure
Next i
ToRange Sheet11.Range("A2")

End Sub
