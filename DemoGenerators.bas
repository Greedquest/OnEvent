Attribute VB_Name = "DemoGenerators"
'@IgnoreModule ImplicitPublicMember, ProcedureNotUsed
'@Folder "Generators"
Option Explicit


Sub testNumberRange()
    Dim c As New NumberRange
    c.generatorTo 27

    Dim idx As Long: idx = 1
    Dim val

    For Each val In c                            'btw you cannot loop over ienumvariant, it must have IDispatch + dispid -4
        Debug.Print val
        If idx > 100 Then Exit Sub               ' Just in case of infinite loops
        idx = idx + 1
    Next val
End Sub

Sub doenumerate()
    '    Dim c As New Collection
    '    Dim x As Long
    '    For x = 1 To 20
    '        c.Add "item" & x
    '    Next x

    Dim c As New NumberRange
    c.generatorTo 27


    Dim idx As Long: idx = 1
    Dim i As ReturnLong
    Dim val

    For Each val In Enumerate(c, outIndex:=i)
        Debug.Print i, val
        If idx > 100 Then Exit Sub               ' Just in case of infinite loops
        idx = idx + 1
    Next val
    
    
End Sub

Sub cellsPerhaps()


    'Dim r As Range
    'Set r = ThisWorkbook.Sheets(1).Range("A1:B10")

    Dim r As New NumberRange
    r.generatorTo 27


    Dim x

    For Each x In TQDM(r, 30)
        Debug.Print x
    Next x

End Sub

Sub AllTested()

    Dim r As New NumberRange
    r.generatorTo 27


    Dim i As ReturnLong
    Dim x

    For Each x In Enumerate(r, i) 'Enumerate(TQDM(r, 30), i)
        Debug.Print i, x
        If i = 13 Then End
    Next x

End Sub


Sub test_3()
    Dim r As Range
    Set r = [A1:B3]

    Dim x As Range, i As ReturnLong
    For Each x In Enumerate(TQDM(r), i)
        Debug.Print "i="; i, x.Address
    Next x
End Sub


Sub testRawEnum()
    Dim y As New CountDownGenerator
    y.startFrom = 500
    
    Dim x
    For Each x In y
        'Debug.Print x
        If x = 2 Then End
    Next x
    

End Sub
