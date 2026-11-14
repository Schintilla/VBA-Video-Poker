Attribute VB_Name = "TestCode"

Sub colourchange()
    Range("c12").Select
    a$ = Range("C12").Value
    n = Len(a$)
    With Range("c12")
        For x = 1 To n / 2 + 1 Step 2
            .Characters(start:=x, Length:=2).Font.ColorIndex = x + 8
            .Characters(start:=n + 1 - x, Length:=2).Font.ColorIndex = x + 8
        Next x
        For x = n / 2 + 1 To 1 Step -2
            .Characters(start:=x, Length:=2).Font.ColorIndex = x + 10
            .Characters(start:=n + 1 - x, Length:=2).Font.ColorIndex = x + 10
        Next x
        .Characters(start:=1, Length:=n).Font.ColorIndex = 3
    End With
End Sub

Sub removedeal()
    Worksheets("Display").DrawingObjects("Deal Draw").Delete
End Sub

Sub dealback()
    Worksheets("Display").DrawingObjects(Array("Up Button", "Down Button")).Delete
    Range("C12").ClearContents
    Range("AH18").ClearContents
    Worksheets("Cards").DrawingObjects("Deal Draw").Copy
    ActiveSheet.Paste Destination:=Worksheets("Display").Cells(15, 34)
End Sub

Sub test3()
    sus$ = "spade"
    Worksheets("Cards").DrawingObjects(sus$).Copy
    ycc = 32
    xcc = 26
    z = 1
    cnumber = 8
    For x = 1 To 3
        seq = Worksheets("Deck Data").Cells(1 + cnumber, 13 + x).Value
        If seq <> 0 Then
            Select Case seq
                Case 1
                    For Y = 1 To 4
                        ActiveSheet.Paste Destination:=ActiveSheet.Cells(ycc + 1 + (Y - 1) * 2, xcc + x)
                        zz$ = sus$ + LTrim(Str(z))
                        ActiveSheet.Shapes(sus$).Select
                        Selection.name = zz$
                        z = z + 1
                        Next Y
                Case 2
                    For Y = 1 To 3
                        ActiveSheet.Paste Destination:=ActiveSheet.Cells(ycc + 1 + (Y - 1) * 3, xcc + x)
                        zz$ = sus$ + LTrim(Str(z))
                        ActiveSheet.Shapes(sus$).Select
                        Selection.name = zz$
                        z = z + 1
                        Next Y
                Case 3
                    For Y = 1 To 2
                        ActiveSheet.Paste Destination:=ActiveSheet.Cells(ycc + 1 + (Y - 1) * 6, xcc + x)
                        zz$ = sus$ + LTrim(Str(z))
                        ActiveSheet.Shapes(sus$).Select
                        Selection.name = zz$
                        z = z + 1
                        Next Y
                Case 4
                    ActiveSheet.Paste Destination:=ActiveSheet.Cells(ycc + 1 + 3, xcc + x)
                    zz$ = sus$ + LTrim(Str(z))
                    ActiveSheet.Shapes(sus$).Select
                    Selection.name = zz$
                    z = z + 1
                Case 5
                    ActiveSheet.Paste Destination:=ActiveSheet.Cells(ycc + 1 + 1, xcc + x)
                    zz$ = sus$ + LTrim(Str(z))
                    ActiveSheet.Shapes(sus$).Select
                    Selection.name = zz$
                    z = z + 1
                Case 6
                    ActiveSheet.Paste Destination:=ActiveSheet.Cells(ycc + 1 + 1, xcc + x)
                    zz$ = sus$ + LTrim(Str(z))
                    ActiveSheet.Shapes(sus$).Select
                    Selection.name = zz$
                    z = z + 1
                    ActiveSheet.Paste Destination:=ActiveSheet.Cells(ycc + 1 + 5, xcc + x)
                    zz$ = sus$ + LTrim(Str(z))
                    ActiveSheet.Shapes(sus$).Select
                    Selection.name = zz$
                    z = z + 1
            End Select
        End If
        Next x
    Select Case cnumber
        Case 1
            ActiveSheet.Shapes.Range(Array("spade1")).Select
        Case 2
            ActiveSheet.Shapes.Range(Array("spade1", "spade2")).Select
        Case 3
            ActiveSheet.Shapes.Range(Array("spade1", "spade2", "spade3")).Select
        Case 4
            ActiveSheet.Shapes.Range(Array("spade1", "spade2", "spade3", "spade4")).Select
        Case 5
            ActiveSheet.Shapes.Range(Array("spade1", "spade2", "spade3", "spade4", "spade5")).Select
        Case 6
            ActiveSheet.Shapes.Range(Array("spade1", "spade2", "spade3", "spade4", "spade5", "spade6")).Select
        Case 7
            ActiveSheet.Shapes.Range(Array("spade1", "spade2", "spade3", "spade4", "spade5", "spade6", "spade7")).Select
        Case 8
            ActiveSheet.Shapes.Range(Array("spade1", "spade2", "spade3", "spade4", "spade5", "spade6", "spade7", "spade8")).Select
        Case 9
            ActiveSheet.Shapes.Range(Array("spade1", "spade2", "spade3", "spade4", "spade5", "spade6", "spade7", "spade8", "spade9")).Select
        Case 10
            ActiveSheet.Shapes.Range(Array("spade1", "spade2", "spade3", "spade4", "spade5", "spade6", "spade7", "spade8", "spade9", "spade10")).Select
    End Select
    If cnumber <> 1 Then
        Selection.Shapes.Select
        Selection.Cut
    Else
        Selection.name = "spade0"
        Selection.Cut
    End If
End Sub
Sub soundtest()
    sound
End Sub
Sub DislaySuits()
    Dim z As Integer, ycc As Integer, xcc As Integer, x As Integer
    Dim c As Integer, b As Integer
    Dim Y As Integer, seq As Integer
    Application.ScreenUpdating = False
    sy = "ª"
    sy = "¨"
    'sy = "©"
    'sy = "§"
    If sy = "ª" Or sy = "¨" Then col = 9
    If sy = "§" Or sy = "©" Then col = 1
    Set ws = Worksheets("Cards")
    ycc = 21
    xcc = 26
    With ws
        For cnumber = 4 To 7
            xcc = 14 + 6 * (cnumber - 3)
            For x = 1 To 3
                seq = Worksheets("Deck Data").Cells(1 + cnumber, 13 + x).Value
                If seq <> 0 Then
                    Select Case seq
                        Case 1
                            For Y = 1 To 4
                                ypos = ycc + (Y - 1) * 3
                                With .Cells(ypos, xcc + x)
                                    .Value = sy
                                    .MergeCells = False
                                    .Font.name = "Symbol"
                                    .Font.Size = 28
                                    .HorizontalAlignment = xlCenter
                                    .VerticalAlignment = xlCenter
                                    .Font.ColorIndex = col
                                End With
                                .Range(.Cells(ypos, xcc + x), .Cells(ypos + 1, xcc + x)).MergeCells = True
                            Next Y
                        Case 2
                            For Y = 1 To 3
                                ypos = ycc - 5 * (Y = 2) - 9 * (Y = 3)
                                With .Cells(ypos, xcc + x)
                                    .Value = sy
                                    .MergeCells = False
                                    .Font.name = "Symbol"
                                    .Font.Size = 28
                                    .HorizontalAlignment = xlCenter
                                    .VerticalAlignment = xlCenter
                                    .Font.ColorIndex = col
                                End With
                                .Range(.Cells(ypos, xcc + x), .Cells(ypos + 1, xcc + x)).MergeCells = True
                            Next Y
                        Case 3
                            For Y = 1 To 2
                                ypos = ycc + (Y - 1) * 9
                                With .Cells(ypos, xcc + x)
                                    .Value = sy
                                    .MergeCells = False
                                    .Font.name = "Symbol"
                                    .Font.Size = 28
                                    .HorizontalAlignment = xlCenter
                                    .VerticalAlignment = xlCenter
                                    .Font.ColorIndex = col
                                End With
                                .Range(.Cells(ypos, xcc + x), .Cells(ypos + 1, xcc + x)).MergeCells = True
                            Next Y
                        Case 4
                            .Cells(ycc + 5, xcc + x) = sy
                        Case 5
                            .Cells(ycc + 2, xcc + x) = sy
                        Case 6
                            .Cells(ycc + 2, xcc + x) = sy
                            .Cells(ycc + 8, xcc + x) = sy
                    End Select

                End If
                With .Cells(ycc, xcc)
                    .Value = cnumber
                    .Font.name = "Ariel Narrow"
                    .Font.Size = 14
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Font.ColorIndex = col
                End With
                .Cells(ycc + 2, xcc) = sy
                With .Cells(ycc + 10, xcc + 4)
                    Value = cnumber
                    .Font.name = "Ariel Narrow"
                    .Font.Size = 14
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Font.ColorIndex = col
                End With
                .Cells(ycc + 9, xcc + 4) = sy
            Next
        Next
        'Range("AR21:AV32").Copy
        'Range("Z21:AD32").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        '   SkipBlanks:=False, Transpose:=False
    End With
    Application.ScreenUpdating = True
End Sub
Sub DislayCards()
    Dim z As Integer, ycc As Integer, xcc As Integer, x As Integer
    Dim c As Integer, b As Integer
    Dim Y As Integer, seq As Integer
    Dim pArray(10) As Variant
    Application.ScreenUpdating = False
    Set ws = Worksheets("Cards")
    With ws
        For s = 1 To 4
            Select Case s
                Case 1
                    cSuit = "spade"
                Case 2
                    cSuit = "diamond"
                Case 3
                    cSuit = "heart"
                Case 4
                    cSuit = "club"
            End Select
            ycc = 1 + 10 * s
            .DrawingObjects(cSuit).Copy
            b = 1
            c = 1
            For cnumber = 1 To 10
                xcc = 14 + 6 * cnumber
                For x = 1 To 3
                    seq = Worksheets("Deck Data").Cells(1 + cnumber, 13 + x).Value
                    If seq <> 0 Then
                        Select Case seq
                            Case 1
                                For Y = 1 To 4
                                    .Paste Destination:=.Cells(ycc + 1 + (Y - 1) * 2, xcc + x)
                                    Set pc = .Shapes(.Shapes.count)
                                    pc.name = CStr(b) + cSuit
                                    b = b + 1
                                    DoEvents
                                Next Y
                            Case 2
                                For Y = 1 To 3
                                    .Paste Destination:=.Cells(ycc + 1 + (Y - 1) * 3, xcc + x)
                                    Set pc = .Shapes(.Shapes.count)
                                    pc.name = CStr(b) + cSuit
                                    b = b + 1
                                    DoEvents
                                Next Y
                            Case 3
                                For Y = 1 To 2
                                    .Paste Destination:=.Cells(ycc + 1 + (Y - 1) * 6, xcc + x)
                                    Set pc = .Shapes(.Shapes.count)
                                    pc.name = CStr(b) + cSuit
                                    b = b + 1
                                    DoEvents
                                Next Y
                            Case 4
                                .Paste Destination:=.Cells(ycc + 1 + 3, xcc + x)
                                Set pc = .Shapes(.Shapes.count)
                                pc.name = CStr(b) + cSuit
                                b = b + 1
                            Case 5
                                .Paste Destination:=.Cells(ycc + 1 + 1, xcc + x)
                                Set pc = .Shapes(.Shapes.count)
                                pc.name = CStr(b) + cSuit
                                b = b + 1
                            Case 6
                                .Paste Destination:=.Cells(ycc + 1 + 1, xcc + x)
                                Set pc = .Shapes(.Shapes.count)
                                pc.name = CStr(b) + cSuit
                                b = b + 1
                                .Paste Destination:=.Cells(ycc + 1 + 5, xcc + x)
                                Set pc = .Shapes(.Shapes.count)
                                pc.name = CStr(b) + cSuit
                                b = b + 1
                        End Select
                    End If
                Next
                Erase pArray()
                For a = 1 To cnumber
                   pArray(a) = CStr(c + a - 1) + cSuit
                Next
                If cnumber > 1 Then
                    Set symb = ActiveSheet.Shapes.Range((pArray)).Group
                    symb.name = CStr(cnumber) + " of " + cSuit
                Else
                    ActiveSheet.Shapes("1" + cSuit).name = CStr(cnumber) + " of " + cSuit
                End If
                c = b
            Next
        Next
    End With
    Application.ScreenUpdating = True
End Sub
