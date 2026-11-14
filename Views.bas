Attribute VB_Name = "Views"
Sub gotograph()
    Sheets("Analysis").Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    Range("E11").Select
End Sub
Sub gotoprob()
    'Sheets("Tables").Select
    'Range("C2:E2").Select
    Worksheets("Tables").Range("B3:G14").Copy
    With ProbTables
        Set .Picture = PastePicture(xlBitmap)
        .Show 0
        .Top = 120
        .Left = 600
    End With
End Sub
Sub gotoplay()
    Sheets("Display").Select
    Application.GoTo Reference:=Range("B1"), Scroll:=True
    Worksheets("Display").Range(Cells(14 + 1, 3 + 31), Cells(14 + 4, 3 + 31)).Select
    ActiveWindow.DisplayVerticalScrollBar = False
End Sub
Sub PayoutCht()
    Application.GoTo Reference:=Range("N1"), Scroll:=True
    Range("P3").Select
End Sub
Sub WinHandChart()
    Application.GoTo Reference:=Range("A44"), Scroll:=True
    Range("E46").Select
End Sub
Sub gotoshed()
    'Sheets("Tables").Select
    'Application.Goto Reference:=Range("A19"), Scroll:=True
    'Range("C19:F19").Select
    Worksheets("Tables").Range("B20:H41").Copy
    With PayoutTables
        Set .Picture = PastePicture(xlBitmap)
        .Show 0
        .Top = 110
        .Left = 510
    End With
End Sub
Sub gotohand()
    Sheets("Analysis").Select
    Application.GoTo Reference:=Range("A26"), Scroll:=True
    Range("D40").Select
End Sub
Sub gotohandedit()
    Sheets("Tables").Select
    Application.GoTo Reference:=Range("A45"), Scroll:=True
    Range("B60").Select
'    Worksheets("Tables").Range("B46:G74").Copy
'    With AdjustAuto
'        Set .Picture = PastePicture(xlBitmap)
'        .Show 0
'        .Top = 100
'        .Left = 500
'    End With
End Sub
Sub contents()
'    Sheets("Help").Select
'    Application.Goto Reference:=Range("A1"), Scroll:=True
'    Range("D1").Select
'    ActiveWindow.DisplayVerticalScrollBar = True
    With HelpDialog
        Worksheets("Help").Range("B3:D11").Copy
        Set .MultiPage1.Page1.IPage1.Picture = PastePicture(xlBitmap)
        Worksheets("Help").Range("B12:D21").Copy
        Set .MultiPage1.Page2.IPage2.Picture = PastePicture(xlBitmap)
        Worksheets("Help").Range("B22:D32").Copy
        Set .MultiPage1.Page3.IPage3.Picture = PastePicture(xlBitmap)
        .Show 0
        .Top = 50
        .Left = 500
    End With
    
End Sub
Sub HandStrategy()
    ActiveWindow.DisplayVerticalScrollBar = False
    Application.GoTo Reference:=Range("F82"), Scroll:=True
    Range("K95").Select
End Sub
Sub PayScehdule()
    ActiveWindow.DisplayVerticalScrollBar = False
    Application.GoTo Reference:=Range("F64"), Scroll:=True
    Range("K76").Select
End Sub
Sub JokerScehdule()
    ActiveWindow.DisplayVerticalScrollBar = False
    Application.GoTo Reference:=Range("F45"), Scroll:=True
    Range("K58").Select
End Sub
Sub Graphical()
    ActiveWindow.DisplayVerticalScrollBar = False
    Application.GoTo Reference:=Range("F3"), Scroll:=True
    Range("K7").Select
End Sub
Sub Propability()
    ActiveWindow.DisplayVerticalScrollBar = False
    Application.GoTo Reference:=Range("F11"), Scroll:=True
    Range("K16").Select
End Sub
Sub ViewDeck()
    ActiveWindow.DisplayVerticalScrollBar = False
    Application.GoTo Reference:=Range("F18"), Scroll:=True
    Range("K25").Select
End Sub
Sub ViewNext5()
    ActiveWindow.DisplayVerticalScrollBar = False
    Application.GoTo Reference:=Range("F29"), Scroll:=True
    Range("K38").Select
End Sub
Sub WinningHands()
    ActiveWindow.DisplayVerticalScrollBar = False
    Application.GoTo Reference:=Range("F101"), Scroll:=True
    Range("K115").Select
End Sub

