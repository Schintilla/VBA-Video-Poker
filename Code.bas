Attribute VB_Name = "Code"
Option Base 1
Dim carddeck(53) As Integer, shuffle(53) As Integer, deal As Integer, hand(5) As Integer
Dim pic$(5), suit$(5), winx As Integer, xc As Integer, yc As Integer, hidecard As Boolean
Dim something As Integer, msg As String, cardhold(5) As Integer, straighthold(5) As Integer
Dim handhold As String, flushhold(5) As Integer, pairhold(5) As Integer, highcard As Integer
Dim autop As Boolean, aplay As Integer, autohold As Boolean, coins As Integer
Dim doublego As Integer, cardnumber1 As Integer, cardnumber2 As Integer, decksize As Integer
Dim doubleagain As Integer

Public Declare PtrSafe Function sndPlaySound Lib "winmm.dll" _
    Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long

Sub startgame()
    Dim n As Integer, d As Integer, noplays As Integer
    'Application.ScreenUpdating = False
    Application.ErrorCheckingOptions.NumberAsText = False
    yc = [CardStart].Row
    xc = [CardStart].Column
    coins = 5
    hidecard = False
    ActiveWindow.Zoom = 101
    With ActiveSheet
        If deal = 0 Then
            d = .Range("AJ7").Value
            decksize = 52 - (d = 7)
        End If
        autop = .Range("AJ10").Value
        noplays = 1
        If autop = True And decksize = 52 Then
            hidecard = .Range("AJ11").Value
            deal = 0
            n = .Range("AJ9").Value
            noplays = Worksheets("Deck Data").Cells(1 + n, 17).Value
            noplays = noplays * 2
        End If
        'If noplays > 1 Then Application.OnKey "{ESCAPE}", "haltauto"
        .UnProtect
        .Cells(3, 30).Value = ""
        For aplay = 1 To noplays
            play
            'If deal = 0 Then
                'LiveUpdate
            'End If
        Next aplay
        .Range("AJ10").Value = False
        .Protect DrawingObjects:=True, contents:=False, Scenarios:=True
    End With
    'Application.ScreenUpdating = True
End Sub
Sub LiveUpdate()
    If ChartHands.Visible = True Then
        'Application.ScreenUpdating = True
        HandsChart
        'Application.ScreenUpdating = False
    End If
End Sub

Sub UnProtect()
    ActiveSheet.UnProtect
End Sub
Sub haltauto()
    resp = MsgBox("Auto Play Halted" + Chr$(13) + Chr$(13) + "Continue or Cancel?", 1)
    Application.OnKey "{ESCAPE}"
End Sub
Sub PlayWavFile(WavFileName As String, Wait As Boolean)
    If Dir(WavFileName) = "" Then Exit Sub ' no file to play
    If Wait Then ' play sound before running any more code
        sndPlaySound WavFileName, 0
    Else ' play sound while code is running
        sndPlaySound WavFileName, 1
    End If
End Sub


Sub play()
    Dim n As Integer, xcInc As Integer, i As Integer
    Dim s As Shape, aCard As String
    Set ws = ActiveSheet
    If deal = 0 Then
        PlayNo = ws.Cells(3, 29).Value
        ws.Cells(3, 29).Value = "'" + LTrim(Str$(PlayNo + 1))
        If aplay = 1 Or hidecard = False Then
            ws.Cells(yc - 2, xc - 1).ClearContents
            ws.Range(Cells(yc, xc), Cells(yc + 8, xc + 28)).ClearContents
            ws.GroupObjects.Delete
            allSuits = "ClubSpadeDiamondHeart"
            For Each s In ws.Shapes
                If InStr(s.name, "ard") > 0 Or InStr(allSuits, s.name) > 0 Then
                    s.Delete
                End If
            Next
            'ws.Pictures.Delete
            With ws
                For n = 1 To 5
                    xcInc = (n - 1) * 6
                    .Cells(yc - 1, xc + 2 + xcInc).Font.ColorIndex = 28
                    aCard = "Card Back " + LTrim(Str(n))
                    Worksheets("Cards").DrawingObjects(aCard).Copy
                    'Pause (0.25)
                    On Error Resume Next
back:
                    .Paste Destination:=.Cells(yc, xc + xcInc)
                    If Err.Number = 1004 Then
                        On Error GoTo -1
                        GoTo back
                    End If
                    .Range("F5").Select
                    '.Shapes("Hold" + CStr(n)).Select
                Next n
            End With
        End If
        If aplay > 1 Then ws.Range("C13:AE13").Font.ColorIndex = 28
        ws.Cells(3, 6).Value = 0
        cash = ws.Cells(7, 6).Value - coins
        If cash < 0 Then
            ws.Cells(7, 6).Value = 0
            credit = ws.Cells(5, 6).Value
            If credit = 0 Then
                ws.Cells(7, 6).Value = "'95"
                MsgBox ("No Cash or Credit left. Added another 100 to Cash")
                ws.Range("AJ3") = ws.Range("AJ3") + 100
                With Worksheets("Shuffle")
                    lastplay = Range("PlayNo")
                    For i = 1 To lastplay
                        .Cells(i, 4) = .Cells(i, 4) + 100
                    Next
                End With
            Else
                ws.Cells(5, 6).Value = "'" + LTrim(Str$(credit - coins))
            End If
        Else
            ws.Cells(7, 6).Value = "'" + LTrim(Str$(cash))
        End If
        Application.Calculation = xlCalculationManual
        deck
        deals
        assesshand
        Application.Calculation = xlCalculationAutomatic
        deal = 1
        If hidecard = False Then
            ws.DrawingObjects("Deal Draw").Characters.Text = "Draw"
            For n = 1 To 5
                holdbutton$ = "Hold " + LTrim(Str$(n))
                ws.DrawingObjects(holdbutton$).Font.ColorIndex = 3
            Next n
        End If
    Else
        Application.Calculation = xlCalculationManual
        deals
        assesshand
        Application.Calculation = xlCalculationAutomatic
        deal = 0
        If hidecard = False Then
            ws.DrawingObjects("Deal Draw").Characters.Text = "Deal"
            For n = 1 To 5
                holdbutton$ = "Hold " + LTrim(Str$(n))
                ws.DrawingObjects(holdbutton$).Font.ColorIndex = 1
            Next n
        End If
        If winx <> 0 Then
            ws.Cells(3, 6).Value = "'" + LTrim(Str$(winx))
            If Range("AJ6").Value = True And aplay = 1 And decksize = 52 Then
                doublegame
            Else
                updatescore
            End If
        End If
        PlayNo = ws.Cells(3, 29).Value
        Wins = ws.Cells(4, 29).Value
        If aplay > 1 Then
            ws.Cells(3, 30).Value = "'(" + CStr(aplay / 2) + ")"
        End If
        percwins = Int(Wins * 100 / PlayNo)
        ws.Cells(4, 30).Value = "'(" + LTrim(Str$(percwins)) + "%)"
        payout = ws.Cells(2, 36).Value
        perc = Int(payout * 20 * 100 / PlayNo) / 100
        perpayout$ = "'" + LTrim(Str$(perc)) + "%"
        ws.Cells(5, 29).Value = perpayout$
        With Worksheets("Shuffle")
            .Cells(1 + PlayNo, 4).Value = .Cells(PlayNo, 4).Value + winx - 5
            .Cells(1 + PlayNo, 22).Value = ws.Range("AJ3")
            .Cells(1 + PlayNo, 7).Value = perc / 100
            .Cells(1 + PlayNo, 21).Value = Replace(ws.Cells(6, 29).Value, "%", "") / 100
            Call xaxismax
            Call xaxismaxPayout
        End With
    End If
    ws.Shapes("Outerbox").Fill.Visible = msoFalse
    If Range("AJ6").Value = False Or winx = 0 Or deal = 1 Or decksize = 53 Then
        ws.Range(Cells(yc + 1, xc + 31), Cells(yc + 4, xc + 31)).Select
    Else
        ws.Range("B1").Select
    End If
End Sub



Sub updatescore()
    Dim n As Integer
    With ActiveSheet
        .Cells(3, 6).Value = "'" + CStr(winx)
        Wins = .Cells(4, 29).Value
        .Cells(4, 29).Value = "'" + CStr(Wins + 1)
        .Cells(5, 6).Select
        For n = 1 To winx
            credit = .Cells(5, 6).Value
            .Cells(5, 6).Value = "'" + CStr(credit + 1)
            For x = 1 To 10000
            Next x
            'Beep
            PlayWavFile "c:\windows\media\tada.wav", False
        Next n
        .Cells(2, 36).Value = .Cells(2, 36).Value + winx
    End With
End Sub

Sub deck()
    Dim p As Integer, x As Integer, n As Integer, jcard As Boolean
    If ActiveSheet.Cells(8, 36) = 2 And shuffle(1) > 0 Then
        jcard = False
        For p = 1 To decksize
            carddeck(p) = shuffle(p)
            If shuffle(p) = 53 And decksize = 52 Then
                carddeck(p) = shuffle(53)
            End If
            If shuffle(p) = 53 And decksize = 53 Then
                jcard = True
            End If
        Next p
        If decksize = 53 And jcard = False Then
            carddeck(53) = 53
        End If
        'Worksheets("Shuffle").Range("J48").Value = "shuffled deck"
    Else
        For p = 1 To decksize
            carddeck(p) = p
        Next p
        'Worksheets("Shuffle").Range("J48").Value = "new deck"
    End If
    For x = decksize To 1 Step -1
        Randomize
        randomno = Int((x) * Rnd + 1)
        shuffle(x) = carddeck(randomno)
        For n = randomno To decksize - (decksize + 1 - x)
            carddeck(n) = carddeck(n + 1)
        Next n
    Next x
    'carddeck = ShuffleDeckOfCards
End Sub
Function ShuffleDeckOfCards() As Variant
    Dim suits As Variant
    Dim ranks As Variant
    Dim deck(1 To 53) As String
    Dim i As Integer, j As Integer
    Dim temp As String

    ' Initialize suits and ranks
    suits = Array("Hearts", "Diamonds", "Clubs", "Spades")
    ranks = Array("2", "3", "4", "5", "6", "7", "8", "9", "10", "Jack", "Queen", "King", "Ace")

    ' Create the deck of cards
    For i = 1 To 4
        For j = 1 To 13
            deck((i - 1) * 13 + j) = ranks(j) & " of " & suits(i)
        Next j
    Next i

    ' Add Jokers
    deck(53) = "Joker"

    ' Shuffle the deck using Fisher-Yates algorithm
    For i = 53 To 2 Step -1
        j = Int((i - 1) * Rnd) + 1
        temp = deck(i)
        deck(i) = deck(j)
        deck(j) = temp
    Next i

    ' Return the shuffled deck
    ShuffleDeckOfCards = deck
End Function


Sub deals()
    Dim k As Integer, j As Integer, cardnumber As Integer, n As Integer, m As Integer
    Dim start As Integer, finish As Integer
    'Application.Calculation = xlCalculationManual
    If Range("AJ6").Value = True And deal = 3 Then
        start = doublego
        finish = start
        j = doublego
    Else
        j = -5 * (deal = 1) + 1
        finish = 5
        start = 1
    End If
    For k = start To finish
        If deal = 1 And ActiveSheet.Cells(yc - 1, xc + 2 + (k - 1) * 6).Font.ColorIndex = 7 Then GoTo nextcard
        cardsuit$ = Worksheets("Deck Data").Cells(shuffle(j), 4).Value
        suit$(k) = cardsuit$
        If cardsuit$ = "J" Then
            cardnumber = -15
        Else
            cardnumber = ((shuffle(j) - 1) Mod 13) + 1
        End If
        If hidecard = True Then GoTo bypassdisp
        Application.ScreenUpdating = False
        If deal <> 1 Then
            a$ = "Card Back " + LTrim(Str(k))
            ActiveSheet.Pictures(a$).Delete
        Else
            If pic$(k) <> "99" Then
                ActiveSheet.DrawingObjects(pic$(k)).Delete
            Else
            cardname$ = "card set " + LTrim(Str$(k))
                ActiveSheet.DrawingObjects(cardname$).Delete
            End If
        End If
        If cardnumber > 10 Or cardnumber = -15 Then
            If cardnumber = -15 Then
                cardname$ = "Joker Card 1"
            Else
                n = (cardnumber - 10) * 4 - 3 + 170 - Asc(cardsuit$)
                cardname$ = "Card " + LTrim(Str$(n))
            End If
            Worksheets("Cards").DrawingObjects(cardname$).Copy
            'Pause (0.25)
            On Error Resume Next
back:
            ActiveSheet.Paste Destination:=ActiveSheet.Cells(yc, xc + (k - 1) * 6)
            If Err.Number = 1004 Then
                On Error GoTo -1
                GoTo back
            End If
            ActiveSheet.DrawingObjects(cardname$).SendToBack
            pic$(k) = cardname$
        Else
            symbol$ = Worksheets("Deck Data").Cells(cardnumber, 3).Value
            If cardsuit$ = "©" Or cardsuit$ = "¨" Then
                n = 9
            Else
                n = 1
            End If
            With ActiveSheet
                xinc = (k - 1) * 6
                .Cells(yc, xc + xinc).Value = symbol$
                .Cells(yc, xc + xinc).Font.ColorIndex = n
                .Cells(yc + 8, xc + 4 + xinc).Value = symbol$
                .Cells(yc + 8, xc + 4 + xinc).Font.ColorIndex = n
                .Cells(yc + 1, xc + xinc).Value = cardsuit$
                .Cells(yc + 1, xc + xinc).Font.ColorIndex = n
                .Cells(yc + 7, xc + 4 + xinc).Value = cardsuit$
                .Cells(yc + 7, xc + 4 + xinc).Font.ColorIndex = n
            End With
            Call cardface(cardsuit$, cardnumber, k)
            pic$(k) = "99"
        End If
        'Worksheets("Display").DrawingObjects("Deal Draw").Select
        'SendKeys "{Esc}"
        'DoEvents
        Range("F5").Select
        'ActiveSheet.Shapes("Hold" + CStr(cardnumber)).Select
        'Application.CutCopyMode = False
        Application.ScreenUpdating = True
bypassdisp:
        'SendKeys "{Esc}"
        DoEvents
        hand(k) = shuffle(j)
        j = j + 1
nextcard:
    Next k
    'ActiveSheet.Range("AH16").Select
    'Application.Calculation = xlCalculationAutomatic
    'ActiveSheet.Shapes("Outerbox").Fill.Visible = msoFalse
End Sub
Sub cardface(sy As String, cnumber As Integer, slot As Integer)
    'Application.ScreenUpdating = False
    Dim p As Integer, m As Integer, xcInc As Integer, cName As String
    If sy = "ª" Then sus$ = "spade"
    If sy = "¨" Then sus$ = "diamond"
    If sy = "©" Then sus$ = "heart"
    If sy = "§" Then sus$ = "club"
    cName = CStr(cnumber) + " of " + sus$
    Worksheets("Cards").Shapes(cName).Copy
    p = -(cnumber < 4)
    m = -3 * (cnumber = 1) + 1
    xcInc = p + 1 + (slot - 1) * 6
    'Pause (0.25)
    On Error Resume Next
back:
    ActiveSheet.Paste Destination:=Cells(yc + m, xc + xcInc)
    If Err.Number = 1004 Then
        On Error GoTo -1
        GoTo back
    End If
    With ActiveSheet.Shapes(cName)
        .IncrementTop -2
        cardname$ = "card set " + LTrim(Str$(slot))
        .name = cardname$
        .ZOrder 1
    End With
    'Application.ScreenUpdating = True
End Sub
Sub ErrorCheck()
    
End Sub
Sub assesshand()
    Dim x As Integer, Y As Integer, d As Integer, e As Integer, f As Integer, c As Integer
    Dim jokpos As Integer, aceno As Integer, mn As Integer, mnn As Integer, mx As Integer
    Set ws = Worksheets("Shuffle")
    'MsgBox "Assesshand"
    joker = False
    For x = 1 To 5
        ws.Cells(47 + x, 16).Value = ((hand(x) - 1) Mod 13) + 1
        ws.Cells(47 + x, 14).Value = Asc(suit$(x))
        If suit$(x) = "J" Then
            ws.Cells(47 + x, 16).Value = 16
            joker = True
            jokpos = x
        End If
    Next x
    pairfound$ = ""
    flushfound$ = ""
    straightfound$ = ""
    highcard = 1
    checkhand ("Flush")
    If something = 1 Then
        flushfound$ = msg
        fhold$ = ""
        For x = 1 To 5
            flushhold(x) = cardhold(x)
            If cardhold(x) <> 0 Then fhold$ = fhold$ + LTrim(Str$(cardhold(x)))
        Next x
    End If
    checkhand ("pair")
    If something = 1 Then
        pairfound$ = msg
        phold$ = ""
        sumcards = 0
        For x = 1 To 5
            pairhold(x) = cardhold(x)
            sumcards = sumcards + cardhold(x)
            If cardhold(x) <> 0 Then phold$ = phold$ + LTrim(Str$(cardhold(x)))
        Next
        If pairfound$ = "4 Of A Kind" Then
            kickerpos = 15 - sumcards
            kicker = ws.Cells(47 + kickerpos, 16).Value
            quadcard = ws.Cells(47 + cardhold(1), 16).Value
        End If
    End If
    acepos = ws.Cells(47 + 1, 18).Value
    aceno = ws.Cells(47 + 1, 19).Value
    msg = ""
    If joker = True Then
        mn = ws.Cells(47 + 1, 18).Value + 1
        mx = ws.Cells(47 + 4, 18).Value - 1
        mnn = ws.Cells(47 + 2, 18).Value + 1
        If mx - mn = 2 Then
            For x = mn To mx
                ws.Cells(47 + jokpos, 16).Value = x
                checkhand ("Straight")
                If msg = "Straight" Then Exit For
            Next x
            ws.Cells(47 + jokpos, 16).Value = 16
        End If
        If acepos = 1 And msg <> "Straight" Then
            ws.Cells(47 + aceno, 16).Value = 14
            mx = 14 - 1
            If mx - mnn = 2 Then
                For x = mnn To mx
                    ws.Cells(47 + jokpos, 16).Value = x
                    checkhand ("Straight")
                    If msg = "Straight" Then Exit For
                Next x
                ws.Cells(47 + jokpos, 16).Value = 16
            End If
            ws.Cells(47 + aceno, 16).Value = 1
        End If
    End If
    If msg <> "Straight" Then
        msg = ""
        checkhand ("Straight")
        msgone = msg
        somegone = something
        If acepos = 1 And (msg = "3 Card Straight" Or msg = "4 Card Straight" Or something <> 1) Then
            ws.Cells(47 + aceno, 16).Value = 14
            checkhand ("Straight")
            If something <> 1 Then
                msg = msgone
                something = somegone
            End If
            ws.Cells(47 + aceno, 16).Value = 1
        End If
    End If
    ws.Range("P48:Q52").Copy Destination:=ws.Range("R48")
    ws.Range("R48:S52").Sort Key1:=ws.Range("R48"), Order1:=xlAscending
    'MsgBox msg + " " + Str(something)
    'ws.Range("H46").Value = msg
    If something = 1 Then
        straightfound$ = msg
        shold$ = ""
        For x = 1 To 5
            straighthold(x) = cardhold(x)
            If cardhold(x) <> 0 Then shold$ = shold$ + LTrim(Str(cardhold(x)))
        Next x
    End If
    msg = Chr$(13)
    If highcard = 1 Or straightfound$ = "3 Card Straight" Or straightfound$ = "4 Card Straight" Or flushfound$ = "3 Card Flush" Then
        nohighcards = 0
        If ws.Cells(47 + 5 + (joker = True), 18).Value > 10 Then nohighcards = 1
        If ws.Cells(47 + 4 + (joker = True), 18).Value > 10 Then nohighcards = nohighcards + 1
        If ws.Cells(47 + 3 + (joker = True), 18).Value > 10 Then nohighcards = nohighcards + 1
        'MsgBox Str(nohighcards)
        If acepos = 1 Then
            cardhold(1) = ws.Cells(47 + 1, 19).Value
            Select Case nohighcards
            Case 0
                msg = msg + "Ace Only:" + LTrim(Str$(cardhold(1))) + Chr$(13)
            Case 1
                cardhold(2) = ws.Cells(47 + 5 + (joker = True), 19).Value
                msg = msg + "Ace + Face Card:" + LTrim(Str$(cardhold(1))) + LTrim(Str$(cardhold(2))) + Chr$(13)
            Case 2
                w = ws.Cells(47 + cardhold(1), 14).Value
                cardhold(2) = ws.Cells(47 + 4 + (joker = True), 19).Value
                cardhold(3) = ws.Cells(47 + 5 + (joker = True), 19).Value
                e = ws.Cells(47 + cardhold(2), 14).Value
                f = ws.Cells(47 + cardhold(3), 14).Value
                If w = f Then msg = msg + "Ace + Face Card (suit):" + LTrim(Str$(cardhold(1))) + LTrim(Str$(cardhold(3))) + Chr$(13)
                If w = e Then msg = msg + "Ace + Face Card (suit):" + LTrim(Str$(cardhold(1))) + LTrim(Str$(cardhold(2))) + Chr$(13)
                e = ws.Cells(47 + 5 + (joker = True), 18).Value
                If e = 13 Then msg = msg + "Ace + Face Card (straight):" + LTrim(Str$(cardhold(1))) + LTrim(Str$(cardhold(3))) + Chr$(13)
            End Select
        End If
        If nohighcards <> 0 Then
            For x = 1 To nohighcards
                cardhold(x) = ws.Cells(47 + 6 - x + (joker = True), 19).Value
            Next x
            Select Case nohighcards
                Case 1
                    cc$ = LTrim(Str$(cardhold(1)))
                    cc$ = "1 Face Card:" + cc$
                Case 2
                    d = ws.Cells(47 + cardhold(1), 14).Value
                    e = ws.Cells(47 + cardhold(2), 14).Value
                    cc$ = LTrim(Str(cardhold(1))) + LTrim(Str(cardhold(2)))
                    If d = e Then
                        cc$ = "2 Face Cards (suit):" + cc$
                    Else
                        cc$ = "2 Face Cards:" + cc$
                    End If
                Case 3
                    d = ws.Cells(47 + cardhold(1), 14).Value
                    e = ws.Cells(47 + cardhold(2), 14).Value
                    f = ws.Cells(47 + cardhold(3), 14).Value
                    cc$ = "2 Face Cards:" + LTrim(Str(cardhold(2))) + LTrim(Str(cardhold(3)))
                    If d = f Then cc$ = "2 Face Cards (suit):" + LTrim(Str(cardhold(1))) + LTrim(Str(cardhold(3)))
                    If d = e Then cc$ = "2 Face Cards (suit):" + LTrim(Str(cardhold(1))) + LTrim(Str(cardhold(2)))
                    If e = f Then cc$ = "2 Face Cards (suit):" + LTrim(Str(cardhold(2))) + LTrim(Str(cardhold(3)))
            End Select
            msg = msg + Chr$(13) + cc$ + Chr(13)
            If ws.Cells(47 + 5 - nohighcards, 18).Value = 10 Then
                cardhold(1) = ws.Cells(47 + 5 - nohighcards, 19).Value
                d = ws.Cells(47 + cardhold(1), 14).Value
                For x = nohighcards To 1 Step -1
                    cardhold(2) = ws.Cells(47 + 6 - x + (joker = True), 19).Value
                    e = ws.Cells(47 + cardhold(2), 14).Value
                    f = ws.Cells(47 + 6 - x + (joker = True), 18).Value
                    If d = e And f <> 13 Then
                        cc$ = LTrim(Str$(cardhold(1))) + LTrim(Str$(cardhold(2)))
                        msg$ = msg + "Face + 10 (suit):" + cc$
                        Exit For
                    End If
                Next
            End If
        End If
        If msg = Chr$(13) Then msg = msg + "Nothing"
        msg = msg + Chr$(13)
    End If
    'MsgBox msg
    If highcard = 0 Then
        rr$ = "Straight Flush"
        e = 0
        If flushfound$ <> "" Then
            cc$ = ""
            For x = 1 To Len(fhold$)
                d = ws.Cells(47 + flushhold(x), 16).Value
                If d > 9 Or d = 1 Then
                    cc$ = cc$ + LTrim(Str(flushhold(x)))
                    e = e + 1
                End If
            Next
            If e > 2 And e < 5 Then msg = msg + LTrim(Str(e)) + " Card Royal Flush" + ":" + cc$ + Chr$(13)
        End If
        If straightfound$ = "Straight" And flushfound$ = "Flush" Then
            If acepos = 1 And e = 5 Then
                msg = msg + "Royal Flush:" + shold$ + Chr$(13)
            Else
                msg = msg + rr$ + ":" + shold$ + Chr$(13)
            End If
        End If
        If straightfound$ = "4 Card Straight" And flushfound$ = "4 Card Flush" Then
            matchcard = 1
            For x = 1 To 4
                mc = 0
                For Y = 1 To 4
                    If straighthold(x) = flushhold(Y) Then
                        mc = 1
                    End If
                Next Y
                If mc = 0 Then matchcard = 0
            Next x
            If matchcard = 1 Then msg = msg + "4 Card " + rr$ + ":" + shold$ + Chr$(13)
        End If
        If straightfound$ = "3 Card Straight" And flushfound$ = "3 Card Flush" Then
            matchcard = 1
            For x = 1 To 3
                mc = 0
                For Y = 1 To 3
                    If straighthold(x) = flushhold(Y) Then
                        mc = 1
                    End If
                Next Y
                If mc = 0 Then matchcard = 0
            Next x
            If matchcard = 1 Then msg = msg + "3 Card " + rr$ + ":" + shold$ + Chr$(13)
        End If
        If straightfound$ = "3 Card Straight" And flushfound$ = "4 Card Flush" Then
            matchcard = 1
            For x = 1 To 3
                mc = 0
                For Y = 1 To 4
                    If straighthold(x) = flushhold(Y) Then
                        mc = 1
                    End If
                Next Y
                If mc = 0 Then matchcard = 0
            Next x
            If matchcard = 1 Then msg = msg + "3 Card " + rr$ + ":" + shold$ + Chr$(13)
        End If
        If straightfound$ = "4 Card Straight" And flushfound$ = "3 Card Flush" Then
            matchcard = 1
            For x = 1 To 3
                mc = 0
                For Y = 1 To 4
                    If straighthold(Y) = flushhold(x) Then
                        mc = 1
                    End If
                Next Y
                If mc = 0 Then matchcard = 0
            Next x
            If matchcard = 1 Then msg = msg + "3 Card " + rr$ + ":" + fhold$ + Chr$(13)
        End If
        If straightfound$ <> "" Then msg = msg + straightfound$ + ":" + shold$ + Chr$(13)
        If flushfound$ <> "" Then msg = msg + flushfound$ + ":" + fhold$ + Chr$(13)
        If pairfound$ <> "" Then msg = msg + pairfound$ + ":" + phold$
    End If
    'MsgBox flushfound$ + " " + straightfound$
    If joker = True And straightfound$ = "Straight" Then
            If flushfound$ = "4 Card Flush" Then msg = msg + "4 Card Straight Flush:" + "2345" + Chr$(13)
            If InStr(msg, "4 Card Royal Flush") > 0 Then msg = msg + "4 Card Royal Flush:" + "2345" + Chr$(13)
    End If
    'MsgBox msg
    msg = msg + Chr$(13)
    handhold = ""
    If joker = True Then
        For Y = 1 To 30
            c = Worksheets("Joker").Cells(2 + Y, 3).Font.ColorIndex
            If deal = 1 And c = 2 Then
                GoTo nxthd
            End If
            chand$ = RTrim(Worksheets("Joker").Cells(2 + Y, 3).Value)
            'chxx$ = RTrim(Worksheets("Joker").Cells(2 + y, 4).Value)
            'chxy$ = RTrim(Worksheets("Joker").Cells(2 + y, 5).Value)
            If chand$ = "Nothing" Or chand$ = "" Then
                chand$ = "Nothing"
                chxy$ = ""
                Exit For
            End If
            h$ = Chr$(13) + chand$
            d = InStr(msg, h$)
            If d > 0 Then
                'chxx$ = RTrim(Worksheets("Joker").Cells(2 + y, 4).Value)
                chand$ = RTrim(Worksheets("Joker").Cells(2 + Y, 4).Value)
                chxy$ = RTrim(Worksheets("Joker").Cells(2 + Y, 5).Value)
                Exit For
            End If
nxthd:
        Next Y
    Else
        For Y = 1 To 30
            c = Worksheets("Tables").Cells(46 + Y, 2).Font.ColorIndex
            If deal = 1 And c = 2 Then
                GoTo nxthand
            End If
            chand$ = RTrim(Worksheets("Tables").Cells(46 + Y, 2).Value)
            If chand$ = "Nothing" Or chand$ = "" Then
                chand$ = "Nothing"
                Exit For
            End If
            h$ = Chr$(13) + chand$
            d = InStr(msg, h$)
            If d > 0 Then
                e = InStr(d + 1, msg, Chr$(13))
                handhold = Mid(msg, d + Len(h$) + 1, e - (d + Len(h$) + 1))
                Exit For
            End If
nxthand:
        Next Y
    End If
    handdetail = ActiveSheet.Range("AJ12").Value
    If handdetail = False And c = 2 Then
        handwin$ = "Nothing"
    Else
        handwin$ = chand$
    End If
    If joker = True Then
        handwin$ = chand$
        If deal = 0 Then
            ActiveSheet.Cells(yc - 2, xc - 1).Value = handwin$ + " " + chxy$
        Else
            ActiveSheet.Cells(yc - 2, xc - 1).Value = handwin$
        End If
    Else
        If hidecard = False Then ActiveSheet.Cells(yc - 2, xc - 1).Value = handwin$
    End If
    'ActiveSheet.Range("A1").Select
    'If deal = 0 Then
        'MsgBox msg + "Best hand =" + chand$ + ", Cards to hold " + handhold + Chr$(13) + "Winning Hand - " + handwin$
        'End
    'End If
    autohold = ActiveSheet.Range("AJ13").Value
    If deal = 0 And (autop = True Or autohold = True) And decksize = 52 Then
        For x = 1 To Len(handhold)
            d = Val(Mid(handhold, x, 1))
            Select Case d
            Case 1
                bbut1
            Case 2
                bbut2
            Case 3
                bbut3
            Case 4
                bbut4
            Case 5
                bbut5
            End Select
        Next
    End If
    winx = 0
    For x = 1 To 9
        If ActiveSheet.Cells(2 + x, 15).Font.ColorIndex = 5 Then
            With ActiveSheet.Range(Cells(2 + x, 14), Cells(2 + x, 22))
                .Font.ColorIndex = 6
                .Interior.ColorIndex = 5
            End With
        End If
        If ActiveSheet.Cells(2 + x, 15).Value = handwin$ Then
            If deal = 1 Then
                Worksheets("Analysis").Cells(30 + x, 9).Value = Worksheets("Analysis").Cells(30 + x, 9).Value + 1
            End If
            With ActiveSheet.Range(Cells(2 + x, 14), Cells(2 + x, 22))
                .Font.ColorIndex = 5
                .Interior.ColorIndex = 6
            End With
            winx = ActiveSheet.Cells(2 + x, 21).Value
            If deal = 0 Then
                'Beep
                PlayWavFile "c:\windows\media\chimes.wav", False
            End If
        End If
    Next x
    If joker = True And deal = 1 Then
        If chand$ = "5 Of A Kind" Then
            'Beep
            PlayWavFile "c:\windows\media\wmpaud2.wav", False
            winx = 1000
            MsgBox "<<<<<<< 5 Of A Kind >>>>>>>" + chr13 + chr13 + "YOU WIN 1000!!!!!!!"
        End If
        If chand$ = "Royal Flush (Joker)" Then
            'Beep
            PlayWavFile "c:\windows\media\wmpaud5.wav", False
            winx = 500
            MsgBox "******* ROYAL FLUSH *****" + chr13 + chr13 + "WITH JOKER!!!!!"
        End If
    End If
    If winx = 0 And deal = 1 Then Worksheets("Analysis").Cells(30 + 10, 9).Value = Worksheets("Analysis").Cells(30 + 10, 9).Value + 1
    d = ActiveSheet.Range("AJ7").Value
    If deal = 1 And handwin$ = "Royal Flush" Then
        'Beep
        PlayWavFile "c:\windows\media\wmpaud9.wav", False
        MsgBox "****** C O N G R A T U L A T I O N S *********" + Chr(13) + Chr(13) + "You've hit the JACKPOT !!!!!"
        If decksize = 53 Then winx = 4000
    End If
    If decksize = 52 And joker = False Then
        If pairfound$ = "4 Of A Kind" And d > 1 Then
            If d < 5 Then
                Select Case quadcard
                Case 1
                     winx = Worksheets("Tables").Cells(33, 2 + d).Value
                Case 2, 3, 4
                    winx = Worksheets("Tables").Cells(35, 2 + d).Value
                Case 5 To 13
                    winx = Worksheets("Tables").Cells(37, 2 + d).Value
                End Select
                If quadcard < 5 And d = 4 And kicker < 5 Then winx = winx * 2
            End If
            If d > 4 Then
                Select Case quadcard
                Case 1
                    winx = Worksheets("Tables").Cells(33, 2 + d).Value
                Case 11, 12, 13
                    winx = Worksheets("Tables").Cells(34, 2 + d).Value
                Case 2 To 10
                    winx = Worksheets("Tables").Cells(36, 2 + d).Value
                End Select
                If (quadcard = 1 Or quadcard > 10) And (kicker > 10 Or kicker = 1) Then winx = winx * 2
            End If
            If autop = False Then MsgBox "Well Done!!!!!!!!!!!!!!" + Chr$(13) + "You win - " + LTrim(Str(winx)) + Chr$(13)
        End If
    End If
End Sub



Sub checkhand(chktype As String)
    Dim stno(5, 5) As Integer, patternlength(5) As Integer, d As Integer, e As Integer
    Dim h As Integer, x As Integer, Y As Integer, f As Integer, k As Integer
    For x = 1 To 5
        cardhold(x) = 0
    Next x
    With Worksheets("Shuffle")
        If chktype = "Flush" Then
            .Range("N48:O52").Copy Destination:=.Range("R48")
        Else
            .Range("P48:Q52").Copy Destination:=.Range("R48")
        End If
        .Range("R48:S52").Sort Key1:=.Range("R48"), Order1:=xlAscending
    End With
    f = 1
    patternno = 1
    For x = 1 To 5
        With Worksheets("Shuffle")
            d = .Cells(47 + x, 18).Value
            e = .Cells(48 + x, 18).Value
            If chktype = "Straight" Then e = e - 1
            h = .Cells(47 + x, 19).Value
        End With
        If d = e Then
            stno(patternno, f) = h
            f = f + 1
            patternlength(patternno) = patternlength(patternno) + 1
        Else
            stno(patternno, f) = h
            f = 1
            patternno = patternno + 1
        End If
    Next x
    'Debug.Print "-------------------------"
    'For x = 1 To 5
    '    For y = 1 To 5
    '        Debug.Print stno(x, y);
    '    Next y
    '    Debug.Print "Length="; patternlength(x) + 1
    'Next x
    If chktype = "pair" Then
        f = 1
    Else
        f = 2
    End If
    something = 0
    nocards = 0
    k = 0
    For x = 1 To patternno - 1
        If patternlength(x) + 1 > f Then
            'Debug.Print "Cards to hold:";
            For Y = 1 To patternlength(x) + 1
                'Debug.Print stno(x, y);
                cardhold(Y + k) = stno(x, Y)
            Next Y
            k = k + Y - 1
            If nocards < patternlength(x) + 1 Then nocards = patternlength(x) + 1
            something = 1
        End If
        'Debug.Print
    Next x
    If something = 1 Then
        highcard = 0
        If chktype = "pair" Then
            d = Worksheets("Shuffle").Cells(47 + cardhold(1), 16).Value
            If d = 1 Or d > 10 Then
                msg = "Jacks Or Better"
            Else
                msg = "Tens Or Lower"
            End If
            If nocards = 3 Then msg = "3 Of A Kind"
            If nocards = 3 And patternno = 3 Then msg = "Full House"
            If nocards = 2 And patternno = 4 Then msg = "2 Pairs"
            If nocards = 4 Then msg = "4 Of A Kind"
        Else
            If nocards = 5 Then
                msg = chktype
            Else
                msg = LTrim(Str$(nocards)) + " Card " + chktype
            End If
        End If
    End If
End Sub
Sub doublegame()
    Dim n As Integer, x As Integer
    Range("E13:AC13").Font.ColorIndex = 28
    doubleagain = 0
    With ActiveSheet
        .DrawingObjects("Deal Draw").Visible = False
        .DrawingObjects("No Button").Visible = True
        .DrawingObjects("Yes Button").Visible = True
    End With
    a$ = "Play double to win " + Str(winx * (doubleagain + 1) * 2) + " ?"
    Range("B12").Value = a$
    n = Len(a$)
    With Range("B12")
        For x = 1 To n / 2 + 1 Step 2
            .Characters(start:=x, Length:=2).Font.ColorIndex = x + 3
            .Characters(start:=n + 1 - x, Length:=2).Font.ColorIndex = x + 3
            Beep
        Next x
        For x = n / 2 + 1 To 1 Step -2
            .Characters(start:=x, Length:=2).Font.ColorIndex = x + 7
            .Characters(start:=n + 1 - x, Length:=2).Font.ColorIndex = x + 7
            Beep
        Next x
        .Characters(start:=1, Length:=n).Font.ColorIndex = 3
    End With
End Sub

Sub yesdouble()
    Dim n As Integer
    Set ws = ActiveSheet
    doubleagain = doubleagain + 1
    doublego = 1
    deck
    With ws
        .UnProtect
        .Range(.Cells(yc, xc), .Cells(yc + 8, xc + 28)).ClearContents
        .GroupObjects.Delete
        .Pictures.Delete
        For n = 1 To 5
            .Cells(yc - 1, xc + 2 + (n - 1) * 6).Font.ColorIndex = 28
            a$ = "Card Back " + LTrim(Str(n))
            Worksheets("Cards").DrawingObjects(a$).Copy
            .Paste Destination:=.Cells(yc, xc + (n - 1) * 6)
            .DrawingObjects(a$).SendToBack
        Next n
        deal = 3
        deals
        .DrawingObjects("No Button").Visible = False
        .DrawingObjects("Yes Button").Visible = False
        .Range("B12").Value = "Pick a card"
        .Range("B1").Select
        .Protect DrawingObjects:=True, contents:=False, Scenarios:=True
    End With
End Sub

Sub nodouble()
    With ActiveSheet
        .UnProtect
        .DrawingObjects("No Button").Visible = False
        .DrawingObjects("Yes Button").Visible = False
        .Range("B12").Value = "Game Over"
        .Range("AH18").ClearContents
        double_return
    End With
End Sub

Sub double_check(held As Integer)
    Dim n As Integer
    ActiveSheet.UnProtect
    doublego = held
    deals
    cardnumber1 = ((shuffle(1) - 1) Mod 13) + 1
    cardnumber2 = ((shuffle(doublego) - 1) Mod 13) + 1
    cardnumber1 = cardnumber1 - (cardnumber1 = 1) * 13
    cardnumber2 = cardnumber2 - (cardnumber2 = 1) * 13
    If cardnumber2 < cardnumber1 Then
        Range("B12").Value = "Sorry! You Lose!!!"
        winx = 0
    Else
        If cardnumber2 > cardnumber1 Then
            Beep
            a$ = "You win !!! Double"
            winx = winx * 2
        Else
            a$ = "Draw! Try double"
        End If
        ActiveSheet.DrawingObjects("No Button").Visible = True
        ActiveSheet.DrawingObjects("Yes Button").Visible = True
        Range("B12").Value = a$ + " again to win " + Str(winx * 2) + " ?"
    End If
    For doublego = 2 To 5
        If doublego <> held Then deals
    Next doublego
    If cardnumber2 < cardnumber1 Then double_return
    ActiveSheet.Protect DrawingObjects:=True, contents:=False, Scenarios:=True
    Range("B1").Select
End Sub

Sub double_return()
    With ActiveSheet
        .DrawingObjects("Deal Draw").Visible = True
        deal = 0
        updatescore
        .Range(.Cells(yc + 1, xc + 31), .Cells(yc + 4, xc + 31)).Select
        .Protect DrawingObjects:=True, contents:=False, Scenarios:=True
    End With
End Sub
Sub bbut1()
    Call heldbutton(2)
End Sub

Sub bbut2()
    heldbutton (8)
    If deal = 3 Then
        double_check (2)
    End If
End Sub

Sub bbut3()
    heldbutton (14)
    If deal = 3 Then
        double_check (3)
    End If
End Sub

Sub bbut4()
    heldbutton (20)
    If deal = 3 Then
        double_check (4)
    End If
End Sub

Sub bbut5()
    heldbutton (26)
    If deal = 3 Then
        double_check (5)
    End If
End Sub

Sub heldbutton(cardhold As Integer)
    With ActiveSheet
        If deal > 0 Or autop = True Or autohold = True Then
            If .Cells(yc - 1, xc + cardhold).Font.ColorIndex = "28" Then
                .Cells(yc - 1, xc + cardhold).Font.ColorIndex = "7"
            Else
                .Cells(yc - 1, xc + cardhold).Font.ColorIndex = "28"
            End If
        End If
    End With
End Sub


Sub xaxismax()
    Dim n As Integer
    n = Worksheets("Display").Range("AC3").Value + 10
    If n > 5 Then
        a$ = "=Shuffle!R1C4:R" + LTrim(Str$(n)) + "C4"
        b$ = "=Shuffle!R1C22:R" + LTrim(Str$(n)) + "C22"
        With Worksheets("Analysis")
            .ChartObjects("Chart 1").Chart.SeriesCollection(1).Values = a$
            .ChartObjects("Chart 1").Chart.SeriesCollection(2).Values = b$
            .Spinners("xmax").Value = n
        End With
    End If
End Sub
Sub xaxismaxPayout()
    Dim n As Integer
    n = Worksheets("Display").Range("AC3").Value + 10
    If n > 50 Then
        a$ = "=Shuffle!R1C7:R" + LTrim(Str$(n)) + "C7"
        b$ = "=Shuffle!R1C21:R" + LTrim(Str$(n)) + "C21"
        With Worksheets("Analysis").ChartObjects("Chart 3").Chart
            .SeriesCollection(1).Values = a$
            .SeriesCollection(2).Values = b$
        End With
    End If
End Sub

Sub yaxismax()
    Dim m As Integer, n As Integer
    With Worksheets("Shuffle")
        m = WorksheetFunction.Max(.Range("D:D"))
        n = WorksheetFunction.Min(.Range("D:D"))
    End With
    With Worksheets("Analysis")
        If n < 0 Then
            .Spinners("ymin").Value = Abs(n)
            .ChartObjects("Chart 1").Chart.Axes(xlValue).MinimumScale = n
        End If
        If m > 300 Then
            .Spinners("ymax").Value = m
            .ChartObjects("Chart 1").Chart.Axes(xlValue).MaximumScale = m
        End If
    End With
End Sub

Sub xscale()
    Dim n As Integer
    With Worksheets("Analysis")
        n = .Spinners("xmax").Value
        a$ = "=Shuffle!R1C4:R" + LTrim(Str$(n)) + "C4"
        .ChartObjects("Chart 1").Chart.SeriesCollection(1).Values = a$
    End With
End Sub

Sub yscalemin()
    Dim n As Integer
    With Worksheets("Analysis")
        n = .Spinners("ymin").Value
        .ChartObjects("Chart 1").Chart.Axes(xlValue).MinimumScale = -n
    End With
End Sub

Sub yscalemax()
    Dim n As Integer
    With Worksheets("Analysis")
        n = .Spinners("ymax").Value
        .ChartObjects("Chart 1").Chart.Axes(xlValue).MaximumScale = n
    End With
End Sub

Sub reset()
    If deal = 0 Then
        With Worksheets("Display")
            .Cells(3, 29).Value = 0
            .Cells(4, 29).Value = 0
            .Cells(4, 30).Value = "(0%)"
            .Cells(7, 6).Value = "'100"
            .Cells(5, 6).Value = 0
            .Cells(5, 29).Value = "'0%"
            .Cells(2, 36).Value = 0
            .Cells(3, 30) = ""
            .Cells(3, 36) = 100
        End With
        With Worksheets("Shuffle")
            .Range("D:D").ClearContents
            .Range("G:G").ClearContents
            .Range("U:U").ClearContents
            .Range("V:V").ClearContents
            .Cells(1, 4).Value = 100
            .Cells(1, 22).Value = 100
        End With
        With Worksheets("Analysis")
            a$ = "=Shuffle!R1C7:R" + CStr(50) + "C7"
            b$ = "=Shuffle!R1C21:R" + CStr(50) + "C21"
            With .ChartObjects("Chart 3").Chart
                .SeriesCollection(1).Values = a$
                .SeriesCollection(2).Values = b$
            End With
            .Range("I31:I40").Value = 0
            .Spinners("xmax").Value = 50
            xscale
            .Spinners("ymin").Value = 0
            yscalemin
            .Spinners("ymax").Value = 300
            yscalemax
        End With
    End If
End Sub


Sub gametype()
    If deal = 0 Then
        d = ActiveSheet.Range("AJ7").Value
        schedule$ = Chr$(Asc("B") + d) + "21:" + Chr$(Asc("B") + d) + "29"
        Worksheets("Tables").Range(schedule$).Copy
        ActiveSheet.Range("U3").Select
        Selection.PasteSpecial Paste:=xlValues
        payout$ = Chr$(Asc("B") + d) + "40"
        perpayout$ = "'" + LTrim(Str(Int(100 * Worksheets("Tables").Range(payout$).Value) / 100)) + "%"
        ActiveSheet.Range("AC6").Value = perpayout$
        a$ = Worksheets("Deck Data").Cells(2 + d, 12).Value
        ActiveSheet.Range("L2").Value = a$
        ActiveSheet.Range(Cells(14 + 1, 3 + 31), Cells(14 + 4, 3 + 31)).Select
    Else
        MsgBox "Not set. Finish game first!"
    End If
End Sub

Sub JokerTable()
    If ActiveCell <> "A1" Then
        'Sheets("Joker").Select
        'Range("A1").Select
        'Range("H10").Select
    End If
    Worksheets("Joker").Range("C2:F30").Copy
    With JokerTables
        Set .Picture = PastePicture(xlBitmap)
        .Show 0
        .Top = 110
        .Left = 800
    End With
End Sub

Sub uphand()
    selposn$ = Selection.Address(ReferenceStyle:=xlR1C1)
    f = InStr(selposn$, "C")
    numrows = Val(Mid(selposn$, 2, f - 2))
    numcols = Val(Right(selposn$, Len(selposn$) - f))
    If numcols = 2 And numrows > 48 And numrows < 72 Then
        Selection.Resize(1, 2).Select
        Selection.Cut
        ActiveCell.Offset(-1, 0).Select
        Selection.Insert Shift:=xlDown
        Selection.Resize(1, 1).Select
    End If
End Sub

Sub downhand()
    Dim numcols As Integer, numrows As Integer
    selposn$ = Selection.Address(ReferenceStyle:=xlR1C1)
    f = InStr(selposn$, "C")
    numrows = Val(Mid(selposn$, 2, f - 2))
    numcols = Val(Right(selposn$, Len(selposn$) - f))
    If numcols = 2 And numrows > 47 And numrows < 71 Then
        Selection.Resize(1, 2).Select
        Selection.Cut
        ActiveCell.Offset(2, 0).Select
        Selection.Insert Shift:=xlDown
        Selection.Resize(1, 1).Select
    End If
End Sub

Sub defhand()
    Range("B47:C72").Sort Key1:=Range("C48"), Order1:=xlAscending, Header:= _
    xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
    xlTopToBottom
    Range("B60").Select
End Sub

Sub perpay()
    Dim x As Integer, Y As Integer
    With Worksheets("Tables")
        For x = 1 To 6
            sump = 0
            For Y = 1 To 10
                w = .Cells(20 + Y, 2 + x).Value
                v = .Cells(3 + Y, 4).Value
                sump = sump + w * 20 / v
            Next Y
            .Cells(31, 2 + x).Value = sump
        Next x
        For x = 1 To 5
            sump = 0
            For Y = 1 To 5
                w = .Cells(32 + Y, 3 + x).Value
                v = .Cells(6, 4).Value
                h = .Cells(32 + Y, 9).Value
                v = v * 13 / h
                sump = sump + w * 20 / v
            Next Y
            .Cells(38, 3 + x).Value = sump
        Next x
    End With
End Sub


Sub nextfivecards()
    Dim x As Integer, Y As Integer
    Dim cdno As Integer
    If shuffle(1) <> 0 Then
        'Sheets("Shuffle").Select
        'Application.Goto Reference:=Range("E24"), Scroll:=True
        For Y = 1 To 2
            For x = 1 To 5
                cdsu$ = Application.VLookup(shuffle(x + 5 * (Y - 1)), Worksheets("Deck Data").Range("B1:D53"), 3)
                cdno = ((shuffle(x + 5 * (Y - 1)) - 1) Mod 13) + 1
                If cdsu$ = "J" Then
                    symbol$ = "Joker"
                    cdsu$ = "y"
                Else
                    symbol$ = Application.VLookup(cdno, Worksheets("Deck Data").Range("B1:D53"), 2)
                End If
                With Worksheets("Shuffle")
                    xz = 8 + (x - 1) * 2
                    yz = (Y - 1) * 3
                    .Cells(26 + yz, xz).Value = symbol$
                    .Cells(27 + yz, xz).Value = cdsu$
                    If cdsu$ = "©" Or cdsu$ = "¨" Then
                        .Cells(27 + yz, xz).Font.ColorIndex = 9
                    Else
                        .Cells(27 + yz, xz).Font.ColorIndex = 1
                    End If
                End With
            Next x
        Next Y
        'Range("L33").Select
        ShowNext5
    End If
End Sub
Sub ShowNext5()
    Worksheets("Shuffle").Range("F25:Q32").Copy
    With Next5
        Set .Picture = PastePicture(xlBitmap)
        If .Visible = False Then
            .Show 0
            .Top = 150
            .Left = 660
        Else
            t = .Top
            l = .Left
            .Show 0
            .Top = t
            .Left = l
        End If
    End With
End Sub
Sub BestHand()
    
End Sub
Sub Pause(sec As Single)
    Dim startTime As Single
    startTime = Timer
    'Do
    'Loop While startTime + sec > Timer
End Sub
Sub displaydeck()
    Dim x As Integer, n As Integer, Y As Integer
    If ActiveCell <> "G1" Then
        'Worksheets("Shuffle").Select
        'Application.Goto Reference:=Range("G1"), Scroll:=True
    End If
    With Worksheets("Shuffle")
        .Range("AK1:AO53").ClearContents
        For n = 1 To decksize
            x = shuffle(n)
            Y = (x - 1) Mod 13 + 1
            .Cells(n, 1).Value = x
            .Cells(n, 2).Value = Y
            If x < 14 Then
                .Cells(n, 37).Value = Y
            End If
            If x > 13 And x < 27 Then
                .Cells(n, 38).Value = Y
            End If
            If x > 26 And x < 40 Then
                .Cells(n, 39).Value = Y
            End If
            If x > 39 And x < 53 Then
                .Cells(n, 40).Value = Y
            End If
            If x = 53 Then
                .Cells(n, 41).Value = Y
            End If
        Next n
    End With
    'Range("I2").Select
    DoEvents
    ShowDeck
End Sub
Sub ShowDeck()
    Set ws2 = Worksheets("Shuffle")
    Set myChart = ws2.ChartObjects("shuffled deck").Chart
    myChart.ChartArea.Copy
    With DeckDisplay
        Set .Picture = PastePicture(xlBitmap)
        If .Visible = False Then
            .Show 0
            .Top = 150
            .Left = 710
        Else
            t = .Top
            l = .Left
            .Show 0
            .Top = t
            .Left = l
        End If
    End With
End Sub

Sub spade()
    With Worksheets("Shuffle")
        d$ = .DrawingObjects("spade data").Characters.Text
        If d$ = "Y" Then
            .ChartObjects("shuffled deck").Chart.SeriesCollection(1).MarkerStyle = xlNone
            .DrawingObjects("spade data").Characters.Text = "N"
        Else
            With .ChartObjects("shuffled deck").Chart.SeriesCollection(1)
                .MarkerBackgroundColorIndex = 1
                .MarkerForegroundColorIndex = 1
                .MarkerStyle = xlSquare
                .Smooth = False
            End With
            .DrawingObjects("spade data").Characters.Text = "Y"
        End If
    End With
End Sub

Sub diamond()
    With Worksheets("Shuffle")
        d$ = .DrawingObjects("diamond data").Characters.Text
        If d$ = "Y" Then
            .ChartObjects("shuffled deck").Chart.SeriesCollection(2).MarkerStyle = xlNone
            .DrawingObjects("diamond data").Characters.Text = "N"
        Else
            With .ChartObjects("shuffled deck").Chart.SeriesCollection(2)
                .MarkerBackgroundColorIndex = 3
                .MarkerForegroundColorIndex = 3
                .MarkerStyle = xlDiamond
                .Smooth = False
            End With
            .DrawingObjects("diamond data").Characters.Text = "Y"
        End If
    End With
End Sub

Sub heart()
    With Worksheets("Shuffle")
        d$ = .DrawingObjects("heart data").Characters.Text
        If d$ = "Y" Then
            .ChartObjects("shuffled deck").Chart.SeriesCollection(3).MarkerStyle = xlNone
            .DrawingObjects("heart data").Characters.Text = "N"
        Else
            With .ChartObjects("shuffled deck").Chart.SeriesCollection(3)
                .MarkerBackgroundColorIndex = 7
                .MarkerForegroundColorIndex = 7
                .MarkerStyle = xlTriangle
                .Smooth = False
            End With
            .DrawingObjects("heart data").Characters.Text = "Y"
        End If
    End With
End Sub

Sub club()
    With Worksheets("Shuffle")
        d$ = .DrawingObjects("club data").Characters.Text
        If d$ = "Y" Then
            .ChartObjects("shuffled deck").Chart.SeriesCollection(4).MarkerStyle = xlNone
            .DrawingObjects("club data").Characters.Text = "N"
        Else
            With .ChartObjects("shuffled deck").Chart.SeriesCollection(4)
                .MarkerBackgroundColorIndex = 5
                .MarkerForegroundColorIndex = 5
                .MarkerStyle = xlCircle
                .Smooth = False
            End With
            .DrawingObjects("club data").Characters.Text = "Y"
        End If
    End With
End Sub
