Attribute VB_Name = "Charting"
Sub GotoChart()
    CashChart
    Charting.PayoutChart
End Sub
Sub PayoutChart()
    Set ws2 = Worksheets("Analysis")
    Set myChart = ws2.ChartObjects("Chart 3").Chart
    myChart.ChartArea.Copy
    With ChartPayout
        Set .IPayout.Picture = PastePicture(xlBitmap)
        If .Visible = False Then
            .Show 0
            .Top = 200
            .Left = 700
            t = .Top
        Else
            t = .Top
            l = .Left
            .Show 0
            .Top = t
            .Left = l
        End If
    End With
End Sub
Sub CashChart()
    Set ws2 = Worksheets("Analysis")
    Set myChart = ws2.ChartObjects("Chart 1").Chart
    myChart.ChartArea.Copy
    With ChartCash
        Set .ICash.Picture = PastePicture(xlBitmap)
        If .Visible = False Then
            .Show 0
            .Top = 180
            .Left = 720
            t = .Top
        Else
            t = .Top
            l = .Left
            .Show 0
            .Top = t
            .Left = l
        End If
    End With
End Sub
Sub HandsChart()
    Set ws2 = Worksheets("Analysis")
    Set myChart = ws2.ChartObjects("Chart 4").Chart
    myChart.ChartArea.Copy
    With ChartHands
        Set .IHands.Picture = PastePicture(xlBitmap)
        Set myChart = ws2.ChartObjects("Chart 2").Chart
        myChart.ChartArea.Copy
        Set .IWins.Picture = PastePicture(xlBitmap)
        Set myChart = ws2.ChartObjects("Chart 5").Chart
        myChart.ChartArea.Copy
        Set .IAmt.Picture = PastePicture(xlBitmap)
        If .Visible = False Then
            .Show 0
            .Top = 160
            .Left = 680
            t = .Top
        Else
            t = .Top
            l = .Left
            .Show 0
            .Top = t
            .Left = l
        End If
    End With
End Sub
