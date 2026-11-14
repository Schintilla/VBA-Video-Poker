VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChartCash 
   Caption         =   "Winnings"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   8445.001
   OleObjectBlob   =   "ChartCash.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChartCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bCashRefresh_Click()
    CashChart
End Sub

Private Sub bCloseCash_Click()
    ChartCash.Hide
End Sub

Private Sub bXMax_Click()
    xaxismax
    CashChart
End Sub

Private Sub bYMax_Click()
    yaxismax
    CashChart
End Sub

Private Sub sXMax_SpinDown()
    Dim n As Integer
    With Worksheets("Analysis")
        n = .Spinners("ymax").Value
        p = Worksheets("Display").Range("PlayNo")
        'If n - 25 > p Then
            .Spinners("ymax").Value = n - 25
            a$ = "=Shuffle!R1C4:R" + CStr(n - 25) + "C4"
            b$ = "=Shuffle!R1C22:R" + CStr(n - 25) + "C22"
            .ChartObjects("Chart 1").Chart.SeriesCollection(1).Values = a$
            .ChartObjects("Chart 1").Chart.SeriesCollection(2).Values = b$
            CashChart
        'End If
    End With
End Sub

Private Sub sXMax_SpinUp()
    Dim n As Integer
    With Worksheets("Analysis")
        n = .Spinners("ymax").Value
        .Spinners("ymax").Value = n + 25
        a$ = "=Shuffle!R1C4:R" + LTrim(Str$(n + 25)) + "C4"
        b$ = "=Shuffle!R1C22:R" + LTrim(Str$(n + 25)) + "C22"
        .ChartObjects("Chart 1").Chart.SeriesCollection(1).Values = a$
        .ChartObjects("Chart 1").Chart.SeriesCollection(2).Values = b$
    End With
    CashChart
End Sub

Private Sub sYMax_SpinDown()
    With Worksheets("Analysis")
        n = .Spinners("ymax").Value
        If n > 50 Then
            .Spinners("ymax").Value = n - 25
            .ChartObjects("Chart 1").Chart.Axes(xlValue).MaximumScale = n - 25
        End If
    End With
    CashChart
End Sub

Private Sub sYMax_SpinUp()
    Dim n As Integer
    With Worksheets("Analysis")
        n = .Spinners("ymax").Value
        .Spinners("ymax").Value = n + 25
        .ChartObjects("Chart 1").Chart.Axes(xlValue).MaximumScale = n + 25
    End With
    CashChart
End Sub

Private Sub sYMin_SpinDown()
    Dim n As Integer
    With Worksheets("Analysis")
        n = .Spinners("ymin").Value
        If n >= 25 Then
            .Spinners("ymin").Value = n - 25
            .ChartObjects("Chart 1").Chart.Axes(xlValue).MinimumScale = -n + 25
        End If
    End With
    CashChart
End Sub

Private Sub sYMin_SpinUp()
    Dim n As Integer
    With Worksheets("Analysis")
        n = .Spinners("ymin").Value
        .Spinners("ymin").Value = n + 25
        .ChartObjects("Chart 1").Chart.Axes(xlValue).MinimumScale = -n - 25
    End With
    CashChart
End Sub
