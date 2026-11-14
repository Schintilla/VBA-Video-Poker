VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChartPayout 
   Caption         =   "Payout"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   8565.001
   OleObjectBlob   =   "ChartPayout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChartPayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bClosePayout_Click()
    ChartPayout.Hide
End Sub

Private Sub bRefresh_Click()
    PayoutChart
End Sub

Private Sub bZoom_Click()
    If bZoom.Caption = "Zoom In" Then
        n = Worksheets("Display").Range("AC3").Value + 10
        If n > 50 Then
            a$ = "=Shuffle!R" + CStr(n - 50) + "C7:R" + CStr(n + 10) + "C7"
            b$ = "=Shuffle!R" + CStr(n - 50) + "C21:R" + CStr(n + 10) + "C21"
            With Worksheets("Analysis").ChartObjects("Chart 3").Chart
                .SeriesCollection(1).Values = a$
                .SeriesCollection(2).Values = b$
            End With
            Charting.PayoutChart
            bZoom.Caption = "Zoom Out"
        End If
    Else
        Call Code.xaxismaxPayout
        Charting.PayoutChart
        bZoom.Caption = "Zoom In"
    End If
End Sub
