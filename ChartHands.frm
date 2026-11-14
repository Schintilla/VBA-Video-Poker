VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChartHands 
   Caption         =   "Hands Won"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   7995
   OleObjectBlob   =   "ChartHands.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChartHands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bCloseHands_Click()
    ChartHands.Hide
End Sub

Private Sub bHandsRefresh_Click()
    HandsChart
End Sub
Private Sub bTable_Click()
    Worksheets("Analysis").Range("G29:L42").Copy
    With WinsTable
        Set .IWinsTable.Picture = PastePicture(xlBitmap)
        .Show 0
        .Top = 200
        .Left = 760
    End With
End Sub
