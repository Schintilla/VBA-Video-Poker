VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PayoutTables 
   Caption         =   "Payout Tables"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   8250.001
   OleObjectBlob   =   "PayoutTables.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PayoutTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bClosePayTables_Click()
    PayoutTables.Hide
End Sub

Private Sub bGotoJoker_Click()
    JokerTable
End Sub
