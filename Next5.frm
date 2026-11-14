VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Next5 
   Caption         =   "Next 5 Cards"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   7860
   OleObjectBlob   =   "Next5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Next5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bCloseNext5_Click()
    Next5.Hide
End Sub

Private Sub bDeck_Click()
    displaydeck
End Sub

Private Sub bNextRefresh_Click()
    nextfivecards
End Sub
