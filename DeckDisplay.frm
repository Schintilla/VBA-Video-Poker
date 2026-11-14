VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeckDisplay 
   Caption         =   "Shuffled Deck"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   9915.001
   OleObjectBlob   =   "DeckDisplay.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeckDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bCloseDeck_Click()
    DeckDisplay.Hide
End Sub

Private Sub bClub_Click()
    club
    ShowDeck
End Sub

Private Sub bDeckRefresh_Click()
    displaydeck
End Sub

Private Sub bDiamond_Click()
    diamond
    ShowDeck
End Sub

Private Sub bHeart_Click()
    heart
    ShowDeck
End Sub

Private Sub bSpade_Click()
    spade
    ShowDeck
End Sub
