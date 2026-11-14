VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HelpDialog 
   Caption         =   "Help"
   ClientHeight    =   9345.001
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   8880.001
   OleObjectBlob   =   "HelpDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HelpDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bGraph_Click()
    Worksheets("Help").Shapes("Graphical").Copy
    HelpDialog.FInfo.Picture = PastePicture(xlBitmap)
End Sub

Private Sub bHold2_Click()
    Worksheets("Help").Shapes("Text 9").Copy
    HelpDialog.FInfo.Picture = PastePicture(xlBitmap)
End Sub

Private Sub bInfo_Click()
    Worksheets("Help").Shapes("Text 9").Copy
    'HelpDialog.IDetail.Picture = PastePicture(xlBitmap)
    HelpDialog.FInfo.Picture = PastePicture(xlBitmap)
End Sub

Private Sub bJoker_Click()
    Worksheets("Help").Shapes("Text 8").Copy
    HelpDialog.FInfo.Picture = PastePicture(xlBitmap)
End Sub

Private Sub bPayout_Click()
    Worksheets("Help").Shapes("Text 5").Copy
    HelpDialog.FInfo.Picture = PastePicture(xlBitmap)
End Sub

Private Sub bPayout2_Click()
    Worksheets("Help").Shapes("Text 5").Copy
    HelpDialog.FInfo.Picture = PastePicture(xlBitmap)
End Sub

Private Sub bProb_Click()
    Worksheets("Help").Shapes("Probabilities").Copy
    HelpDialog.FInfo.Picture = PastePicture(xlBitmap)
End Sub

Private Sub bViewDeck_Click()
    Worksheets("Help").Shapes("Text 6").Copy
    HelpDialog.FInfo.Picture = PastePicture(xlBitmap)
End Sub

Private Sub bViewNext_Click()
    Worksheets("Help").Shapes("Text 7").Copy
    HelpDialog.FInfo.Picture = PastePicture(xlBitmap)
End Sub

Private Sub bWins_Click()
    Worksheets("Help").Shapes("Text 4").Copy
    HelpDialog.FInfo.Picture = PastePicture(xlBitmap)
End Sub

Private Sub CommandButton1_Click()
    HelpDialog.Hide
End Sub
