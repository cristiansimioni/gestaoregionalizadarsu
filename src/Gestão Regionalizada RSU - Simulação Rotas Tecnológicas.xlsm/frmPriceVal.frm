VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPriceVal 
   Caption         =   "Preços para Valorização"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7188
   OleObjectBlob   =   "frmPriceVal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPriceVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnPriceValAutoconsumo_Click()
    frmPriceValAutoconsumo.Show
End Sub

Private Sub btnPriceValMarket_Click()
    frmPriceValMarket.Show
End Sub

Private Sub btnPriceValPublic_Click()
    frmPriceValPublic.Show
End Sub

Private Sub btnPriceValRevenue_Click()
    frmPriceValRevenue.Show
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "AAAAAAAAAAAAAAA")
End Sub
