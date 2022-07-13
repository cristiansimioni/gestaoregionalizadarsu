VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuantitativeVal 
   Caption         =   "UserForm1"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7170
   OleObjectBlob   =   "frmQuantitativeVal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmQuantitativeVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnQuantitativeValAutoconsumo_Click()
    frmQuantitativeValAutoconsumo.Show
End Sub

Private Sub btnQuantitativeValMarket_Click()
    frmQuantitativeValMarket.Show
End Sub

Private Sub btnQuantitativeValPublic_Click()
    frmQuantitativeValPublic.Show
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Me.Caption = APPNAME & " - Selectionar Cidades"
    Me.BackColor = ApplicationColors.frmBgColorLevel3
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.bgColorLevel3
         End If
    Next Ctrl
End Sub
