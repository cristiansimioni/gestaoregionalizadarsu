VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPriceValMarket 
   Caption         =   "UserForm2"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "frmPriceValMarket.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPriceValMarket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Me.Caption = APPNAME & " - Selectionar Cidades"
    Me.BackColor = ApplicationColors.frmBgColorLevel4
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.bgColorLevel4
         End If
    Next Ctrl
End Sub
