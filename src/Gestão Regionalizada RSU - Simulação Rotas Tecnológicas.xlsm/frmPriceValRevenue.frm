VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPriceValRevenue 
   Caption         =   "UserForm1"
   ClientHeight    =   2280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "frmPriceValRevenue.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPriceValRevenue"
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
