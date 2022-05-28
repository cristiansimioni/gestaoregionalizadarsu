VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} step2 
   Caption         =   "Passo 2"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8355.001
   OleObjectBlob   =   "step2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "step2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnRunAlgorithm_Click()
    Util.RunPythonScript
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = APPNAME & " - Passo 2"
    Me.BackColor = ApplicationColors.bgColorLevel2
End Sub
