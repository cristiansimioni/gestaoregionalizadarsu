VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFinancialAssumptions 
   Caption         =   "UserForm1"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14295
   OleObjectBlob   =   "frmFinancialAssumptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFinancialAssumptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FormChanged As Boolean

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Premissas Financeiras")
    
    FormChanged = False
End Sub

