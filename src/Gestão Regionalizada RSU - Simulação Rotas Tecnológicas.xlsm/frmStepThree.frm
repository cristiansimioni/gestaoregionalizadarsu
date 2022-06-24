VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepThree 
   Caption         =   "UserForm1"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7170
   OleObjectBlob   =   "frmStepThree.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStepThree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnCapexData_Click()
    frmCapexData.Show
End Sub

Private Sub btnContract_Click()
    frmContract.Show
End Sub

Private Sub btnFinancialAssumptions_Click()
    frmFinancialAssumptions.Show
End Sub

Private Sub btnOpexData_Click()
    frmOpexData.Show
End Sub

Private Sub btnRouteDefinition_Click()
    frmRoute.Show
End Sub

Private Sub btnTaxes_Click()
    frmTaxes.Show
End Sub

Private Sub btnUserBase_Click()
    frmUserBase.Show
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = APPNAME & " - Passo 3"
    Me.BackColor = ApplicationColors.bgColorLevel2
    
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.btColorLevel2
         End If
    Next Ctrl
End Sub
