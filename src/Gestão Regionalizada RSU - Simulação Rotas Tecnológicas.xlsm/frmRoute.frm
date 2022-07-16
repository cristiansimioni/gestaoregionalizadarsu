VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRoute 
   Caption         =   "UserForm1"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8265.001
   OleObjectBlob   =   "frmRoute.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FormChanged As Boolean

Private Sub btnBack_Click()
    If FormChanged Then
        answer = MsgBox(MSG_CHANGED_NOT_SAVED, vbQuestion + vbYesNo + vbDefaultButton2, "Salvar Alterações")
        If answer = vbYes Then
          Call btnSave_Click
        Else
          Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub btnDefault_Click()
    txtMixedWasteToBeSorted = Database.GetDatabaseValue("MixedWasteToBeSorted", colDefaultValue)
    txtMechanizedSortingEfficiency = Database.GetDatabaseValue("MechanizedSortingEfficiency", colDefaultValue)
End Sub

Private Sub btnSave_Click()
    If modForm.validateForm() Then
        Call Database.SetDatabaseValue("MixedWasteToBeSorted", colUserValue, CDbl(txtMixedWasteToBeSorted.Text))
        Call Database.SetDatabaseValue("MechanizedSortingEfficiency", colUserValue, CDbl(txtMechanizedSortingEfficiency.Text))
        FormChanged = False
        Unload Me
    Else
        answer = MsgBox("Valores inválidos. Favor verificar!", vbExclamation, "Dados inválidos")
    End If
End Sub

Private Sub txtMechanizedSortingEfficiency_Change()
    Call modForm.textBoxChange(txtMechanizedSortingEfficiency, "MechanizedSortingEfficiency", FormChanged)
End Sub

Private Sub txtMixedWasteToBeSorted_Change()
    Call modForm.textBoxChange(txtMixedWasteToBeSorted, "MixedWasteToBeSorted", FormChanged)
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Gravimetria do RSU")
    
    txtMixedWasteToBeSorted = Database.GetDatabaseValue("MixedWasteToBeSorted", colUserValue)
    txtMechanizedSortingEfficiency = Database.GetDatabaseValue("MechanizedSortingEfficiency", colUserValue)
    
    FormChanged = False
End Sub
