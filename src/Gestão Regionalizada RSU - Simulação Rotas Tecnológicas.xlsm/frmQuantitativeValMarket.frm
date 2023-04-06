VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuantitativeValMarket 
   Caption         =   "UserForm1"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9948.001
   OleObjectBlob   =   "frmQuantitativeValMarket.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmQuantitativeValMarket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim FormChanged As Boolean

Private Sub btnBack_Click()
    If FormChanged Then
        answer = MsgBox(MSG_CHANGED_NOT_SAVED, vbQuestion + vbYesNo + vbDefaultButton2, MSG_CHANGED_NOT_SAVED_TITLE)
        If answer = vbYes Then
          Call btnSave_Click
        Else
          Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub txtBiomethaneSaleBase_Change()
Call modForm.textBoxChange(txtBiomethaneSaleBase, "BiomethaneSaleBase", FormChanged)
End Sub
Private Sub txtInfrastructureCTVRBase_Change()
Call modForm.textBoxChange(txtInfrastructureCTVRBase, "InfrastructureCTVRBase", FormChanged)
End Sub
Private Sub txtBiomethaneSaleOptimized_Change()
Call modForm.textBoxChange(txtBiomethaneSaleOptimized, "BiomethaneSaleOptimized", FormChanged)
End Sub
Private Sub txtInfrastructureCTVROptimized_Change()
Call modForm.textBoxChange(txtInfrastructureCTVROptimized, "InfrastructureCTVROptimized", FormChanged)
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Comercialização Mercado")
    
    txtBiomethaneSaleBase = Database.GetDatabaseValue("BiomethaneSaleBase", colUserValue)
    txtInfrastructureCTVRBase = Database.GetDatabaseValue("InfrastructureCTVRBase", colUserValue)
    txtBiomethaneSaleOptimized = Database.GetDatabaseValue("BiomethaneSaleOptimized", colUserValue)
    txtInfrastructureCTVROptimized = Database.GetDatabaseValue("InfrastructureCTVROptimized", colUserValue)

    FormChanged = False
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrorHandler
        Call Database.SetDatabaseValue("BiomethaneSaleBase", colUserValue, CDbl(txtBiomethaneSaleBase.Text))
        Call Database.SetDatabaseValue("InfrastructureCTVRBase", colUserValue, CDbl(txtInfrastructureCTVRBase.Text))
        Call Database.SetDatabaseValue("BiomethaneSaleOptimized", colUserValue, CDbl(txtBiomethaneSaleOptimized.Text))
        Call Database.SetDatabaseValue("InfrastructureCTVROptimized", colUserValue, CDbl(txtInfrastructureCTVROptimized.Text))
        FormChanged = False
        frmStepFour.updateForm
        Unload Me
        ThisWorkbook.Save
    Exit Sub
    
ErrorHandler:
    Call MsgBox(MSG_INVALID_DATA, vbCritical, MSG_INVALID_DATA_TITLE)
End Sub

Private Sub btnDefault_Click()
    txtBiomethaneSaleBase = Database.GetDatabaseValue("BiomethaneSaleBase", colDefaultValue)
    txtInfrastructureCTVRBase = Database.GetDatabaseValue("InfrastructureCTVRBase", colDefaultValue)
    txtBiomethaneSaleOptimized = Database.GetDatabaseValue("BiomethaneSaleOptimized", colDefaultValue)
    txtInfrastructureCTVROptimized = Database.GetDatabaseValue("InfrastructureCTVROptimized", colDefaultValue)
End Sub
