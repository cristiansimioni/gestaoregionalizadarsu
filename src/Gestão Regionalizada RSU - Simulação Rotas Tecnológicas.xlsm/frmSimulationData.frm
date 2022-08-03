VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSimulationData 
   Caption         =   "Metas para a Simulação do Estudo de Caso"
   ClientHeight    =   5235
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   9600.001
   OleObjectBlob   =   "frmSimulationData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSimulationData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LandfillDeviationTarget As Double
Dim ExpectedDeadline As Double
Dim MixedRecyclingIndex As Double
Dim TargetExpectation As Double
Dim CurrentLandfillCost As Double
Dim CurrentCostRSU As Double
Dim LandfillCurrentDeviation As Double
Dim ValuationEfficiency As Double

Dim FormChanged As Boolean

Function ValidateForm() As Boolean
    ValidateForm = True
End Function

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

Private Sub btnSave_Click()
    If ValidateForm() Then
        Call Database.SetDatabaseValue("LandfillDeviationTarget", colUserValue, CDbl(txtLandfillDeviationTarget.Text))
        Call Database.SetDatabaseValue("ExpectedDeadline", colUserValue, CDbl(txtExpectedDeadline.Text))
        Call Database.SetDatabaseValue("MixedRecyclingIndex", colUserValue, CDbl(txtMixedRecyclingIndex.Text))
        Call Database.SetDatabaseValue("TargetExpectation", colUserValue, CDbl(txtTargetExpectation.Text))
        Call Database.SetDatabaseValue("CurrentLandfillCost", colUserValue, CDbl(txtCurrentLandfillCost.Text))
        Call Database.SetDatabaseValue("CurrentCostRSU", colUserValue, CDbl(txtCurrentCostRSU.Text))
        Call Database.SetDatabaseValue("LandfillCurrentDeviation", colUserValue, CDbl(txtLandfillCurrentDeviation.Text))
        
        FormChanged = False
        Unload Me
        frmStepOne.updateForm
    Else
        answer = MsgBox("Valores inválidos. Favor verificar!", vbExclamation, "Dados inválidos")
    End If
End Sub

Private Sub textBoxChange(ByRef txtBox, ByVal varName As String)
    Dim errorMsg As String
    If Database.Validate(varName, txtBox.Text, errorMsg) Then
        txtBox.BackColor = ApplicationColors.bgColorValidTextBox
        txtBox.ControlTipText = errorMsg
    Else
        txtBox.BackColor = ApplicationColors.bgColorInvalidTextBox
        txtBox.ControlTipText = errorMsg
    End If
    FormChanged = True
End Sub

Private Sub txtCurrentCostRSU_Change()
    Call textBoxChange(txtCurrentCostRSU, "CurrentCostRSU")
    Call calculateValuationEfficiency
End Sub

Private Sub txtCurrentLandfillCost_Change()
    Call textBoxChange(txtCurrentLandfillCost, "CurrentLandfillCost")
    Call calculateValuationEfficiency
End Sub

Private Sub txtExpectedDeadline_Change()
    Call textBoxChange(txtExpectedDeadline, "ExpectedDeadline")
End Sub

Private Sub txtLandfillCurrentDeviation_Change()
    Call textBoxChange(txtLandfillCurrentDeviation, "LandfillCurrentDeviation")
    Call calculateValuationEfficiency
End Sub

Private Sub txtLandfillDeviationTarget_Change()
    Call textBoxChange(txtLandfillDeviationTarget, "LandfillDeviationTarget")
    Call calculateValuationEfficiency
End Sub

Private Sub txtMixedRecyclingIndex_Change()
    Call textBoxChange(txtMixedRecyclingIndex, "MixedRecyclingIndex")
End Sub

Private Sub txtTargetExpectation_Change()
    Call textBoxChange(txtTargetExpectation, "TargetExpectation")
    Call calculateValuationEfficiency
End Sub

Private Sub txtValuationEfficiency_Change()
    'Call textBoxChange(txtValuationEfficiency, "ValuationEfficiency")
End Sub


Private Sub calculateValuationEfficiency()
    If IsNumeric(txtTargetExpectation.Text) And IsNumeric(txtCurrentCostRSU.Text) And IsNumeric(txtLandfillDeviationTarget.Text) And IsNumeric(txtLandfillCurrentDeviation.Text) And _
       IsNumeric(txtCurrentLandfillCost.Text) Then
       ValuationEfficiency = (((CDbl(txtTargetExpectation.Text) - CDbl(txtCurrentCostRSU.Text)) + ((CDbl(txtLandfillDeviationTarget.Text) / 100) - (CDbl(txtLandfillCurrentDeviation.Text) / 100)) * CDbl(txtCurrentLandfillCost.Text)) / CDbl(txtTargetExpectation.Text)) * 100
       txtValuationEfficiency = ValuationEfficiency
    Else
       txtValuationEfficiency = 0
    End If
    
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Metas para a Simulação do Estudo de Caso")
    txtValuationEfficiency.ForeColor = RGB(0, 0, 0)
    
    txtLandfillDeviationTarget.Text = Database.GetDatabaseValue("LandfillDeviationTarget", colUserValue)
    txtExpectedDeadline.Text = Database.GetDatabaseValue("ExpectedDeadline", colUserValue)
    txtMixedRecyclingIndex.Text = Database.GetDatabaseValue("MixedRecyclingIndex", colUserValue)
    txtTargetExpectation.Text = Database.GetDatabaseValue("TargetExpectation", colUserValue)
    txtCurrentLandfillCost.Text = Database.GetDatabaseValue("CurrentLandfillCost", colUserValue)
    txtCurrentCostRSU.Text = Database.GetDatabaseValue("CurrentCostRSU", colUserValue)
    txtLandfillCurrentDeviation.Text = Database.GetDatabaseValue("LandfillCurrentDeviation", colUserValue)
    txtValuationEfficiency.Text = Database.GetDatabaseValue("ValuationEfficiency", colUserValue)
    
    FormChanged = False
End Sub
