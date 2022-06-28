VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSimulationData 
   Caption         =   "Metas para a Simulação do Estudo de Caso"
   ClientHeight    =   3390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
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
Dim FormChanged As Boolean

Function validateForm() As Boolean
    validateForm = True
End Function

Private Sub btnBack_Click()
    If FormChanged Then
        answer = MsgBox("Você realizou alterações, gostaria de salvar?", vbQuestion + vbYesNo + vbDefaultButton2, "Salvar Alterações")
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
    If validateForm() Then
        Call Database.SetDatabaseValue("LandfillDeviationTarget", colUserValue, CDbl(txtLandfillDeviationTarget.Text))
        Call Database.SetDatabaseValue("ExpectedDeadline", colUserValue, CDbl(txtExpectedDeadline.Text))
        Call Database.SetDatabaseValue("MixedRecyclingIndex", colUserValue, CDbl(txtMixedRecyclingIndex.Text))
        Call Database.SetDatabaseValue("TargetExpectation", colUserValue, CDbl(txtTargetExpectation.Text))
        FormChanged = False
        Unload Me
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

Private Sub txtExpectedDeadline_Change()
    Call textBoxChange(txtExpectedDeadline, "ExpectedDeadline")
End Sub

Private Sub txtLandfillDeviationTarget_Change()
    Call textBoxChange(txtLandfillDeviationTarget, "LandfillDeviationTarget")
End Sub

Private Sub txtMixedRecyclingIndex_Change()
    Call textBoxChange(txtMixedRecyclingIndex, "MixedRecyclingIndex")
End Sub

Private Sub txtTargetExpectation_Change()
    Call textBoxChange(txtTargetExpectation, "TargetExpectation")
End Sub


Private Sub UserForm_Initialize()
    Me.Caption = APPNAME & " - Metas para a Simulação do Estudo de Caso"
    Me.BackColor = ApplicationColors.bgColorLevel3
    
    'Form Appearance
    Me.Caption = APPNAME & " - Gravimetria do RSU"
    Me.BackColor = ApplicationColors.bgColorLevel3
    
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.btColorLevel3
            Ctrl.ForeColor = ApplicationColors.fgColorLevel3
         End If
    Next Ctrl
    
    LandfillDeviationTarget = Database.GetDatabaseValue("LandfillDeviationTarget", colUserValue)
    ExpectedDeadline = Database.GetDatabaseValue("ExpectedDeadline", colUserValue)
    MixedRecyclingIndex = Database.GetDatabaseValue("MixedRecyclingIndex", colUserValue)
    TargetExpectation = Database.GetDatabaseValue("TargetExpectation", colUserValue)
    
    If LandfillDeviationTarget + ExpectedDeadline + MixedRecyclingIndex + TargetExpectation > 0 Then
        txtLandfillDeviationTarget.Text = LandfillDeviationTarget
        txtExpectedDeadline.Text = ExpectedDeadline
        txtMixedRecyclingIndex.Text = MixedRecyclingIndex
        txtTargetExpectation.Text = TargetExpectation
    End If
    
    FormChanged = False
End Sub
