VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} simulationData 
   Caption         =   "Metas para a Simulação do Estudo de Caso"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   OleObjectBlob   =   "simulationData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "simulationData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LandfillDeviationTarget As Double
Dim ExpectedDeadline As Double
Dim MixedRecyclingIndex As Double
Dim TargetExpectation As Double

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    Database.setLandfillDeviationTarget (CDbl(txtLandfillDeviationTarget.Text))
    Database.setExpectedDeadline (CDbl(txtExpectedDeadline.Text))
    Database.setMixedRecyclingIndex (CDbl(txtMixedRecyclingIndex.Text))
    Database.setTargetExpectation (CDbl(txtTargetExpectation.Text))
End Sub

Private Sub txtExpectedDeadline_Change()
    Dim errorMsg As String
    If Util.validateRange(txtExpectedDeadline.Text, 20, 35, errorMsg) Then
        txtExpectedDeadline.BackColor = Util.xColorGreen
        txtExpectedDeadline.ControlTipText = errorMsg
    Else
        txtExpectedDeadline.BackColor = Util.xColorRed
        txtExpectedDeadline.ControlTipText = errorMsg
    End If
End Sub

Private Sub txtLandfillDeviationTarget_Change()
    Dim errorMsg As String
    If Util.validateRange(txtLandfillDeviationTarget.Text, 20, 90, errorMsg) Then
        txtLandfillDeviationTarget.BackColor = Util.xColorGreen
        txtLandfillDeviationTarget.ControlTipText = errorMsg
    Else
        txtLandfillDeviationTarget.BackColor = Util.xColorRed
        txtLandfillDeviationTarget.ControlTipText = errorMsg
    End If
End Sub

Private Sub txtMixedRecyclingIndex_Change()
    Dim errorMsg As String
    If Util.validateRange(txtMixedRecyclingIndex.Text, 20, 100, errorMsg) Then
        txtMixedRecyclingIndex.BackColor = Util.xColorGreen
        txtMixedRecyclingIndex.ControlTipText = errorMsg
    Else
        txtMixedRecyclingIndex.BackColor = Util.xColorRed
        txtMixedRecyclingIndex.ControlTipText = errorMsg
    End If
End Sub

Private Sub txtTargetExpectation_Change()
    Dim errorMsg As String
    If Util.validateRange(txtTargetExpectation.Text, 20, 500, errorMsg) Then
        txtTargetExpectation.BackColor = Util.xColorGreen
        txtTargetExpectation.ControlTipText = errorMsg
    Else
        txtTargetExpectation.BackColor = Util.xColorRed
        txtTargetExpectation.ControlTipText = errorMsg
    End If
End Sub


Private Sub UserForm_Initialize()
    Me.BackColor = Util.xColorLevel3
    
    LandfillDeviationTarget = Database.getLandfillDeviationTarget()
    ExpectedDeadline = Database.getExpectedDeadline()
    MixedRecyclingIndex = Database.getMixedRecyclingIndex()
    TargetExpectation = Database.getTargetExpectation()
    
    If LandfillDeviationTarget + ExpectedDeadline + MixedRecyclingIndex + TargetExpectation > 0 Then
        txtLandfillDeviationTarget.Text = LandfillDeviationTarget
        txtExpectedDeadline.Text = ExpectedDeadline
        txtMixedRecyclingIndex.Text = MixedRecyclingIndex
        txtTargetExpectation.Text = TargetExpectation
    End If
    
End Sub
