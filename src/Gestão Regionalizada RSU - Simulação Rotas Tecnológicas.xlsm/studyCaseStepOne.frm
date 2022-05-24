VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} studyCaseStepOne 
   Caption         =   "Dados de Definição do Estudo de Caso"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7425
   OleObjectBlob   =   "studyCaseStepOne.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "studyCaseStepOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GenerationPerCapitaRDO As Double
Dim IndexSelectiveColletionRSU As Double
Dim AnnualGrowthPopulation As Double
Dim AnnualGrowthCollect As Double
Dim COEmission As Double
Dim AverageCostTransportation As Double
Dim ReducingCostMovimentation As Double

Private Sub btnBack_Click()
    Unload Me
End Sub

Function validateForm() As Boolean
    validateForm = True
End Function

Private Sub btnSave_Click()
    If validateForm() Then
        Database.setGenerationPerCapitaRDO (CDbl(txtGenerationPerCapitaRDO.Text))
        Database.setIndexSelectiveColletionRSU (CDbl(txtIndexSelectiveColletionRSU.Text))
        Database.setAnnualGrowthPopulation (CDbl(txtAnnualGrowthPopulation.Text))
        Database.setAnnualGrowthCollect (CDbl(txtAnnualGrowthCollect.Text))
        Database.setCOEmission (CDbl(txtCOEmission.Text))
        Database.setAverageCostTransportation (CDbl(txtAverageCostTransportation.Text))
        Database.setReducingCostMovimentation (CDbl(txtReducingCostMovimentation.Text))
        'Unload Me
    Else
        MsgBox "Valores inválidos. Favor verificar!"
    End If
End Sub


Private Sub txtGenerationPerCapitaRDO_Change()
    Dim errorMsg As String
    If Util.validateRange(txtGenerationPerCapitaRDO.Text, 0.75, 1.25, errorMsg) Then
        txtGenerationPerCapitaRDO.BackColor = Util.xColorGreen
        txtGenerationPerCapitaRDO.ControlTipText = errorMsg
    Else
        txtGenerationPerCapitaRDO.BackColor = Util.xColorRed
        txtGenerationPerCapitaRDO.ControlTipText = errorMsg
    End If
End Sub


Private Sub txtAnnualGrowthPopulation_Change()
    Dim errorMsg As String
    If Util.validateRange(txtAnnualGrowthPopulation.Text, 0#, 100#, errorMsg) Then
        txtAnnualGrowthPopulation.BackColor = Util.xColorGreen
        txtAnnualGrowthPopulation.ControlTipText = errorMsg
    Else
        txtAnnualGrowthPopulation.BackColor = Util.xColorRed
        txtAnnualGrowthPopulation.ControlTipText = errorMsg
    End If
End Sub

Private Sub txtIndexSelectiveColletionRSU_Change()
    Dim errorMsg As String
    If Util.validateRange(txtIndexSelectiveColletionRSU.Text, 0#, 100#, errorMsg) Then
        txtIndexSelectiveColletionRSU.BackColor = Util.xColorGreen
        txtIndexSelectiveColletionRSU.ControlTipText = errorMsg
    Else
        txtIndexSelectiveColletionRSU.BackColor = Util.xColorRed
        txtIndexSelectiveColletionRSU.ControlTipText = errorMsg
    End If
End Sub

Private Sub txtAnnualGrowthCollect_Change()
    Dim errorMsg As String
    If Util.validateRange(txtAnnualGrowthCollect.Text, 0#, 100#, errorMsg) Then
        txtAnnualGrowthCollect.BackColor = Util.xColorGreen
        txtAnnualGrowthCollect.ControlTipText = errorMsg
    Else
        txtAnnualGrowthCollect.BackColor = Util.xColorRed
        txtAnnualGrowthCollect.ControlTipText = errorMsg
    End If
End Sub

Private Sub txtCOEmission_Change()
    Dim errorMsg As String
    If Util.validateRange(txtCOEmission.Text, 0.5, 2.5, errorMsg) Then
        txtCOEmission.BackColor = Util.xColorGreen
        txtCOEmission.ControlTipText = errorMsg
    Else
        txtCOEmission.BackColor = Util.xColorRed
        txtCOEmission.ControlTipText = errorMsg
    End If
End Sub

Private Sub txtAverageCostTransportation_Change()
    Dim errorMsg As String
    If Util.validateRange(txtAverageCostTransportation.Text, 0.5, 10#, errorMsg) Then
        txtAverageCostTransportation.BackColor = Util.xColorGreen
        txtAverageCostTransportation.ControlTipText = errorMsg
    Else
        txtAverageCostTransportation.BackColor = Util.xColorRed
        txtAverageCostTransportation.ControlTipText = errorMsg
    End If
End Sub

Private Sub txtReducingCostMovimentation_Change()
    Dim errorMsg As String
    If Util.validateRange(txtAverageCostTransportation.Text, 0#, 100#, errorMsg) Then
        txtReducingCostMovimentation.BackColor = Util.xColorGreen
        txtReducingCostMovimentation.ControlTipText = errorMsg
    Else
        txtReducingCostMovimentation.BackColor = Util.xColorRed
        txtReducingCostMovimentation.ControlTipText = errorMsg
    End If
End Sub

Private Sub UserForm_Initialize()
    studyCaseStepOne.BackColor = xColorLevel3

    txtGenerationPerCapitaRDO.Text = Database.getGenerationPerCapitaRDO
    txtIndexSelectiveColletionRSU.Text = Database.getIndexSelectiveColletionRSU
    txtAnnualGrowthPopulation.Text = Database.getAnnualGrowthPopulation
    txtAnnualGrowthCollect.Text = Database.getAnnualGrowthCollect
    txtCOEmission.Text = Database.getCOEmission
    txtAverageCostTransportation.Text = Database.getAverageCostTransportation
    txtReducingCostMovimentation.Text = Database.getReducingCostMovimentation
End Sub
