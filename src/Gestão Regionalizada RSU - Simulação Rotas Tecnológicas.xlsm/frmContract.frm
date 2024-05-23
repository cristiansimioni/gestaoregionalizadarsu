VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmContract 
   Caption         =   "UserForm1"
   ClientHeight    =   4560
   ClientLeft      =   75
   ClientTop       =   300
   ClientWidth     =   5295
   OleObjectBlob   =   "frmContract.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmContract"
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

Private Sub txtCostCollectionTransportSelectiveDry_Change()
Call modForm.textBoxChange(txtCostCollectionTransportSelectiveDry, "CostCollectionTransportSelectiveDry", FormChanged)
End Sub
Private Sub txtCostCollectionTransportOrganicSelection_Change()
Call modForm.textBoxChange(txtCostCollectionTransportOrganicSelection, "CostCollectionTransportOrganicSelection", FormChanged)
End Sub
Private Sub txtCostCollectionTransportMixedTailings_Change()
Call modForm.textBoxChange(txtCostCollectionTransportMixedTailings, "CostCollectionTransportMixedTailings", FormChanged)
End Sub
Private Sub txtCostLandfillHazardousWasteDisposal_Change()
Call modForm.textBoxChange(txtCostLandfillHazardousWasteDisposal, "CostLandfillHazardousWasteDisposal", FormChanged)
End Sub
Private Sub txtAnnualPopulationGrowthEstimate_Change()
Call modForm.textBoxChange(txtAnnualPopulationGrowthEstimate, "AnnualPopulationGrowthEstimate", FormChanged)
End Sub
Private Sub txtAnnualExpenseManagementContract_Change()
Call modForm.textBoxChange(txtAnnualExpenseManagementContract, "AnnualExpenseManagementContract", FormChanged)
End Sub
Private Sub txtInvestmentCostsSocialEnvPrograms_Change()
Call modForm.textBoxChange(txtInvestmentCostsSocialEnvPrograms, "InvestmentCostsSocialEnvPrograms", FormChanged)
End Sub
Private Sub txtInvestmentCostsContractSpecificItems_Change()
Call modForm.textBoxChange(txtInvestmentCostsContractSpecificItems, "InvestmentCostsContractSpecificItems", FormChanged)
End Sub
Private Sub txtAmountRPUPublicCleaningDisposal_Change()
Call modForm.textBoxChange(txtAmountRPUPublicCleaningDisposal, "AmountRPUPublicCleaningDisposal", FormChanged)
End Sub


Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Contrato")
    
    txtCostCollectionTransportSelectiveDry = Database.GetDatabaseValue("CostCollectionTransportSelectiveDry", colUserValue)
    txtCostCollectionTransportOrganicSelection = Database.GetDatabaseValue("CostCollectionTransportOrganicSelection", colUserValue)
    txtCostCollectionTransportMixedTailings = Database.GetDatabaseValue("CostCollectionTransportMixedTailings", colUserValue)
    txtCostLandfillHazardousWasteDisposal = Database.GetDatabaseValue("CostLandfillHazardousWasteDisposal", colUserValue)
    txtAnnualPopulationGrowthEstimate = Database.GetDatabaseValue("AnnualPopulationGrowthEstimate", colUserValue)
    txtAnnualExpenseManagementContract = Database.GetDatabaseValue("AnnualExpenseManagementContract", colUserValue)
    txtInvestmentCostsSocialEnvPrograms = Database.GetDatabaseValue("InvestmentCostsSocialEnvPrograms", colUserValue)
    txtInvestmentCostsContractSpecificItems = Database.GetDatabaseValue("InvestmentCostsContractSpecificItems", colUserValue)
    txtAmountRPUPublicCleaningDisposal = Database.GetDatabaseValue("AmountRPUPublicCleaningDisposal", colUserValue)

    FormChanged = False
    
    Me.Height = 313
    Me.width = 425
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrorHandler
        Call Database.SetDatabaseValue("CostCollectionTransportSelectiveDry", colUserValue, CDbl(txtCostCollectionTransportSelectiveDry.Text))
        Call Database.SetDatabaseValue("CostCollectionTransportOrganicSelection", colUserValue, CDbl(txtCostCollectionTransportOrganicSelection.Text))
        Call Database.SetDatabaseValue("CostCollectionTransportMixedTailings", colUserValue, CDbl(txtCostCollectionTransportMixedTailings.Text))
        Call Database.SetDatabaseValue("CostLandfillHazardousWasteDisposal", colUserValue, CDbl(txtCostLandfillHazardousWasteDisposal.Text))
        Call Database.SetDatabaseValue("AnnualPopulationGrowthEstimate", colUserValue, CDbl(txtAnnualPopulationGrowthEstimate.Text))
        Call Database.SetDatabaseValue("AnnualExpenseManagementContract", colUserValue, CDbl(txtAnnualExpenseManagementContract.Text))
        Call Database.SetDatabaseValue("InvestmentCostsSocialEnvPrograms", colUserValue, CDbl(txtInvestmentCostsSocialEnvPrograms.Text))
        Call Database.SetDatabaseValue("InvestmentCostsContractSpecificItems", colUserValue, CDbl(txtInvestmentCostsContractSpecificItems.Text))
        Call Database.SetDatabaseValue("AmountRPUPublicCleaningDisposal", colUserValue, CDbl(txtAmountRPUPublicCleaningDisposal.Text))
        FormChanged = False
        frmStepThree.updateForm
        Unload Me
        ThisWorkbook.Save
    Exit Sub
    
ErrorHandler:
    Call MsgBox(MSG_INVALID_DATA, vbCritical, MSG_INVALID_DATA_TITLE)
    
End Sub

Private Sub btnDefault_Click()
    txtCostCollectionTransportSelectiveDry = Database.GetDatabaseValue("CostCollectionTransportSelectiveDry", colDefaultValue)
    txtCostCollectionTransportOrganicSelection = Database.GetDatabaseValue("CostCollectionTransportOrganicSelection", colDefaultValue)
    txtCostCollectionTransportMixedTailings = Database.GetDatabaseValue("CostCollectionTransportMixedTailings", colDefaultValue)
    txtCostLandfillHazardousWasteDisposal = Database.GetDatabaseValue("CostLandfillHazardousWasteDisposal", colDefaultValue)
    txtAnnualPopulationGrowthEstimate = Database.GetDatabaseValue("AnnualPopulationGrowthEstimate", colDefaultValue)
    txtAnnualExpenseManagementContract = Database.GetDatabaseValue("AnnualExpenseManagementContract", colDefaultValue)
    txtInvestmentCostsSocialEnvPrograms = Database.GetDatabaseValue("InvestmentCostsSocialEnvPrograms", colDefaultValue)
    txtInvestmentCostsContractSpecificItems = Database.GetDatabaseValue("InvestmentCostsContractSpecificItems", colDefaultValue)
    txtAmountRPUPublicCleaningDisposal = Database.GetDatabaseValue("AmountRPUPublicCleaningDisposal", colDefaultValue)
End Sub

