VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOpexData 
   Caption         =   "UserForm1"
   ClientHeight    =   8445.001
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   11055
   OleObjectBlob   =   "frmOpexData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOpexData"
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

Function ValidateForm() As Boolean
    ValidateForm = True
End Function

Private Sub cbxHiringRegimeManualOperator_Change()
FormChanged = True
End Sub

Private Sub txtAverageSalaryManager_Change()
Call modForm.textBoxChange(txtAverageSalaryManager, "AverageSalaryManager", FormChanged)
End Sub
Private Sub txtAverageSalarySupervision_Change()
Call modForm.textBoxChange(txtAverageSalarySupervision, "AverageSalarySupervision", FormChanged)
End Sub
Private Sub txtAverageSalaryOperational_Change()
Call modForm.textBoxChange(txtAverageSalaryOperational, "AverageSalaryOperational", FormChanged)
End Sub
Private Sub txtHiringRegimeManualOperator_Change()
Call modForm.textBoxChange(txtHiringRegimeManualOperator, "HiringRegimeManualOperator", FormChanged)
End Sub
Private Sub txtAverageSalaryManualOperator_Change()
Call modForm.textBoxChange(txtAverageSalaryManualOperator, "AverageSalaryManualOperator", FormChanged)
End Sub
Private Sub txtAverageCostElectricityConsumption_Change()
Call modForm.textBoxChange(txtAverageCostElectricityConsumption, "AverageCostElectricityConsumption", FormChanged)
End Sub
Private Sub txtFixedCostDemandContractedElectricity_Change()
Call modForm.textBoxChange(txtFixedCostDemandContractedElectricity, "FixedCostDemandContractedElectricity", FormChanged)
End Sub
Private Sub txtAverageServiceCostFixedAuxiliary_Change()
Call modForm.textBoxChange(txtAverageServiceCostFixedAuxiliary, "AverageServiceCostFixedAuxiliary", FormChanged)
End Sub
Private Sub txtAverageCostWheelLoaderRental_Change()
Call modForm.textBoxChange(txtAverageCostWheelLoaderRental, "AverageCostWheelLoaderRental", FormChanged)
End Sub
Private Sub txtAverageRentalCostEquipmentLandfill_Change()
Call modForm.textBoxChange(txtAverageRentalCostEquipmentLandfill, "AverageRentalCostEquipmentLandfill", FormChanged)
End Sub
Private Sub txtAverageCostDisposalLiquidEffluen_Change()
Call modForm.textBoxChange(txtAverageCostDisposalLiquidEffluen, "AverageCostDisposalLiquidEffluen", FormChanged)
End Sub
Private Sub txtAverageCostDieselOilInternalMovement_Change()
Call modForm.textBoxChange(txtAverageCostDieselOilInternalMovement, "AverageCostDieselOilInternalMovement", FormChanged)
End Sub
Private Sub txtAverageUreaCost_Change()
Call modForm.textBoxChange(txtAverageUreaCost, "AverageUreaCost", FormChanged)
End Sub
Private Sub txtAverageCostHydrated_Change()
Call modForm.textBoxChange(txtAverageCostHydrated, "AverageCostHydrated", FormChanged)
End Sub
Private Sub txtAverageCostActivatedCarbon_Change()
Call modForm.textBoxChange(txtAverageCostActivatedCarbon, "AverageCostActivatedCarbon", FormChanged)
End Sub
Private Sub txtBoilerCleaningWaterConsumption_Change()
Call modForm.textBoxChange(txtBoilerCleaningWaterConsumption, "BoilerCleaningWaterConsumption", FormChanged)
End Sub
Private Sub txtAverageCostIndustrialWaterConsumption_Change()
Call modForm.textBoxChange(txtAverageCostIndustrialWaterConsumption, "AverageCostIndustrialWaterConsumption", FormChanged)
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Dados Indexadores de Opex")
    
    'Combo box
    Dim index As Integer
    index = 0
    Dim valuesHiringRegimeManualOperator
    valuesHiringRegimeManualOperator = Split(Database.GetDatabaseValue("HiringRegimeManualOperator", colUnit), ",")
    For Each v In valuesHiringRegimeManualOperator
        cbxHiringRegimeManualOperator.AddItem v
        If v = Database.GetDatabaseValue("HiringRegimeManualOperator", colUserValue) Then
            cbxHiringRegimeManualOperator.ListIndex = index
        End If
        index = index + 1
    Next v
    index = 0
    
    txtAverageSalaryManager = Database.GetDatabaseValue("AverageSalaryManager", colUserValue)
    txtAverageSalarySupervision = Database.GetDatabaseValue("AverageSalarySupervision", colUserValue)
    txtAverageSalaryOperational = Database.GetDatabaseValue("AverageSalaryOperational", colUserValue)
    txtHiringRegimeManualOperator = Database.GetDatabaseValue("HiringRegimeManualOperator", colUserValue)
    txtAverageSalaryManualOperator = Database.GetDatabaseValue("AverageSalaryManualOperator", colUserValue)
    txtAverageCostElectricityConsumption = Database.GetDatabaseValue("AverageCostElectricityConsumption", colUserValue)
    txtFixedCostDemandContractedElectricity = Database.GetDatabaseValue("FixedCostDemandContractedElectricity", colUserValue)
    txtAverageServiceCostFixedAuxiliary = Database.GetDatabaseValue("AverageServiceCostFixedAuxiliary", colUserValue)
    txtAverageCostWheelLoaderRental = Database.GetDatabaseValue("AverageCostWheelLoaderRental", colUserValue)
    txtAverageRentalCostEquipmentLandfill = Database.GetDatabaseValue("AverageRentalCostEquipmentLandfill", colUserValue)
    txtAverageCostDisposalLiquidEffluen = Database.GetDatabaseValue("AverageCostDisposalLiquidEffluen", colUserValue)
    txtAverageCostDieselOilInternalMovement = Database.GetDatabaseValue("AverageCostDieselOilInternalMovement", colUserValue)
    txtAverageUreaCost = Database.GetDatabaseValue("AverageUreaCost", colUserValue)
    txtAverageCostHydrated = Database.GetDatabaseValue("AverageCostHydrated", colUserValue)
    txtAverageCostActivatedCarbon = Database.GetDatabaseValue("AverageCostActivatedCarbon", colUserValue)
    txtBoilerCleaningWaterConsumption = Database.GetDatabaseValue("BoilerCleaningWaterConsumption", colUserValue)
    txtAverageCostIndustrialWaterConsumption = Database.GetDatabaseValue("AverageCostIndustrialWaterConsumption", colUserValue)
    
    FormChanged = False
End Sub

Private Sub btnSave_Click()
    If modForm.ValidateForm() Then
        Call Database.SetDatabaseValue("AverageSalaryManager", colUserValue, CDbl(txtAverageSalaryManager.Text))
        Call Database.SetDatabaseValue("AverageSalarySupervision", colUserValue, CDbl(txtAverageSalarySupervision.Text))
        Call Database.SetDatabaseValue("AverageSalaryOperational", colUserValue, CDbl(txtAverageSalaryOperational.Text))
        Call Database.SetDatabaseValue("HiringRegimeManualOperator", colUserValue, cbxHiringRegimeManualOperator.value)
        Call Database.SetDatabaseValue("AverageSalaryManualOperator", colUserValue, CDbl(txtAverageSalaryManualOperator.Text))
        Call Database.SetDatabaseValue("AverageCostElectricityConsumption", colUserValue, CDbl(txtAverageCostElectricityConsumption.Text))
        Call Database.SetDatabaseValue("FixedCostDemandContractedElectricity", colUserValue, CDbl(txtFixedCostDemandContractedElectricity.Text))
        Call Database.SetDatabaseValue("AverageServiceCostFixedAuxiliary", colUserValue, CDbl(txtAverageServiceCostFixedAuxiliary.Text))
        Call Database.SetDatabaseValue("AverageCostWheelLoaderRental", colUserValue, CDbl(txtAverageCostWheelLoaderRental.Text))
        Call Database.SetDatabaseValue("AverageRentalCostEquipmentLandfill", colUserValue, CDbl(txtAverageRentalCostEquipmentLandfill.Text))
        Call Database.SetDatabaseValue("AverageCostDisposalLiquidEffluen", colUserValue, CDbl(txtAverageCostDisposalLiquidEffluen.Text))
        Call Database.SetDatabaseValue("AverageCostDieselOilInternalMovement", colUserValue, CDbl(txtAverageCostDieselOilInternalMovement.Text))
        Call Database.SetDatabaseValue("AverageUreaCost", colUserValue, CDbl(txtAverageUreaCost.Text))
        Call Database.SetDatabaseValue("AverageCostHydrated", colUserValue, CDbl(txtAverageCostHydrated.Text))
        Call Database.SetDatabaseValue("AverageCostActivatedCarbon", colUserValue, CDbl(txtAverageCostActivatedCarbon.Text))
        Call Database.SetDatabaseValue("BoilerCleaningWaterConsumption", colUserValue, CDbl(txtBoilerCleaningWaterConsumption.Text))
        Call Database.SetDatabaseValue("AverageCostIndustrialWaterConsumption", colUserValue, CDbl(txtAverageCostIndustrialWaterConsumption.Text))
        FormChanged = False
        frmStepThree.updateForm
        Unload Me
    Else
        answer = MsgBox(MSG_INVALID_DATA, vbExclamation, MSG_INVALID_DATA_TITLE)
    End If
End Sub

Private Sub btnDefault_Click()
    Dim index As Integer
    index = 0
    Dim valuesHiringRegimeManualOperator
    valuesHiringRegimeManualOperator = Split(Database.GetDatabaseValue("HiringRegimeManualOperator", colUnit), ",")
    For Each v In valuesHiringRegimeManualOperator
        If v = Database.GetDatabaseValue("HiringRegimeManualOperator", colDefaultValue) Then
            cbxHiringRegimeManualOperator.ListIndex = index
        End If
        index = index + 1
    Next v
    index = 0
    
    txtAverageSalaryManager = Database.GetDatabaseValue("AverageSalaryManager", colDefaultValue)
    txtAverageSalarySupervision = Database.GetDatabaseValue("AverageSalarySupervision", colDefaultValue)
    txtAverageSalaryOperational = Database.GetDatabaseValue("AverageSalaryOperational", colDefaultValue)
    txtAverageSalaryManualOperator = Database.GetDatabaseValue("AverageSalaryManualOperator", colDefaultValue)
    txtAverageCostElectricityConsumption = Database.GetDatabaseValue("AverageCostElectricityConsumption", colDefaultValue)
    txtFixedCostDemandContractedElectricity = Database.GetDatabaseValue("FixedCostDemandContractedElectricity", colDefaultValue)
    txtAverageServiceCostFixedAuxiliary = Database.GetDatabaseValue("AverageServiceCostFixedAuxiliary", colDefaultValue)
    txtAverageCostWheelLoaderRental = Database.GetDatabaseValue("AverageCostWheelLoaderRental", colDefaultValue)
    txtAverageRentalCostEquipmentLandfill = Database.GetDatabaseValue("AverageRentalCostEquipmentLandfill", colDefaultValue)
    txtAverageCostDisposalLiquidEffluen = Database.GetDatabaseValue("AverageCostDisposalLiquidEffluen", colDefaultValue)
    txtAverageCostDieselOilInternalMovement = Database.GetDatabaseValue("AverageCostDieselOilInternalMovement", colDefaultValue)
    txtAverageUreaCost = Database.GetDatabaseValue("AverageUreaCost", colDefaultValue)
    txtAverageCostHydrated = Database.GetDatabaseValue("AverageCostHydrated", colDefaultValue)
    txtAverageCostActivatedCarbon = Database.GetDatabaseValue("AverageCostActivatedCarbon", colDefaultValue)
    txtBoilerCleaningWaterConsumption = Database.GetDatabaseValue("BoilerCleaningWaterConsumption", colDefaultValue)
    txtAverageCostIndustrialWaterConsumption = Database.GetDatabaseValue("AverageCostIndustrialWaterConsumption", colDefaultValue)
End Sub
