VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCapexData 
   Caption         =   "UserForm1"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11025
   OleObjectBlob   =   "frmCapexData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCapexData"
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

Function validateForm() As Boolean
    validateForm = True
End Function

Private Sub txtRealEuro_Change()
Call modForm.textBoxChange(txtRealEuro, "RealEuro", FormChanged)
End Sub
Private Sub txtRealDollar_Change()
Call modForm.textBoxChange(txtRealDollar, "RealDollar", FormChanged)
End Sub
Private Sub txtTaxesImportationEquipment_Change()
Call modForm.textBoxChange(txtTaxesImportationEquipment, "TaxesImportationEquipment", FormChanged)
End Sub
Private Sub txtAveragePriceBuildingLand_Change()
Call modForm.textBoxChange(txtAveragePriceBuildingLand, "AveragePriceBuildingLand", FormChanged)
End Sub
Private Sub txtAverageLandPriceLandfill_Change()
Call modForm.textBoxChange(txtAverageLandPriceLandfill, "AverageLandPriceLandfill", FormChanged)
End Sub
Private Sub txtAverageLandscapingPrice_Change()
Call modForm.textBoxChange(txtAverageLandscapingPrice, "AverageLandscapingPrice", FormChanged)
End Sub
Private Sub txtIncinerationEurBra_Change()
Call modForm.textBoxChange(txtIncinerationEurBra, "IncinerationEurBra", FormChanged)
End Sub
Private Sub txtAveragePriceIndustrialConcrete_Change()
Call modForm.textBoxChange(txtAveragePriceIndustrialConcrete, "AveragePriceIndustrialConcrete", FormChanged)
End Sub
Private Sub txtAveragePriceConstructionIndustrial_Change()
Call modForm.textBoxChange(txtAveragePriceConstructionIndustrial, "AveragePriceConstructionIndustrial", FormChanged)
End Sub
Private Sub txtNationalPrices_Change()
Call modForm.textBoxChange(txtNationalPrices, "NationalPrices", FormChanged)
End Sub
Private Sub txtMechanizedTechnologyNationalization_Change()
Call modForm.textBoxChange(txtMechanizedTechnologyNationalization, "MechanizedTechnologyNationalization", FormChanged)
End Sub
Private Sub txtMechanizedTechnologyOvercapacity_Change()
Call modForm.textBoxChange(txtMechanizedTechnologyOvercapacity, "MechanizedTechnologyOvercapacity", FormChanged)
End Sub
Private Sub txtProductionTechnologyCDRNationalization_Change()
Call modForm.textBoxChange(txtProductionTechnologyCDRNationalization, "ProductionTechnologyCDRNationalization", FormChanged)
End Sub
Private Sub txtProductionTechnologyCDROvercapacity_Change()
Call modForm.textBoxChange(txtProductionTechnologyCDROvercapacity, "ProductionTechnologyCDROvercapacity", FormChanged)
End Sub
Private Sub txtProductionTechnologyBIONationalization_Change()
Call modForm.textBoxChange(txtProductionTechnologyBIONationalization, "ProductionTechnologyBIONationalization", FormChanged)
End Sub
Private Sub txtProductionTechnologyBIOOvercapacity_Change()
Call modForm.textBoxChange(txtProductionTechnologyBIOOvercapacity, "ProductionTechnologyBIOOvercapacity", FormChanged)
End Sub
Private Sub txtTechnologyAnaerobicNationalization_Change()
Call modForm.textBoxChange(txtTechnologyAnaerobicNationalization, "TechnologyAnaerobicNationalization", FormChanged)
End Sub
Private Sub txtTechnologyAnaerobicOvercapacity_Change()
Call modForm.textBoxChange(txtTechnologyAnaerobicOvercapacity, "TechnologyAnaerobicOvercapacity", FormChanged)
End Sub
Private Sub txtTechnologyCompostingNationalization_Change()
Call modForm.textBoxChange(txtTechnologyCompostingNationalization, "TechnologyCompostingNationalization", FormChanged)
End Sub
Private Sub txtTechnologyCompostingOvercapacity_Change()
Call modForm.textBoxChange(txtTechnologyCompostingOvercapacity, "TechnologyCompostingOvercapacity", FormChanged)
End Sub
Private Sub txtTechnologyIncinerationNationalization_Change()
Call modForm.textBoxChange(txtTechnologyIncinerationNationalization, "TechnologyIncinerationNationalization", FormChanged)
End Sub
Private Sub txtTechnologyIncinerationOvercapacity_Change()
Call modForm.textBoxChange(txtTechnologyIncinerationOvercapacity, "TechnologyIncinerationOvercapacity", FormChanged)
End Sub
Private Sub txtTechnologyLandfillOvercapacity_Change()
Call modForm.textBoxChange(txtTechnologyLandfillOvercapacity, "TechnologyLandfillOvercapacity", FormChanged)
End Sub


Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Dados Indexadores de Capex")
    
    txtRealEuro = Database.GetDatabaseValue("RealEuro", colUserValue)
    txtRealDollar = Database.GetDatabaseValue("RealDollar", colUserValue)
    txtTaxesImportationEquipment = Database.GetDatabaseValue("TaxesImportationEquipment", colUserValue)
    txtAveragePriceBuildingLand = Database.GetDatabaseValue("AveragePriceBuildingLand", colUserValue)
    txtAverageLandPriceLandfill = Database.GetDatabaseValue("AverageLandPriceLandfill", colUserValue)
    txtAverageLandscapingPrice = Database.GetDatabaseValue("AverageLandscapingPrice", colUserValue)
    txtIncinerationEurBra = Database.GetDatabaseValue("IncinerationEurBra", colUserValue)
    txtAveragePriceIndustrialConcrete = Database.GetDatabaseValue("AveragePriceIndustrialConcrete", colUserValue)
    txtAveragePriceConstructionIndustrial = Database.GetDatabaseValue("AveragePriceConstructionIndustrial", colUserValue)
    txtNationalPrices = Database.GetDatabaseValue("NationalPrices", colUserValue)
    txtMechanizedTechnologyNationalization = Database.GetDatabaseValue("MechanizedTechnologyNationalization", colUserValue)
    txtMechanizedTechnologyOvercapacity = Database.GetDatabaseValue("MechanizedTechnologyOvercapacity", colUserValue)
    txtProductionTechnologyCDRNationalization = Database.GetDatabaseValue("ProductionTechnologyCDRNationalization", colUserValue)
    txtProductionTechnologyCDROvercapacity = Database.GetDatabaseValue("ProductionTechnologyCDROvercapacity", colUserValue)
    txtProductionTechnologyBIONationalization = Database.GetDatabaseValue("ProductionTechnologyBIONationalization", colUserValue)
    txtProductionTechnologyBIOOvercapacity = Database.GetDatabaseValue("ProductionTechnologyBIOOvercapacity", colUserValue)
    txtTechnologyAnaerobicNationalization = Database.GetDatabaseValue("TechnologyAnaerobicNationalization", colUserValue)
    txtTechnologyAnaerobicOvercapacity = Database.GetDatabaseValue("TechnologyAnaerobicOvercapacity", colUserValue)
    txtTechnologyCompostingNationalization = Database.GetDatabaseValue("TechnologyCompostingNationalization", colUserValue)
    txtTechnologyCompostingOvercapacity = Database.GetDatabaseValue("TechnologyCompostingOvercapacity", colUserValue)
    txtTechnologyIncinerationNationalization = Database.GetDatabaseValue("TechnologyIncinerationNationalization", colUserValue)
    txtTechnologyIncinerationOvercapacity = Database.GetDatabaseValue("TechnologyIncinerationOvercapacity", colUserValue)
    txtTechnologyLandfillOvercapacity = Database.GetDatabaseValue("TechnologyLandfillOvercapacity", colUserValue)
    
    FormChanged = False
End Sub

Private Sub btnSave_Click()
    If modForm.validateForm() Then
        Call Database.SetDatabaseValue("RealEuro", colUserValue, CDbl(txtRealEuro.Text))
        Call Database.SetDatabaseValue("RealDollar", colUserValue, CDbl(txtRealDollar.Text))
        Call Database.SetDatabaseValue("TaxesImportationEquipment", colUserValue, CDbl(txtTaxesImportationEquipment.Text))
        Call Database.SetDatabaseValue("AveragePriceBuildingLand", colUserValue, CDbl(txtAveragePriceBuildingLand.Text))
        Call Database.SetDatabaseValue("AverageLandPriceLandfill", colUserValue, CDbl(txtAverageLandPriceLandfill.Text))
        Call Database.SetDatabaseValue("AverageLandscapingPrice", colUserValue, CDbl(txtAverageLandscapingPrice.Text))
        Call Database.SetDatabaseValue("IncinerationEurBra", colUserValue, CDbl(txtIncinerationEurBra.Text))
        Call Database.SetDatabaseValue("AveragePriceIndustrialConcrete", colUserValue, CDbl(txtAveragePriceIndustrialConcrete.Text))
        Call Database.SetDatabaseValue("AveragePriceConstructionIndustrial", colUserValue, CDbl(txtAveragePriceConstructionIndustrial.Text))
        Call Database.SetDatabaseValue("NationalPrices", colUserValue, CDbl(txtNationalPrices.Text))
        Call Database.SetDatabaseValue("MechanizedTechnologyNationalization", colUserValue, CDbl(txtMechanizedTechnologyNationalization.Text))
        Call Database.SetDatabaseValue("MechanizedTechnologyOvercapacity", colUserValue, CDbl(txtMechanizedTechnologyOvercapacity.Text))
        Call Database.SetDatabaseValue("ProductionTechnologyCDRNationalization", colUserValue, CDbl(txtProductionTechnologyCDRNationalization.Text))
        Call Database.SetDatabaseValue("ProductionTechnologyCDROvercapacity", colUserValue, CDbl(txtProductionTechnologyCDROvercapacity.Text))
        Call Database.SetDatabaseValue("ProductionTechnologyBIONationalization", colUserValue, CDbl(txtProductionTechnologyBIONationalization.Text))
        Call Database.SetDatabaseValue("ProductionTechnologyBIOOvercapacity", colUserValue, CDbl(txtProductionTechnologyBIOOvercapacity.Text))
        Call Database.SetDatabaseValue("TechnologyAnaerobicNationalization", colUserValue, CDbl(txtTechnologyAnaerobicNationalization.Text))
        Call Database.SetDatabaseValue("TechnologyAnaerobicOvercapacity", colUserValue, CDbl(txtTechnologyAnaerobicOvercapacity.Text))
        Call Database.SetDatabaseValue("TechnologyCompostingNationalization", colUserValue, CDbl(txtTechnologyCompostingNationalization.Text))
        Call Database.SetDatabaseValue("TechnologyCompostingOvercapacity", colUserValue, CDbl(txtTechnologyCompostingOvercapacity.Text))
        Call Database.SetDatabaseValue("TechnologyIncinerationNationalization", colUserValue, CDbl(txtTechnologyIncinerationNationalization.Text))
        Call Database.SetDatabaseValue("TechnologyIncinerationOvercapacity", colUserValue, CDbl(txtTechnologyIncinerationOvercapacity.Text))
        Call Database.SetDatabaseValue("TechnologyLandfillOvercapacity", colUserValue, CDbl(txtTechnologyLandfillOvercapacity.Text))
        FormChanged = False
        Unload Me
    Else
        answer = MsgBox(MSG_INVALID_DATA, vbExclamation, MSG_INVALID_DATA_TITLE)
    End If
End Sub

Private Sub btnDefault_Click()
    txtRealEuro = Database.GetDatabaseValue("RealEuro", colDefaultValue)
    txtRealDollar = Database.GetDatabaseValue("RealDollar", colDefaultValue)
    txtTaxesImportationEquipment = Database.GetDatabaseValue("TaxesImportationEquipment", colDefaultValue)
    txtAveragePriceBuildingLand = Database.GetDatabaseValue("AveragePriceBuildingLand", colDefaultValue)
    txtAverageLandPriceLandfill = Database.GetDatabaseValue("AverageLandPriceLandfill", colDefaultValue)
    txtAverageLandscapingPrice = Database.GetDatabaseValue("AverageLandscapingPrice", colDefaultValue)
    txtIncinerationEurBra = Database.GetDatabaseValue("IncinerationEurBra", colDefaultValue)
    txtAveragePriceIndustrialConcrete = Database.GetDatabaseValue("AveragePriceIndustrialConcrete", colDefaultValue)
    txtAveragePriceConstructionIndustrial = Database.GetDatabaseValue("AveragePriceConstructionIndustrial", colDefaultValue)
    txtNationalPrices = Database.GetDatabaseValue("NationalPrices", colDefaultValue)
    txtMechanizedTechnologyNationalization = Database.GetDatabaseValue("MechanizedTechnologyNationalization", colDefaultValue)
    txtMechanizedTechnologyOvercapacity = Database.GetDatabaseValue("MechanizedTechnologyOvercapacity", colDefaultValue)
    txtProductionTechnologyCDRNationalization = Database.GetDatabaseValue("ProductionTechnologyCDRNationalization", colDefaultValue)
    txtProductionTechnologyCDROvercapacity = Database.GetDatabaseValue("ProductionTechnologyCDROvercapacity", colDefaultValue)
    txtProductionTechnologyBIONationalization = Database.GetDatabaseValue("ProductionTechnologyBIONationalization", colDefaultValue)
    txtProductionTechnologyBIOOvercapacity = Database.GetDatabaseValue("ProductionTechnologyBIOOvercapacity", colDefaultValue)
    txtTechnologyAnaerobicNationalization = Database.GetDatabaseValue("TechnologyAnaerobicNationalization", colDefaultValue)
    txtTechnologyAnaerobicOvercapacity = Database.GetDatabaseValue("TechnologyAnaerobicOvercapacity", colDefaultValue)
    txtTechnologyCompostingNationalization = Database.GetDatabaseValue("TechnologyCompostingNationalization", colDefaultValue)
    txtTechnologyCompostingOvercapacity = Database.GetDatabaseValue("TechnologyCompostingOvercapacity", colDefaultValue)
    txtTechnologyIncinerationNationalization = Database.GetDatabaseValue("TechnologyIncinerationNationalization", colDefaultValue)
    txtTechnologyIncinerationOvercapacity = Database.GetDatabaseValue("TechnologyIncinerationOvercapacity", colDefaultValue)
    txtTechnologyLandfillOvercapacity = Database.GetDatabaseValue("TechnologyLandfillOvercapacity", colDefaultValue)
End Sub
