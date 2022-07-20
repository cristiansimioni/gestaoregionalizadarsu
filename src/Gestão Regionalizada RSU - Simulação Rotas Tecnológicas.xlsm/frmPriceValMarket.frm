VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPriceValMarket 
   Caption         =   "UserForm2"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600.001
   OleObjectBlob   =   "frmPriceValMarket.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPriceValMarket"
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

Private Sub txtElectricEnergyBiomassBase_Change()
Call modForm.textBoxChange(txtElectricEnergyBiomassBase, "ElectricEnergyBiomassBase", FormChanged)
End Sub
Private Sub txtElectricEnergySolidWasteBase_Change()
Call modForm.textBoxChange(txtElectricEnergySolidWasteBase, "ElectricEnergySolidWasteBase", FormChanged)
End Sub
Private Sub txtBiomethaneBase_Change()
Call modForm.textBoxChange(txtBiomethaneBase, "BiomethaneBase", FormChanged)
End Sub
Private Sub txtCDRBase_Change()
Call modForm.textBoxChange(txtCDRBase, "CDRBase", FormChanged)
End Sub
Private Sub txtOrganicCompoundBase_Change()
Call modForm.textBoxChange(txtOrganicCompoundBase, "OrganicCompoundBase", FormChanged)
End Sub
Private Sub txtDonationRevenueRecyclablesBase_Change()
Call modForm.textBoxChange(txtDonationRevenueRecyclablesBase, "DonationRevenueRecyclablesBase", FormChanged)
End Sub
Private Sub txtDonationRevenueCollectorsBase_Change()
Call modForm.textBoxChange(txtDonationRevenueCollectorsBase, "DonationRevenueCollectorsBase", FormChanged)
End Sub
Private Sub txtSalesRecyclablesOutsideStateBase_Change()
Call modForm.textBoxChange(txtSalesRecyclablesOutsideStateBase, "SalesRecyclablesOutsideStateBase", FormChanged)
End Sub
Private Sub txtSalePricePaperBase_Change()
Call modForm.textBoxChange(txtSalePricePaperBase, "SalePricePaperBase", FormChanged)
End Sub
Private Sub txtSalePricePlasticFilmeBase_Change()
Call modForm.textBoxChange(txtSalePricePlasticFilmeBase, "SalePricePlasticFilmeBase", FormChanged)
End Sub
Private Sub txtSalePriceRigidPlasticBase_Change()
Call modForm.textBoxChange(txtSalePriceRigidPlasticBase, "SalePriceRigidPlasticBase", FormChanged)
End Sub
Private Sub txtSalePriceGlassBase_Change()
Call modForm.textBoxChange(txtSalePriceGlassBase, "SalePriceGlassBase", FormChanged)
End Sub
Private Sub txtSalePriceFerrousMetalsBase_Change()
Call modForm.textBoxChange(txtSalePriceFerrousMetalsBase, "SalePriceFerrousMetalsBase", FormChanged)
End Sub
Private Sub txtSalePriceNonFerrousMetalsBase_Change()
Call modForm.textBoxChange(txtSalePriceNonFerrousMetalsBase, "SalePriceNonFerrousMetalsBase", FormChanged)
End Sub
Private Sub txtElectricEnergyBiomassOptimized_Change()
Call modForm.textBoxChange(txtElectricEnergyBiomassOptimized, "ElectricEnergyBiomassOptimized", FormChanged)
End Sub
Private Sub txtElectricEnergySolidWasteOptimized_Change()
Call modForm.textBoxChange(txtElectricEnergySolidWasteOptimized, "ElectricEnergySolidWasteOptimized", FormChanged)
End Sub
Private Sub txtBiomethaneOptimized_Change()
Call modForm.textBoxChange(txtBiomethaneOptimized, "BiomethaneOptimized", FormChanged)
End Sub
Private Sub txtCDROptimized_Change()
Call modForm.textBoxChange(txtCDROptimized, "CDROptimized", FormChanged)
End Sub
Private Sub txtOrganicCompoundOptimized_Change()
Call modForm.textBoxChange(txtOrganicCompoundOptimized, "OrganicCompoundOptimized", FormChanged)
End Sub
Private Sub txtDonationRevenueRecyclablesOptimized_Change()
Call modForm.textBoxChange(txtDonationRevenueRecyclablesOptimized, "DonationRevenueRecyclablesOptimized", FormChanged)
End Sub
Private Sub txtDonationRevenueCollectorsOptimized_Change()
Call modForm.textBoxChange(txtDonationRevenueCollectorsOptimized, "DonationRevenueCollectorsOptimized", FormChanged)
End Sub
Private Sub txtSalesRecyclablesOutsideStateOptimized_Change()
Call modForm.textBoxChange(txtSalesRecyclablesOutsideStateOptimized, "SalesRecyclablesOutsideStateOptimized", FormChanged)
End Sub
Private Sub txtSalePricePaperOptimized_Change()
Call modForm.textBoxChange(txtSalePricePaperOptimized, "SalePricePaperOptimized", FormChanged)
End Sub
Private Sub txtSalePricePlasticFilmeOptimized_Change()
Call modForm.textBoxChange(txtSalePricePlasticFilmeOptimized, "SalePricePlasticFilmeOptimized", FormChanged)
End Sub
Private Sub txtSalePriceRigidPlasticOptimized_Change()
Call modForm.textBoxChange(txtSalePriceRigidPlasticOptimized, "SalePriceRigidPlasticOptimized", FormChanged)
End Sub
Private Sub txtSalePriceGlassOptimized_Change()
Call modForm.textBoxChange(txtSalePriceGlassOptimized, "SalePriceGlassOptimized", FormChanged)
End Sub
Private Sub txtSalePriceFerrousMetalsOptimized_Change()
Call modForm.textBoxChange(txtSalePriceFerrousMetalsOptimized, "SalePriceFerrousMetalsOptimized", FormChanged)
End Sub
Private Sub txtSalePriceNonFerrousMetalsOptimized_Change()
Call modForm.textBoxChange(txtSalePriceNonFerrousMetalsOptimized, "SalePriceNonFerrousMetalsOptimized", FormChanged)
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Comercialização Mercado")
    
    txtElectricEnergyBiomassBase = Database.GetDatabaseValue("ElectricEnergyBiomassBase", colUserValue)
    txtElectricEnergySolidWasteBase = Database.GetDatabaseValue("ElectricEnergySolidWasteBase", colUserValue)
    txtBiomethaneBase = Database.GetDatabaseValue("BiomethaneBase", colUserValue)
    txtCDRBase = Database.GetDatabaseValue("CDRBase", colUserValue)
    txtOrganicCompoundBase = Database.GetDatabaseValue("OrganicCompoundBase", colUserValue)
    txtDonationRevenueRecyclablesBase = Database.GetDatabaseValue("DonationRevenueRecyclablesBase", colUserValue)
    txtDonationRevenueCollectorsBase = Database.GetDatabaseValue("DonationRevenueCollectorsBase", colUserValue)
    txtSalesRecyclablesOutsideStateBase = Database.GetDatabaseValue("SalesRecyclablesOutsideStateBase", colUserValue)
    txtSalePricePaperBase = Database.GetDatabaseValue("SalePricePaperBase", colUserValue)
    txtSalePricePlasticFilmeBase = Database.GetDatabaseValue("SalePricePlasticFilmeBase", colUserValue)
    txtSalePriceRigidPlasticBase = Database.GetDatabaseValue("SalePriceRigidPlasticBase", colUserValue)
    txtSalePriceGlassBase = Database.GetDatabaseValue("SalePriceGlassBase", colUserValue)
    txtSalePriceFerrousMetalsBase = Database.GetDatabaseValue("SalePriceFerrousMetalsBase", colUserValue)
    txtSalePriceNonFerrousMetalsBase = Database.GetDatabaseValue("SalePriceNonFerrousMetalsBase", colUserValue)
    txtElectricEnergyBiomassOptimized = Database.GetDatabaseValue("ElectricEnergyBiomassOptimized", colUserValue)
    txtElectricEnergySolidWasteOptimized = Database.GetDatabaseValue("ElectricEnergySolidWasteOptimized", colUserValue)
    txtBiomethaneOptimized = Database.GetDatabaseValue("BiomethaneOptimized", colUserValue)
    txtCDROptimized = Database.GetDatabaseValue("CDROptimized", colUserValue)
    txtOrganicCompoundOptimized = Database.GetDatabaseValue("OrganicCompoundOptimized", colUserValue)
    txtDonationRevenueRecyclablesOptimized = Database.GetDatabaseValue("DonationRevenueRecyclablesOptimized", colUserValue)
    txtDonationRevenueCollectorsOptimized = Database.GetDatabaseValue("DonationRevenueCollectorsOptimized", colUserValue)
    txtSalesRecyclablesOutsideStateOptimized = Database.GetDatabaseValue("SalesRecyclablesOutsideStateOptimized", colUserValue)
    txtSalePricePaperOptimized = Database.GetDatabaseValue("SalePricePaperOptimized", colUserValue)
    txtSalePricePlasticFilmeOptimized = Database.GetDatabaseValue("SalePricePlasticFilmeOptimized", colUserValue)
    txtSalePriceRigidPlasticOptimized = Database.GetDatabaseValue("SalePriceRigidPlasticOptimized", colUserValue)
    txtSalePriceGlassOptimized = Database.GetDatabaseValue("SalePriceGlassOptimized", colUserValue)
    txtSalePriceFerrousMetalsOptimized = Database.GetDatabaseValue("SalePriceFerrousMetalsOptimized", colUserValue)
    txtSalePriceNonFerrousMetalsOptimized = Database.GetDatabaseValue("SalePriceNonFerrousMetalsOptimized", colUserValue)

    FormChanged = False
End Sub

Private Sub btnSave_Click()
    If modForm.validateForm() Then
        Call Database.SetDatabaseValue("ElectricEnergyBiomassBase", colUserValue, CDbl(txtElectricEnergyBiomassBase.Text))
        Call Database.SetDatabaseValue("ElectricEnergySolidWasteBase", colUserValue, CDbl(txtElectricEnergySolidWasteBase.Text))
        Call Database.SetDatabaseValue("BiomethaneBase", colUserValue, CDbl(txtBiomethaneBase.Text))
        Call Database.SetDatabaseValue("CDRBase", colUserValue, CDbl(txtCDRBase.Text))
        Call Database.SetDatabaseValue("OrganicCompoundBase", colUserValue, CDbl(txtOrganicCompoundBase.Text))
        Call Database.SetDatabaseValue("DonationRevenueRecyclablesBase", colUserValue, txtDonationRevenueRecyclablesBase.Text)
        Call Database.SetDatabaseValue("DonationRevenueCollectorsBase", colUserValue, CDbl(txtDonationRevenueCollectorsBase.Text))
        Call Database.SetDatabaseValue("SalesRecyclablesOutsideStateBase", colUserValue, CDbl(txtSalesRecyclablesOutsideStateBase.Text))
        Call Database.SetDatabaseValue("SalePricePaperBase", colUserValue, CDbl(txtSalePricePaperBase.Text))
        Call Database.SetDatabaseValue("SalePricePlasticFilmeBase", colUserValue, CDbl(txtSalePricePlasticFilmeBase.Text))
        Call Database.SetDatabaseValue("SalePriceRigidPlasticBase", colUserValue, CDbl(txtSalePriceRigidPlasticBase.Text))
        Call Database.SetDatabaseValue("SalePriceGlassBase", colUserValue, CDbl(txtSalePriceGlassBase.Text))
        Call Database.SetDatabaseValue("SalePriceFerrousMetalsBase", colUserValue, CDbl(txtSalePriceFerrousMetalsBase.Text))
        Call Database.SetDatabaseValue("SalePriceNonFerrousMetalsBase", colUserValue, CDbl(txtSalePriceNonFerrousMetalsBase.Text))
        Call Database.SetDatabaseValue("ElectricEnergyBiomassOptimized", colUserValue, CDbl(txtElectricEnergyBiomassOptimized.Text))
        Call Database.SetDatabaseValue("ElectricEnergySolidWasteOptimized", colUserValue, CDbl(txtElectricEnergySolidWasteOptimized.Text))
        Call Database.SetDatabaseValue("BiomethaneOptimized", colUserValue, CDbl(txtBiomethaneOptimized.Text))
        Call Database.SetDatabaseValue("CDROptimized", colUserValue, CDbl(txtCDROptimized.Text))
        Call Database.SetDatabaseValue("OrganicCompoundOptimized", colUserValue, CDbl(txtOrganicCompoundOptimized.Text))
        Call Database.SetDatabaseValue("DonationRevenueRecyclablesOptimized", colUserValue, txtDonationRevenueRecyclablesOptimized.Text)
        Call Database.SetDatabaseValue("DonationRevenueCollectorsOptimized", colUserValue, CDbl(txtDonationRevenueCollectorsOptimized.Text))
        Call Database.SetDatabaseValue("SalesRecyclablesOutsideStateOptimized", colUserValue, CDbl(txtSalesRecyclablesOutsideStateOptimized.Text))
        Call Database.SetDatabaseValue("SalePricePaperOptimized", colUserValue, CDbl(txtSalePricePaperOptimized.Text))
        Call Database.SetDatabaseValue("SalePricePlasticFilmeOptimized", colUserValue, CDbl(txtSalePricePlasticFilmeOptimized.Text))
        Call Database.SetDatabaseValue("SalePriceRigidPlasticOptimized", colUserValue, CDbl(txtSalePriceRigidPlasticOptimized.Text))
        Call Database.SetDatabaseValue("SalePriceGlassOptimized", colUserValue, CDbl(txtSalePriceGlassOptimized.Text))
        Call Database.SetDatabaseValue("SalePriceFerrousMetalsOptimized", colUserValue, CDbl(txtSalePriceFerrousMetalsOptimized.Text))
        Call Database.SetDatabaseValue("SalePriceNonFerrousMetalsOptimized", colUserValue, CDbl(txtSalePriceNonFerrousMetalsOptimized.Text))
        FormChanged = False
        Unload Me
    Else
        answer = MsgBox(MSG_INVALID_DATA, vbExclamation, MSG_INVALID_DATA_TITLE)
    End If
End Sub

Private Sub btnDefault_Click()
    txtElectricEnergyBiomassBase = Database.GetDatabaseValue("ElectricEnergyBiomassBase", colDefaultValue)
    txtElectricEnergySolidWasteBase = Database.GetDatabaseValue("ElectricEnergySolidWasteBase", colDefaultValue)
    txtBiomethaneBase = Database.GetDatabaseValue("BiomethaneBase", colDefaultValue)
    txtCDRBase = Database.GetDatabaseValue("CDRBase", colDefaultValue)
    txtOrganicCompoundBase = Database.GetDatabaseValue("OrganicCompoundBase", colDefaultValue)
    txtDonationRevenueRecyclablesBase = Database.GetDatabaseValue("DonationRevenueRecyclablesBase", colDefaultValue)
    txtDonationRevenueCollectorsBase = Database.GetDatabaseValue("DonationRevenueCollectorsBase", colDefaultValue)
    txtSalesRecyclablesOutsideStateBase = Database.GetDatabaseValue("SalesRecyclablesOutsideStateBase", colDefaultValue)
    txtSalePricePaperBase = Database.GetDatabaseValue("SalePricePaperBase", colDefaultValue)
    txtSalePricePlasticFilmeBase = Database.GetDatabaseValue("SalePricePlasticFilmeBase", colDefaultValue)
    txtSalePriceRigidPlasticBase = Database.GetDatabaseValue("SalePriceRigidPlasticBase", colDefaultValue)
    txtSalePriceGlassBase = Database.GetDatabaseValue("SalePriceGlassBase", colDefaultValue)
    txtSalePriceFerrousMetalsBase = Database.GetDatabaseValue("SalePriceFerrousMetalsBase", colDefaultValue)
    txtSalePriceNonFerrousMetalsBase = Database.GetDatabaseValue("SalePriceNonFerrousMetalsBase", colDefaultValue)
    txtElectricEnergyBiomassOptimized = Database.GetDatabaseValue("ElectricEnergyBiomassOptimized", colDefaultValue)
    txtElectricEnergySolidWasteOptimized = Database.GetDatabaseValue("ElectricEnergySolidWasteOptimized", colDefaultValue)
    txtBiomethaneOptimized = Database.GetDatabaseValue("BiomethaneOptimized", colDefaultValue)
    txtCDROptimized = Database.GetDatabaseValue("CDROptimized", colDefaultValue)
    txtOrganicCompoundOptimized = Database.GetDatabaseValue("OrganicCompoundOptimized", colDefaultValue)
    txtDonationRevenueRecyclablesOptimized = Database.GetDatabaseValue("DonationRevenueRecyclablesOptimized", colDefaultValue)
    txtDonationRevenueCollectorsOptimized = Database.GetDatabaseValue("DonationRevenueCollectorsOptimized", colDefaultValue)
    txtSalesRecyclablesOutsideStateOptimized = Database.GetDatabaseValue("SalesRecyclablesOutsideStateOptimized", colDefaultValue)
    txtSalePricePaperOptimized = Database.GetDatabaseValue("SalePricePaperOptimized", colDefaultValue)
    txtSalePricePlasticFilmeOptimized = Database.GetDatabaseValue("SalePricePlasticFilmeOptimized", colDefaultValue)
    txtSalePriceRigidPlasticOptimized = Database.GetDatabaseValue("SalePriceRigidPlasticOptimized", colDefaultValue)
    txtSalePriceGlassOptimized = Database.GetDatabaseValue("SalePriceGlassOptimized", colDefaultValue)
    txtSalePriceFerrousMetalsOptimized = Database.GetDatabaseValue("SalePriceFerrousMetalsOptimized", colDefaultValue)
    txtSalePriceNonFerrousMetalsOptimized = Database.GetDatabaseValue("SalePriceNonFerrousMetalsOptimized", colDefaultValue)
End Sub

