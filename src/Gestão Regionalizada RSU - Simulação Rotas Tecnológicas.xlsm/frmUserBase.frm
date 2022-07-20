VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUserBase 
   Caption         =   "UserForm1"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "frmUserBase.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUserBase"
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

Private Sub txtHistoricalWaterConsumption_Change()
Call modForm.textBoxChange(txtHistoricalWaterConsumption, "HistoricalWaterConsumption", FormChanged)
End Sub
Private Sub txtMSWManagementUsers_Change()
Call modForm.textBoxChange(txtMSWManagementUsers, "MSWManagementUsers", FormChanged)
End Sub
Private Sub txtRegulatoryServiceCost_Change()
Call modForm.textBoxChange(txtRegulatoryServiceCost, "RegulatoryServiceCost", FormChanged)
End Sub
Private Sub txtBadDebtWaterCollectionSystem_Change()
Call modForm.textBoxChange(txtBadDebtWaterCollectionSystem, "BadDebtWaterCollectionSystem", FormChanged)
End Sub
Private Sub txtCostCollectionService_Change()
Call modForm.textBoxChange(txtCostCollectionService, "CostCollectionService", FormChanged)
End Sub
Private Sub txtUserSavingsTotalResidential_Change()
Call modForm.textBoxChange(txtUserSavingsTotalResidential, "UserSavingsTotalResidential", FormChanged)
End Sub
Private Sub txtUserSavingsResidentialSocial_Change()
Call modForm.textBoxChange(txtUserSavingsResidentialSocial, "UserSavingsResidentialSocial", FormChanged)
End Sub
Private Sub txtSocialTariffSubsidy_Change()
Call modForm.textBoxChange(txtSocialTariffSubsidy, "SocialTariffSubsidy", FormChanged)
End Sub
Private Sub txtUserSavingsCommercialCategory_Change()
Call modForm.textBoxChange(txtUserSavingsCommercialCategory, "UserSavingsCommercialCategory", FormChanged)
End Sub
Private Sub txtUserSavingsPublicPhilanthropicCategory_Change()
Call modForm.textBoxChange(txtUserSavingsPublicPhilanthropicCategory, "UserSavingsPublicPhilanthropicCategory", FormChanged)
End Sub
Private Sub txtUserSavingsIndustrialCategory_Change()
Call modForm.textBoxChange(txtUserSavingsIndustrialCategory, "UserSavingsIndustrialCategory", FormChanged)
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Base Usuários Distribuição Tarifa RSU")
    
    txtHistoricalWaterConsumption = Database.GetDatabaseValue("HistoricalWaterConsumption", colUserValue)
    txtMSWManagementUsers = Database.GetDatabaseValue("MSWManagementUsers", colUserValue)
    txtRegulatoryServiceCost = Database.GetDatabaseValue("RegulatoryServiceCost", colUserValue)
    txtBadDebtWaterCollectionSystem = Database.GetDatabaseValue("BadDebtWaterCollectionSystem", colUserValue)
    txtCostCollectionService = Database.GetDatabaseValue("CostCollectionService", colUserValue)
    txtUserSavingsTotalResidential = Database.GetDatabaseValue("UserSavingsTotalResidential", colUserValue)
    txtUserSavingsResidentialSocial = Database.GetDatabaseValue("UserSavingsResidentialSocial", colUserValue)
    txtSocialTariffSubsidy = Database.GetDatabaseValue("SocialTariffSubsidy", colUserValue)
    txtUserSavingsCommercialCategory = Database.GetDatabaseValue("UserSavingsCommercialCategory", colUserValue)
    txtUserSavingsPublicPhilanthropicCategory = Database.GetDatabaseValue("UserSavingsPublicPhilanthropicCategory", colUserValue)
    txtUserSavingsIndustrialCategory = Database.GetDatabaseValue("UserSavingsIndustrialCategory", colUserValue)
    
    FormChanged = False
End Sub

Private Sub btnSave_Click()
    If modForm.validateForm() Then
        Call Database.SetDatabaseValue("HistoricalWaterConsumption", colUserValue, CDbl(txtHistoricalWaterConsumption.Text))
        Call Database.SetDatabaseValue("MSWManagementUsers", colUserValue, CDbl(txtMSWManagementUsers.Text))
        Call Database.SetDatabaseValue("RegulatoryServiceCost", colUserValue, CDbl(txtRegulatoryServiceCost.Text))
        Call Database.SetDatabaseValue("BadDebtWaterCollectionSystem", colUserValue, CDbl(txtBadDebtWaterCollectionSystem.Text))
        Call Database.SetDatabaseValue("CostCollectionService", colUserValue, CDbl(txtCostCollectionService.Text))
        Call Database.SetDatabaseValue("UserSavingsTotalResidential", colUserValue, CDbl(txtUserSavingsTotalResidential.Text))
        Call Database.SetDatabaseValue("UserSavingsResidentialSocial", colUserValue, CDbl(txtUserSavingsResidentialSocial.Text))
        Call Database.SetDatabaseValue("SocialTariffSubsidy", colUserValue, CDbl(txtSocialTariffSubsidy.Text))
        Call Database.SetDatabaseValue("UserSavingsCommercialCategory", colUserValue, CDbl(txtUserSavingsCommercialCategory.Text))
        Call Database.SetDatabaseValue("UserSavingsPublicPhilanthropicCategory", colUserValue, CDbl(txtUserSavingsPublicPhilanthropicCategory.Text))
        Call Database.SetDatabaseValue("UserSavingsIndustrialCategory", colUserValue, CDbl(txtUserSavingsIndustrialCategory.Text))
        FormChanged = False
        Unload Me
    Else
        answer = MsgBox(MSG_INVALID_DATA, vbExclamation, MSG_INVALID_DATA_TITLE)
    End If
End Sub

Private Sub btnDefault_Click()
    txtHistoricalWaterConsumption = Database.GetDatabaseValue("HistoricalWaterConsumption", colDefaultValue)
    txtMSWManagementUsers = Database.GetDatabaseValue("MSWManagementUsers", colDefaultValue)
    txtRegulatoryServiceCost = Database.GetDatabaseValue("RegulatoryServiceCost", colDefaultValue)
    txtBadDebtWaterCollectionSystem = Database.GetDatabaseValue("BadDebtWaterCollectionSystem", colDefaultValue)
    txtCostCollectionService = Database.GetDatabaseValue("CostCollectionService", colDefaultValue)
    txtUserSavingsTotalResidential = Database.GetDatabaseValue("UserSavingsTotalResidential", colDefaultValue)
    txtUserSavingsResidentialSocial = Database.GetDatabaseValue("UserSavingsResidentialSocial", colDefaultValue)
    txtSocialTariffSubsidy = Database.GetDatabaseValue("SocialTariffSubsidy", colDefaultValue)
    txtUserSavingsCommercialCategory = Database.GetDatabaseValue("UserSavingsCommercialCategory", colDefaultValue)
    txtUserSavingsPublicPhilanthropicCategory = Database.GetDatabaseValue("UserSavingsPublicPhilanthropicCategory", colDefaultValue)
    txtUserSavingsIndustrialCategory = Database.GetDatabaseValue("UserSavingsIndustrialCategory", colDefaultValue)
End Sub

