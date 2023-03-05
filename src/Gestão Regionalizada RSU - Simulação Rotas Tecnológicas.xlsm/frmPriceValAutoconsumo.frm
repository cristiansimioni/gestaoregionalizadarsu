VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPriceValAutoconsumo 
   Caption         =   "Preços para Valorização - Autoconsumo"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9252.001
   OleObjectBlob   =   "frmPriceValAutoconsumo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPriceValAutoconsumo"
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


Private Sub txtCostPurchaseElectricityConcessionaireBase_Change()
Call modForm.textBoxChange(txtCostPurchaseElectricityConcessionaireBase, "CostPurchaseElectricityConcessionaireBase", FormChanged)
End Sub
Private Sub txtReferencePublicFuelCostAutBase_Change()
Call modForm.textBoxChange(txtReferencePublicFuelCostAutBase, "ReferencePublicFuelCostAutBase", FormChanged)
End Sub
Private Sub txtProposedPriceBiofuelAutBase_Change()
Call modForm.textBoxChange(txtProposedPriceBiofuelAutBase, "ProposedPriceBiofuelAutBase", FormChanged)
End Sub
Private Sub txtCostPurchaseElectricityConcessionaireOptimized_Change()
Call modForm.textBoxChange(txtCostPurchaseElectricityConcessionaireOptimized, "CostPurchaseElectricityConcessionaireOptimized", FormChanged)
End Sub
Private Sub txtReferencePublicFuelCostAutOptimized_Change()
Call modForm.textBoxChange(txtReferencePublicFuelCostAutOptimized, "ReferencePublicFuelCostAutOptimized", FormChanged)
End Sub
Private Sub txtProposedPriceBiofuelAutOptimized_Change()
Call modForm.textBoxChange(txtProposedPriceBiofuelAutOptimized, "ProposedPriceBiofuelAutOptimized", FormChanged)
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Autoconsumo")
    
    txtCostPurchaseElectricityConcessionaireBase = Database.GetDatabaseValue("CostPurchaseElectricityConcessionaireBase", colUserValue)
    txtReferencePublicFuelCostAutBase = Database.GetDatabaseValue("ReferencePublicFuelCostAutBase", colUserValue)
    txtProposedPriceBiofuelAutBase = Database.GetDatabaseValue("ProposedPriceBiofuelAutBase", colUserValue)
    txtCostPurchaseElectricityConcessionaireOptimized = Database.GetDatabaseValue("CostPurchaseElectricityConcessionaireOptimized", colUserValue)
    txtReferencePublicFuelCostAutOptimized = Database.GetDatabaseValue("ReferencePublicFuelCostAutOptimized", colUserValue)
    txtProposedPriceBiofuelAutOptimized = Database.GetDatabaseValue("ProposedPriceBiofuelAutOptimized", colUserValue)

    FormChanged = False
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrorHandler
        Call Database.SetDatabaseValue("CostPurchaseElectricityConcessionaireBase", colUserValue, CDbl(txtCostPurchaseElectricityConcessionaireBase.Text))
        Call Database.SetDatabaseValue("ReferencePublicFuelCostAutBase", colUserValue, CDbl(txtReferencePublicFuelCostAutBase.Text))
        Call Database.SetDatabaseValue("ProposedPriceBiofuelAutBase", colUserValue, CDbl(txtProposedPriceBiofuelAutBase.Text))
        Call Database.SetDatabaseValue("CostPurchaseElectricityConcessionaireOptimized", colUserValue, CDbl(txtCostPurchaseElectricityConcessionaireOptimized.Text))
        Call Database.SetDatabaseValue("ReferencePublicFuelCostAutOptimized", colUserValue, CDbl(txtReferencePublicFuelCostAutOptimized.Text))
        Call Database.SetDatabaseValue("ProposedPriceBiofuelAutOptimized", colUserValue, CDbl(txtProposedPriceBiofuelAutOptimized.Text))
        FormChanged = False
        frmStepFour.updateForm
        Unload Me
        ThisWorkbook.Save
    Exit Sub
    
ErrorHandler:
    Call MsgBox(MSG_INVALID_DATA, vbCritical, MSG_INVALID_DATA_TITLE)
End Sub

Private Sub btnDefault_Click()
    txtCostPurchaseElectricityConcessionaireBase = Database.GetDatabaseValue("CostPurchaseElectricityConcessionaireBase", colDefaultValue)
    txtReferencePublicFuelCostAutBase = Database.GetDatabaseValue("ReferencePublicFuelCostAutBase", colDefaultValue)
    txtProposedPriceBiofuelAutBase = Database.GetDatabaseValue("ProposedPriceBiofuelAutBase", colDefaultValue)
    txtCostPurchaseElectricityConcessionaireOptimized = Database.GetDatabaseValue("CostPurchaseElectricityConcessionaireOptimized", colDefaultValue)
    txtReferencePublicFuelCostAutOptimized = Database.GetDatabaseValue("ReferencePublicFuelCostAutOptimized", colDefaultValue)
    txtProposedPriceBiofuelAutOptimized = Database.GetDatabaseValue("ProposedPriceBiofuelAutOptimized", colDefaultValue)
End Sub


