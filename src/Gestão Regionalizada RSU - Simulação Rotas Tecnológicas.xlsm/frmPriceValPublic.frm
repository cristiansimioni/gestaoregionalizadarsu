VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPriceValPublic 
   Caption         =   "UserForm1"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8880.001
   OleObjectBlob   =   "frmPriceValPublic.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPriceValPublic"
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

Private Sub txtReferencePublicCostElectricityBase_Change()
Call modForm.textBoxChange(txtReferencePublicCostElectricityBase, "ReferencePublicCostElectricityBase", FormChanged)
End Sub
Private Sub txtReferenceProposedPriceElectricityBase_Change()
Call modForm.textBoxChange(txtReferenceProposedPriceElectricityBase, "ReferenceProposedPriceElectricityBase", FormChanged)
End Sub
Private Sub txtReferencePublicFuelCostBase_Change()
Call modForm.textBoxChange(txtReferencePublicFuelCostBase, "ReferencePublicFuelCostBase", FormChanged)
End Sub
Private Sub txtProposedPriceBiofuelBase_Change()
Call modForm.textBoxChange(txtProposedPriceBiofuelBase, "ProposedPriceBiofuelBase", FormChanged)
End Sub
Private Sub txtReferencePublicCostElectricityOptimized_Change()
Call modForm.textBoxChange(txtReferencePublicCostElectricityOptimized, "ReferencePublicCostElectricityOptimized", FormChanged)
End Sub
Private Sub txtReferenceProposedPriceElectricityOptimized_Change()
Call modForm.textBoxChange(txtReferenceProposedPriceElectricityOptimized, "ReferenceProposedPriceElectricityOptimized", FormChanged)
End Sub
Private Sub txtReferencePublicFuelCostOptimized_Change()
Call modForm.textBoxChange(txtReferencePublicFuelCostOptimized, "ReferencePublicFuelCostOptimized", FormChanged)
End Sub
Private Sub txtProposedPriceBiofuelOptimized_Change()
Call modForm.textBoxChange(txtProposedPriceBiofuelOptimized, "ProposedPriceBiofuelOptimized", FormChanged)
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Utilidade Pública")
    
    txtReferencePublicCostElectricityBase = Database.GetDatabaseValue("ReferencePublicCostElectricityBase", colUserValue)
    txtReferenceProposedPriceElectricityBase = Database.GetDatabaseValue("ReferenceProposedPriceElectricityBase", colUserValue)
    txtReferencePublicFuelCostBase = Database.GetDatabaseValue("ReferencePublicFuelCostBase", colUserValue)
    txtProposedPriceBiofuelBase = Database.GetDatabaseValue("ProposedPriceBiofuelBase", colUserValue)
    txtReferencePublicCostElectricityOptimized = Database.GetDatabaseValue("ReferencePublicCostElectricityOptimized", colUserValue)
    txtReferenceProposedPriceElectricityOptimized = Database.GetDatabaseValue("ReferenceProposedPriceElectricityOptimized", colUserValue)
    txtReferencePublicFuelCostOptimized = Database.GetDatabaseValue("ReferencePublicFuelCostOptimized", colUserValue)
    txtProposedPriceBiofuelOptimized = Database.GetDatabaseValue("ProposedPriceBiofuelOptimized", colUserValue)

    FormChanged = False
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrorHandler
        Call Database.SetDatabaseValue("ReferencePublicCostElectricityBase", colUserValue, CDbl(txtReferencePublicCostElectricityBase.Text))
        Call Database.SetDatabaseValue("ReferenceProposedPriceElectricityBase", colUserValue, CDbl(txtReferenceProposedPriceElectricityBase.Text))
        Call Database.SetDatabaseValue("ReferencePublicFuelCostBase", colUserValue, CDbl(txtReferencePublicFuelCostBase.Text))
        Call Database.SetDatabaseValue("ProposedPriceBiofuelBase", colUserValue, CDbl(txtProposedPriceBiofuelBase.Text))
        Call Database.SetDatabaseValue("ReferencePublicCostElectricityOptimized", colUserValue, CDbl(txtReferencePublicCostElectricityOptimized.Text))
        Call Database.SetDatabaseValue("ReferenceProposedPriceElectricityOptimized", colUserValue, CDbl(txtReferenceProposedPriceElectricityOptimized.Text))
        Call Database.SetDatabaseValue("ReferencePublicFuelCostOptimized", colUserValue, CDbl(txtReferencePublicFuelCostOptimized.Text))
        Call Database.SetDatabaseValue("ProposedPriceBiofuelOptimized", colUserValue, CDbl(txtProposedPriceBiofuelOptimized.Text))
        FormChanged = False
        frmStepFour.updateForm
        Unload Me
        ThisWorkbook.Save
    Exit Sub
    
ErrorHandler:
    Call MsgBox(MSG_INVALID_DATA, vbCritical, MSG_INVALID_DATA_TITLE)
End Sub

Private Sub btnDefault_Click()
    txtReferencePublicCostElectricityBase = Database.GetDatabaseValue("ReferencePublicCostElectricityBase", colDefaultValue)
    txtReferenceProposedPriceElectricityBase = Database.GetDatabaseValue("ReferenceProposedPriceElectricityBase", colDefaultValue)
    txtReferencePublicFuelCostBase = Database.GetDatabaseValue("ReferencePublicFuelCostBase", colDefaultValue)
    txtProposedPriceBiofuelBase = Database.GetDatabaseValue("ProposedPriceBiofuelBase", colDefaultValue)
    txtReferencePublicCostElectricityOptimized = Database.GetDatabaseValue("ReferencePublicCostElectricityOptimized", colDefaultValue)
    txtReferenceProposedPriceElectricityOptimized = Database.GetDatabaseValue("ReferenceProposedPriceElectricityOptimized", colDefaultValue)
    txtReferencePublicFuelCostOptimized = Database.GetDatabaseValue("ReferencePublicFuelCostOptimized", colDefaultValue)
    txtProposedPriceBiofuelOptimized = Database.GetDatabaseValue("ProposedPriceBiofuelOptimized", colDefaultValue)
End Sub
