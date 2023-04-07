VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuantitativeValAutoconsumo 
   Caption         =   "UserForm1"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10080
   OleObjectBlob   =   "frmQuantitativeValAutoconsumo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmQuantitativeValAutoconsumo"
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

Private Sub txtQuantityTrucksBase_Change()
Call modForm.textBoxChange(txtQuantityTrucksBase, "QuantityTrucksBase", FormChanged)
End Sub
Private Sub txtMileageTruckBase_Change()
Call modForm.textBoxChange(txtMileageTruckBase, "MileageTruckBase", FormChanged)
End Sub
Private Sub txtFleetRenewalTermTruckBase_Change()
Call modForm.textBoxChange(txtFleetRenewalTermTruckBase, "FleetRenewalTermTruckBase", FormChanged)
End Sub
Private Sub txtInfrastructureBiomethaneBase_Change()
Call modForm.textBoxChange(txtInfrastructureBiomethaneBase, "InfrastructureBiomethaneBase", FormChanged)
End Sub
Private Sub txtQuantityTrucksOptimized_Change()
Call modForm.textBoxChange(txtQuantityTrucksOptimized, "QuantityTrucksOptimized", FormChanged)
End Sub
Private Sub txtMileageTruckOptimized_Change()
Call modForm.textBoxChange(txtMileageTruckOptimized, "MileageTruckOptimized", FormChanged)
End Sub
Private Sub txtFleetRenewalTermTruckOptimized_Change()
Call modForm.textBoxChange(txtFleetRenewalTermTruckOptimized, "FleetRenewalTermTruckOptimized", FormChanged)
End Sub
Private Sub txtInfrastructureBiomethaneOptimized_Change()
Call modForm.textBoxChange(txtInfrastructureBiomethaneOptimized, "InfrastructureBiomethaneOptimized", FormChanged)
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Autoconsumo")
    
    txtQuantityTrucksBase = Database.GetDatabaseValue("QuantityTrucksBase", colUserValue)
    txtMileageTruckBase = Database.GetDatabaseValue("MileageTruckBase", colUserValue)
    txtFleetRenewalTermTruckBase = Database.GetDatabaseValue("FleetRenewalTermTruckBase", colUserValue)
    txtInfrastructureBiomethaneBase = Database.GetDatabaseValue("InfrastructureBiomethaneBase", colUserValue)
    txtQuantityTrucksOptimized = Database.GetDatabaseValue("QuantityTrucksOptimized", colUserValue)
    txtMileageTruckOptimized = Database.GetDatabaseValue("MileageTruckOptimized", colUserValue)
    txtFleetRenewalTermTruckOptimized = Database.GetDatabaseValue("FleetRenewalTermTruckOptimized", colUserValue)
    txtInfrastructureBiomethaneOptimized = Database.GetDatabaseValue("InfrastructureBiomethaneOptimized", colUserValue)
    
    FormChanged = False
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrorHandler
        Call Database.SetDatabaseValue("QuantityTrucksBase", colUserValue, CDbl(txtQuantityTrucksBase.Text))
        Call Database.SetDatabaseValue("MileageTruckBase", colUserValue, CDbl(txtMileageTruckBase.Text))
        Call Database.SetDatabaseValue("FleetRenewalTermTruckBase", colUserValue, CDbl(txtFleetRenewalTermTruckBase.Text))
        Call Database.SetDatabaseValue("InfrastructureBiomethaneBase", colUserValue, CDbl(txtInfrastructureBiomethaneBase.Text))
        Call Database.SetDatabaseValue("QuantityTrucksOptimized", colUserValue, CDbl(txtQuantityTrucksOptimized.Text))
        Call Database.SetDatabaseValue("MileageTruckOptimized", colUserValue, CDbl(txtMileageTruckOptimized.Text))
        Call Database.SetDatabaseValue("FleetRenewalTermTruckOptimized", colUserValue, CDbl(txtFleetRenewalTermTruckOptimized.Text))
        Call Database.SetDatabaseValue("InfrastructureBiomethaneOptimized", colUserValue, CDbl(txtInfrastructureBiomethaneOptimized.Text))
        FormChanged = False
        frmStepFour.updateForm
        Unload Me
        ThisWorkbook.Save
    Exit Sub
    
ErrorHandler:
    Call MsgBox(MSG_INVALID_DATA, vbCritical, MSG_INVALID_DATA_TITLE)
End Sub

Private Sub btnDefault_Click()
    txtQuantityTrucksBase = Database.GetDatabaseValue("QuantityTrucksBase", colDefaultValue)
    txtMileageTruckBase = Database.GetDatabaseValue("MileageTruckBase", colDefaultValue)
    txtFleetRenewalTermTruckBase = Database.GetDatabaseValue("FleetRenewalTermTruckBase", colDefaultValue)
    txtInfrastructureBiomethaneBase = Database.GetDatabaseValue("InfrastructureBiomethaneBase", colDefaultValue)
    txtQuantityTrucksOptimized = Database.GetDatabaseValue("QuantityTrucksOptimized", colDefaultValue)
    txtMileageTruckOptimized = Database.GetDatabaseValue("MileageTruckOptimized", colDefaultValue)
    txtFleetRenewalTermTruckOptimized = Database.GetDatabaseValue("FleetRenewalTermTruckOptimized", colDefaultValue)
    txtInfrastructureBiomethaneOptimized = Database.GetDatabaseValue("InfrastructureBiomethaneOptimized", colDefaultValue)
End Sub

