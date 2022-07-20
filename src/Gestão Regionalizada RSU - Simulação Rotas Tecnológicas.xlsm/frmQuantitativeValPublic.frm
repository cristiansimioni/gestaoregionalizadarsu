VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuantitativeValPublic 
   Caption         =   "UserForm1"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10065
   OleObjectBlob   =   "frmQuantitativeValPublic.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmQuantitativeValPublic"
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

Private Sub txtVehicleQuantityBase_Change()
Call modForm.textBoxChange(txtVehicleQuantityBase, "VehicleQuantityBase", FormChanged)
End Sub
Private Sub txtMileageVehicleBase_Change()
Call modForm.textBoxChange(txtMileageVehicleBase, "MileageVehicleBase", FormChanged)
End Sub
Private Sub txtFleetRenewalTermVehicleBase_Change()
Call modForm.textBoxChange(txtFleetRenewalTermVehicleBase, "FleetRenewalTermVehicleBase", FormChanged)
End Sub
Private Sub txtInfrastructureCommercialBase_Change()
Call modForm.textBoxChange(txtInfrastructureCommercialBase, "InfrastructureCommercialBase", FormChanged)
End Sub
Private Sub txtVehicleQuantityOptimized_Change()
Call modForm.textBoxChange(txtVehicleQuantityOptimized, "VehicleQuantityOptimized", FormChanged)
End Sub
Private Sub txtMileageVehicleOptimized_Change()
Call modForm.textBoxChange(txtMileageVehicleOptimized, "MileageVehicleOptimized", FormChanged)
End Sub
Private Sub txtFleetRenewalTermVehicleOptimized_Change()
Call modForm.textBoxChange(txtFleetRenewalTermVehicleOptimized, "FleetRenewalTermVehicleOptimized", FormChanged)
End Sub
Private Sub txtInfrastructureCommercialOptimized_Change()
Call modForm.textBoxChange(txtInfrastructureCommercialOptimized, "InfrastructureCommercialOptimized", FormChanged)
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Quantitivo para Valorização > Utilidade Pública")
    
    txtVehicleQuantityBase = Database.GetDatabaseValue("VehicleQuantityBase", colUserValue)
    txtMileageVehicleBase = Database.GetDatabaseValue("MileageVehicleBase", colUserValue)
    txtFleetRenewalTermVehicleBase = Database.GetDatabaseValue("FleetRenewalTermVehicleBase", colUserValue)
    txtInfrastructureCommercialBase = Database.GetDatabaseValue("InfrastructureCommercialBase", colUserValue)
    txtVehicleQuantityOptimized = Database.GetDatabaseValue("VehicleQuantityOptimized", colUserValue)
    txtMileageVehicleOptimized = Database.GetDatabaseValue("MileageVehicleOptimized", colUserValue)
    txtFleetRenewalTermVehicleOptimized = Database.GetDatabaseValue("FleetRenewalTermVehicleOptimized", colUserValue)
    txtInfrastructureCommercialOptimized = Database.GetDatabaseValue("InfrastructureCommercialOptimized", colUserValue)

    FormChanged = False
End Sub

Private Sub btnSave_Click()
    If modForm.validateForm() Then
        Call Database.SetDatabaseValue("VehicleQuantityBase", colUserValue, CDbl(txtVehicleQuantityBase.Text))
        Call Database.SetDatabaseValue("MileageVehicleBase", colUserValue, CDbl(txtMileageVehicleBase.Text))
        Call Database.SetDatabaseValue("FleetRenewalTermVehicleBase", colUserValue, CDbl(txtFleetRenewalTermVehicleBase.Text))
        Call Database.SetDatabaseValue("InfrastructureCommercialBase", colUserValue, CDbl(txtInfrastructureCommercialBase.Text))
        Call Database.SetDatabaseValue("VehicleQuantityOptimized", colUserValue, CDbl(txtVehicleQuantityOptimized.Text))
        Call Database.SetDatabaseValue("MileageVehicleOptimized", colUserValue, CDbl(txtMileageVehicleOptimized.Text))
        Call Database.SetDatabaseValue("FleetRenewalTermVehicleOptimized", colUserValue, CDbl(txtFleetRenewalTermVehicleOptimized.Text))
        Call Database.SetDatabaseValue("InfrastructureCommercialOptimized", colUserValue, CDbl(txtInfrastructureCommercialOptimized.Text))
        FormChanged = False
        Unload Me
    Else
        answer = MsgBox(MSG_INVALID_DATA, vbExclamation, MSG_INVALID_DATA_TITLE)
    End If
End Sub

Private Sub btnDefault_Click()
    txtVehicleQuantityBase = Database.GetDatabaseValue("VehicleQuantityBase", colDefaultValue)
    txtMileageVehicleBase = Database.GetDatabaseValue("MileageVehicleBase", colDefaultValue)
    txtFleetRenewalTermVehicleBase = Database.GetDatabaseValue("FleetRenewalTermVehicleBase", colDefaultValue)
    txtInfrastructureCommercialBase = Database.GetDatabaseValue("InfrastructureCommercialBase", colDefaultValue)
    txtVehicleQuantityOptimized = Database.GetDatabaseValue("VehicleQuantityOptimized", colDefaultValue)
    txtMileageVehicleOptimized = Database.GetDatabaseValue("MileageVehicleOptimized", colDefaultValue)
    txtFleetRenewalTermVehicleOptimized = Database.GetDatabaseValue("FleetRenewalTermVehicleOptimized", colDefaultValue)
    txtInfrastructureCommercialOptimized = Database.GetDatabaseValue("InfrastructureCommercialOptimized", colDefaultValue)
End Sub





