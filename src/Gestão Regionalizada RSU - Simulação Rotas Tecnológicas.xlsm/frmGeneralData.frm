VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGeneralData 
   Caption         =   "UserForm1"
   ClientHeight    =   3780
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   7800
   OleObjectBlob   =   "frmGeneralData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGeneralData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COEmission As Double
Dim AverageCostTransportation As Double
Dim ReducingCostMovimentation As Double
Dim CostWasteExistingLandfills As Double
Dim FormChanged As Boolean

Private Sub btnBack_Click()
    If FormChanged Then
        answer = MsgBox("Você realizou alterações, gostaria de salvar?", vbQuestion + vbYesNo + vbDefaultButton2, "Salvar Alterações")
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

Private Sub btnSave_Click()
    If ValidateForm() Then
        Call Database.SetDatabaseValue("COEmission", colUserValue, CDbl(txtCOEmission.Text))
        Call Database.SetDatabaseValue("AverageCostTransportation", colUserValue, CDbl(txtAverageCostTransportation.Text))
        Call Database.SetDatabaseValue("ReducingCostMovimentation", colUserValue, CDbl(txtReducingCostMovimentation.Text))
        Call Database.SetDatabaseValue("CostWasteExistingLandfills", colUserValue, CDbl(txtCostWasteExistingLandfills.Text))
        FormChanged = False
        Unload Me
    Else
        answer = MsgBox("Valores inválidos. Favor verificar!", vbExclamation, "Dados inválidos")
    End If
End Sub

Private Sub textBoxChange(ByRef txtBox, ByVal varName As String)
    Dim errorMsg As String
    If Database.Validate(varName, txtBox.Text, errorMsg) Then
        txtBox.BackColor = ApplicationColors.bgColorValidTextBox
        txtBox.ControlTipText = errorMsg
    Else
        txtBox.BackColor = ApplicationColors.bgColorInvalidTextBox
        txtBox.ControlTipText = errorMsg
    End If
    FormChanged = True
End Sub

Private Sub txtCOEmission_Change()
    Call textBoxChange(txtCOEmission, "COEmission")
End Sub

Private Sub txtAverageCostTransportation_Change()
    Call textBoxChange(txtAverageCostTransportation, "AverageCostTransportation")
End Sub

Private Sub txtCostWasteExistingLandfills_Change()
    Call textBoxChange(txtCostWasteExistingLandfills, "CostWasteExistingLandfills")
End Sub

Private Sub txtReducingCostMovimentation_Change()
    Call textBoxChange(txtReducingCostMovimentation, "ReducingCostMovimentation")
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Definição do Estudo de Caso")
    
    'Read database values
    COEmission = Database.GetDatabaseValue("COEmission", colUserValue)
    AverageCostTransportation = Database.GetDatabaseValue("AverageCostTransportation", colUserValue)
    ReducingCostMovimentation = Database.GetDatabaseValue("ReducingCostMovimentation", colUserValue)
    CostWasteExistingLandfills = Database.GetDatabaseValue("CostWasteExistingLandfills", colUserValue)

    'Only show the data if it's available
    If COEmission + AverageCostTransportation + ReducingCostMovimentation + CostWasteExistingLandfills <> 0 Then
        txtCOEmission.Text = COEmission
        txtAverageCostTransportation.Text = AverageCostTransportation
        txtReducingCostMovimentation.Text = ReducingCostMovimentation
        txtCostWasteExistingLandfills.Text = CostWasteExistingLandfills
    End If
    
    FormChanged = False
End Sub


