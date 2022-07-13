VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGeneralData 
   Caption         =   "UserForm1"
   ClientHeight    =   3345
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   7755
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
        answer = MsgBox("Voc� realizou altera��es, gostaria de salvar?", vbQuestion + vbYesNo + vbDefaultButton2, "Salvar Altera��es")
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

Private Sub btnSave_Click()
    If validateForm() Then
        Call Database.SetDatabaseValue("COEmission", colUserValue, CDbl(txtCOEmission.Text))
        Call Database.SetDatabaseValue("AverageCostTransportation", colUserValue, CDbl(txtAverageCostTransportation.Text))
        Call Database.SetDatabaseValue("ReducingCostMovimentation", colUserValue, CDbl(txtReducingCostMovimentation.Text))
        Call Database.SetDatabaseValue("CostWasteExistingLandfills", colUserValue, CDbl(txtCostWasteExistingLandfills.Text))
        FormChanged = False
        Unload Me
    Else
        answer = MsgBox("Valores inv�lidos. Favor verificar!", vbExclamation, "Dados inv�lidos")
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
    Me.Caption = APPNAME & " - Defini��o do Estudo de Caso"
    Me.BackColor = ApplicationColors.frmBgColorLevel3
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.bgColorLevel3
         End If
    Next Ctrl
    
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


