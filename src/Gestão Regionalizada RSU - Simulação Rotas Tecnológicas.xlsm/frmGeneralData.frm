VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGeneralData 
   Caption         =   "UserForm1"
   ClientHeight    =   2412
   ClientLeft      =   204
   ClientTop       =   744
   ClientWidth     =   6240
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

Private Sub btnDefault_Click()
    txtCOEmission.Text = Database.GetDatabaseValue("COEmission", colDefaultValue)
    txtReducingCostMovimentation.Text = Database.GetDatabaseValue("ReducingCostMovimentation", colDefaultValue)
    txtCapexInbound = Database.GetDatabaseValue("CapexInbound", colDefaultValue)
    txtCapexOutbound = Database.GetDatabaseValue("CapexOutbound", colDefaultValue)
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrorHandler
        Call Database.SetDatabaseValue("COEmission", colUserValue, CDbl(txtCOEmission.Text))
        Call Database.SetDatabaseValue("ReducingCostMovimentation", colUserValue, CDbl(txtReducingCostMovimentation.Text))
        Call Database.SetDatabaseValue("CapexInbound", colUserValue, CDbl(txtCapexInbound.Text))
        Call Database.SetDatabaseValue("CapexOutbound", colUserValue, CDbl(txtCapexOutbound.Text))
        FormChanged = False
        frmStepTwo.updateForm
        Unload Me
        ThisWorkbook.Save
    Exit Sub
    
ErrorHandler:
    Call MsgBox(MSG_INVALID_DATA, vbCritical, MSG_INVALID_DATA_TITLE)
    
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


Private Sub txtReducingCostMovimentation_Change()
    Call textBoxChange(txtReducingCostMovimentation, "ReducingCostMovimentation")
End Sub

Private Sub txtCapexInbound_Change()
    Call textBoxChange(txtCapexInbound, "CapexInbound")
End Sub

Private Sub txtCapexOutbound_Change()
    Call textBoxChange(txtCapexOutbound, "CapexOutbound")
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Dados Gerais")
    
    'Read database values
    COEmission = Database.GetDatabaseValue("COEmission", colUserValue)
    ReducingCostMovimentation = Database.GetDatabaseValue("ReducingCostMovimentation", colUserValue)
    capexInbound = Database.GetDatabaseValue("CapexInbound", colUserValue)
    capexOutbound = Database.GetDatabaseValue("CapexOutbound", colUserValue)

    'Only show the data if it's available
    If COEmission + ReducingCostMovimentation + CostWasteExistingLandfills + capexInbound + capexOutbound <> 0 Then
        txtCOEmission.Text = COEmission
        txtReducingCostMovimentation.Text = ReducingCostMovimentation
        txtCapexInbound = capexInbound
        txtCapexOutbound = capexOutbound
    End If
    
    FormChanged = False

    Me.Height = 217
    Me.width = 400
End Sub


