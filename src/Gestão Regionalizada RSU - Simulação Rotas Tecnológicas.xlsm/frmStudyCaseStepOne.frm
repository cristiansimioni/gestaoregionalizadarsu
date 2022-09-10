VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStudyCaseStepOne 
   Caption         =   "Dados de Definição do Estudo de Caso"
   ClientHeight    =   3915
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   8625.001
   OleObjectBlob   =   "frmStudyCaseStepOne.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStudyCaseStepOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim GenerationPerCapitaRDO As Double
Dim IndexSelectiveColletionRSU As Double
Dim AnnualGrowthPopulation As Double
Dim AnnualGrowthCollect As Double
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

Function ValidateForm() As Boolean
    ValidateForm = True
End Function

Private Sub btnSave_Click()
    If ValidateForm() Then
        Call Database.SetDatabaseValue("GenerationPerCapitaRDO", colUserValue, CDbl(txtGenerationPerCapitaRDO.Text))
        Call Database.SetDatabaseValue("IndexSelectiveColletionRSU", colUserValue, CDbl(txtIndexSelectiveColletionRSU.Text))
        Call Database.SetDatabaseValue("AnnualGrowthPopulation", colUserValue, CDbl(txtAnnualGrowthPopulation.Text))
        Call Database.SetDatabaseValue("AnnualGrowthCollect", colUserValue, CDbl(txtAnnualGrowthCollect.Text))
        FormChanged = False
        Unload Me
        frmStepOne.updateForm
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

Private Sub txtGenerationPerCapitaRDO_Change()
    Call textBoxChange(txtGenerationPerCapitaRDO, "GenerationPerCapitaRDO")
End Sub

Private Sub txtAnnualGrowthPopulation_Change()
    Call textBoxChange(txtAnnualGrowthPopulation, "AnnualGrowthPopulation")
End Sub

Private Sub txtIndexSelectiveColletionRSU_Change()
    Call textBoxChange(txtIndexSelectiveColletionRSU, "IndexSelectiveColletionRSU")
End Sub

Private Sub txtAnnualGrowthCollect_Change()
    Call textBoxChange(txtAnnualGrowthCollect, "AnnualGrowthCollect")
End Sub


Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Definição do Estudo de Caso")
    
    'Read database values
    GenerationPerCapitaRDO = Database.GetDatabaseValue("GenerationPerCapitaRDO", colUserValue)
    IndexSelectiveColletionRSU = Database.GetDatabaseValue("IndexSelectiveColletionRSU", colUserValue)
    AnnualGrowthPopulation = Database.GetDatabaseValue("AnnualGrowthPopulation", colUserValue)
    AnnualGrowthCollect = Database.GetDatabaseValue("AnnualGrowthCollect", colUserValue)

    'Only show the data if it's available
    If GenerationPerCapitaRDO + IndexSelectiveColletionRSU + AnnualGrowthPopulation + _
       AnnualGrowthCollect <> 0 Then
        txtGenerationPerCapitaRDO.Text = GenerationPerCapitaRDO
        txtIndexSelectiveColletionRSU.Text = IndexSelectiveColletionRSU
        txtAnnualGrowthPopulation.Text = AnnualGrowthPopulation
        txtAnnualGrowthCollect.Text = AnnualGrowthCollect
    End If
    
    FormChanged = False
End Sub
