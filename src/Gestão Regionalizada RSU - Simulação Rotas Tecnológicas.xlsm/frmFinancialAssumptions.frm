VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFinancialAssumptions 
   Caption         =   "UserForm1"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600.001
   OleObjectBlob   =   "frmFinancialAssumptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFinancialAssumptions"
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

Private Sub cbxFinancingInstitutionProject_Change()
    FormChanged = True
End Sub

Private Sub cbxFinancingInstitutionShareholder_Change()
    FormChanged = True
End Sub

Private Sub cbxVariableProject_Change()
    FormChanged = True
End Sub

Private Sub cbxVariableShareholder_Change()
    FormChanged = True
End Sub

Private Sub txtContractTerm_Change()
Call modForm.textBoxChange(txtContractTerm, "ContractTerm", FormChanged)
End Sub
Private Sub txtContractTermEquityProject_Change()
Call modForm.textBoxChange(txtContractTermEquityProject, "ContractTermEquityProject", FormChanged)
End Sub
Private Sub txtFinancingInstitutionProject_Change()
Call modForm.textBoxChange(txtFinancingInstitutionProject, "FinancingInstitutionProject", FormChanged)
End Sub
Private Sub txtOwnCapitalCostProject_Change()
Call modForm.textBoxChange(txtOwnCapitalCostProject, "OwnCapitalCostProject", FormChanged)
End Sub
Private Sub txtRealInterestRateProject_Change()
Call modForm.textBoxChange(txtRealInterestRateProject, "RealInterestRateProject", FormChanged)
End Sub
Private Sub txtLoanAmortizationPeriodProject_Change()
Call modForm.textBoxChange(txtLoanAmortizationPeriodProject, "LoanAmortizationPeriodProject", FormChanged)
End Sub
Private Sub txtGracePeriodPaymentProject_Change()
Call modForm.textBoxChange(txtGracePeriodPaymentProject, "GracePeriodPaymentProject", FormChanged)
End Sub
Private Sub txtInterestRateProject_Change()
Call modForm.textBoxChange(txtInterestRateProject, "InterestRateProject", FormChanged)
End Sub
Private Sub txtVariableProject_Change()
Call modForm.textBoxChange(txtVariableProject, "VariableProject", FormChanged)
End Sub
Private Sub txtTargetProject_Change()
Call modForm.textBoxChange(txtTargetProject, "TargetProject", FormChanged)
End Sub
Private Sub txtContractTermEquityShareholder_Change()
Call modForm.textBoxChange(txtContractTermEquityShareholder, "ContractTermEquityShareholder", FormChanged)
End Sub
Private Sub txtFinancingInstitutionShareholder_Change()
Call modForm.textBoxChange(txtFinancingInstitutionShareholder, "FinancingInstitutionShareholder", FormChanged)
End Sub
Private Sub txtOwnCapitalCostShareholder_Change()
Call modForm.textBoxChange(txtOwnCapitalCostShareholder, "OwnCapitalCostShareholder", FormChanged)
End Sub
Private Sub txtRealInterestRateShareholder_Change()
Call modForm.textBoxChange(txtRealInterestRateShareholder, "RealInterestRateShareholder", FormChanged)
End Sub
Private Sub txtLoanAmortizationPeriodShareholder_Change()
Call modForm.textBoxChange(txtLoanAmortizationPeriodShareholder, "LoanAmortizationPeriodShareholder", FormChanged)
End Sub
Private Sub txtGracePeriodPaymentShareholder_Change()
Call modForm.textBoxChange(txtGracePeriodPaymentShareholder, "GracePeriodPaymentShareholder", FormChanged)
End Sub
Private Sub txtInterestRateShareholder_Change()
Call modForm.textBoxChange(txtInterestRateShareholder, "InterestRateShareholder", FormChanged)
End Sub
Private Sub txtVariableShareholder_Change()
Call modForm.textBoxChange(txtVariableShareholder, "VariableShareholder", FormChanged)
End Sub
Private Sub txtTargetShareholder_Change()
Call modForm.textBoxChange(txtTargetShareholder, "TargetShareholder", FormChanged)
End Sub


Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Premissas Financeiras")
    
    
    'Combo box
    Dim index As Integer
    index = 0
    Dim valuesFinancingInstitutionProject
    valuesFinancingInstitutionProject = Split(Database.GetDatabaseValue("FinancingInstitutionProject", colUnit), ",")
    For Each v In valuesFinancingInstitutionProject
        cbxFinancingInstitutionProject.AddItem v
        If v = Database.GetDatabaseValue("FinancingInstitutionProject", colUserValue) Then
            cbxFinancingInstitutionProject.ListIndex = index
        End If
        index = index + 1
    Next v
    index = 0
    Dim valuesFinancingInstitutionShareholder
    valuesFinancingInstitutionShareholder = Split(Database.GetDatabaseValue("FinancingInstitutionShareholder", colUnit), ",")
    For Each v In valuesFinancingInstitutionShareholder
        cbxFinancingInstitutionShareholder.AddItem v
        If v = Database.GetDatabaseValue("FinancingInstitutionShareholder", colUserValue) Then
            cbxFinancingInstitutionShareholder.ListIndex = index
        End If
        index = index + 1
    Next v
    index = 0
    Dim valuesVariableProject
    valuesVariableProject = Split(Database.GetDatabaseValue("VariableProject", colUnit), ",")
    For Each v In valuesVariableProject
        cbxVariableProject.AddItem v
        If v = Database.GetDatabaseValue("VariableProject", colUserValue) Then
            cbxVariableProject.ListIndex = index
        End If
        index = index + 1
    Next v
    index = 0
    Dim valuesVariableShareholder
    valuesVariableShareholder = Split(Database.GetDatabaseValue("VariableShareholder", colUnit), ",")
    For Each v In valuesVariableShareholder
        cbxVariableShareholder.AddItem v
        If v = Database.GetDatabaseValue("VariableShareholder", colUserValue) Then
            cbxVariableShareholder.ListIndex = index
        End If
        index = index + 1
    Next v
    
    txtContractTerm = Database.GetDatabaseValue("ContractTerm", colUserValue)
    txtContractTermEquityProject = Database.GetDatabaseValue("ContractTermEquityProject", colUserValue)
    txtOwnCapitalCostProject = Database.GetDatabaseValue("OwnCapitalCostProject", colUserValue)
    txtRealInterestRateProject = Database.GetDatabaseValue("RealInterestRateProject", colUserValue)
    txtLoanAmortizationPeriodProject = Database.GetDatabaseValue("LoanAmortizationPeriodProject", colUserValue)
    txtGracePeriodPaymentProject = Database.GetDatabaseValue("GracePeriodPaymentProject", colUserValue)
    txtInterestRateProject = Database.GetDatabaseValue("InterestRateProject", colUserValue)
    txtTargetProject = Database.GetDatabaseValue("TargetProject", colUserValue)
    txtContractTermEquityShareholder = Database.GetDatabaseValue("ContractTermEquityShareholder", colUserValue)
    txtOwnCapitalCostShareholder = Database.GetDatabaseValue("OwnCapitalCostShareholder", colUserValue)
    txtRealInterestRateShareholder = Database.GetDatabaseValue("RealInterestRateShareholder", colUserValue)
    txtLoanAmortizationPeriodShareholder = Database.GetDatabaseValue("LoanAmortizationPeriodShareholder", colUserValue)
    txtGracePeriodPaymentShareholder = Database.GetDatabaseValue("GracePeriodPaymentShareholder", colUserValue)
    txtInterestRateShareholder = Database.GetDatabaseValue("InterestRateShareholder", colUserValue)
    txtTargetShareholder = Database.GetDatabaseValue("TargetShareholder", colUserValue)

    FormChanged = False
End Sub

Private Sub btnSave_Click()
    If modForm.validateForm() Then
        Call Database.SetDatabaseValue("ContractTerm", colUserValue, CDbl(txtContractTerm.Text))
        Call Database.SetDatabaseValue("ContractTermEquityProject", colUserValue, CDbl(txtContractTermEquityProject.Text))
        Call Database.SetDatabaseValue("FinancingInstitutionProject", colUserValue, cbxFinancingInstitutionProject.value)
        Call Database.SetDatabaseValue("OwnCapitalCostProject", colUserValue, CDbl(txtOwnCapitalCostProject.Text))
        Call Database.SetDatabaseValue("RealInterestRateProject", colUserValue, CDbl(txtRealInterestRateProject.Text))
        Call Database.SetDatabaseValue("LoanAmortizationPeriodProject", colUserValue, CDbl(txtLoanAmortizationPeriodProject.Text))
        Call Database.SetDatabaseValue("GracePeriodPaymentProject", colUserValue, CDbl(txtGracePeriodPaymentProject.Text))
        Call Database.SetDatabaseValue("InterestRateProject", colUserValue, CDbl(txtInterestRateProject.Text))
        Call Database.SetDatabaseValue("VariableProject", colUserValue, cbxVariableProject.value)
        Call Database.SetDatabaseValue("TargetProject", colUserValue, CDbl(txtTargetProject.Text))
        Call Database.SetDatabaseValue("ContractTermEquityShareholder", colUserValue, CDbl(txtContractTermEquityShareholder.Text))
        Call Database.SetDatabaseValue("FinancingInstitutionShareholder", colUserValue, cbxFinancingInstitutionShareholder.value)
        Call Database.SetDatabaseValue("OwnCapitalCostShareholder", colUserValue, CDbl(txtOwnCapitalCostShareholder.Text))
        Call Database.SetDatabaseValue("RealInterestRateShareholder", colUserValue, CDbl(txtRealInterestRateShareholder.Text))
        Call Database.SetDatabaseValue("LoanAmortizationPeriodShareholder", colUserValue, CDbl(txtLoanAmortizationPeriodShareholder.Text))
        Call Database.SetDatabaseValue("GracePeriodPaymentShareholder", colUserValue, CDbl(txtGracePeriodPaymentShareholder.Text))
        Call Database.SetDatabaseValue("InterestRateShareholder", colUserValue, CDbl(txtInterestRateShareholder.Text))
        Call Database.SetDatabaseValue("VariableShareholder", colUserValue, cbxVariableShareholder.value)
        Call Database.SetDatabaseValue("TargetShareholder", colUserValue, CDbl(txtTargetShareholder.Text))
        FormChanged = False
        Unload Me
    Else
        answer = MsgBox(MSG_INVALID_DATA, vbExclamation, MSG_INVALID_DATA_TITLE)
    End If
End Sub

Private Sub btnDefault_Click()

    'Combo box
    Dim index As Integer
    index = 0
    Dim valuesFinancingInstitutionProject
    valuesFinancingInstitutionProject = Split(Database.GetDatabaseValue("FinancingInstitutionProject", colUnit), ",")
    For Each v In valuesFinancingInstitutionProject
        If v = Database.GetDatabaseValue("FinancingInstitutionProject", colDefaultValue) Then
            cbxFinancingInstitutionProject.ListIndex = index
        End If
        index = index + 1
    Next v
    index = 0
    Dim valuesFinancingInstitutionShareholder
    valuesFinancingInstitutionShareholder = Split(Database.GetDatabaseValue("FinancingInstitutionShareholder", colUnit), ",")
    For Each v In valuesFinancingInstitutionShareholder
        If v = Database.GetDatabaseValue("FinancingInstitutionShareholder", colDefaultValue) Then
            cbxFinancingInstitutionShareholder.ListIndex = index
        End If
        index = index + 1
    Next v
    index = 0
    Dim valuesVariableProject
    valuesVariableProject = Split(Database.GetDatabaseValue("VariableProject", colUnit), ",")
    For Each v In valuesVariableProject
        If v = Database.GetDatabaseValue("VariableProject", colDefaultValue) Then
            cbxVariableProject.ListIndex = index
        End If
        index = index + 1
    Next v
    index = 0
    Dim valuesVariableShareholder
    valuesVariableShareholder = Split(Database.GetDatabaseValue("VariableShareholder", colUnit), ",")
    For Each v In valuesVariableShareholder
        If v = Database.GetDatabaseValue("VariableShareholder", colDefaultValue) Then
            cbxVariableShareholder.ListIndex = index
        End If
        index = index + 1
    Next v
    
    txtContractTerm = Database.GetDatabaseValue("ContractTerm", colDefaultValue)
    txtContractTermEquityProject = Database.GetDatabaseValue("ContractTermEquityProject", colDefaultValue)
    txtOwnCapitalCostProject = Database.GetDatabaseValue("OwnCapitalCostProject", colDefaultValue)
    txtRealInterestRateProject = Database.GetDatabaseValue("RealInterestRateProject", colDefaultValue)
    txtLoanAmortizationPeriodProject = Database.GetDatabaseValue("LoanAmortizationPeriodProject", colDefaultValue)
    txtGracePeriodPaymentProject = Database.GetDatabaseValue("GracePeriodPaymentProject", colDefaultValue)
    txtInterestRateProject = Database.GetDatabaseValue("InterestRateProject", colDefaultValue)
    txtTargetProject = Database.GetDatabaseValue("TargetProject", colDefaultValue)
    txtContractTermEquityShareholder = Database.GetDatabaseValue("ContractTermEquityShareholder", colDefaultValue)
    txtOwnCapitalCostShareholder = Database.GetDatabaseValue("OwnCapitalCostShareholder", colDefaultValue)
    txtRealInterestRateShareholder = Database.GetDatabaseValue("RealInterestRateShareholder", colDefaultValue)
    txtLoanAmortizationPeriodShareholder = Database.GetDatabaseValue("LoanAmortizationPeriodShareholder", colDefaultValue)
    txtGracePeriodPaymentShareholder = Database.GetDatabaseValue("GracePeriodPaymentShareholder", colDefaultValue)
    txtInterestRateShareholder = Database.GetDatabaseValue("InterestRateShareholder", colDefaultValue)
    txtTargetShareholder = Database.GetDatabaseValue("TargetShareholder", colDefaultValue)
End Sub

