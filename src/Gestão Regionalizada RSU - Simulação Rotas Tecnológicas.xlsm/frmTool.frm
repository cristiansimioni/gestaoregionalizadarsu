VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTool 
   ClientHeight    =   8595.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15225
   OleObjectBlob   =   "frmTool.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnAbout_Click()
    frmAbout.Show
End Sub

Private Sub btnClean_Click()
    Dim answer As Integer
    answer = MsgBox(MSG_CLEAN_DATABASE, vbExclamation + vbYesNo + vbDefaultButton2, MSG_ATTENTION)
    If answer = vbYes Then
        Database.Clean
        Call updateForm
    End If
End Sub

Private Sub btnHelp_Click()
    On Error Resume Next
        ActiveWorkbook.FollowHyperlink (Application.ThisWorkbook.Path & "\" & FOLDERMANUAL & "\" & FILEMANUAL)
    On Error GoTo 0
End Sub

Private Sub btnStepFive_Click()
    frmStepFive.Show
End Sub

Private Sub btnStepFour_Click()
    frmStepFour.Show
End Sub

Private Sub btnStepOne_Click()
    frmStepOne.Show
End Sub

Private Sub btnStepSix_Click()
    frmStepSix.Show
End Sub

Private Sub btnStepThree_Click()
    frmStepThree.Show
End Sub

Private Sub btnStepTwo_Click()
    frmStepTwo.Show
End Sub

Public Function updateForm()
    imgStepOneStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgStepTwoStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgStepThreeStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgStepFourStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgStepFiveStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgStepSixStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    btnStepOne.Enabled = True
    btnStepTwo.Enabled = False
    btnStepThree.Enabled = False
    btnStepFour.Enabled = False
    btnStepFive.Enabled = False
    btnStepSix.Enabled = False
    
    If ValidateFormRules("frmRSUGravimetry") And _
        ValidateFormRules("frmStudyCaseStepOne") And _
        ValidateFormRules("frmSimulationData") And _
        ValidateFormRules("frmStepOne") And _
        readSelectedCities.count >= 2 Then
        imgStepOneStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
        btnStepTwo.Enabled = True
    End If
    
    If ValidateFormRules("frmGeneralData") And _
        ValidateFormRules("frmAlgorithmParameter") And _
        ValidateFormRules("frmStepTwo") And _
        btnStepTwo.Enabled = True Then
        imgStepTwoStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
        btnStepThree.Enabled = True
    End If
    
    If ValidateFormRules("frmRoute") And _
        ValidateFormRules("frmCapexData") And _
        ValidateFormRules("frmOpexData") And _
        ValidateFormRules("frmTaxes") And _
        ValidateFormRules("frmContract") And _
        ValidateFormRules("frmFinancialAssumptions") And _
        ValidateFormRules("frmUserBase") And _
        btnStepThree.Enabled = True Then
        imgStepThreeStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
        btnStepFour.Enabled = True
    End If
    
    If ValidateFormRules("frmPriceValRevenue") And _
        ValidateFormRules("frmPriceValMarket") And _
        ValidateFormRules("frmPriceValAutoconsumo") And _
        ValidateFormRules("frmPriceValPublic") And _
        ValidateFormRules("frmQuantitativeValMarket") And _
        ValidateFormRules("frmQuantitativeValAutoconsumo") And _
        ValidateFormRules("frmQuantitativeValPublic") And _
        btnStepFour.Enabled = True Then
        imgStepFourStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
        btnStepFive.Enabled = True
        imgStepFiveStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
        'btnStepSix.Enabled = True
    End If
    
End Function

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 1, APPVERSION)
    
    lblApplicationName = APPSHORTNAME
    lblApplicationSubName = APPSUBNAME
    lblApplicationVersion = "Versão: " & APPVERSION
    
    updateForm
    
End Sub
