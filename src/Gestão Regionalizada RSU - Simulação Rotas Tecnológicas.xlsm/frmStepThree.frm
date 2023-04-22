VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepThree 
   Caption         =   "UserForm1"
   ClientHeight    =   5904
   ClientLeft      =   195
   ClientTop       =   765
   ClientWidth     =   6135
   OleObjectBlob   =   "frmStepThree.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStepThree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnBack_Click()
    frmTool.updateForm
    Unload Me
End Sub

Private Sub btnCapexData_Click()
    frmCapexData.Show
End Sub

Private Sub btnContract_Click()
    frmContract.Show
End Sub

Private Sub btnFinancialAssumptions_Click()
    frmFinancialAssumptions.Show
End Sub

Private Sub btnHelpStep_Click()
    On Error Resume Next
        ThisWorkbook.FollowHyperlink (Application.ThisWorkbook.Path & "\" & FOLDERMANUAL & "\" & FILEMANUALSTEP3)
    On Error GoTo 0
End Sub

Private Sub btnOpexData_Click()
    frmOpexData.Show
End Sub

Private Sub btnRouteDefinition_Click()
    frmRoute.Show
End Sub

Private Sub btnTaxes_Click()
    frmTaxes.Show
End Sub

Private Sub btnUserBase_Click()
    frmUserBase.Show
End Sub

Public Function updateForm()
    imgRouteDefinition.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgCapexData.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgOpexData.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgTaxes.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgContract.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgFinancialAssumptions.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgUserBase.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    
    If ValidateFormRules("frmRoute") Then imgRouteDefinition.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmCapexData") Then imgCapexData.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmOpexData") Then imgOpexData.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmTaxes") Then imgTaxes.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmContract") Then imgContract.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmFinancialAssumptions") Then imgFinancialAssumptions.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmUserBase") Then imgUserBase.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
End Function

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 2, "Passo 3")
    
    Call updateForm
    
    Me.Height = 397
    Me.width = 490
End Sub
