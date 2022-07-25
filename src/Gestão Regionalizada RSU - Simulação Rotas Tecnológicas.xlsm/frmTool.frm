VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTool 
   ClientHeight    =   8595.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15240
   OleObjectBlob   =   "frmTool.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
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
    End If
End Sub

Private Sub btnHelp_Click()
    ActiveWorkbook.FollowHyperlink (Application.ThisWorkbook.Path & "\assets\manual\Manual da Ferramenta.pdf")
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
    
    If ValidateFormRules("frmRSUGravimetry") And _
        ValidateFormRules("frmStudyCaseStepOne") And _
        ValidateFormRules("frmSimulationData") And _
        ValidateFormRules("frmStepOne") And _
        readSelectedCities.Count >= 2 Then
        imgStepOneStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
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
