VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTool 
   Caption         =   "GEF Biogás Brasil"
   ClientHeight    =   8655.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12720
   OleObjectBlob   =   "frmTool.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "frmTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClean_Click()
    Database.CleanDatabase
End Sub


Private Sub btnHelp_Click()
    ActiveWorkbook.FollowHyperlink (Application.ActiveWorkbook.Path & "\assets\manual\Manual da Ferramenta.pdf")
End Sub

Private Sub StepFiveButton_Click()
    frmStepFive.Show
End Sub

Private Sub StepFourButton_Click()
    frmStepFour.Show
End Sub

Private Sub StepOneButton_Click()
    frmStepOne.Show
End Sub

Private Sub StepSixButton_Click()
    frmStepSix.Show
End Sub

Private Sub StepThreeButton_Click()
    frmStepThree.Show
End Sub

Private Sub StepTwoButton_Click()
    frmStepTwo.Show
End Sub

Private Sub UserForm_Activate()
    lblApplicationName.Caption = APPNAME
    'lblApplicationVersion.Caption = APPVERSION
    'lblApplicationLastUpdate.Caption = APPLASTUPDATED
End Sub


Private Sub UserForm_Initialize()
    Me.Caption = "GEF Biogás Brasil - " & APPNAME & " - " & APPVERSION
    Me.BackColor = ApplicationColors.bgColorLevel1
    Me.StepOneButton.BackColor = ApplicationColors.btColorLevel1
    Me.StepTwoButton.BackColor = ApplicationColors.btColorLevel1
    Me.StepThreeButton.BackColor = ApplicationColors.btColorLevel1
    Me.StepFourButton.BackColor = ApplicationColors.btColorLevel1
    Me.StepFiveButton.BackColor = ApplicationColors.btColorLevel1
    Me.StepSixButton.BackColor = ApplicationColors.btColorLevel1
End Sub
