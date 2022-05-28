VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTool 
   Caption         =   "GEF Biogás Brasil"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
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


Private Sub StepOneButton_Click()
    frmStepOne.Show
End Sub

Private Sub StepTwoButton_Click()
    step2.Show
End Sub

Private Sub UserForm_Activate()
    lblApplicationName.Caption = APPNAME
    'lblApplicationVersion.Caption = APPVERSION
    'lblApplicationLastUpdate.Caption = APPLASTUPDATED
    
    StepThreeButton.Enabled = False
    StepFourButton.Enabled = False
    StepFiveButton.Enabled = False
    StepSixButton.Enabled = False

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
