VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tool 
   Caption         =   "GEF Biogás Brasil"
   ClientHeight    =   9240.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11640
   OleObjectBlob   =   "tool.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "tool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btnClean_Click()
    Database.cleanDB
End Sub

Private Sub StepOneButton_Click()
    Step1.Show
End Sub

Private Sub StepTwoButton_Click()
    step2.Show
End Sub

Private Sub UserForm_Activate()
    lblApplicationName.Caption = Util.xApplicationName
    lblApplicationVersion.Caption = Util.xApplicationVersion
    lblApplicationLastUpdate.Caption = Util.xApplicationLastUpdate
    
    StepThreeButton.Enabled = False
    StepFourButton.Enabled = False
    StepFiveButton.Enabled = False
    StepSixButton.Enabled = False

End Sub


Private Sub UserForm_Initialize()
    Util.initializeDefinitions
    Database.initializeDB
    tool.Caption = "GEF Biogás Brasil - " & Util.xApplicationName & " - " & Util.xApplicationVersion
    tool.BackColor = Util.xColorLevel1
End Sub
