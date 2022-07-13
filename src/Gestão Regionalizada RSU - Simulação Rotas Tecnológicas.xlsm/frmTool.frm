VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTool 
   Caption         =   "GEF Biogás Brasil"
   ClientHeight    =   8655.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12720
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
Private Sub btnClean_Click()
    Database.CleanDatabase
End Sub

Private Sub btnHelp_Click()
    ActiveWorkbook.FollowHyperlink (Application.ActiveWorkbook.Path & "\assets\manual\Manual da Ferramenta.pdf")
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

Private Sub UserForm_Initialize()
    'Form Appearance
    Me.Caption = "GEF Biogás Brasil - " & APPNAME & " - " & APPVERSION
    Me.BackColor = ApplicationColors.frmBgColorLevel1
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.bgColorLevel1
            Ctrl.ForeColor = ApplicationColors.fgColorLevel1
         End If
    Next Ctrl
    
    lblApplicationName.Caption = APPNAME
    
    
    imgStepOneStatus.Picture = LoadPicture(Application.ActiveWorkbook.Path & "\" & FOLDERASSETS & "\" & FOLDERICONS & "\check-icon.jpg")
    imgStepSixStatus.Picture = LoadPicture(Application.ActiveWorkbook.Path & "\" & FOLDERASSETS & "\" & FOLDERICONS & "\check-icon.jpg")
    
    
End Sub
