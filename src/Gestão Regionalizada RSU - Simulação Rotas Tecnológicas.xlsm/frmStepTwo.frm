VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepTwo 
   Caption         =   "Passo 2"
   ClientHeight    =   4965
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   6930
   OleObjectBlob   =   "frmStepTwo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStepTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGeneralData_Click()
    frmGeneralData.Show
End Sub

Private Sub btnRunAlgorithm_Click()
    'Calculate cities distance
    Call modCity.calculateDistances
    
    'Save cities to csv
    Call Util.saveAsCSV("CIRSOP", "C:\Users\cristiansimioni\Downloads")

    'Save distance to csv
    Call Util.saveAsCSV("CIRSOP", "C:\Users\cristiansimioni\Downloads")
    
    'Run the algorithm
    Util.RunPythonScript
    
    'Load the result into the workbook
    
End Sub

Private Sub CommandButton4_Click()
    frmEditCities.Show
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = APPNAME & " - Passo 2"
    Me.BackColor = ApplicationColors.bgColorLevel2
    
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.btColorLevel2
         End If
    Next Ctrl
    
End Sub
