VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepTwo 
   Caption         =   "Passo 2"
   ClientHeight    =   4800
   ClientLeft      =   360
   ClientTop       =   1395
   ClientWidth     =   6945
   OleObjectBlob   =   "frmStepTwo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStepTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnGeneralData_Click()
    frmGeneralData.Show
End Sub

Private Sub btnRunAlgorithm_Click()
    btnRunAlgorithm.Enabled = False
    'Calculate cities distance
    Call modCity.calculateDistances
    
    'Create project folder
    Dim prjPath As String
    Dim prjName As String
    prjPath = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
    prjName = Database.GetDatabaseValue("ProjectName", colUserValue)
    prjPath = Util.FolderCreate(prjPath, prjName)
    
    'Create algorithm folder
    Dim algPath As String
    algPath = Util.FolderCreate(prjPath, FOLDERALGORITHM)
    
    'Save cities to csv
    Call Util.saveAsCSV(prjName, algPath, "city")

    'Save distance to csv
    Call Util.saveAsCSV(prjName, algPath, "distance")
    
    'Run the algorithm
    Call Util.RunPythonScript(algPath, prjName)
    
    'Load the result into the workbook
    Call Util.CSVImport(algPath, prjName)
    
    btnRunAlgorithm.Enabled = True
    
End Sub

Private Sub btnSelectArrays_Click()
    frmSelectArrays.Show
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
