VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepTwo 
   Caption         =   "Passo 2"
   ClientHeight    =   5835
   ClientLeft      =   480
   ClientTop       =   1860
   ClientWidth     =   9960.001
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

Private Sub btnAlgorithParameter_Click()
    frmAlgorithmParameter.Show
End Sub

Private Sub btnRunAlgorithm_Click()
    btnRunAlgorithm.Enabled = False
    'Calculate cities distance
    'Call modCity.calculateDistances
    
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
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 2, "Passo 2")
    
End Sub
