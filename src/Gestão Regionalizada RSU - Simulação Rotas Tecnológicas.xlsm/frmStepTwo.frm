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
    frmTool.updateForm
    Unload Me
End Sub

Private Sub btnGeneralData_Click()
    frmGeneralData.Show
End Sub

Private Sub btnAlgorithParameter_Click()
    frmAlgorithmParameter.Show
End Sub

Public Function updateForm()
    imgGeneralData.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgUTVR.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgParameterAlgorithm.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgAlgorithm.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgArrays.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    
    btnRunAlgorithm.Enabled = False
    btnSelectArrays.Enabled = False
    
    If ValidateFormRules("frmGeneralData") Then imgGeneralData.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmAlgorithmParameter") Then
        imgParameterAlgorithm.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    End If
    If Database.GetDatabaseValue("AlgorithmStatus", colUserValue) = "Sim" Then
        imgAlgorithm.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
        btnSelectArrays.Enabled = True
    End If
    If validateDatabaseCities Then
        imgUTVR.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
        Call Database.SetDatabaseValue("CityStatus", colUserValue, "Sim")
    Else
        Call Database.SetDatabaseValue("CityStatus", colUserValue, "")
    End If
    If modArray.countSelectedArrays = 4 Then
        imgArrays.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
        Call Database.SetDatabaseValue("ArrayStatus", colUserValue, "Sim")
    Else
        Call Database.SetDatabaseValue("ArrayStatus", colUserValue, "")
    End If
    
    If ValidateFormRules("frmAlgorithmParameter") And validateDatabaseCities Then
        btnRunAlgorithm.Enabled = True
    End If
End Function

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
    If Util.RunPythonScript(algPath, prjName) Then
        'Load the result into the workbook
        Call Util.CSVImport(algPath, prjName)
    End If
    
    btnRunAlgorithm.Enabled = True
    
    Call updateForm
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
    
    Call updateForm
End Sub
