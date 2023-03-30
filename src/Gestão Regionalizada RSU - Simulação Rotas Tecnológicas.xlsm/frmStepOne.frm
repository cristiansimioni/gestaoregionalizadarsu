VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepOne 
   Caption         =   "Passo 1"
   ClientHeight    =   6660
   ClientLeft      =   240
   ClientTop       =   936
   ClientWidth     =   8772.001
   OleObjectBlob   =   "frmStepOne.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStepOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim FormChanged As Boolean

Private Sub btnBack_Click()
    If FormChanged Then
        answer = MsgBox(MSG_CHANGED_NOT_SAVED, vbQuestion + vbYesNo + vbDefaultButton2, MSG_CHANGED_NOT_SAVED_TITLE)
        If answer = vbYes Then
          Call btnSave_Click
        Else
            Unload Me
            frmTool.updateForm
        End If
    Else
        Unload Me
        frmTool.updateForm
        ThisWorkbook.Save
    End If
End Sub

Private Sub btnFolder_Click()
    Dim sFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Selecione a pasta onde deseja salvar o projeto"
        If .Show = -1 Then
            sFolder = .SelectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then
        txtPath.Text = sFolder
        imgFolder.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
        FormChanged = True
    End If
End Sub

Private Sub btnHelpStep_Click()
    On Error Resume Next
        ThisWorkbook.FollowHyperlink (Application.ThisWorkbook.Path & "\" & FOLDERMANUAL & "\" & FILEMANUALSTEP1)
    On Error GoTo 0
End Sub

Private Sub btnSave_Click()
    Call Database.SetDatabaseValue("ProjectName", DatabaseColumn.colUserValue, txtProjectName.Text)
    Call Database.SetDatabaseValue("ProjectPathFolder", DatabaseColumn.colUserValue, txtPath.Text)
    Unload Me
    frmTool.updateForm
    ThisWorkbook.Save
End Sub

Private Sub btnSelectCities_Click()
    frmSelectCities.Show
End Sub

Private Sub btnRSUGravimetry_Click()
    frmRSUGravimetry.Show
End Sub

Private Sub btnSimulationData_Click()
    frmSimulationData.Show
End Sub

Private Sub btnStudyCaseStepOne_Click()
    frmStudyCaseStepOne.Show
End Sub


Private Sub txtProjectName_Change()
    FormChanged = True
End Sub

Public Function updateForm()
    imgFolder.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgRSUGravimetry.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgSimulation.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgSelectCities.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgStudyCase.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    
    If ValidateFormRules("frmRSUGravimetry") Then imgRSUGravimetry.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmStudyCaseStepOne") Then imgStudyCase.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmSimulationData") Then imgSimulation.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If Dir(txtPath.Text, vbDirectory) <> "." Then imgFolder.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If readSelectedCities.count >= 2 Then imgSelectCities.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
End Function

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 2, "Passo 1")
    
    'Read database values
    txtProjectName.Text = Database.GetDatabaseValue("ProjectName", colUserValue)
    txtPath.Text = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
    
    'Validation
    Call updateForm
    
    FormChanged = False
    
End Sub
