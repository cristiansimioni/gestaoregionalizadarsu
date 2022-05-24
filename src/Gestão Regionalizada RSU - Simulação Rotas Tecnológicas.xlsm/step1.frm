VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} step1 
   Caption         =   "Passo 1"
   ClientHeight    =   10365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12495
   OleObjectBlob   =   "step1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "step1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backButton_Click()
    Unload Me
End Sub

Private Sub btnFolder_Click()
    Dim sFolder As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Selecione a pasta onde deseja salvar o projeto"
        If .Show = -1 Then
            sFolder = .SelectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then
        txtPath.Text = sFolder
    End If
    
End Sub

Private Sub btnSave_Click()
    Database.setProjectName (txtProjectName.Text)
    Database.setProjectPathFolder (txtPath.Text)
End Sub

Private Sub citiesButton_Click()
    cities.Show
End Sub

Private Sub rsuGravimetryButton_Click()
    rsuGravimetry.Show
End Sub

Private Sub simulationDataButton_Click()
    simulationData.Show
End Sub

Private Sub studyCaseStepOneButton_Click()
    studyCaseStepOne.Show
End Sub


Private Sub UserForm_Initialize()
    Step1.Caption = Util.xApplicationName & " - Passo 1"
    Step1.BackColor = xColorLevel2
    txtProjectName.Text = Database.getProjectName()
    txtPath.Text = Database.getProjectPathFolder
End Sub
