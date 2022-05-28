VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepOne 
   Caption         =   "Passo 1"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6915
   OleObjectBlob   =   "frmStepOne.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStepOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
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
    Call Database.SetDatabaseValue("ProjectName", DatabaseColumn.colUserValue, txtProjectName.Text)
    Call Database.SetDatabaseValue("ProjectPathFolder", DatabaseColumn.colUserValue, txtPath.Text)
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


Private Sub UserForm_Initialize()
    Me.Caption = APPNAME & " - Passo 1"
    Me.BackColor = ApplicationColors.bgColorLevel2
    txtProjectName.Text = Database.GetDatabaseValue("ProjectName", colUserValue)
    txtPath.Text = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
    
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.btColorLevel2
         End If
    Next Ctrl
    
End Sub
