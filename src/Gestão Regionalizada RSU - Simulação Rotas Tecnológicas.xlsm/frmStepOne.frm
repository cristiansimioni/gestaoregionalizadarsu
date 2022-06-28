VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepOne 
   Caption         =   "Passo 1"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7425
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
        answer = MsgBox("Você realizou alterações, gostaria de salvar?", vbQuestion + vbYesNo + vbDefaultButton2, "Salvar Alterações")
        If answer = vbYes Then
          Call btnSave_Click
        Else
          Unload Me
        End If
    Else
        Unload Me
    End If
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
        FormChanged = True
    End If
End Sub

Private Sub btnSave_Click()
    Call Database.SetDatabaseValue("ProjectName", DatabaseColumn.colUserValue, txtProjectName.Text)
    Call Database.SetDatabaseValue("ProjectPathFolder", DatabaseColumn.colUserValue, txtPath.Text)
    Unload Me
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

Private Sub UserForm_Initialize()
    'Form Appearance
    Me.Caption = APPNAME & " - Passo 1"
    Me.BackColor = ApplicationColors.bgColorLevel2
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.btColorLevel2
            Ctrl.ForeColor = ApplicationColors.fgColorLevel2
         End If
    Next Ctrl
    
    'Read database values
    txtProjectName.Text = Database.GetDatabaseValue("ProjectName", colUserValue)
    txtPath.Text = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
    
    FormChanged = False
End Sub
