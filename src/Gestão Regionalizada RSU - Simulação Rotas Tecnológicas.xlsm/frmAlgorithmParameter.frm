VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAlgorithmParameter 
   Caption         =   "UserForm1"
   ClientHeight    =   1992
   ClientLeft      =   96
   ClientTop       =   396
   ClientWidth     =   4104
   OleObjectBlob   =   "frmAlgorithmParameter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAlgorithmParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FormChanged As Boolean

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub textBoxChange(ByRef txtBox, ByVal varName As String)
    Dim errorMsg As String
    If Database.Validate(varName, txtBox.Text, errorMsg) Then
        txtBox.BackColor = ApplicationColors.bgColorValidTextBox
        txtBox.ControlTipText = errorMsg
    Else
        txtBox.BackColor = ApplicationColors.bgColorInvalidTextBox
        txtBox.ControlTipText = errorMsg
    End If
    FormChanged = True
End Sub

Private Sub btnDefault_Click()
    txtTrashThreshold = Database.GetDatabaseValue("TrashThreshold", colDefaultValue)
    txtMaxClusters = Database.GetDatabaseValue("MaxClusters", colDefaultValue)
    txtMaxSubarrays = Database.GetDatabaseValue("MaxSubarrays", colDefaultValue)
    txtArrays = Database.GetDatabaseValue("MaxArrays", colDefaultValue)
End Sub

Private Sub btnPythonExecutable_Click()
    Dim sPython As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "Selecione o executável (.exe) do Python"
        .Filters.Add "Python", "*.exe", 1
        .AllowMultiSelect = False
        If .Show = -1 Then
            sPython = .SelectedItems(1)
        End If
    End With
    
    If sPython <> "" Then
        txtPythonPath.Text = sPython
        FormChanged = True
    End If
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrorHandler
        Call Database.SetDatabaseValue("PythonPath", colUserValue, txtPythonPath.Text)
        Call Database.SetDatabaseValue("TrashThreshold", colUserValue, CDbl(txtTrashThreshold.Text))
        Call Database.SetDatabaseValue("MaxClusters", colUserValue, CDbl(txtMaxClusters.Text))
        Call Database.SetDatabaseValue("MaxSubarrays", colUserValue, CDbl(txtMaxSubarrays.Text))
        Call Database.SetDatabaseValue("MaxArrays", colUserValue, CDbl(txtArrays.Text))
        FormChanged = False
        frmStepTwo.updateForm
        Unload Me
        ThisWorkbook.Save
    Exit Sub
    
ErrorHandler:
    Call MsgBox(MSG_INVALID_DATA, vbCritical, MSG_INVALID_DATA_TITLE)
    
End Sub

Private Sub txtArrays_Change()
    Call textBoxChange(txtArrays, "MaxArrays")
End Sub

Private Sub txtMaxClusters_Change()
    Call textBoxChange(txtMaxClusters, "MaxClusters")
End Sub

Private Sub txtTrashThreshold_Change()
    Call textBoxChange(txtTrashThreshold, "TrashThreshold")
End Sub

Private Sub txtMaxSubarrays_Change()
    Call textBoxChange(txtMaxSubarrays, "MaxSubarrays")
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Parametrizar Algoritmo")
    'Special Configuration
    txtPythonPath.ForeColor = RGB(0, 0, 0)
    txtPythonPath.TextAlign = fmTextAlignLeft
    
    'Read database values
    txtPythonPath = Database.GetDatabaseValue("PythonPath", colUserValue)
    txtTrashThreshold = Database.GetDatabaseValue("TrashThreshold", colUserValue)
    txtMaxClusters = Database.GetDatabaseValue("MaxClusters", colUserValue)
    txtMaxSubarrays = Database.GetDatabaseValue("MaxSubarrays", colUserValue)
    txtArrays = Database.GetDatabaseValue("MaxArrays", colUserValue)
    
    If txtPythonPath = "" Then
        pythonVersion = CreateObject("WScript.Shell").Exec("python --version").StdOut.ReadAll
        If pythonVersion <> "" Then
            strPath = CreateObject("WScript.Shell").Exec("where python").StdOut.ReadAll
            strPath = Replace(strPath, vbCrLf, vbCr)
            strPath = Replace(strPath, vbLf, vbCr)
            splitLineBreaks = Split(strPath, vbCr)
            txtPythonPath = splitLineBreaks(0) 'Left(strPath, Len(strPath) - 2)
        Else
            Call MsgBox("O Python não está instalado na sua máquina, favor instalar para poder executar o algoritmo.", vbCritical, "Python não encontrado")
        End If
    End If
    
    Me.Height = 272
    Me.width = 412
    
    FormChanged = False
End Sub
