VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAlgorithmParameter 
   Caption         =   "UserForm1"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10830
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

Private Sub btnSave_Click()
    Call Database.SetDatabaseValue("PythonPath", colUserValue, txtPythonPath.Text)
    Call Database.SetDatabaseValue("TrashThreshold", colUserValue, CDbl(txtTrashThreshold.Text))
    Call Database.SetDatabaseValue("MaxClusters", colUserValue, CDbl(txtMaxClusters.Text))
    FormChanged = False
    Unload Me
End Sub

Private Sub txtMaxClusters_Change()
    Call textBoxChange(txtMaxClusters, "MaxClusters")
End Sub

Private Sub txtTrashThreshold_Change()
    Call textBoxChange(txtTrashThreshold, "TrashThreshold")
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Parametrizar Algoritmo", True)
    
    
    'Read database values
    txtPythonPath = Database.GetDatabaseValue("PythonPath", colUserValue)
    txtTrashThreshold = Database.GetDatabaseValue("TrashThreshold", colUserValue)
    txtMaxClusters = Database.GetDatabaseValue("MaxClusters", colUserValue)
    
    FormChanged = False
    
End Sub
