VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepSix 
   Caption         =   "UserForm1"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "frmStepSix.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStepSix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnHelpStep_Click()
    On Error Resume Next
        ActiveWorkbook.FollowHyperlink (Application.ThisWorkbook.Path & "\" & FOLDERMANUAL & "\" & FILEMANUALSTEP6)
    On Error GoTo 0
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnReport_Click()
    Dim prjPath As String
    Dim prjName As String
        
    prjPath = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
    prjName = Database.GetDatabaseValue("ProjectName", colUserValue)
    prjPath = Util.FolderCreate(prjPath, prjName)
    Dim reportPath As String
    reportPath = Util.FolderCreate(prjPath, FOLDERREPORT)
    
    On Error GoTo ErrorHandler
        Dim wksReportData As Worksheet
        Set wksReportData = Util.GetReportWorksheet
        
        wksReportData.Cells(12, 1).value = txtIntroduction.Text
        wksReportData.Cells(23, 1).value = txtObjectives.Text
        wksReportData.Cells(43, 1).value = txtArray.Text
        wksReportData.Cells(63, 1).value = txtRoutes.Text
        wksReportData.Cells(83, 1).value = txtMarket.Text
        wksReportData.Cells(103, 1).value = txtValuation.Text
        wksReportData.Cells(123, 1).value = txtConclusion.Text
        
        wksReportData.range("A1:K141").ExportAsFixedFormat Type:=xlTypePDF, Quality:=xlQualityStandard, OpenAfterPublish:=True, filename:=reportPath & "\Relatório.pdf"
        'Call MsgBox("Relatório gerado com sucesso!", vbOKOnly, "Relatório")
    Exit Sub
    
ErrorHandler:
    Call MsgBox(MSG_INVALID_DATA, vbExclamation, MSG_INVALID_DATA_TITLE)
    
End Sub


Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 2, "Passo 6")

End Sub
