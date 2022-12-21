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

Private Sub txtArray_Change()
    Call Database.SetDatabaseValue("ArrayText", colUserValue, txtArray.Text)
End Sub

Private Sub txtConclusion_Change()
    Call Database.SetDatabaseValue("ConclusionText", colUserValue, txtConclusion.Text)
End Sub

Private Sub txtIntroduction_Change()
    Call Database.SetDatabaseValue("IntroductionText", colUserValue, txtIntroduction.Text)
End Sub

Private Sub txtMarket_Change()
    Call Database.SetDatabaseValue("MarketText", colUserValue, txtMarket.Text)
End Sub

Private Sub txtObjectives_Change()
    Call Database.SetDatabaseValue("ObjectivesText", colUserValue, txtObjectives.Text)
End Sub

Private Sub txtRoutes_Change()
    Call Database.SetDatabaseValue("RoutesText", colUserValue, txtRoutes.Text)
End Sub

Private Sub txtValuation_Change()
    Call Database.SetDatabaseValue("ValuationText", colUserValue, txtValuation.Text)
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 2, "Passo 6")
    
    'Read database values
    txtIntroduction.Text = Database.GetDatabaseValue("IntroductionText", colUserValue)
    txtObjectives.Text = Database.GetDatabaseValue("ObjectivesText", colUserValue)
    txtArray.Text = Database.GetDatabaseValue("ArrayText", colUserValue)
    txtRoutes.Text = Database.GetDatabaseValue("RoutesText", colUserValue)
    txtMarket.Text = Database.GetDatabaseValue("MarketText", colUserValue)
    txtValuation.Text = Database.GetDatabaseValue("ValuationText", colUserValue)
    txtConclusion.Text = Database.GetDatabaseValue("ConclusionText", colUserValue)

End Sub
