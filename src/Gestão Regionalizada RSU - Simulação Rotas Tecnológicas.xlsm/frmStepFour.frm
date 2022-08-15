VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepFour 
   Caption         =   "UserForm1"
   ClientHeight    =   8955.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   OleObjectBlob   =   "frmStepFour.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStepFour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
    frmTool.updateForm
    Unload Me
End Sub

Private Sub btnPriceVal_Click()
    frmPriceVal.Show
End Sub

Private Sub btnQuantitativeVal_Click()
    frmQuantitativeVal.Show
End Sub

Private Sub btnExecuteSimulation_Click()
    frmProgressBarSimulation.Show
End Sub

Private Sub btnPriceValAutoconsumo_Click()
    frmPriceValAutoconsumo.Show
End Sub

Private Sub btnPriceValMarket_Click()
    frmPriceValMarket.Show
End Sub

Private Sub btnPriceValPublic_Click()
    frmPriceValPublic.Show
End Sub

Private Sub btnPriceValRevenue_Click()
    frmPriceValRevenue.Show
End Sub

Private Sub btnQuantitativeValAutoconsumo_Click()
    frmQuantitativeValAutoconsumo.Show
End Sub

Private Sub btnQuantitativeValMarket_Click()
    frmQuantitativeValMarket.Show
End Sub

Private Sub btnQuantitativeValPublic_Click()
    frmQuantitativeValPublic.Show
End Sub

Public Function updateForm()
    imgPriceValRevenue.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgPriceValMarket.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgPriceValAutoconsumo.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgPriceValPublic.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgQuantitativeValMarket.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgQuantitativeValAutoconsumo.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgQuantitativeValPublic.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgExecuteSimulation.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    
    If ValidateFormRules("frmPriceValRevenue") Then imgPriceValRevenue.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmPriceValMarket") Then imgPriceValMarket.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmPriceValAutoconsumo") Then imgPriceValAutoconsumo.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmPriceValPublic") Then imgPriceValPublic.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmQuantitativeValMarket") Then imgQuantitativeValMarket.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmQuantitativeValAutoconsumo") Then imgQuantitativeValAutoconsumo.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If ValidateFormRules("frmQuantitativeValPublic") Then imgQuantitativeValPublic.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    If Database.GetDatabaseValue("SimulationStatus", colUserValue) = "Sim" Then
        imgExecuteSimulation.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    End If
End Function

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 2, "Passo 4")
    
    Call updateForm
End Sub
