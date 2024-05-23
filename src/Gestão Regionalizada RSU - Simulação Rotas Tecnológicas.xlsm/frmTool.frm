VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTool 
   ClientHeight    =   8715.001
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   13680
   OleObjectBlob   =   "frmTool.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function StartWindow Lib "user32" Alias "GetWindowLongA" ( _
            ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare PtrSafe Function MoveWindow Lib "user32" Alias "SetWindowLongA" ( _
            ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long


Private Const STYLE_CURRENT As Long = (-16)        '// A new window STYLE

'// Window STYLE
Private Const WS_CX_MINIMIZAR As Long = &H20000 '// Minimize button
Private Const WS_CX_MAXIMIZAR As Long = &H10000 '// Maximize button

'// Window status
Private Const SW_EXIBIR_NORMAL = 1
Private Const SW_EXIBIR_MINIMIZADO = 2
Private Const SW_EXIBIR_MAXIMIZADO = 3

Dim Form_Personalized As Long
Dim STYLE As Long

'Abre o formulário "sobre" que apresenta informações sobre a ferramenta
Private Sub btnAbout_Click()
    frmAbout.Show
End Sub

'Limpa a base de dados para uma nova simulação. Uma confirmação é
'requisitada para o usuário para evitar que a base seja apagada sem querer
Private Sub btnClean_Click()
    Dim answer As Integer
    answer = MsgBox(MSG_CLEAN_DATABASE, vbExclamation + vbYesNo + vbDefaultButton2, MSG_ATTENTION)
    If answer = vbYes Then
        Database.Clean
        Call updateForm
    End If
End Sub

'Abre o manual da ferramenta
Private Sub btnHelp_Click()
    On Error Resume Next
        ThisWorkbook.FollowHyperlink (Application.ThisWorkbook.Path & "\" & FOLDERMANUAL & "\" & FILEMANUAL)
    On Error GoTo 0
End Sub

'Abre o passo cinco da ferramenta
Private Sub btnStepFive_Click()
    frmStepFive.Show
End Sub

'Abre o passo quatro da ferramenta
Private Sub btnStepFour_Click()
    frmStepFour.Show
End Sub

'Abre o passo um da ferramenta
Private Sub btnStepOne_Click()
    frmStepOne.Show
End Sub

'Abre o passo seis da ferramenta
Private Sub btnStepSix_Click()
    frmStepSix.Show
End Sub

'Abre o passo três da ferramenta
Private Sub btnStepThree_Click()
    frmStepThree.Show
End Sub

'Abre o passo dois da ferramenta
Private Sub btnStepTwo_Click()
    frmStepTwo.Show
End Sub

'Atualiza os elementos do formulário conforme o estado atual dos passos preenchidos
Public Function updateForm()
    Dim prjPath As String
    Dim prjName As String
    Dim Fso As Object
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    prjPath = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
    prjName = Database.GetDatabaseValue("ProjectName", colUserValue)
    
    'Atualiza o formulário para o modo padrão
    imgStepOneStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgStepTwoStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgStepThreeStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgStepFourStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgStepFiveStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgStepSixStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONWARNING)
    imgPartnersBottom.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMAGEPARTNERS)
    btnStepOne.Enabled = True
    btnStepTwo.Enabled = False
    btnStepThree.Enabled = False
    btnStepFour.Enabled = False
    btnStepFive.Enabled = False
    btnStepSix.Enabled = False
    
    'Se os formuários de gravimetria, estudo de caso, dados de simulação,
    'a pasta para salvar os arquivos foi selecionada e mais de dois municípios foram
    'selecionados, então o passo um está completo e o passo dois é liberado
    If ValidateFormRules("frmRSUGravimetry") And _
        ValidateFormRules("frmStudyCaseStepOne") And _
        ValidateFormRules("frmSimulationData") And _
        ValidateFormRules("frmStepOne") And _
        readSelectedCities.count >= 2 Then
        
        imgStepOneStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
        btnStepTwo.Enabled = True
        
    End If
    
    If ValidateFormRules("frmGeneralData") And _
        ValidateFormRules("frmAlgorithmParameter") And _
        ValidateFormRules("frmStepTwo") And _
        btnStepTwo.Enabled = True Then
        imgStepTwoStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
        btnStepThree.Enabled = True
    End If
    
    If ValidateFormRules("frmRoute") And _
        ValidateFormRules("frmCapexData") And _
        ValidateFormRules("frmOpexData") And _
        ValidateFormRules("frmTaxes") And _
        ValidateFormRules("frmContract") And _
        ValidateFormRules("frmFinancialAssumptions") And _
        ValidateFormRules("frmUserBase") And _
        btnStepThree.Enabled = True Then
        imgStepThreeStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
        btnStepFour.Enabled = True
    End If
    
    If ValidateFormRules("frmPriceValRevenue") And _
        ValidateFormRules("frmPriceValMarket") And _
        ValidateFormRules("frmPriceValAutoconsumo") And _
        ValidateFormRules("frmPriceValPublic") And _
        ValidateFormRules("frmQuantitativeValMarket") And _
        ValidateFormRules("frmQuantitativeValAutoconsumo") And _
        ValidateFormRules("frmQuantitativeValPublic") And _
        ValidateFormRules("frmStepFour") And _
        btnStepFour.Enabled = True And _
        Fso.FolderExists(prjPath & "\" & prjName) Then
        imgStepFourStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
        btnStepFive.Enabled = True
        imgStepFiveStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
        btnStepSix.Enabled = True
        imgStepSixStatus.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERICONS & "\" & ICONCHECK)
    End If
    
End Function

Private Sub UserForm_Initialize()
    'Ajusta a aparência do formulário
    Call modForm.applyLookAndFeel(Me, 1, APPVERSION)
    
    lblApplicationName = APPSHORTNAME
    lblApplicationSubName = APPSUBNAME
    lblApplicationVersion = "Versão: " & APPVERSION
    
    'Form_Personalized = FindWindowA(vbNullString, Me.Caption)
    'STYLE = StartWindow(Form_Personalized, STYLE_CURRENT)
    'STYLE = STYLE Or WS_CX_MINIMIZAR      '// Minimize button
    'STYLE = STYLE Or WS_CX_MAXIMIZAR      '// Maximize button
    'MoveWindow Form_Personalized, STYLE_CURRENT, (STYLE)
    
    Me.Height = 463
    Me.width = 694
    
    updateForm
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ThisWorkbook.Save
End Sub
