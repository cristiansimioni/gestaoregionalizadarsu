VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepFive 
   Caption         =   "UserForm1"
   ClientHeight    =   11715
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   18360
   OleObjectBlob   =   "frmStepFive.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStepFive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arrays As Collection

Public Enum ArrayConsolidado
    colID = 1
    colRota = 2
    colTarifaBruta = 4
    colTarifaLiquida = 5
    colEficienciaValorizacao = 6
    colCapex = 7
    colOpex = 8
    colEntradaPlanta = 9
    colReciclaveis = 10
    colCDR = 11
    colRejeitos = 12
    colComposto = 13
    colPerdaMassa = 14
    colBiometano = 15
    colEnergiaEletrica = 16
End Enum

Private Sub btnBack_Click()
    frmTool.updateForm
    Unload Me
End Sub



Private Sub MultiPage1_Change()
'Purpose: mark current page caption by a checkmark
    With Me.MultiPage1
        Dim pg As MSForms.Page
    'a) de-mark old caption
        Set pg = oldPage(Me.MultiPage1)
        pg.Caption = Replace(pg.Caption, ChkMark, vbNullString)
    'b) mark new caption & remember latest multipage value
        Set pg = .Pages(.value)
        pg.Caption = ChkMark & pg.Caption
        .Tag = .value                         ' << remember latest page index
    End With
End Sub

Function oldPage(mp As MSForms.MultiPage) As MSForms.Page
'Purpose: return currently marked page in given multipage
    With mp
        Set oldPage = .Pages(Val(.Tag))
    End With
End Function

Function ChkMark() As String
'Purpose: return ballot box with check + blank space
    ChkMark = ChrW(&H2611) & ChrW(&HA0)  ' ballot box with check + blank
End Function

Private Sub btnFiles_Click()
    Dim prjPath As String
    Dim prjName As String
    prjPath = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
    prjName = Database.GetDatabaseValue("ProjectName", colUserValue)
    prjPath = Util.FolderCreate(prjPath, prjName)
    ThisWorkbook.FollowHyperlink prjPath
End Sub

Private Sub btnHelpStep_Click()
    On Error Resume Next
        ThisWorkbook.FollowHyperlink (Application.ThisWorkbook.Path & "\" & FOLDERMANUAL & "\" & FILEMANUALSTEP5)
    On Error GoTo 0
End Sub

Private Sub cbxArrayRoute_Change()
    cbxSubArrayRoute.Clear
    
    For Each a In arrays
        If a.vSelected Then
            If a.vCode = cbxArrayRoute.value Then
                For Each S In a.vSubArray
                    cbxSubArrayRoute.AddItem S.vCode
                Next S
                cbxSubArrayRoute.AddItem "Consolidado"
            End If
        End If
    Next
    
    cbxRoute.ListIndex = -1
    Call enableDisableRouteLabels(False, "")
    Call enableDisableConsolidado(False)
    
End Sub

Private Sub cbxArray_Change()
    cbxSubArray.Clear
    
    For Each a In arrays
        If a.vSelected Then
            If a.vCode = cbxArray.value Then
                For Each S In a.vSubArray
                    cbxSubArray.AddItem S.vCode
                Next S
            End If
        End If
    Next
    
    If cbxMarket.value <> "" And cbxArray.value <> "" And cbxSubArray.value <> "" Then
    
        Call PlotGraph
    
    End If
    
End Sub

Private Sub cbxArraySelected_Change()

    'Clear
    t = 1
    While t <= 3
        Me.Controls("txtSubArray" & t).value = ""
        Me.Controls("txtSubArrayLandfill" & t).value = ""
        Me.Controls("txtSubArrayExistentLandfill" & t).value = ""
        Me.Controls("txtSubArrayUTVR" & t).value = ""
        Me.Controls("txtSubArrayTotal" & t).value = ""
        Me.Controls("txtSubArrayTrash" & t).value = ""
        Me.Controls("txtSubArrayTechnology" & t).value = ""
        Me.Controls("txtSubArrayInbound" & t).value = ""
        Me.Controls("txtSubArrayOutbound" & t).value = ""
        Me.Controls("txtSubArrayOutboundExistent" & t).value = ""
        t = t + 1
    Wend
    
    For Each a In arrays
        If a.vSelected Then
            If a.vCode = cbxArraySelected.value Then
                txtArrayTotal.Text = a.vTotal
                txtArrayTrash.Text = a.vTrash
                txtArrayTechnology.Text = a.vTechnology
                txtArrayInbound.Text = a.vInbound
                txtArrayOutbound.Text = a.vOutbound
                txtArrayOutboundExistent.Text = a.vOutboundExistentLandfill
                
                t = 1
                For Each S In a.vSubArray
                    Me.Controls("txtSubArray" & t).value = S.vArrayRaw
                    Me.Controls("txtSubArrayLandfill" & t).value = S.vLandfill
                    Me.Controls("txtSubArrayExistentLandfill" & t).value = S.vExistentLandfill
                    Me.Controls("txtSubArrayUTVR" & t).value = S.vUTVR
                    Me.Controls("txtSubArrayTotal" & t).value = S.vTotal
                    Me.Controls("txtSubArrayTrash" & t).value = S.vTrash
                    Me.Controls("txtSubArrayTechnology" & t).value = S.vTechnology
                    Me.Controls("txtSubArrayInbound" & t).value = S.vInbound
                    Me.Controls("txtSubArrayOutbound" & t).value = S.vOutbound
                    Me.Controls("txtSubArrayOutboundExistent" & t).value = S.vOutboundExistentLandfill
                    t = t + 1
                Next S
            End If
        End If
    Next
End Sub

Private Sub cbxCharts_Change()
    currentChart = cbxCharts
    Dim chartPath As String
    Dim prjPath As String
    Dim prjName As String
    prjPath = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
    prjName = Database.GetDatabaseValue("ProjectName", colUserValue)
    prjPath = Util.FolderCreate(prjPath, prjName)
    chartPath = Util.FolderCreate(prjPath, FOLDERCHART)
    For Each c In ThisWorkbook.Sheets("Dashboard").ChartObjects
        If c.Chart.ChartTitle.Text = currentChart Then
            Fname = chartPath & "\" & c.Chart.ChartTitle.Text & ".jpg"
            Me.Image1.Picture = LoadPicture(Fname)
        End If
    Next c
    
    txtChartDescription.Text = GetDatabaseValue(currentChart, colDefaultValue)
    txtChartDescription.Visible = True
End Sub

Private Sub ChangeRoute()
    Dim wksChartData As Worksheet
    Set wksChartData = Util.GetChartDataWorksheet
    
    wksChartData.Cells(39, 4).value = cbxMarketRoute.value
    wksChartData.Cells(39, 5).value = cbxArrayRoute.value
    wksChartData.Cells(39, 6).value = cbxSubArrayRoute.value
    
    Dim lineData As Integer
    
    If cbxSubArrayRoute.value = "Consolidado" Then
        lineData = 50
    Else
        If cbxRoute.value = "RT1-A" Then
            imgRoute.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMAGESCREENROUTEONEA)
            lineData = 43
        ElseIf cbxRoute.value = "RT1-B" Then
            imgRoute.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMAGESCREENROUTEONEB)
            lineData = 44
        ElseIf cbxRoute.value = "RT1-C" Then
            imgRoute.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMAGESCREENROUTEONEC)
            lineData = 45
        ElseIf cbxRoute.value = "RT2" Then
            imgRoute.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMAGESCREENROUTETWO)
            lineData = 46
        ElseIf cbxRoute.value = "RT3" Then
            imgRoute.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMAGESCREENROUTETHREE)
            lineData = 47
        ElseIf cbxRoute.value = "RT4" Then
            imgRoute.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMAGESCREENROUTEFOUR)
            lineData = 48
        ElseIf cbxRoute.value = "RT5" Then
            imgRoute.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMAGESCREENROUTEFIVE)
            lineData = 49
        End If
    End If
    
    If wksChartData.Cells(50, 2).value = cbxRoute.value And wksChartData.Cells(50, 2).value <> "" Then
        lblSelectedRoute.Visible = True
    Else
        lblSelectedRoute.Visible = False
    End If
    
    If lineData <> 0 Then
        capexRouteData.Caption = Format(wksChartData.Cells(lineData, 4).value, "#.000")
        opexRouteData.Caption = Format(wksChartData.Cells(lineData, 5).value, "#.000")
        inputRouteData.Caption = Format(wksChartData.Cells(lineData, 6).value, "#.000")
        reciclableRouteData.Caption = Format(wksChartData.Cells(lineData, 7).value, "#.000")
        cdrRouteData.Caption = Format(wksChartData.Cells(lineData, 8).value, "#.000")
        landfillDangerRouteData.Caption = Format(wksChartData.Cells(lineData, 11).value, "#.000")
        landfillRouteData.Caption = Format(wksChartData.Cells(lineData, 9).value, "#.000")
        organicCompoundRouteData.Caption = Format(wksChartData.Cells(lineData, 10).value, "#.000")
        lossWeightRouteData.Caption = Format(wksChartData.Cells(lineData, 12).value, "#.000")
        finalUsageRouteData.Caption = Format(wksChartData.Cells(lineData, 13).value, "#,##0.000")
        finalUsage2RouteData.Caption = Format(wksChartData.Cells(lineData, 14).value, "#,##0.000")
        biogasRouteData.Caption = Format(wksChartData.Cells(lineData, 15).value, "#,##0.000")
    End If
    
End Sub

Private Sub PlotGraph()
    Dim prjPath As String
    Dim prjName As String
        
    prjPath = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
    prjName = Database.GetDatabaseValue("ProjectName", colUserValue)
    prjPath = Util.FolderCreate(prjPath, prjName)
        
    'Create base market folder
    Dim chartPath As String
    chartPath = Util.FolderCreate(prjPath, FOLDERCHART)
    
    Dim wksChartData As Worksheet
    Set wksChartData = Util.GetChartDataWorksheet
        
    wksChartData.Cells(27, 4).value = cbxMarket.value
    wksChartData.Cells(27, 5).value = cbxArray.value
    wksChartData.Cells(27, 6).value = cbxSubArray.value
        
    For Each c In ThisWorkbook.Sheets("Dashboard").ChartObjects
        If c.name = "Avaliação" Then
            c.Activate
            c.Chart.ChartTitle.Text = "Avaliação de Custos para o Município de Tratamento de RSU" & " - " & cbxMarket.value & cbxSubArray.value
            Fname = chartPath & "\" & c.Chart.ChartTitle.Text & ".jpg"
            c.Chart.Export filename:=Fname, FilterName:="jpg"
            Me.Image2.Picture = LoadPicture(Fname)
        End If
    Next c
End Sub

Private Sub cbxMarket_Change()
    If cbxMarket.value <> "" And cbxArray.value <> "" And cbxSubArray.value <> "" Then
        Call PlotGraph
    End If
End Sub

Private Sub cbxMarketRoute_Change()
    If cbxMarketRoute.value <> "" And cbxArrayRoute.value <> "" And cbxSubArrayRoute.value <> "" And cbxRoute.value <> "" Then
        Call ChangeRoute
    ElseIf cbxMarketRoute.value <> "" And cbxArrayRoute.value <> "" And cbxSubArrayRoute.value = "Consolidado" Then
        Call ChangeRoute
        Call updateConsolidadoValues
    End If
End Sub

Private Sub cbxMarketValuation_Change()
    If cbxMarketValuation.value <> "" Then
        Dim wksBridgeData As Worksheet
        Set wksBridgeData = Util.GetBridgeDataWorksheet
        wksBridgeData.Cells(2, 1).value = cbxMarketValuation.value
        
        Dim prjPath As String
        Dim prjName As String
        
        prjPath = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
        prjName = Database.GetDatabaseValue("ProjectName", colUserValue)
        prjPath = Util.FolderCreate(prjPath, prjName)
        Dim chartPath As String
        chartPath = Util.FolderCreate(prjPath, FOLDERCHART)
        For Each c In ThisWorkbook.Sheets("Bridges").ChartObjects
            c.Activate
            Fname = chartPath & "\" & c.Chart.ChartTitle.Text & ".bmp"
            c.Chart.Export filename:=Fname, FilterName:="bmp"
        Next c
        
        Me.imgEffort1.Picture = LoadPicture(chartPath & "\" & "Esforço - " & cbxMarketValuation.value & lblArray1.Caption & ".bmp")
        Me.imgEffort2.Picture = LoadPicture(chartPath & "\" & "Esforço - " & cbxMarketValuation.value & lblArray2.Caption & ".bmp")
        Me.imgEffort3.Picture = LoadPicture(chartPath & "\" & "Esforço - " & cbxMarketValuation.value & lblArray3.Caption & ".bmp")
        Me.imgEffort4.Picture = LoadPicture(chartPath & "\" & "Esforço - " & cbxMarketValuation.value & lblArray4.Caption & ".bmp")
        Me.imgIndirect1.Picture = LoadPicture(chartPath & "\" & "Ganhos Indiretos - " & cbxMarketValuation.value & lblArray1.Caption & ".bmp")
        Me.imgIndirect2.Picture = LoadPicture(chartPath & "\" & "Ganhos Indiretos - " & cbxMarketValuation.value & lblArray2.Caption & ".bmp")
        Me.imgIndirect3.Picture = LoadPicture(chartPath & "\" & "Ganhos Indiretos - " & cbxMarketValuation.value & lblArray3.Caption & ".bmp")
        Me.imgIndirect4.Picture = LoadPicture(chartPath & "\" & "Ganhos Indiretos - " & cbxMarketValuation.value & lblArray4.Caption & ".bmp")
        Me.imgPublic1.Picture = LoadPicture(chartPath & "\" & "Desoneração Gestão Pública - " & cbxMarketValuation.value & lblArray1.Caption & ".bmp")
        Me.imgPublic2.Picture = LoadPicture(chartPath & "\" & "Desoneração Gestão Pública - " & cbxMarketValuation.value & lblArray2.Caption & ".bmp")
        Me.imgPublic3.Picture = LoadPicture(chartPath & "\" & "Desoneração Gestão Pública - " & cbxMarketValuation.value & lblArray3.Caption & ".bmp")
        Me.imgPublic4.Picture = LoadPicture(chartPath & "\" & "Desoneração Gestão Pública - " & cbxMarketValuation.value & lblArray4.Caption & ".bmp")
        
        formulaUp1.Caption = wksBridgeData.Cells(3, 25).value
        formulaDown1.Caption = wksBridgeData.Cells(3, 26).value
        formulaResult1 = wksBridgeData.Cells(3, 27).value
        
        formulaUp2.Caption = wksBridgeData.Cells(5, 25).value
        formulaDown2.Caption = wksBridgeData.Cells(5, 26).value
        formulaResult2 = wksBridgeData.Cells(5, 27).value
        
        formulaUp3.Caption = wksBridgeData.Cells(7, 25).value
        formulaDown3.Caption = wksBridgeData.Cells(7, 26).value
        formulaResult3 = wksBridgeData.Cells(7, 27).value
        
        formulaUp4.Caption = wksBridgeData.Cells(9, 25).value
        formulaDown4.Caption = wksBridgeData.Cells(9, 26).value
        formulaResult4 = wksBridgeData.Cells(9, 27).value
        
        For Each Ctrl In Me.Controls
            If InStr(Ctrl.name, "formula") > 0 Then
                Ctrl.Visible = True
            End If
        Next Ctrl
        
    Else
        For Each Ctrl In Me.Controls
            If InStr(Ctrl.name, "formula") > 0 Then
                Ctrl.Visible = False
            ElseIf InStr(Ctrl.name, "imgEffort") > 0 Or InStr(Ctrl.name, "imgIndirect") > 0 Or InStr(Ctrl.name, "imgPublic") > 0 Then
                Ctrl.Picture = Nothing
            End If
        Next Ctrl
    End If
End Sub

Private Sub cbxRoute_Change()
    If cbxMarketRoute.value <> "" And cbxArrayRoute.value <> "" And cbxSubArrayRoute.value <> "" And cbxRoute.value <> "" Then
        Call ChangeRoute
        Call enableDisableRouteLabels(True, cbxRoute.value)
    End If
End Sub

Private Sub cbxSubArray_Change()
    If cbxMarket.value <> "" And cbxArray.value <> "" And cbxSubArray.value <> "" Then
    
        Call PlotGraph
    
    End If
End Sub

Private Sub cbxSubArrayRoute_Change()
    If cbxMarketRoute.value <> "" And cbxArrayRoute.value <> "" And cbxSubArrayRoute.value <> "" Then
    
        Call ChangeRoute
        
        If cbxSubArrayRoute.value = "Consolidado" Then
            cbxRoute.Visible = False
            lblRoute.Visible = False
            Call enableDisableRouteLabels(False, "")
            cbxRoute.ListIndex = -1
            lblSelectedRoute.Visible = False
            Call updateConsolidadoValues
            Call enableDisableConsolidado(True)
        Else
            cbxRoute.Visible = True
            lblRoute.Visible = True
            Call enableDisableConsolidado(False)
        End If
        
        
    End If
End Sub


Sub updateConsolidadoValues()
    Dim wksChartData As Worksheet
    Set wksChartData = Util.GetChartDataWorksheet
    
    Dim rowSubArray1, rowSubArray2, rowSubArray3, rowArray
    rowSubArray1 = 57
    rowSubArray2 = 58
    rowSubArray3 = 59
    rowArray = 60
    
    idSubConsolidado1.Caption = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colID)
    idSubConsolidado2.Caption = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colID)
    idSubConsolidado3.Caption = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colID)
    idArrayConsolidado.Caption = wksChartData.Cells(rowArray, ArrayConsolidado.colID)
    
    routeSubConsolidado1.Text = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colRota)
    routeSubConsolidado2.Text = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colRota)
    routeSubConsolidado3.Text = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colRota)
    
    tarifaBrutaSubConsolidado1.Text = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colTarifaBruta)
    tarifaBrutaSubConsolidado2.Text = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colTarifaBruta)
    tarifaBrutaSubConsolidado3.Text = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colTarifaBruta)
    tarifaBrutaArrayConsolidado.Text = wksChartData.Cells(rowArray, ArrayConsolidado.colTarifaBruta)
    
    tarifaLiquidaSubConsolidado1.Text = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colTarifaLiquida)
    tarifaLiquidaSubConsolidado2.Text = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colTarifaLiquida)
    tarifaLiquidaSubConsolidado3.Text = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colTarifaLiquida)
    tarifaLiquidaArrayConsolidado.Text = wksChartData.Cells(rowArray, ArrayConsolidado.colTarifaLiquida)
    
    eficienciaSubConsolidado1.Text = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colEficienciaValorizacao)
    eficienciaSubConsolidado2.Text = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colEficienciaValorizacao)
    eficienciaSubConsolidado3.Text = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colEficienciaValorizacao)
    eficienciaArrayConsolidado.Text = wksChartData.Cells(rowArray, ArrayConsolidado.colEficienciaValorizacao)
    
    capexSubConsolidado1.Text = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colCapex)
    capexSubConsolidado2.Text = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colCapex)
    capexSubConsolidado3.Text = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colCapex)
    capexArrayConsolidado.Text = wksChartData.Cells(rowArray, ArrayConsolidado.colCapex)
    
    opexSubConsolidado1.Text = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colOpex)
    opexSubConsolidado2.Text = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colOpex)
    opexSubConsolidado3.Text = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colOpex)
    opexArrayConsolidado.Text = wksChartData.Cells(rowArray, ArrayConsolidado.colOpex)
    
    entradaSubConsolidado1.Text = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colEntradaPlanta)
    entradaSubConsolidado2.Text = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colEntradaPlanta)
    entradaSubConsolidado3.Text = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colEntradaPlanta)
    entradaArrayConsolidado.Text = wksChartData.Cells(rowArray, ArrayConsolidado.colEntradaPlanta)
    
    reciclaveisSubConsolidado1.Text = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colReciclaveis)
    reciclaveisSubConsolidado2.Text = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colReciclaveis)
    reciclaveisSubConsolidado3.Text = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colReciclaveis)
    reciclaveisArrayConsolidado.Text = wksChartData.Cells(rowArray, ArrayConsolidado.colReciclaveis)
    
    cdrSubConsolidado1.Text = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colCDR)
    cdrSubConsolidado2.Text = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colCDR)
    cdrSubConsolidado3.Text = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colCDR)
    cdrArrayConsolidado.Text = wksChartData.Cells(rowArray, ArrayConsolidado.colCDR)
    
    rejeitosSubConsolidado1.Text = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colRejeitos)
    rejeitosSubConsolidado2.Text = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colRejeitos)
    rejeitosSubConsolidado3.Text = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colRejeitos)
    rejeitosArrayConsolidado.Text = wksChartData.Cells(rowArray, 9)
    
    compostoSubConsolidado1.Text = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colComposto)
    compostoSubConsolidado2.Text = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colComposto)
    compostoSubConsolidado3.Text = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colComposto)
    compostoArrayConsolidado.Text = wksChartData.Cells(rowArray, ArrayConsolidado.colComposto)
    
    perdaMassaSubConsolidado1.Text = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colPerdaMassa)
    perdaMassaSubConsolidado2.Text = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colPerdaMassa)
    perdaMassaSubConsolidado3.Text = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colPerdaMassa)
    perdaMassaArrayConsolidado.Text = wksChartData.Cells(rowArray, ArrayConsolidado.colPerdaMassa)
    
    biometanoSubConsolidado1.Text = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colBiometano)
    biometanoSubConsolidado2.Text = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colBiometano)
    biometanoSubConsolidado3.Text = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colBiometano)
    biometanoArrayConsolidado.Text = wksChartData.Cells(rowArray, ArrayConsolidado.colBiometano)
    
    energiaSubConsolidado1.Text = wksChartData.Cells(rowSubArray1, ArrayConsolidado.colEnergiaEletrica)
    energiaSubConsolidado2.Text = wksChartData.Cells(rowSubArray2, ArrayConsolidado.colEnergiaEletrica)
    energiaSubConsolidado3.Text = wksChartData.Cells(rowSubArray3, ArrayConsolidado.colEnergiaEletrica)
    energiaArrayConsolidado.Text = wksChartData.Cells(rowArray, ArrayConsolidado.colEnergiaEletrica)
    
End Sub


Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 2, "Passo 5")
    txtChartDescription.TextAlign = fmTextAlignCenter
    
    Dim prjPath As String
    Dim prjName As String
    
    prjPath = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
    prjName = Database.GetDatabaseValue("ProjectName", colUserValue)
    prjPath = Util.FolderCreate(prjPath, prjName)
    
    'Create base market folder
    Dim chartPath As String
    chartPath = Util.FolderCreate(prjPath, FOLDERCHART)
    
    Set arrays = readArrays
    
    'Ajustar arranjos selecionados na aba de "Dados - Gráfico"
    Dim wksChartData, wksBridgeData As Worksheet
    Set wksChartData = Util.GetChartDataWorksheet
    Set wksBridgeData = Util.GetBridgeDataWorksheet
    Dim markets As Variant
    markets = Array(FOLDERBASEMARKET, FOLDEROPTIMIZEDMARKET, FOLDERLANDFILLMARKET)
    Dim row, selected, rowBridge As Integer
    row = 4
    For Each m In markets
        selected = 1
        rowBridge = 3
        For Each a In arrays
            If a.vSelected Then
                wksChartData.Cells(row, 1).value = GetMarketCode(m) & a.vCode
                wksBridgeData.Cells(rowBridge, 1).value = a.vCode
                Me.Controls("lblArray" & selected).Caption = a.vCode
                row = row + 1
                rowBridge = rowBridge + 2
                selected = selected + 1
            End If
        Next
    Next m
    
    Dim MyChart As Chart
    Dim Fname As String
    
    For Each c In ThisWorkbook.Sheets("Dashboard").ChartObjects
        cbxCharts.AddItem c.Chart.ChartTitle.Text
        c.Activate
        Fname = chartPath & "\" & c.Chart.ChartTitle.Text & ".jpg"
        c.Chart.Export filename:=Fname, FilterName:="jpg"
    Next c
    
    cbxMarket.AddItem "M1"
    cbxMarket.AddItem "M2"
    cbxMarket.AddItem "M3"
    cbxMarketRoute.AddItem "M1"
    cbxMarketRoute.AddItem "M2"
    cbxMarketRoute.AddItem "M3"
    cbxMarketValuation.AddItem "M1"
    cbxMarketValuation.AddItem "M2"
    cbxMarketValuation.AddItem "M3"
    
    cbxRoute.AddItem "RT1-A"
    cbxRoute.AddItem "RT1-B"
    cbxRoute.AddItem "RT1-C"
    cbxRoute.AddItem "RT2"
    cbxRoute.AddItem "RT3"
    cbxRoute.AddItem "RT4"
    cbxRoute.AddItem "RT5"
    
    For Each a In arrays
        If a.vSelected Then
            cbxArray.AddItem a.vCode
            cbxArrayRoute.AddItem a.vCode
            cbxArraySelected.AddItem a.vCode
        End If
    Next
    
    
    
    For Each Ctrl In Me.Controls
        If InStr(Ctrl.name, "formula") > 0 Then
            Ctrl.Visible = False
        End If
    Next Ctrl
    
    Call enableDisableRouteLabels(False, "")
    
    Const startIndx As Long = 0
    With Me.MultiPage1
        .Pages(startIndx).Caption = ChkMark & .Pages(startIndx).Caption
        .Tag = startIndx
    End With
    
    Call enableDisableConsolidado(False)
    
    Me.Height = 615
    Me.width = 930
    
End Sub

Sub enableDisableConsolidado(ByVal onoff As Boolean)
    For Each Ctrl In Me.Controls
        If InStr(Ctrl.name, "Consolidado") > 0 Then
            Ctrl.Visible = onoff
        End If
    Next Ctrl
End Sub

Sub enableDisableRouteLabels(ByVal onoff As Boolean, ByVal route As String)
    
    imgRoute.Visible = onoff
    
    For Each Ctrl In Me.Controls
        If InStr(Ctrl.name, "RouteData") > 0 Then
            Ctrl.Visible = onoff
        End If
    Next Ctrl
    
    finalUsage2RouteData.Visible = False
    If route <> "RT4" Or route <> "RT5" Then
        landfillDangerRouteData.Visible = False
    End If
    
    If route = "RT1-A" Then
        inputRouteData.Left = 168
        inputRouteData.Top = 110
        reciclableRouteData.Left = 138
        reciclableRouteData.Top = 212
        cdrRouteData.Left = 138
        cdrRouteData.Top = 255
        lossWeightRouteData.Left = 354
        lossWeightRouteData.Top = 78
        organicCompoundRouteData.Left = 192
        organicCompoundRouteData.Top = 366
        finalUsageRouteData.Left = 708
        finalUsageRouteData.Top = 244
        landfillRouteData.Left = 594
        landfillRouteData.Top = 282
        biogasRouteData.Visible = True
    ElseIf route = "RT1-B" Then
        inputRouteData.Left = 168
        inputRouteData.Top = 110
        reciclableRouteData.Left = 138
        reciclableRouteData.Top = 212
        cdrRouteData.Left = 138
        cdrRouteData.Top = 255
        lossWeightRouteData.Left = 354
        lossWeightRouteData.Top = 78
        organicCompoundRouteData.Left = 192
        organicCompoundRouteData.Top = 366
        finalUsageRouteData.Left = 720
        finalUsageRouteData.Top = 117
        landfillRouteData.Left = 540
        landfillRouteData.Top = 312
        finalUsage2RouteData.Left = 714
        finalUsage2RouteData.Top = 276
        finalUsage2RouteData.Visible = True
        biogasRouteData.Visible = True
    ElseIf route = "RT1-C" Then
        inputRouteData.Left = 168
        inputRouteData.Top = 110
        reciclableRouteData.Left = 138
        reciclableRouteData.Top = 212
        cdrRouteData.Left = 138
        cdrRouteData.Top = 255
        lossWeightRouteData.Left = 354
        lossWeightRouteData.Top = 78
        organicCompoundRouteData.Left = 192
        organicCompoundRouteData.Top = 366
        finalUsageRouteData.Left = 716
        finalUsageRouteData.Top = 238
        landfillRouteData.Left = 540
        landfillRouteData.Top = 312
        biogasRouteData.Visible = True
    ElseIf route = "RT2" Then
        inputRouteData.Left = 198
        inputRouteData.Top = 126
        reciclableRouteData.Left = 187
        reciclableRouteData.Top = 230
        cdrRouteData.Left = 190
        cdrRouteData.Top = 272
        lossWeightRouteData.Left = 390
        lossWeightRouteData.Top = 103
        organicCompoundRouteData.Left = 678
        organicCompoundRouteData.Top = 140
        landfillRouteData.Left = 552
        landfillRouteData.Top = 300
        finalUsageRouteData.Visible = False
        biogasRouteData.Visible = False
    ElseIf route = "RT3" Then
        inputRouteData.Left = 180
        inputRouteData.Top = 128
        reciclableRouteData.Left = 192
        reciclableRouteData.Top = 300
        cdrRouteData.Left = 704
        cdrRouteData.Top = 300
        lossWeightRouteData.Left = 358
        lossWeightRouteData.Top = 102
        landfillRouteData.Left = 542
        landfillRouteData.Top = 302
        organicCompoundRouteData.Visible = False
        finalUsageRouteData.Visible = False
        biogasRouteData.Visible = False
    ElseIf route = "RT4" Or route = "RT5" Then
        inputRouteData.Left = 163
        inputRouteData.Top = 126
        reciclableRouteData.Left = 186
        reciclableRouteData.Top = 275
        lossWeightRouteData.Left = 372
        lossWeightRouteData.Top = 99
        landfillRouteData.Left = 397
        landfillRouteData.Top = 300
        finalUsageRouteData.Left = 710
        finalUsageRouteData.Top = 228
        landfillDangerRouteData.Left = 618
        landfillDangerRouteData.Top = 308
        cdrRouteData.Visible = False
        organicCompoundRouteData.Visible = False
        landfillDangerRouteData.Visible = True
        biogasRouteData.Visible = False
    End If

End Sub
