VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepFive 
   Caption         =   "UserForm1"
   ClientHeight    =   11700
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   18345
   OleObjectBlob   =   "frmStepFive.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStepFive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim arrays As Collection

Private Sub btnBack_Click()
    frmTool.updateForm
    Unload Me
End Sub


Private Sub btnFiles_Click()
    Dim chartPath As String
    Dim prjPath As String
    Dim prjName As String
    prjPath = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
    prjName = Database.GetDatabaseValue("ProjectName", colUserValue)
    prjPath = Util.FolderCreate(prjPath, prjName)
    ThisWorkbook.FollowHyperlink prjPath
End Sub

Private Sub btnHelpStep_Click()
    On Error Resume Next
        ActiveWorkbook.FollowHyperlink (Application.ThisWorkbook.Path & "\" & FOLDERMANUAL & "\" & FILEMANUALSTEP5)
    On Error GoTo 0
End Sub

Private Sub cbxArrayRoute_Change()
    cbxSubArrayRoute.Clear
    
    For Each A In arrays
        If A.vSelected Then
            If A.vCode = cbxArrayRoute.value Then
                For Each s In A.vSubArray
                    cbxSubArrayRoute.AddItem s.vCode
                Next s
                cbxSubArrayRoute.AddItem "Consolidado"
            End If
        End If
    Next
    
    cbxRoute.ListIndex = -1
    Call enableDisableRouteLabels(False, "")
    
    'If cbxMarketRoute.value <> "" And cbxArrayRoute.value <> "" And cbxSubArrayRoute.value <> "" And cbxRoute.value <> "" Then
    '    Call ChangeRoute
    'End If
    
End Sub

Private Sub cbxArray_Change()
    cbxSubArray.Clear
    
    For Each A In arrays
        If A.vSelected Then
            If A.vCode = cbxArray.value Then
                For Each s In A.vSubArray
                    cbxSubArray.AddItem s.vCode
                Next s
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
    
    For Each A In arrays
        If A.vSelected Then
            If A.vCode = cbxArraySelected.value Then
                txtArrayTotal.Text = A.vTotal
                txtArrayTrash.Text = A.vTrash
                txtArrayTechnology.Text = A.vTechnology
                txtArrayInbound.Text = A.vInbound
                txtArrayOutbound.Text = A.vOutbound
                txtArrayOutboundExistent.Text = A.vOutboundExistentLandfill
                
                t = 1
                For Each s In A.vSubArray
                    Me.Controls("txtSubArray" & t).value = s.vArrayRaw
                    Me.Controls("txtSubArrayLandfill" & t).value = s.vLandfill
                    Me.Controls("txtSubArrayExistentLandfill" & t).value = s.vExistentLandfill
                    Me.Controls("txtSubArrayUTVR" & t).value = s.vUTVR
                    Me.Controls("txtSubArrayTotal" & t).value = s.vTotal
                    Me.Controls("txtSubArrayTrash" & t).value = s.vTrash
                    Me.Controls("txtSubArrayTechnology" & t).value = s.vTechnology
                    Me.Controls("txtSubArrayInbound" & t).value = s.vInbound
                    Me.Controls("txtSubArrayOutbound" & t).value = s.vOutbound
                    Me.Controls("txtSubArrayOutboundExistent" & t).value = s.vOutboundExistentLandfill
                    t = t + 1
                Next s
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
    For Each c In Sheets("Dashboard").ChartObjects
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
            lblRouteTitle = "RT1A - Biodigest�o e Produ��o Energia El�trica"
            imgRoute.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMGSCREENROUTEONEA)
            lineData = 43
        ElseIf cbxRoute.value = "RT1-B" Then
            lblRouteTitle = "RT1B - Biodigest�o e Descarboniza��o Frota P�blica"
            imgRoute.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMGSCREENROUTEONEB)
            lineData = 44
        ElseIf cbxRoute.value = "RT1-C" Then
            lblRouteTitle = "RT1C - Biodigest�o e Comercializa��o Biometano"
            imgRoute.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMGSCREENROUTEONEC)
            lineData = 45
        ElseIf cbxRoute.value = "RT2" Then
            lblRouteTitle = "RT2 - Compostagem e Produ��o Composto Org�nico"
            imgRoute.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMGSCREENROUTETWO)
            lineData = 46
        ElseIf cbxRoute.value = "RT3" Then
            lblRouteTitle = "RT3 - Biosecagem e Produ��o BioCDR."
            imgRoute.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMGSCREENROUTETHREE)
            lineData = 47
        ElseIf cbxRoute.value = "RT4" Then
            lblRouteTitle = "RT4 - Incinera��o Mass Burning"
            imgRoute.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMGSCREENROUTEFOUR)
            lineData = 48
        ElseIf cbxRoute.value = "RT5" Then
            lblRouteTitle = "RT5 - Incinera��o Mass Burning Descentralizada"
            imgRoute.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMGSCREENROUTEFIVE)
            lineData = 49
        End If
    End If
    
    If lineData <> 0 Then
        capexRouteData.Caption = wksChartData.Cells(lineData, 4).value
        opexRouteData.Caption = wksChartData.Cells(lineData, 5).value
        inputRouteData.Caption = wksChartData.Cells(lineData, 6).value
        reciclableRouteData.Caption = wksChartData.Cells(lineData, 7).value
        cdrRouteData.Caption = wksChartData.Cells(lineData, 8).value
        landfillDangerRouteData.Caption = wksChartData.Cells(lineData, 11).value
        landfillRouteData.Caption = wksChartData.Cells(lineData, 9).value
        organicCompoundRouteData.Caption = wksChartData.Cells(lineData, 10).value
        lossWeightRouteData.Caption = wksChartData.Cells(lineData, 12).value
        finalUsageRouteData.Caption = Format(wksChartData.Cells(lineData, 13).value, "#,##0")
        finalUsage2RouteData.Caption = Format(wksChartData.Cells(lineData, 14).value, "#,##0")
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
        
    For Each c In Sheets("Dashboard").ChartObjects
        If c.name = "Avalia��o" Then
            c.Activate
            c.Chart.ChartTitle.Text = "Avalia��o de Custos para o Munic�pio de Tratamento de RSU" & " - " & cbxMarket.value & cbxSubArray.value
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
        If cbxSubArrayRoute.value = "Consolidado" Then
            cbxRoute.Visible = False
            lblRoute.Visible = False
            Call enableDisableRouteLabels(False, cbxRoute.value)
            cbxRoute.ListIndex = -1
        Else
            cbxRoute.Visible = True
            lblRoute.Visible = True
        End If
        
        Call ChangeRoute
    End If
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
    
    Dim MyChart As Chart
    Dim Fname As String
    
    For Each c In Sheets("Dashboard").ChartObjects
        cbxCharts.AddItem c.Chart.ChartTitle.Text
        c.Activate
        Fname = chartPath & "\" & c.Chart.ChartTitle.Text & ".jpg"
        c.Chart.Export filename:=Fname, FilterName:="jpg"
    Next c
    
    
    Set arrays = readArrays
    
    cbxMarket.AddItem "M1"
    cbxMarket.AddItem "M2"
    cbxMarket.AddItem "M3"
    
    cbxMarketRoute.AddItem "M1"
    cbxMarketRoute.AddItem "M2"
    cbxMarketRoute.AddItem "M3"
    
    cbxRoute.AddItem "RT1-A"
    cbxRoute.AddItem "RT1-B"
    cbxRoute.AddItem "RT1-C"
    cbxRoute.AddItem "RT2"
    cbxRoute.AddItem "RT3"
    cbxRoute.AddItem "RT4"
    cbxRoute.AddItem "RT5"
    
    For Each A In arrays
        If A.vSelected Then
            cbxArray.AddItem A.vCode
            cbxArrayRoute.AddItem A.vCode
            cbxArraySelected.AddItem A.vCode
        End If
    Next
    
    'Ajustar arranjos selecionados na aba de "Dados - Gr�fico"
    Dim wksChartData As Worksheet
    Set wksChartData = Util.GetChartDataWorksheet
    Dim markets As Variant
    markets = Array(FOLDERBASEMARKET, FOLDEROPTIMIZEDMARKET, FOLDERLANDFILLMARKET)
    Dim row As Integer
    row = 4
    For Each m In markets
        For Each A In arrays
            If A.vSelected Then
                wksChartData.Cells(row, 1).value = GetMarketCode(m) & A.vCode
                row = row + 1
            End If
        Next
    Next m
    
    
    Call enableDisableRouteLabels(False, "")
    
End Sub

Sub enableDisableRouteLabels(ByVal onoff As Boolean, ByVal route As String)
    
    lblRouteTitle.Visible = onoff
    imgRoute.Visible = onoff
    
    For Each Ctrl In Me.Controls
        If InStr(Ctrl.name, "RouteData") > 0 Then
            Ctrl.Visible = onoff
        End If
    Next Ctrl
    
    
    lossWeightRouteData.Left = 492
    lossWeightRouteData.Top = 90
    lossWeightRouteDataUnit.Left = 522
    lossWeightRouteDataUnit.Top = 90
    cdrRouteData.Left = 18
    cdrRouteData.Top = 243
    cdrRouteDataUnit.Left = 48
    cdrRouteDataUnit.Top = 243
    finalUsageRouteData.Top = 192
    finalUsageRouteDataUnit.Top = 192
    landfillRouteData.Top = 360
    landfillRouteDataUnit.Top = 360
    finalUsage2RouteData.Visible = False
    finalUsage2RouteDataUnit.Visible = False
    
    If route <> "RT4" Or route <> "RT5" Then
        landfillDangerRouteData.Visible = False
        landfillDangerRouteDataUnit.Visible = False
    End If
    
    If route = "RT1-A" Then
        finalUsageRouteDataUnit.Caption = "MWh/a"
    ElseIf route = "RT1-B" Then
        finalUsageRouteDataUnit.Caption = "Kl Diesel Equiv./a"
        finalUsage2RouteData.Visible = True
        finalUsage2RouteDataUnit.Visible = True
        finalUsage2RouteDataUnit.Caption = "Nm3/a"
    ElseIf route = "RT1-C" Then
        finalUsageRouteData.Top = 360
        finalUsageRouteDataUnit.Top = 360
        finalUsageRouteDataUnit.Caption = "Nm3/a"
    ElseIf route = "RT2" Then
        finalUsageRouteData.Visible = False
        finalUsageRouteDataUnit.Visible = False
        lossWeightRouteData.Left = 378
        lossWeightRouteData.Top = 216
        lossWeightRouteDataUnit.Left = 408
        lossWeightRouteDataUnit.Top = 216
    ElseIf route = "RT3" Then
        organicCompoundRouteData.Visible = False
        organicCompoundRouteDataUnit.Visible = False
        finalUsageRouteData.Visible = False
        finalUsageRouteDataUnit.Visible = False
        cdrRouteData.Left = 642
        cdrRouteData.Top = 318
        cdrRouteDataUnit.Left = 672
        cdrRouteDataUnit.Top = 318
        lossWeightRouteData.Left = 492
        lossWeightRouteData.Top = 115
        lossWeightRouteDataUnit.Left = 522
        lossWeightRouteDataUnit.Top = 115
    ElseIf route = "RT4" Or route = "RT5" Then
        cdrRouteData.Visible = False
        cdrRouteDataUnit.Visible = False
        organicCompoundRouteData.Visible = False
        organicCompoundRouteDataUnit.Visible = False
        landfillDangerRouteData.Visible = True
        landfillDangerRouteDataUnit.Visible = True
        finalUsageRouteData.Top = 210
        finalUsageRouteDataUnit.Top = 210
        lossWeightRouteData.Left = 612
        lossWeightRouteData.Top = 75
        lossWeightRouteDataUnit.Left = 642
        lossWeightRouteDataUnit.Top = 75
        landfillRouteData.Top = 348
        landfillRouteDataUnit.Top = 348
        finalUsageRouteDataUnit.Caption = "MWh/a"
    End If

End Sub
