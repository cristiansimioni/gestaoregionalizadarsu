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

Private Sub cbxArray_Change()
    cbxSubArray.Clear
    
    For Each a In arrays
        If a.vSelected Then
            If a.vCode = cbxArray.value Then
                For Each s In a.vSubArray
                    cbxSubArray.AddItem s.vCode
                Next s
            End If
        End If
    Next
    
    If cbxMarket.value <> "" And cbxArray.value <> "" And cbxSubArray.value <> "" Then
    
        Call PlotGraph
    
    End If
    
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

Private Sub cbxSubArray_Change()
    If cbxMarket.value <> "" And cbxArray.value <> "" And cbxSubArray.value <> "" Then
    
        Call PlotGraph
    
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
    
    For Each a In arrays
        If a.vSelected Then
            cbxArray.AddItem a.vCode
        End If
    Next

End Sub
