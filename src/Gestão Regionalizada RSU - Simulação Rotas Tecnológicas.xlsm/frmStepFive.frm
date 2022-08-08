VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepFive 
   Caption         =   "UserForm1"
   ClientHeight    =   9915.001
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   17130
   OleObjectBlob   =   "frmStepFive.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStepFive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
    frmTool.updateForm
    Unload Me
End Sub


Private Sub cbxCharts_Change()
    currentChart = cbxCharts
    For Each c In Sheets("Dashboard").ChartObjects
        If c.name = currentChart Then
            Dim chartPath As String
            Dim prjPath As String
            Dim prjName As String
            prjPath = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
            prjName = Database.GetDatabaseValue("ProjectName", colUserValue)
            prjPath = Util.FolderCreate(prjPath, prjName)
            chartPath = Util.FolderCreate(prjPath, FOLDERCHART)
            Fname = chartPath & "\" & c.name & ".jpg"
            Me.Image1.Picture = LoadPicture(Fname)
        End If
    Next c
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 2, "Passo 5")
    
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
        cbxCharts.AddItem c.name
        Fname = chartPath & "\" & c.name & ".jpg"
        c.Chart.Export filename:=Fname, FilterName:="jpg"
    Next c
    

End Sub
