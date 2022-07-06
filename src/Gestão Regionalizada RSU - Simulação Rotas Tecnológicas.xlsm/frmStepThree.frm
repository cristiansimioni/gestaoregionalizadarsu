VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepThree 
   Caption         =   "UserForm1"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7170
   OleObjectBlob   =   "frmStepThree.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStepThree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnCapexData_Click()
    frmCapexData.Show
End Sub

Private Sub btnContract_Click()
    frmContract.Show
End Sub

Private Sub btnExecuteSimulation_Click()
    'Create project folder
    Dim prjPath As String
    Dim prjName As String
    prjPath = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
    prjName = Database.GetDatabaseValue("ProjectName", colUserValue)
    prjPath = Util.FolderCreate(prjPath, prjName)
    
    'Create base market folder
    Dim baseMarketPath As String
    baseMarketPath = Util.FolderCreate(prjPath, FOLDERBASEMARKET)
    Dim optimizedMarketPath As String
    optimizedMarketPath = Util.FolderCreate(prjPath, FOLDEROPTIMIZEDMARKET)
    Dim landfillMarketPath As String
    landfillMarketPath = Util.FolderCreate(prjPath, FOLDERLANDFILLMARKET)
    
    'Process arrays
    Dim arrays As Collection
    Set arrays = readArrays
    
    Dim markets, routes As Variant
    markets = Array(FOLDERBASEMARKET, FOLDEROPTIMIZEDMARKET, FOLDERLANDFILLMARKET)
    routes = Array("RT1-A", "RT1-B", "RT1-C", "RT2", "RT3", "RT4", "RT5")
    
    Dim wksDefinedArrays As Worksheet
    Set wksDefinedArrays = Util.GetDefinedArraysWorksheet
    
    Dim row As Integer
    row = 2
    
    For Each a In arrays
        If a.vSelected Then
            For Each m In markets
                Dim marketPath, arrayMarketPath As String
                marketPath = Util.FolderCreate(prjPath, m)
                arrayMarketPath = Util.FolderCreate(marketPath, a.vCode)
                For Each s In a.vSubArray
                    For Each r In routes
                        Dim subArrayBaseMarketPath, subArrayOptimizedMarketPath, subArrayLandfillMarketPath, newFile As String
                        subArrayMarketPath = Util.FolderCreate(arrayMarketPath, s.vCode)
                        
                        wksDefinedArrays.Cells(row, 1).value = m
                        wksDefinedArrays.Cells(row, 2).value = a.vCode
                        wksDefinedArrays.Cells(row, 3).value = s.vCode
                        wksDefinedArrays.Cells(row, 4).value = r
                        row = row + 1
                        'Create routes from 1 to 5 for all markets
                        newFile = subArrayMarketPath & "\" & r & ".xlsm"
                        FileCopy "C:\Users\cristiansimioni\Desktop\gestaoregionalizadarsu\templates\Base Ferramenta 3 - RT 1.xlsm", newFile
                        
                        Call EditRouteToolData(newFile)
                        
                    Next r
                Next s
                wksDefinedArrays.Cells(row, 1).value = m
                wksDefinedArrays.Cells(row, 2).value = a.vCode
                wksDefinedArrays.Cells(row, 3).value = ""
                wksDefinedArrays.Cells(row, 4).value = ""
                row = row + 1
            Next m
            'Create tool 2 for array
            
            'Read data from tool 2 and insert into sheet
            
        End If
    Next a
End Sub

Private Sub btnFinancialAssumptions_Click()
    frmFinancialAssumptions.Show
End Sub

Private Sub btnOpexData_Click()
    frmOpexData.Show
End Sub

Private Sub btnRouteDefinition_Click()
    frmRoute.Show
End Sub

Private Sub btnTaxes_Click()
    frmTaxes.Show
End Sub

Private Sub btnUserBase_Click()
    frmUserBase.Show
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = APPNAME & " - Passo 3"
    Me.BackColor = ApplicationColors.bgColorLevel2
    
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.btColorLevel2
         End If
    Next Ctrl
End Sub
