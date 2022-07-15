VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStepFour 
   Caption         =   "UserForm1"
   ClientHeight    =   8535.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7920
   OleObjectBlob   =   "frmStepFour.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStepFour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnPriceVal_Click()
    frmPriceVal.Show
End Sub

Private Sub btnQuantitativeVal_Click()
    frmQuantitativeVal.Show
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
    routes = Array("RT1", "RT2", "RT3", "RT4", "RT5")
    
    Dim wksDefinedArrays As Worksheet
    Set wksDefinedArrays = Util.GetDefinedArraysWorksheet
    
    Dim row As Integer
    row = 2
    
    wksDefinedArrays.range("A2:BJ2000").ClearContents
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False
    
    For Each a In arrays
        If a.vSelected Then
            For Each m In markets
                Dim marketPath, arrayMarketPath As String
                marketPath = Util.FolderCreate(prjPath, m)
                arrayMarketPath = Util.FolderCreate(marketPath, a.vCode)
                
                For Each s In a.vSubArray
                    Dim routeFiles As New Collection
                    For Each r In routes
                        Dim subArrayBaseMarketPath, subArrayOptimizedMarketPath, subArrayLandfillMarketPath, newFile, templateFile As String
                        subArrayMarketPath = Util.FolderCreate(arrayMarketPath, s.vCode)
                        
                        If InStr(r, "RT1") Then
                            wksDefinedArrays.Cells(row, 1).value = m
                            wksDefinedArrays.Cells(row, 2).value = a.vCode
                            wksDefinedArrays.Cells(row, 3).value = s.vCode
                            wksDefinedArrays.Cells(row, 4).value = "RT1-A"
                            row = row + 1
                            wksDefinedArrays.Cells(row, 1).value = m
                            wksDefinedArrays.Cells(row, 2).value = a.vCode
                            wksDefinedArrays.Cells(row, 3).value = s.vCode
                            wksDefinedArrays.Cells(row, 4).value = "RT1-B"
                            row = row + 1
                            wksDefinedArrays.Cells(row, 1).value = m
                            wksDefinedArrays.Cells(row, 2).value = a.vCode
                            wksDefinedArrays.Cells(row, 3).value = s.vCode
                            wksDefinedArrays.Cells(row, 4).value = "RT1-C"
                            row = row + 1
                        Else
                            wksDefinedArrays.Cells(row, 1).value = m
                            wksDefinedArrays.Cells(row, 2).value = a.vCode
                            wksDefinedArrays.Cells(row, 3).value = s.vCode
                            wksDefinedArrays.Cells(row, 4).value = r
                            row = row + 1
                        End If
                        
                        'Create routes from 1 to 5 for all markets
                        newFile = subArrayMarketPath & "\" & GetMarketCode(m) & s.vCode & r & ".xlsm"
                        templateFile = Application.ActiveWorkbook.Path & "\templates\Base Ferramenta 3 - " & r & ".xlsm"
                        
                        routeFiles.Add newFile
                        
                        'Only create the file if it's not created yet
                        If Len(Dir(newFile)) = 0 Then
                            FileCopy templateFile, newFile
                        End If
                        
                        Call EditRouteToolData(newFile, s, m)
                        
                    Next r
                    wksDefinedArrays.Cells(row, 1).value = m
                    wksDefinedArrays.Cells(row, 2).value = a.vCode
                    wksDefinedArrays.Cells(row, 3).value = "Consolidado"
                    wksDefinedArrays.Cells(row, 4).value = "NA"
                    row = row + 1
                    
                    'Create tool 2 for array
                    Dim toolTwoFile, templateToolTwoFile As String
                    toolTwoFile = subArrayMarketPath & "\" & GetMarketCode(m) & a.vCode & r & " - Ferramenta 2.xlsm"
                    templateFile = Application.ActiveWorkbook.Path & "\templates\Base Ferramenta 3 - Ferramenta 2.xlsm"
                    
                    'Only create the file if it's not created yet
                     If Len(Dir(toolTwoFile)) = 0 Then
                        FileCopy templateFile, toolTwoFile
                     End If
                    
                    Call EditToolTwoData(toolTwoFile, routeFiles)
                    
                    Call CopyDataFromToolTwo(toolTwoFile, row)
                Next s
                    
                    
                'Read data from tool 2 and insert into sheet
                wksDefinedArrays.Cells(row, 1).value = m
                wksDefinedArrays.Cells(row, 2).value = "Consolidado"
                wksDefinedArrays.Cells(row, 3).value = "NA"
                wksDefinedArrays.Cells(row, 4).value = "NA"
                row = row + 1
            Next m
            
        End If
    Next a
    
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    MsgBox "Simulação finalizada com sucesso!", vbInformation
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

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 2, "Passo 4")
End Sub
