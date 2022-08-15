VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgressBarSimulation 
   Caption         =   "Processando..."
   ClientHeight    =   1545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12435
   OleObjectBlob   =   "frmProgressBarSimulation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgressBarSimulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total As Long
Dim processed As Long
Dim width As Long
Dim percent As Double

Private Sub executeSimulation()
    'Create project folder
    Dim prjPath As String
    Dim prjName As String
    
    
    lblFile = "Criando arquivos..."
    DoEvents
    
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
    
    total = arrays.count
    
    Dim markets, routes As Variant
    markets = Array(FOLDERBASEMARKET, FOLDEROPTIMIZEDMARKET, FOLDERLANDFILLMARKET)
    'markets = Array(FOLDERBASEMARKET)
    routes = Array("RT1", "RT2", "RT3", "RT4", "RT5")
    
    Dim wksDefinedArrays As Worksheet
    Set wksDefinedArrays = Util.GetDefinedArraysWorksheet
    
    Dim row As Integer
    row = 2
    
    wksDefinedArrays.range("A2:BJ2000").ClearContents
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False
    
    total = 0
    For Each a In arrays
        If a.vSelected Then
            total = total + a.vSubArray.count
        End If
    Next a
    total = total * (UBound(markets) - LBound(markets) + 1) * (UBound(routes) - LBound(routes) + 1 + 1) '+1 Ferramenta 2
    
    Dim tarifaLiquida, eficiencia As Double
    tarifaLiquidaBase = Database.GetDatabaseValue("TargetExpectation", colUserValue)
    eficienciaBase = Database.GetDatabaseValue("ValuationEfficiency", colUserValue) / 100
    
    
    processed = 1
    For Each a In arrays
        If a.vSelected Then
            For Each m In markets
                Dim marketPath, arrayMarketPath As String
                marketPath = Util.FolderCreate(prjPath, m)
                arrayMarketPath = Util.FolderCreate(marketPath, a.vCode)
                
                For Each s In a.vSubArray
                    Dim routeFiles As New Collection
                    Dim consolidatedRows As New Collection
                    For Each r In routes
                        Dim subArrayBaseMarketPath, subArrayOptimizedMarketPath, subArrayLandfillMarketPath, newFile, templateFile As String
                        subArrayMarketPath = Util.FolderCreate(arrayMarketPath, s.vCode)
                        
                        If InStr(r, "RT1") Then
                            wksDefinedArrays.Cells(row, 1).value = m
                            wksDefinedArrays.Cells(row, 2).value = a.vCode
                            wksDefinedArrays.Cells(row, 3).value = s.vCode
                            wksDefinedArrays.Cells(row, 4).value = "RT1-A"
                            wksDefinedArrays.Cells(row, 5).value = GetMarketCode(m) & s.vCode & "RT1-A"
                            row = row + 1
                            wksDefinedArrays.Cells(row, 1).value = m
                            wksDefinedArrays.Cells(row, 2).value = a.vCode
                            wksDefinedArrays.Cells(row, 3).value = s.vCode
                            wksDefinedArrays.Cells(row, 4).value = "RT1-B"
                            wksDefinedArrays.Cells(row, 5).value = GetMarketCode(m) & s.vCode & "RT1-B"
                            row = row + 1
                            wksDefinedArrays.Cells(row, 1).value = m
                            wksDefinedArrays.Cells(row, 2).value = a.vCode
                            wksDefinedArrays.Cells(row, 3).value = s.vCode
                            wksDefinedArrays.Cells(row, 4).value = "RT1-C"
                            wksDefinedArrays.Cells(row, 5).value = GetMarketCode(m) & s.vCode & "RT1-C"
                            row = row + 1
                        Else
                            wksDefinedArrays.Cells(row, 1).value = m
                            wksDefinedArrays.Cells(row, 2).value = a.vCode
                            wksDefinedArrays.Cells(row, 3).value = s.vCode
                            wksDefinedArrays.Cells(row, 4).value = r
                            wksDefinedArrays.Cells(row, 5).value = GetMarketCode(m) & s.vCode & r
                            row = row + 1
                        End If
                        
                        'Create routes from 1 to 5 for all markets
                        newFile = subArrayMarketPath & "\" & GetMarketCode(m) & s.vCode & r & ".xlsm"
                        templateFile = Application.ThisWorkbook.Path & "\templates\Base Ferramenta 3 - " & r & ".xlsm"
                        
                        lblFile = "Processando arquivo: " & newFile
                        percent = processed / total
                        lblProgress.width = percent * width
                        lblValue = Round(percent * 100, 1) & "%"
                        processed = processed + 1
                        DoEvents
                        
                        routeFiles.Add newFile
                        
                        'Only create the file if it's not created yet
                        If Len(Dir(newFile)) = 0 Then
                            FileCopy templateFile, newFile
                        End If
                        
                        Call EditRouteToolData(newFile, s, m)
                        
                    Next r

                    'Create tool 2 for array
                    Dim toolTwoFile, templateToolTwoFile As String
                    toolTwoFile = subArrayMarketPath & "\" & GetMarketCode(m) & s.vCode & " - Ferramenta 2.xlsm"
                    templateFile = Application.ThisWorkbook.Path & "\templates\Base Ferramenta 3 - Ferramenta 2.xlsm"
                    
                    
                    lblFile = "Processando arquivo: " & toolTwoFile
                    DoEvents
                    
                    'Only create the file if it's not created yet
                     If Len(Dir(toolTwoFile)) = 0 Then
                        FileCopy templateFile, toolTwoFile
                     End If
                    
                    Call EditToolTwoData(toolTwoFile, routeFiles, s, m)
                    
                    Call CopyDataFromToolTwo(toolTwoFile, row)
                    
                    
                    'Verificar qual é a melhor rota
                    Dim rowRoute As Integer
                    rowRoute = row - 7
                    Dim selectedRow As Integer
                    Dim minTarifa, maxEficiencia As Double
                    minTarifa = 999999
                    maxEficiencia = -100#
                    
                    For rowRoute = row - 7 To row - 1

                        If minTarifa > wksDefinedArrays.Cells(rowRoute, 9).value Then
                            minTarifa = wksDefinedArrays.Cells(rowRoute, 9).value
                            selectedRow = rowRoute
                        End If
                        If maxEficiencia < wksDefinedArrays.Cells(rowRoute, 10).value Then
                            maxEficiencia = wksDefinedArrays.Cells(rowRoute, 10).value
                        End If
                    Next rowRoute
                    
                    wksDefinedArrays.Cells(row, 1).value = m
                    wksDefinedArrays.Cells(row, 2).value = a.vCode
                    wksDefinedArrays.Cells(row, 3).value = s.vCode & "(Consolidado)"
                    wksDefinedArrays.Cells(row, 4).value = "NA"
                    wksDefinedArrays.Cells(row, 5).value = GetMarketCode(m) & s.vCode
                    
                    For x = 6 To 65
                        wksDefinedArrays.Cells(row, x).value = wksDefinedArrays.Cells(selectedRow, x).value
                    Next x
                    
                    If tarifaLiquidaBase > minTarifa Then
                        wksDefinedArrays.Cells(row, 9).Interior.Color = ApplicationColors.bgColorValidTextBox
                    Else
                        wksDefinedArrays.Cells(row, 9).Interior.Color = ApplicationColors.bgColorInvalidTextBox
                    End If
                    
                    If eficienciaBase < maxEficiencia Then
                        wksDefinedArrays.Cells(row, 10).Interior.Color = ApplicationColors.bgColorValidTextBox
                    Else
                        wksDefinedArrays.Cells(row, 10).Interior.Color = ApplicationColors.bgColorInvalidTextBox
                    End If
                    
                    consolidatedRows.Add row
                    
                    row = row + 1
                    
                    Set routeFiles = Nothing
                    
                    percent = processed / total
                    lblProgress.width = percent * width
                    lblValue = Round(percent * 100, 1) & "%"
                    processed = processed + 1
                    DoEvents
                
                Next s
                    
                    
                'Read data from tool 2 and insert into sheet
                wksDefinedArrays.Cells(row, 1).value = m
                wksDefinedArrays.Cells(row, 2).value = a.vCode & "(Consolidado)"
                wksDefinedArrays.Cells(row, 3).value = "NA"
                wksDefinedArrays.Cells(row, 4).value = "NA"
                wksDefinedArrays.Cells(row, 5).value = GetMarketCode(m) & a.vCode
                
                For x = 6 To 65
                    Dim strFormula As String
                    Dim ColumnLetter As String
                    strFormula = "="
                    Dim element As Integer
                    element = 1
                    ColumnLetter = Split(Cells(1, x).Address, "$")(1)
                    
                    If x = 11 Or x = 5 Or x = 6 Or x = 7 Then 'Fixas
                        strFormula = strFormula & ColumnLetter & consolidatedRows(1)
                    ElseIf x >= 12 And x <= 23 Then 'Somatório
                        For Each r In consolidatedRows
                            If element <> 1 Then
                                strFormula = strFormula & "+"
                            End If
                            strFormula = strFormula & ColumnLetter & r
                            element = element + 1
                        Next r
                    Else 'Soma Ponderada
                        Dim divisionPart As String
                        Dim sumPart As String
                        sumPart = "("
                        divisionPart = "("
                        For Each r In consolidatedRows
                            If element <> 1 Then
                                divisionPart = divisionPart & "+"
                                sumPart = sumPart & "+"
                            End If
                            sumPart = sumPart & ColumnLetter & r & "*N" & r
                            divisionPart = divisionPart & "N" & r
                            element = element + 1
                        Next r
                        sumPart = sumPart & ")"
                        divisionPart = divisionPart & ")"
                        
                        strFormula = strFormula & sumPart & "/" & divisionPart
                    End If
                    wksDefinedArrays.Cells(row, x).Formula = strFormula
                Next x
                
                row = row + 1
                
                Set consolidatedRows = Nothing
            Next m
            
        End If
    Next a
    
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    
    Call Database.SetDatabaseValue("SimulationStatus", colUserValue, "Sim")
    
    Me.Caption = "Concluído"
    lblFile.Visible = False
    frmStepFour.updateForm
    
End Sub

Private Sub UserForm_Activate()
    Call executeSimulation
End Sub

Private Sub UserForm_Initialize()
    width = lblProgress.width
    lblProgress.width = 0
    lblFile.Visible = True
    lblProgress.BackColor = ApplicationColors.bgColorLevel2
End Sub
