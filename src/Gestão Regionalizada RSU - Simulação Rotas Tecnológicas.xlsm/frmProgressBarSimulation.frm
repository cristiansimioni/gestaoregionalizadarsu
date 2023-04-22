VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgressBarSimulation 
   Caption         =   "Processando..."
   ClientHeight    =   1500
   ClientLeft      =   30
   ClientTop       =   135
   ClientWidth     =   9975.001
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
    Dim StartTimeTotal As Double
    Dim SecondsElapsedTotal As Double
    
    StartTimeTotal = Timer
    
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
    
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    

    total = arrays.count
    
    Dim markets, routes As Variant
    markets = Array(FOLDERBASEMARKET, FOLDEROPTIMIZEDMARKET, FOLDERLANDFILLMARKET)
    'markets = Array(FOLDERBASEMARKET)
    routes = Array("RT1", "RT2", "RT3", "RT4", "RT5")
    
    Dim wksDefinedArrays As Worksheet
    Set wksDefinedArrays = Util.GetDefinedArraysWorksheet
    
    Dim row As Integer
    row = 3
    
    wksDefinedArrays.range("A3:BJ2000").ClearContents
    wksDefinedArrays.range("A3:BJ2000").Interior.Color = xlNone
    wksDefinedArrays.range("A3:BJ2000").Font.Bold = False
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    total = 0
    For Each A In arrays
        If A.vSelected Then
            total = total + A.vSubArray.count
        End If
    Next A
    total = total * (UBound(markets) - LBound(markets) + 1) * (UBound(routes) - LBound(routes) + 1 + 1) '+1 Ferramenta 2
    
    Dim tarifaLiquida, eficiencia As Double
    tarifaLiquidaBase = Database.GetDatabaseValue("TargetExpectation", colUserValue)
    eficienciaBase = Database.GetDatabaseValue("ValuationEfficiency", colUserValue) / 100
    
    Dim selectedBaseRoute As String
    
    processed = 1
    For Each A In arrays
        If A.vSelected Then
            For Each m In markets
                selectedBaseRoute = ""
                Dim marketPath, arrayMarketPath As String
                marketPath = Util.FolderCreate(prjPath, m)
                arrayMarketPath = Util.FolderCreate(marketPath, A.vCode)
                
                For Each S In A.vSubArray
                    Dim routeFiles As New Collection
                    Dim consolidatedRows As New Collection
                    For Each r In routes
                        Dim subArrayBaseMarketPath, subArrayOptimizedMarketPath, subArrayLandfillMarketPath, newFile, templateFile As String
                        subArrayMarketPath = Util.FolderCreate(arrayMarketPath, S.vCode)
                        
                        If InStr(r, "RT1") Then
                            wksDefinedArrays.Cells(row, 1).value = m
                            wksDefinedArrays.Cells(row, 2).value = A.vCode
                            wksDefinedArrays.Cells(row, 3).value = S.vCode
                            wksDefinedArrays.Cells(row, 4).value = "RT1-A"
                            wksDefinedArrays.Cells(row, 5).value = GetMarketCode(m) & S.vCode & "RT1-A"
                            row = row + 1
                            wksDefinedArrays.Cells(row, 1).value = m
                            wksDefinedArrays.Cells(row, 2).value = A.vCode
                            wksDefinedArrays.Cells(row, 3).value = S.vCode
                            wksDefinedArrays.Cells(row, 4).value = "RT1-B"
                            wksDefinedArrays.Cells(row, 5).value = GetMarketCode(m) & S.vCode & "RT1-B"
                            row = row + 1
                            wksDefinedArrays.Cells(row, 1).value = m
                            wksDefinedArrays.Cells(row, 2).value = A.vCode
                            wksDefinedArrays.Cells(row, 3).value = S.vCode
                            wksDefinedArrays.Cells(row, 4).value = "RT1-C"
                            wksDefinedArrays.Cells(row, 5).value = GetMarketCode(m) & S.vCode & "RT1-C"
                            row = row + 1
                        Else
                            wksDefinedArrays.Cells(row, 1).value = m
                            wksDefinedArrays.Cells(row, 2).value = A.vCode
                            wksDefinedArrays.Cells(row, 3).value = S.vCode
                            wksDefinedArrays.Cells(row, 4).value = r
                            wksDefinedArrays.Cells(row, 5).value = GetMarketCode(m) & S.vCode & r
                            row = row + 1
                        End If
                        
                        StartTime = Timer
                        
                        'Create routes from 1 to 5 for all markets
                        newFile = subArrayMarketPath & "\" & GetMarketCode(m) & S.vCode & r & ".xlsm"
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
                        
                        Call EditRouteToolData(newFile, S, m)
                        
                        SecondsElapsed = Round(Timer - StartTime, 2)
                        
                        Debug.Print "Criar e editar: " & newFile & " - Tempo: " & SecondsElapsed
                        
                    Next r

                    'Create tool 2 for array
                    Dim toolTwoFile, templateToolTwoFile As String
                    toolTwoFile = subArrayMarketPath & "\" & GetMarketCode(m) & S.vCode & " - Ferramenta 2.xlsm"
                    templateFile = Application.ThisWorkbook.Path & "\templates\Base Ferramenta 3 - Ferramenta 2.xlsm"
                    
                    StartTime = Timer
                    
                    lblFile = "Processando arquivo: " & toolTwoFile
                    DoEvents
                    
                    'Only create the file if it's not created yet
                     If Len(Dir(toolTwoFile)) = 0 Then
                        FileCopy templateFile, toolTwoFile
                     End If
                    
                    Call EditToolTwoData(toolTwoFile, routeFiles, S, m)
                    SecondsElapsed = Round(Timer - StartTime, 2)
                    Debug.Print "Criar e editar: " & toolTwoFile & " - Tempo: " & SecondsElapsed
                    
                    StartTime = Timer
                    Call CopyDataFromToolTwo(toolTwoFile, row)
                    SecondsElapsed = Round(Timer - StartTime, 2)
                    Debug.Print "Copiar: " & toolTwoFile & " - Tempo: " & SecondsElapsed
                    
                    
                    'Verificar qual é a melhor rota
                    Dim rowRoute As Integer
                    rowRoute = row - 7
                    Dim selectedRow As Integer
                    Dim minTarifa, bestEficiencia As Double
                    minTarifa = 999999
                    bestEficiencia = -100#
                    Dim foundTarifa As Boolean
                    foundTarifa = False
                    
                    For rowRoute = row - 7 To row - 1
                        'Se for o mercado base que estamos processando, então buscamos a melhor rota seguindo o critério
                        'que se não tiver nenhuma rota com a tarifa líquida abaixo do determinando, escolhemos a de menor,
                        'valor. Porém, se existir uma ou mais rotas com a tarifa líquida abaixo do determinado, adotamos a mais
                        'eficiente, para os demais mercados, utilazaremos a rota selecionada no mercado base.
                        If m = FOLDERBASEMARKET Then
                            If tarifaLiquidaBase > wksDefinedArrays.Cells(rowRoute, 9).value Then
                                foundTarifa = True
                                If wksDefinedArrays.Cells(rowRoute, 10) > bestEficiencia Then
                                    bestEficiencia = wksDefinedArrays.Cells(rowRoute, 10)
                                    S.vSelectedRouteRow = rowRoute
                                    S.vSelectedRoute = wksDefinedArrays.Cells(rowRoute, 4).value
                                End If
                            End If
                            
                            
                            If foundTarifa = False Then
                                If minTarifa > wksDefinedArrays.Cells(rowRoute, 9).value Then
                                    minTarifa = wksDefinedArrays.Cells(rowRoute, 9).value
                                    S.vSelectedRouteRow = rowRoute
                                    S.vSelectedRoute = wksDefinedArrays.Cells(rowRoute, 4).value
                                End If
                            End If
                            
                        Else
                            If S.vSelectedRoute = wksDefinedArrays.Cells(rowRoute, 4).value Then
                                S.vSelectedRouteRow = rowRoute
                            End If
                        End If
                        
                        If tarifaLiquidaBase > wksDefinedArrays.Cells(rowRoute, 9).value Then
                            wksDefinedArrays.Cells(rowRoute, 9).Interior.Color = ApplicationColors.bgColorValidTextBox
                        Else
                            wksDefinedArrays.Cells(rowRoute, 9).Interior.Color = ApplicationColors.bgColorInvalidTextBox
                        End If
                        If eficienciaBase < wksDefinedArrays.Cells(rowRoute, 10) Then
                            wksDefinedArrays.Cells(rowRoute, 10).Interior.Color = ApplicationColors.bgColorValidTextBox
                        Else
                            wksDefinedArrays.Cells(rowRoute, 10).Interior.Color = ApplicationColors.bgColorInvalidTextBox
                        End If
                        
                        
                    Next rowRoute
                    
                    wksDefinedArrays.Cells(row, 1).value = m
                    wksDefinedArrays.Cells(row, 2).value = A.vCode
                    wksDefinedArrays.Cells(row, 3).value = S.vCode & "(Consolidado)"
                    wksDefinedArrays.Cells(row, 4).value = wksDefinedArrays.Cells(S.vSelectedRouteRow, 4).value 'Salvar o valor da rota selecionada na coluna tecnologia
                    wksDefinedArrays.Cells(row, 5).value = GetMarketCode(m) & S.vCode
                    
                    Dim rngRow As range
                    Set rngRow = wksDefinedArrays.Rows(row)
                    rngRow.EntireRow.Interior.Color = RGB(255, 242, 204)
                    
                    For x = 6 To 65
                        wksDefinedArrays.Cells(row, x).value = wksDefinedArrays.Cells(S.vSelectedRouteRow, x).value
                    Next x
                    
                    consolidatedRows.Add row
                    
                    row = row + 1
                    
                    Set routeFiles = Nothing
                    
                    percent = processed / total
                    lblProgress.width = percent * width
                    lblValue = Round(percent * 100, 1) & "%"
                    processed = processed + 1
                    DoEvents
                
                Next S
                    
                    
                'Read data from tool 2 and insert into sheet
                wksDefinedArrays.Cells(row, 1).value = m
                wksDefinedArrays.Cells(row, 2).value = A.vCode & "(Consolidado)"
                wksDefinedArrays.Cells(row, 3).value = "NA"
                wksDefinedArrays.Cells(row, 4).value = "NA"
                wksDefinedArrays.Cells(row, 5).value = GetMarketCode(m) & A.vCode
                
                
                Set rngRow = wksDefinedArrays.Rows(row)
                rngRow.EntireRow.Font.Bold = True
                rngRow.EntireRow.Interior.Color = RGB(233, 196, 106)
                
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
    Next A
      
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    Call Database.SetDatabaseValue("SimulationStatus", colUserValue, "Sim")
    
    Me.Caption = "Concluído"
    lblFile.Visible = False
    frmStepFour.updateForm
    
    SecondsElapsedTotal = Round(Timer - StartTimeTotal, 2)
    Debug.Print "Tempo total: " & SecondsElapsedTotal
    
End Sub

Private Sub UserForm_Activate()
    Call executeSimulation
End Sub

Private Sub UserForm_Initialize()
    width = lblProgress.width
    lblProgress.width = 0
    lblFile.Visible = True
    lblProgress.BackColor = ApplicationColors.bgColorLevel2
    
    Me.Height = 122
    Me.width = 634
End Sub
