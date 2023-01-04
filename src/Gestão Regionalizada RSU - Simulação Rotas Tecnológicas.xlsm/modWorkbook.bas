Attribute VB_Name = "modWorkbook"
Option Explicit

Public Sub EditRouteToolData(ByVal filename, ByVal arr, ByVal market As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(filename)

    ' Valores sub-arranjo
    wb.Sheets("R-Entrada").range("E10") = arr.vTrash
    wb.Sheets("R-Entrada").range("E8") = arr.vPopulation
    wb.Sheets("R&C-Painel de Controle").range("D84") = arr.vInbound
    wb.Sheets("R&C-Painel de Controle").range("D88") = arr.vOutbound
    
    If market = FOLDERLANDFILLMARKET Then
        wb.Sheets("R-Defini豫o").range("E121") = "Existente"
        wb.Sheets("R&C-Painel de Controle").range("D88") = arr.vOutboundExistentLandfill
    End If
    
    'Valores da ferramenta
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.GetDatabaseWorksheet
    Dim lastRow As Integer
    Dim r As Integer
    lastRow = wksDatabase.Cells(Rows.count, DatabaseColumn.colName).End(xlUp).row
    For r = 2 To lastRow
        If wksDatabase.Cells(r, DatabaseColumn.colWorkbook).value = "Ferramenta 1" Then
            Dim var, sheet, range, unit As String
            var = wksDatabase.Cells(r, DatabaseColumn.colName)
            sheet = Database.GetDatabaseValue(var, colSheet)
            range = Database.GetDatabaseValue(var, colCell)
            unit = Database.GetDatabaseValue(var, colUnit)
            If unit = "%" Then
                wb.Sheets(sheet).range(range) = Database.GetDatabaseValue(var, colUserValue) / 100#
            Else
                wb.Sheets(sheet).range(range) = Database.GetDatabaseValue(var, colUserValue)
            End If
        End If
    Next r
    
    Dim MacroName As String
    MacroName = "calculateRoutes"
    Dim varProject As String
    Dim targetProject As Double
    Dim varShareholder As String
    Dim targetShareholder As Double
    
    varProject = Database.GetDatabaseValue("VariableProject", colUserValue)
    targetProject = Database.GetDatabaseValue("TargetProject", colUserValue)
    varShareholder = Database.GetDatabaseValue("VariableShareholder", colUserValue)
    targetShareholder = Database.GetDatabaseValue("TargetShareholder", colUserValue)
    
    Run "'" & filename & "'!" & MacroName, varProject, targetProject, varShareholder, targetShareholder
    
    wb.Save
    wb.Close
End Sub

Public Sub EditToolTwoData(ByVal filename, ByVal routeFiles, ByVal arr, ByVal market As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(filename)
    
    ' Valores sub-arranjo
    wb.Sheets("RESUMO GERAL Valoriz. RT큦").range("C30") = arr.vInbound
    wb.Sheets("RESUMO GERAL Valoriz. RT큦").range("C31") = arr.vOutbound
    
    If market = FOLDERLANDFILLMARKET Then
        wb.Sheets("RESUMO GERAL Valoriz. RT큦").range("C31") = arr.vOutboundExistentLandfill
    End If
    
    'Valores da ferramenta
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.GetDatabaseWorksheet
    Dim lastRow As Integer
    Dim r As Integer
    lastRow = wksDatabase.Cells(Rows.count, DatabaseColumn.colName).End(xlUp).row
    For r = 2 To lastRow
        If wksDatabase.Cells(r, DatabaseColumn.colWorkbook).value = "Ferramenta 2" Then
            Dim var, sheet, range, unit, description As String
            var = wksDatabase.Cells(r, DatabaseColumn.colName)
            sheet = Database.GetDatabaseValue(var, colSheet)
            range = Database.GetDatabaseValue(var, colCell)
            unit = Database.GetDatabaseValue(var, colUnit)
            description = Database.GetDatabaseValue(var, colDescription)
            If InStr(description, "- Base") = 0 And InStr(description, "- Otimizado") = 0 Then
                If unit = "%" Then
                    wb.Sheets(sheet).range(range) = Database.GetDatabaseValue(var, colUserValue) / 100#
                Else
                    wb.Sheets(sheet).range(range) = Database.GetDatabaseValue(var, colUserValue)
                End If
            ElseIf InStr(description, "- Otimizado") > 0 And market = FOLDEROPTIMIZEDMARKET Then
                If unit = "%" Then
                    wb.Sheets(sheet).range(range) = Database.GetDatabaseValue(var, colUserValue) / 100#
                Else
                    wb.Sheets(sheet).range(range) = Database.GetDatabaseValue(var, colUserValue)
                End If
            ElseIf InStr(description, "- Base") > 0 And (market = FOLDERBASEMARKET Or market = FOLDERLANDFILLMARKET) Then
                If unit = "%" Then
                    wb.Sheets(sheet).range(range) = Database.GetDatabaseValue(var, colUserValue) / 100#
                Else
                    wb.Sheets(sheet).range(range) = Database.GetDatabaseValue(var, colUserValue)
                End If
            End If
        End If
    Next r
    
    Dim MacroName As String
    MacroName = "updateRoutesData"
    Run "'" & filename & "'!" & MacroName, routeFiles(1), routeFiles(2), routeFiles(3), routeFiles(4), routeFiles(5)
    
    wb.Save
    wb.Close
End Sub

Public Sub CopyDataFromToolTwo(ByVal filename, ByVal row)
    Dim wbk As Workbook
    Dim wks As Worksheet
    Set wbk = Workbooks.Open(filename:=filename, ReadOnly:=True)
    
    Dim route1ARow, route1BRow, route1CRow, route2Row, route3Row, route4Row, route5Row As Integer
    Dim rowStartToolTwo, rowLastToolTwo, colStartTool As Integer
    
    route1ARow = row - 7
    route1BRow = row - 6
    route1CRow = row - 5
    route2Row = row - 4
    route3Row = row - 3
    route4Row = row - 2
    route5Row = row - 1
    
    colStartTool = 6
    rowLastToolTwo = 68
    
    Dim tarifaLiquida, eficiencia As Double
    tarifaLiquida = Database.GetDatabaseValue("TargetExpectation", colUserValue)
    eficiencia = Database.GetDatabaseValue("ValuationEfficiency", colUserValue) / 100
    
    Set wks = GetDefinedArraysWorksheet
    For rowStartToolTwo = 9 To rowLastToolTwo
        
            
        wks.Cells(route1ARow, colStartTool) = wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 3).value
        wks.Cells(route1BRow, colStartTool) = wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 4).value
        wks.Cells(route1CRow, colStartTool) = wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 5).value
        wks.Cells(route2Row, colStartTool) = wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 6).value
        wks.Cells(route3Row, colStartTool) = wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 7).value
        wks.Cells(route4Row, colStartTool) = wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 8).value
        wks.Cells(route5Row, colStartTool) = wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 9).value
        
        
        If rowStartToolTwo = 12 And colStartTool = 6 And wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 3).value < tarifaLiquida Then
            wks.Cells(route1ARow, colStartTool).Interior.Color = ApplicationColors.bgColorValidTextBox
        ElseIf rowStartToolTwo = 12 And colStartTool = 6 Then
            wks.Cells(route1ARow, colStartTool).Interior.Color = ApplicationColors.bgColorInvalidTextBox
        End If
        If rowStartToolTwo = 13 And colStartTool = 7 And wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 3).value > eficiencia Then
            wks.Cells(route1ARow, colStartTool).Interior.Color = ApplicationColors.bgColorValidTextBox
        ElseIf rowStartToolTwo = 13 And colStartTool = 7 Then
            wks.Cells(route1ARow, colStartTool).Interior.Color = ApplicationColors.bgColorInvalidTextBox
        End If
        
        If rowStartToolTwo = 12 And colStartTool = 6 And wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 4).value < tarifaLiquida Then
            wks.Cells(route1BRow, colStartTool).Interior.Color = ApplicationColors.bgColorValidTextBox
        ElseIf rowStartToolTwo = 12 And colStartTool = 6 Then
            wks.Cells(route1BRow, colStartTool).Interior.Color = ApplicationColors.bgColorInvalidTextBox
        End If
        If rowStartToolTwo = 13 And colStartTool = 7 And wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 4).value > eficiencia Then
            wks.Cells(route1BRow, colStartTool).Interior.Color = ApplicationColors.bgColorValidTextBox
        ElseIf rowStartToolTwo = 13 And colStartTool = 7 Then
            wks.Cells(route1BRow, colStartTool).Interior.Color = ApplicationColors.bgColorInvalidTextBox
        End If
        
        If rowStartToolTwo = 12 And colStartTool = 6 And wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 5).value < tarifaLiquida Then
            wks.Cells(route1CRow, colStartTool).Interior.Color = ApplicationColors.bgColorValidTextBox
        ElseIf rowStartToolTwo = 12 And colStartTool = 6 Then
            wks.Cells(route1CRow, colStartTool).Interior.Color = ApplicationColors.bgColorInvalidTextBox
        End If
        If rowStartToolTwo = 13 And colStartTool = 7 And wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 5).value > eficiencia Then
            wks.Cells(route1CRow, colStartTool).Interior.Color = ApplicationColors.bgColorValidTextBox
        ElseIf rowStartToolTwo = 13 And colStartTool = 7 Then
            wks.Cells(route1CRow, colStartTool).Interior.Color = ApplicationColors.bgColorInvalidTextBox
        End If
        
        If rowStartToolTwo = 12 And colStartTool = 6 And wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 6).value < tarifaLiquida Then
            wks.Cells(route2Row, colStartTool).Interior.Color = ApplicationColors.bgColorValidTextBox
        ElseIf rowStartToolTwo = 12 And colStartTool = 6 Then
            wks.Cells(route2Row, colStartTool).Interior.Color = ApplicationColors.bgColorInvalidTextBox
        End If
        If rowStartToolTwo = 13 And colStartTool = 7 And wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 6).value > eficiencia Then
            wks.Cells(route2Row, colStartTool).Interior.Color = ApplicationColors.bgColorValidTextBox
        ElseIf rowStartToolTwo = 13 And colStartTool = 7 Then
            wks.Cells(route2Row, colStartTool).Interior.Color = ApplicationColors.bgColorInvalidTextBox
        End If
        
        If rowStartToolTwo = 12 And colStartTool = 6 And wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 7).value < tarifaLiquida Then
            wks.Cells(route3Row, colStartTool).Interior.Color = ApplicationColors.bgColorValidTextBox
        ElseIf rowStartToolTwo = 12 And colStartTool = 6 Then
            wks.Cells(route3Row, colStartTool).Interior.Color = ApplicationColors.bgColorInvalidTextBox
        End If
        If rowStartToolTwo = 13 And colStartTool = 7 And wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 7).value > eficiencia Then
            wks.Cells(route3Row, colStartTool).Interior.Color = ApplicationColors.bgColorValidTextBox
        ElseIf rowStartToolTwo = 13 And colStartTool = 7 Then
            wks.Cells(route3Row, colStartTool).Interior.Color = ApplicationColors.bgColorInvalidTextBox
        End If
        
        If rowStartToolTwo = 12 And colStartTool = 6 And wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 8).value < tarifaLiquida Then
            wks.Cells(route4Row, colStartTool).Interior.Color = ApplicationColors.bgColorValidTextBox
        ElseIf rowStartToolTwo = 12 And colStartTool = 6 Then
            wks.Cells(route4Row, colStartTool).Interior.Color = ApplicationColors.bgColorInvalidTextBox
        End If
        If rowStartToolTwo = 13 And colStartTool = 7 And wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 8).value > eficiencia Then
            wks.Cells(route4Row, colStartTool).Interior.Color = ApplicationColors.bgColorValidTextBox
        ElseIf rowStartToolTwo = 13 And colStartTool = 7 Then
            wks.Cells(route4Row, colStartTool).Interior.Color = ApplicationColors.bgColorInvalidTextBox
        End If
        
        If rowStartToolTwo = 12 And colStartTool = 6 And wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 9).value < tarifaLiquida Then
            wks.Cells(route5Row, colStartTool).Interior.Color = ApplicationColors.bgColorValidTextBox
        ElseIf rowStartToolTwo = 12 And colStartTool = 6 Then
            wks.Cells(route5Row, colStartTool).Interior.Color = ApplicationColors.bgColorInvalidTextBox
        End If
        If rowStartToolTwo = 13 And colStartTool = 7 And wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 9).value > eficiencia Then
            wks.Cells(route5Row, colStartTool).Interior.Color = ApplicationColors.bgColorValidTextBox
        ElseIf rowStartToolTwo = 13 And colStartTool = 7 Then
            wks.Cells(route5Row, colStartTool).Interior.Color = ApplicationColors.bgColorInvalidTextBox
        End If
        
        
        colStartTool = colStartTool + 1
    Next
    
    wbk.Close
End Sub
