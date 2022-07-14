Attribute VB_Name = "modWorkbook"
Option Explicit

Public Sub EditRouteToolData(ByVal filename, ByVal arr, ByVal market As String)
    Workbooks.Open filename
    
    ' Valores sub-arranjo
    ActiveWorkbook.Sheets("R-Entrada").range("E10") = arr.vTrash
    ActiveWorkbook.Sheets("R&C-Painel de Controle").range("D84") = arr.vInbound
    ActiveWorkbook.Sheets("R&C-Painel de Controle").range("D88") = arr.vOutbound
    
    If market = FOLDERLANDFILLMARKET Then
        ActiveWorkbook.Sheets("R-Defini豫o").range("E121") = "Existente"
    End If
    
    'Valores da ferramenta
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.GetDatabaseWorksheet
    Dim lastRow As Integer
    Dim r As Integer
    lastRow = wksDatabase.Cells(Rows.Count, DatabaseColumn.colName).End(xlUp).row
    For r = 2 To lastRow
        If wksDatabase.Cells(r, DatabaseColumn.colWorkbook).value = "Ferramenta 1" Then
            Dim var, sheet, range, unit As String
            var = wksDatabase.Cells(r, DatabaseColumn.colName)
            sheet = Database.GetDatabaseValue(var, colSheet)
            range = Database.GetDatabaseValue(var, colCell)
            unit = Database.GetDatabaseValue(var, colUnit)
            If unit = "%" Then
                ActiveWorkbook.Sheets(sheet).range(range) = Database.GetDatabaseValue(var, colUserValue) / 100#
            Else
                ActiveWorkbook.Sheets(sheet).range(range) = Database.GetDatabaseValue(var, colUserValue)
            End If
        End If
    Next r
    
    
    ActiveWorkbook.Save
    ActiveWindow.Close
End Sub

Public Sub EditToolTwoData(ByVal filename, ByVal routeFiles)
    Workbooks.Open filename
    
    'Valores da ferramenta
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.GetDatabaseWorksheet
    Dim lastRow As Integer
    Dim r As Integer
    lastRow = wksDatabase.Cells(Rows.Count, DatabaseColumn.colName).End(xlUp).row
    For r = 2 To lastRow
        If wksDatabase.Cells(r, DatabaseColumn.colWorkbook).value = "Ferramenta 2" Then
            Dim var, sheet, range, unit As String
            var = wksDatabase.Cells(r, DatabaseColumn.colName)
            sheet = Database.GetDatabaseValue(var, colSheet)
            range = Database.GetDatabaseValue(var, colCell)
            unit = Database.GetDatabaseValue(var, colUnit)
            If unit = "%" Then
                ActiveWorkbook.Sheets(sheet).range(range) = Database.GetDatabaseValue(var, colUserValue) / 100#
            Else
                ActiveWorkbook.Sheets(sheet).range(range) = Database.GetDatabaseValue(var, colUserValue)
            End If
        End If
    Next r
    
    Dim MacroName As String
    MacroName = "updateRoutesData"
    Run "'" & filename & "'!" & MacroName, routeFiles(1), routeFiles(2), routeFiles(3), routeFiles(4)
    
    ActiveWorkbook.Save
    ActiveWindow.Close
End Sub

Public Sub CopyDataFromToolTwo(ByVal filename, ByVal row)
    Dim wbk As Workbook
    Dim wks As Worksheet
    Set wbk = Workbooks.Open(filename:=filename, ReadOnly:=True)
    
    Dim route1ARow, route1BRow, route1CRow, route2Row, route3Row, route4Row, route5Row As Integer
    Dim rowStartToolTwo, rowLastToolTwo, colStartTool As Integer
    
    route1ARow = row - 8
    route1BRow = row - 7
    route1CRow = row - 6
    route2Row = row - 5
    route3Row = row - 4
    route4Row = row - 3
    route5Row = row - 2
    
    colStartTool = 6
    rowLastToolTwo = 65
    
    Set wks = GetDefinedArraysWorksheet
    For rowStartToolTwo = 9 To rowLastToolTwo
        wks.Cells(route1ARow, colStartTool) = wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 3).value
        wks.Cells(route1BRow, colStartTool) = wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 4).value
        wks.Cells(route1CRow, colStartTool) = wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 5).value
        wks.Cells(route2Row, colStartTool) = wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 6).value
        wks.Cells(route3Row, colStartTool) = wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 7).value
        wks.Cells(route4Row, colStartTool) = wbk.Worksheets("RESUMO GERAL Valoriz. RT큦").Cells(rowStartToolTwo, 8).value
        colStartTool = colStartTool + 1
    Next
    
    wbk.Close
End Sub
