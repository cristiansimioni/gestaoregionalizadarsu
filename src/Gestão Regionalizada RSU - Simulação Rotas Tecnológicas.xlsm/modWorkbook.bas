Attribute VB_Name = "modWorkbook"
Option Explicit

Public Sub EditRouteToolData(ByVal filename, ByVal arr, ByVal market As String)
    Application.DisplayAlerts = False
    
    Dim value As Integer
    value = 200
    
    Workbooks.Open filename
    
    ' Valores sub-arranjo
    ActiveWorkbook.Sheets("R-Entrada").range("E10") = arr.vTrash
    ActiveWorkbook.Sheets("R&C-Painel de Controle").range("D84") = arr.vInbound
    ActiveWorkbook.Sheets("R&C-Painel de Controle").range("D88") = arr.vOutbound
    
    If market = FOLDERLANDFILLMARKET Then
        ActiveWorkbook.Sheets("R-Definição").range("E121") = "Existente"
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
    
    Application.DisplayAlerts = True
End Sub

Public Sub EditToolTwoData(ByVal filename)
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
    
    ActiveWorkbook.Save
    ActiveWindow.Close
End Sub
