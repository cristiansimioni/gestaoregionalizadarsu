Attribute VB_Name = "modReport"
Public Sub generateReport()
    Dim prjPath As String, prjName As String, reportPath As String
    prjName = Database.GetDatabaseValue("ProjectName", colUserValue)

    'Abre o arquivo template do relatório
    Set WordObj = CreateObject("Word.Application")
    WordObj.Visible = True
    Set Report = WordObj.Documents.Open(ThisWorkbook.Path & "\" & FOLDERTEMPLATES & "\" & "Relatório.docx")
    Set WordCursor = Report.Application.Selection
    
    'Altera o nome do projeto no corpo do texto
    WordCursor.Find.Text = "#NOME_DO_PROJETO#"
    WordCursor.Find.Replacement.Text = prjName
    WordCursor.Find.Execute Replace:=wdReplaceAll
    
    'Altera a conclusão inserida pelo usuário no formulário
    WordCursor.Find.Text = "#CONCLUSÃO#"
    WordCursor.Find.Replacement.Text = Database.GetDatabaseValue("ConclusionText", colUserValue)
    WordCursor.Find.Execute Replace:=wdReplaceAll
    
    'Altera o valores inseridos pelo usuário
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.GetDatabaseWorksheet
    Dim lastRow As Integer
    Dim r As Integer
    lastRow = wksDatabase.Cells(Rows.count, DatabaseColumn.colName).End(xlUp).row
    For r = 2 To lastRow
        WordCursor.Find.Text = "#" & wksDatabase.Cells(r, DatabaseColumn.colName).value & "#"
        WordCursor.Find.Replacement.Text = wksDatabase.Cells(r, DatabaseColumn.colUserValue).value
        WordCursor.Find.Execute Replace:=wdReplaceAll
    Next r
    
    'Insere os gráficos do dashboard
    Dim wksDashboard As Worksheet
    Set wksDashboard = Util.GetDashboardWorksheet
    For Each dChart In wksDashboard.ChartObjects
        If InStr(dChart.name, "GRÁFICO_DASHBOARD") <> 0 Then
             Set WordCursor = Report.Bookmarks(dChart.name).range
             dChart.Copy
             WordCursor.PasteSpecial
        End If
    Next dChart
    
    'Preenche os valores dos arranjos selecionados
    Dim arrays As Collection
    Set arrays = readArrays
    Dim array_id As Integer
    array_id = 1
    For Each A In arrays
        If A.vSelected Then
            'Preenche o título
            WordCursor.Find.Text = "#A" & array_id & "#"
            WordCursor.Find.Replacement.Text = A.vCode
            WordCursor.Find.Execute Replace:=wdReplaceAll
            'Preenche o resultado
            Set Table = getTable("Resultado A" & array_id, Report)
            Set row = Table.Rows.Add
            row.Cells(1).range.Text = A.vTotal
            row.Cells(2).range.Text = A.vTrash
            row.Cells(3).range.Text = A.vTechnology
            row.Cells(4).range.Text = A.vInbound
            row.Cells(5).range.Text = A.vOutbound
            row.Cells(6).range.Text = A.vOutboundExistentLandfill
            'Preenche Sub arranjos
            Set Table = getTable("Sub arranjos A" & array_id, Report)
            For Each s In A.vSubArray
                Set row = Table.Rows.Add
                row.Cells(1).range.Text = s.vCode
                row.Cells(2).range.Text = s.vArrayRaw
                row.Cells(3).range.Text = s.vLandfill
                row.Cells(4).range.Text = s.vExistentLandfill
                row.Cells(5).range.Text = s.vUTVR
                row.Cells(6).range.Text = s.vTotal
                row.Cells(7).range.Text = s.vTrash
                row.Cells(8).range.Text = s.vTechnology
                row.Cells(9).range.Text = s.vInbound
                row.Cells(10).range.Text = s.vOutbound
                row.Cells(11).range.Text = s.vOutboundExistentLandfill
            Next s
            array_id = array_id + 1
        End If
    Next A
    
    'Preenche as cidades na tabela
    Set cities = modCity.readSelectedCities
    Set Table = getTable("Municípios", Report)
    For Each c In cities
        Set row = Table.Rows.Add
        row.Cells(1).range.Text = c.vCityName
        row.Cells(2).range.Text = c.vPopulation
        row.Cells(3).range.Text = c.vTrash
        row.Cells(4).range.Text = c.vConventionalCost
        row.Cells(5).range.Text = c.vTransshipmentCost
        row.Cells(6).range.Text = c.vCostPostTransshipment
        row.Cells(7).range.Text = c.vUTVRAsString
        row.Cells(8).range.Text = c.vExistentLandfillAsString
        row.Cells(9).range.Text = c.vPotentialLandfillAsString
    Next c
    
    'Alterar o nome do projeto na capa
    Report.Shapes.range(Array("Text Box 3")).Select
    Set WordCursor = Report.Application.Selection
    WordCursor.Find.Text = "#NOME_DO_PROJETO#"
    WordCursor.Find.Replacement.Text = prjName
    WordCursor.Find.Execute Replace:=wdReplaceAll
    
    'Alterar informações da ficha técnica
    Report.Shapes.range(Array("Text Box 99")).Select
    Set WordCursor = Report.Application.Selection
    With WordCursor
        .Find.Text = "#NOME_DO_PROJETO#"
        .Find.Replacement.Text = prjName
        .Find.Execute
        .Find.Text = "#DIA#"
        .Find.Replacement.Text = Format(Now(), "dd")
        .Find.Execute
        .Find.Text = "#MÊS#"
        .Find.Replacement.Text = Format(Now(), "mmm")
        .Find.Execute
        .Find.Text = "#ANO#"
        .Find.Replacement.Text = Format(Now(), "yyyy")
        .Find.Execute
    End With
    
    'Atualizar o sumário do arquivo
    Report.Fields.Update
    
    'Criar o path para salvar o relatório se o mesmo ainda não existir
    prjPath = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
    prjPath = Util.FolderCreate(prjPath, prjName)
    reportPath = Util.FolderCreate(prjPath, FOLDERREPORT)
    
    'Salva o relatório final como Word para posterior alteração
    Report.SaveAs2 reportPath & "\Relatório do Projeto " & prjName & ".docx"
    
    'Salvar o relatório final como PDF
    Report.ExportAsFixedFormat OutputFileName:= _
            reportPath & "\Relatório do Projeto " & prjName & ".pdf", _
            ExportFormat:=wdExportFormatPDF, _
            OpenAfterExport:=True, _
            OptimizeFor:=wdExportOptimizeForPrint, _
            range:=wdExportAllDocument, _
            IncludeDocProps:=True, _
            CreateBookmarks:=wdExportCreateWordBookmarks, _
            BitmapMissingFonts:=True
    
    'Fechar o template
    Report.Close SaveChanges:=wdDoNotSaveChanges
    
    'Fechar o Word
    WordObj.Quit
    Set WordObj = Nothing
    
End Sub


Public Function getTable(s As String, r As Variant) As Table
    Dim tbl As Table
    For Each tbl In r.Tables
    If tbl.title = s Then
       Set getTable = tbl
       Exit Function
    End If
    Next
End Function
