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
    'WordCursor.Find.Text = "#NOME_DO_PROJETO"
    'WordCursor.Find.Replacement.Text = prjName
    'WordCursor.Find.Execute Replace:=wdReplaceAll
    
    'Altera a conclusão inserida pelo usuário no formulário
    WordCursor.Find.Text = "#CONCLUSÃO"
    WordCursor.Find.Replacement.Text = Database.GetDatabaseValue("ConclusionText", colUserValue)
    WordCursor.Find.Execute Replace:=wdReplaceAll
    
    'Insere os gráficos do dashboard
    Dim wksDashboard As Worksheet
    Set wksDashboard = Util.GetDashboardWorksheet
    For Each dChart In wksDashboard.ChartObjects
        If InStr(dChart.name, "GRÁFICO_DASHBOARD") <> 0 Then
            WordCursor.Find.MatchWholeWord = True
            WordCursor.Find.Text = "#" & dChart.name
            WordCursor.Find.Replacement.Text = ""
            WordCursor.Find.Execute Replace:=wdReplaceOne
            dChart.Copy
            WordCursor.PasteSpecial
        End If
    Next dChart
    
    'Preenche as cidades na tabela
    Set cities = modCity.readSelectedCities
    Set table = Report.Tables.Item(1)
    For Each c In cities
        Set row = table.Rows.Add
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
    WordCursor.Find.Text = "#NOME_DO_PROJETO"
    WordCursor.Find.Replacement.Text = prjName
    WordCursor.Find.Execute Replace:=wdReplaceAll
    
    'Alterar informações da ficha técnica
    '1Report.Shapes.range(Array("Text Box 99")).Select
    'With WordCursor
    '    .Find.Text = "#NOME_DO_PROJETO"
    '    .Find.Replacement.Text = prjName
    '    .Find.Execute
    '    .Find.Text = "#DIA"
    '    .Find.Replacement.Text = Format(Now(), "dd")
    '    .Find.Execute
    '    .Find.Text = "#MÊs"
    '    .Find.Replacement.Text = Format(Now(), "mmm")
    '    .Find.Execute
    '    .Find.Text = "#ANO"
    '    .Find.Replacement.Text = Format(Now(), "yyyy")
    '    .Find.Execute
    'End With
    
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
