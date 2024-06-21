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
    For Each a In arrays
        If a.vSelected Then
            'Preenche o título
            WordCursor.Find.Text = "#A" & array_id & "#"
            WordCursor.Find.Replacement.Text = a.vCode
            WordCursor.Find.Execute Replace:=wdReplaceAll
            'Preenche o resultado
            Set Table = getTable("Resultado A" & array_id, Report)
            Set row = Table.Rows.Add
            row.Cells(1).range.Text = a.vTotal
            row.Cells(2).range.Text = a.vTrash
            row.Cells(3).range.Text = a.vTechnology
            row.Cells(4).range.Text = a.vInbound
            row.Cells(5).range.Text = a.vOutbound
            row.Cells(6).range.Text = a.vOutboundExistentLandfill
            'Preenche Sub arranjos
            Set Table = getTable("Sub arranjos A" & array_id, Report)
            For Each s In a.vSubArray
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
    Next a
    
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


Public Function generatePresentation() As String
    ' Declaração das variáveis
    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object

    Dim prjPath As String, prjName, presentationPath As String
    prjName = Database.GetDatabaseValue("ProjectName", colUserValue)
    
    Dim wksInfograph As Worksheet
    Set wksInfograph = Util.GetInfographsWorksheet
    
    ' Inicializa o PowerPoint
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    
    ' Abre o arquivo PowerPoint
    Set pptPres = pptApp.Presentations.Open(ThisWorkbook.Path & "\" & FOLDERTEMPLATES & "\" & "Infográficos.pptx")
    
    ' Slide 1 - Capa - OK
    Set pptSlide = pptPres.Slides(1)
    pptSlide.Shapes("Title").TextFrame.TextRange.Text = "- " & prjName & " -"
    
    ' Slide 2 - Estático - OK
    
    ' Slide 3 - Estático - OK
    
    ' Slide 4 - Cenário
    Set pptSlide = pptPres.Slides(4)
    pptSlide.Shapes("QuantidadeMunicípios").TextFrame.TextRange.Text = wksInfograph.range("C4").value
    pptSlide.Shapes("População").TextFrame.TextRange.Text = Format(wksInfograph.range("C5").value, "#,##0")
    pptSlide.Shapes("Resíduos").TextFrame.TextRange.Text = wksInfograph.range("C6").value
    pptSlide.Shapes("TaxaReciclagemAtual").TextFrame.TextRange.Text = "< " & FormatPercent(wksInfograph.range("E47").value, 0)
    pptSlide.Shapes("IRRAtual").TextFrame.TextRange.Text = "< " & FormatPercent(wksInfograph.range("O37").value, 0)
    pptSlide.Shapes("DesvioAterroAtual").TextFrame.TextRange.Text = "< " & FormatPercent(wksInfograph.range("N37").value, 0)
    pptSlide.Shapes("EmissõesAtual").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("N99").value, 0)
    pptSlide.Shapes("EmpregosDiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("N42").value, "#,##0") & " a " & Format(wksInfograph.range("O42").value, "#,##0")
    pptSlide.Shapes("EmpregosIndiretos").TextFrame.TextRange.Text = "+ " & Format(wksInfograph.range("N43").value, "#,##0") & " a " & Format(wksInfograph.range("O43").value, "#,##0")
    pptSlide.Shapes("TaxaReciclagemFuturoPercentagem").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("N39").value, 0) & " a " & FormatPercent(wksInfograph.range("O39").value, 0)
    pptSlide.Shapes("TaxaReciclagemFuturo").TextFrame.TextRange.Text = Format(wksInfograph.range("N88").value, "#,##0") & " a " & Format(wksInfograph.range("O88").value, "#,##0") & " Kt/a"
    pptSlide.Shapes("IRRFuturo").TextFrame.TextRange.Text = "> " & FormatPercent(wksInfograph.range("N38").value, 0) & " a " & FormatPercent(wksInfograph.range("O38").value, 0)
    pptSlide.Shapes("DesvioDeAterroFuturo").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("N38").value, 0) & " a " & FormatPercent(wksInfograph.range("O38").value, 0)
    pptSlide.Shapes("EmissõesFuturo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("O101").value, 0) & " a " & FormatNumber(wksInfograph.range("N101").value, 0)
    
    ' Slide 5 - Análise de Arranjos
    Set pptSlide = pptPres.Slides(5)
    pptSlide.Shapes("QuantidadeMunicípios").TextFrame.TextRange.Text = wksInfograph.range("C4").value
    pptSlide.Shapes("QuantidadeCombinações").TextFrame.TextRange.Text = 16
    
    ' Slide 6 - Arranjo Centralizado - OK
    Set pptSlide = pptPres.Slides(6)
    pptSlide.Shapes("MunicípiosCentralizado").TextFrame.TextRange.Text = wksInfograph.range("E12").value & "."
    pptSlide.Shapes("UTVRCentralizado").TextFrame.TextRange.Text = wksInfograph.range("E20").value & "."
    pptSlide.Shapes("AterroCentralizado").TextFrame.TextRange.Text = wksInfograph.range("E16").value & "."
    pptSlide.Shapes("QuantitivoRSUCentralizado").TextFrame.TextRange.Text = wksInfograph.range("E24").value & " t/d | " & Round(wksInfograph.range("E24").value * 312 / 1000, 2) & " Kt/a"
    pptSlide.Shapes("CustoAlgoritmo").TextFrame.TextRange.Text = wksInfograph.range("E31").value & " R$/t RSU"
    
    ' Slide 7 - Arranjo com dois subarranjos - OK
    Set pptSlide = pptPres.Slides(7)
    pptSlide.Shapes("MunicípiosSub1").TextFrame.TextRange.Text = wksInfograph.range("K12").value & "."
    pptSlide.Shapes("MunicípiosSub2").TextFrame.TextRange.Text = wksInfograph.range("K13").value & "."
    pptSlide.Shapes("UTVRSub1").TextFrame.TextRange.Text = wksInfograph.range("K20").value & "."
    pptSlide.Shapes("UTVRSub2").TextFrame.TextRange.Text = wksInfograph.range("K21").value & "."
    pptSlide.Shapes("AterroSub1").TextFrame.TextRange.Text = wksInfograph.range("K16").value & "."
    pptSlide.Shapes("AterroSub2").TextFrame.TextRange.Text = wksInfograph.range("K17").value & "."
    pptSlide.Shapes("QuantitativoRSUSub1").TextFrame.TextRange.Text = wksInfograph.range("K24").value & " t/d"
    pptSlide.Shapes("QuantitativoRSUSub2").TextFrame.TextRange.Text = wksInfograph.range("K25").value & " t/d"
    pptSlide.Shapes("QuantitativoRSUTotal").TextFrame.TextRange.Text = wksInfograph.range("E24").value & " t/d | " & Round(wksInfograph.range("E24").value * 312 / 1000, 2) & " Kt/a"
    pptSlide.Shapes("CustoAlgoritmoSub1").TextFrame.TextRange.Text = wksInfograph.range("K28").value & " R$/t RSU"
    pptSlide.Shapes("CustoAlgoritmoSub2").TextFrame.TextRange.Text = wksInfograph.range("K29").value & " R$/t RSU"
    pptSlide.Shapes("CustoAlgoritmoTotal").TextFrame.TextRange.Text = wksInfograph.range("K31").value & " R$/t RSU"
    
    ' Slide 8 - Arranjo com três subarranjos - OK
    Set pptSlide = pptPres.Slides(8)
    pptSlide.Shapes("MunicípiosSub1").TextFrame.TextRange.Text = wksInfograph.range("L12").value & "."
    pptSlide.Shapes("MunicípiosSub2").TextFrame.TextRange.Text = wksInfograph.range("L13").value & "."
    pptSlide.Shapes("MunicípiosSub3").TextFrame.TextRange.Text = wksInfograph.range("L14").value & "."
    pptSlide.Shapes("UTVRSub1").TextFrame.TextRange.Text = wksInfograph.range("L20").value & "."
    pptSlide.Shapes("UTVRSub2").TextFrame.TextRange.Text = wksInfograph.range("L21").value & "."
    pptSlide.Shapes("UTVRSub3").TextFrame.TextRange.Text = wksInfograph.range("L22").value & "."
    pptSlide.Shapes("AterroSub1").TextFrame.TextRange.Text = wksInfograph.range("L16").value & "."
    pptSlide.Shapes("AterroSub2").TextFrame.TextRange.Text = wksInfograph.range("L17").value & "."
    pptSlide.Shapes("AterroSub3").TextFrame.TextRange.Text = wksInfograph.range("L18").value & "."
    pptSlide.Shapes("QuantitativoRSUSub1").TextFrame.TextRange.Text = wksInfograph.range("L24").value & " t/d"
    pptSlide.Shapes("QuantitativoRSUSub2").TextFrame.TextRange.Text = wksInfograph.range("L25").value & " t/d"
    pptSlide.Shapes("QuantitativoRSUSub3").TextFrame.TextRange.Text = wksInfograph.range("L26").value & " t/d"
    pptSlide.Shapes("QuantitativoTotalRSU").TextFrame.TextRange.Text = wksInfograph.range("E24").value & " t/d | " & Round(wksInfograph.range("E24").value * 312 / 1000, 2) & " Kt/a"
    pptSlide.Shapes("CustoAlgoritmoSub1").TextFrame.TextRange.Text = wksInfograph.range("L28").value & " R$/t RSU"
    pptSlide.Shapes("CustoAlgoritmoSub2").TextFrame.TextRange.Text = wksInfograph.range("L29").value & " R$/t RSU"
    pptSlide.Shapes("CustoAlgoritmoSub3").TextFrame.TextRange.Text = wksInfograph.range("L30").value & " R$/t RSU"
    pptSlide.Shapes("CustoAlgoritimoTotal").TextFrame.TextRange.Text = wksInfograph.range("L31").value & " R$/t RSU"
    
    ' Slide 9 - Biodigestão Energia Elétrica
    Set pptSlide = pptPres.Slides(9)
    pptSlide.Shapes("InvestimentoDireto").TextFrame.TextRange.Text = "R$ " & FormatNumber(wksInfograph.range("E97").value, 2) & " Bi"
    pptSlide.Shapes("EmpregosDiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("E42").value, "#,##0")
    pptSlide.Shapes("EmpregosIndiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("E43").value, "#,##0")
    pptSlide.Shapes("RepasseMaterialReciclável").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("E40").value, 0)
    pptSlide.Shapes("RepasseMaterialReciclávelReal").TextFrame.TextRange.Text = "(" & Format(wksInfograph.range("E41").value, "#,##0.0") & " Milhões R$/a)"
    pptSlide.Shapes("ReduçãoEmissões").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("E44").value, 0)
    pptSlide.Shapes("IRR").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("E38").value, 0)
    pptSlide.Shapes("SeparaçãoMateriais").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("E39").value, 1)
    pptSlide.Shapes("EstruturaCapital").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("E45").value, 0)
    pptSlide.Shapes("Wacc").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("E46").value, 1)
    pptSlide.Shapes("VendaMaterialRecicláveis").TextFrame.TextRange.Text = Format(wksInfograph.range("E49").value, "0")
    pptSlide.Shapes("VendaMaterialRecicláveisReal").TextFrame.TextRange.Text = Format(wksInfograph.range("E50").value, "0")
    pptSlide.Shapes("VendaCDR").TextFrame.TextRange.Text = Format(wksInfograph.range("E53").value, "0")
    pptSlide.Shapes("VendaCDRReal").TextFrame.TextRange.Text = Format(wksInfograph.range("E54").value, "0")
    pptSlide.Shapes("VendaEnergia").TextFrame.TextRange.Text = Format(wksInfograph.range("E55").value, "0")
    pptSlide.Shapes("VendaEnergiaReal").TextFrame.TextRange.Text = Format(wksInfograph.range("E56").value, "0")
    Set Chart = pptPres.Slides(9).Shapes("Gráfico").Chart
    Set ChartData = Chart.ChartData
    ChartData.Activate
    ChartData.Workbook.Application.Windows(1).Visible = False
    ChartData.Workbook.Sheets(1).range("C2").value = Format(wksInfograph.range("E62").value, "0") 'Custo Atual
    ChartData.Workbook.Sheets(1).range("C3").value = Format(wksInfograph.range("E65").value, "0") 'Custo Máximo
    ChartData.Workbook.Sheets(1).range("B4").value = Format(wksInfograph.range("E68").value, "0") 'Custo Consórcio
    ChartData.Workbook.Sheets(1).range("B5").value = Format(wksInfograph.range("E71").value, "0") 'Custo Mínimo
    ChartData.Workbook.Close True
    

    ' Slide 10 - Biodigestão Comercialização de Biometano
    Set pptSlide = pptPres.Slides(10)
    pptSlide.Shapes("InvestimentoDireto").TextFrame.TextRange.Text = "R$ " & FormatNumber(wksInfograph.range("F97").value, 2) & " Bi"
    pptSlide.Shapes("EmpregosDiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("F42").value, "#,##0")
    pptSlide.Shapes("EmpregosIndiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("F43").value, "#,##0")
    pptSlide.Shapes("RepasseMaterialReciclável").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("F40").value, 0)
    pptSlide.Shapes("RepasseMaterialReciclávelReal").TextFrame.TextRange.Text = "(" & Format(wksInfograph.range("F41").value, "#,##0.0") & " Milhões R$/a)"
    pptSlide.Shapes("ReduçãoEmissões").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("F44").value, 0)
    pptSlide.Shapes("IRR").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("F38").value, 0)
    pptSlide.Shapes("SeparaçãoMateriais").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("F39").value, 1)
    pptSlide.Shapes("EstruturaCapital").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("F45").value, 0)
    pptSlide.Shapes("Wacc").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("F46").value, 1)
    pptSlide.Shapes("VendaMaterialRecicláveis").TextFrame.TextRange.Text = Format(wksInfograph.range("F49").value, "0")
    pptSlide.Shapes("VendaMaterialRecicláveisReal").TextFrame.TextRange.Text = Format(wksInfograph.range("F50").value, "0")
    pptSlide.Shapes("VendaCDR").TextFrame.TextRange.Text = Format(wksInfograph.range("F53").value, "0")
    pptSlide.Shapes("VendaCDRReal").TextFrame.TextRange.Text = Format(wksInfograph.range("F54").value, "0")
    pptSlide.Shapes("VendaBiometano").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("F51").value, 2)
    pptSlide.Shapes("VendaBiometanoReal").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("F52").value, 2)
    Set Chart = pptPres.Slides(10).Shapes("Gráfico").Chart
    Set ChartData = Chart.ChartData
    ChartData.Activate
    ChartData.Workbook.Application.Windows(1).Visible = False
    ChartData.Workbook.Sheets(1).range("C2").value = Format(wksInfograph.range("F62").value, "0") 'Custo Atual
    ChartData.Workbook.Sheets(1).range("C3").value = Format(wksInfograph.range("F65").value, "0") 'Custo Máximo
    ChartData.Workbook.Sheets(1).range("B4").value = Format(wksInfograph.range("F68").value, "0") 'Custo Consórcio
    ChartData.Workbook.Sheets(1).range("B5").value = Format(wksInfograph.range("F71").value, "0") 'Custo Mínimo
    ChartData.Workbook.Close True
    
    ' Slide 11 - Biosecagem
    Set pptSlide = pptPres.Slides(11)
    pptSlide.Shapes("InvestimentoDireto").TextFrame.TextRange.Text = "R$ " & FormatNumber(wksInfograph.range("H97").value, 2) & " Bi"
    pptSlide.Shapes("EmpregosDiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("H42").value, "#,##0")
    pptSlide.Shapes("EmpregosIndiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("H43").value, "#,##0")
    pptSlide.Shapes("RepasseMaterialReciclável").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("H40").value, 0)
    pptSlide.Shapes("RepasseMaterialReciclávelReal").TextFrame.TextRange.Text = "(" & Format(wksInfograph.range("H41").value, "#,##0.0") & " Milhões R$/a)"
    pptSlide.Shapes("ReduçãoEmissões").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("H44").value, 0)
    pptSlide.Shapes("IRR").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("H38").value, 0)
    pptSlide.Shapes("SeparaçãoMateriais").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("H39").value, 1)
    pptSlide.Shapes("EstruturaCapital").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("H45").value, 0)
    pptSlide.Shapes("Wacc").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("H46").value, 1)
    pptSlide.Shapes("VendaMaterialRecicláveis").TextFrame.TextRange.Text = Format(wksInfograph.range("H49").value, "0")
    pptSlide.Shapes("VendaMaterialRecicláveisReal").TextFrame.TextRange.Text = Format(wksInfograph.range("H50").value, "0")
    pptSlide.Shapes("VendaCDR").TextFrame.TextRange.Text = Format(wksInfograph.range("H53").value, "0")
    pptSlide.Shapes("VendaCDRReal").TextFrame.TextRange.Text = Format(wksInfograph.range("H54").value, "0")
    Set Chart = pptPres.Slides(9).Shapes("Gráfico").Chart
    Set ChartData = Chart.ChartData
    ChartData.Activate
    ChartData.Workbook.Application.Windows(1).Visible = False
    ChartData.Workbook.Sheets(1).range("C2").value = Format(wksInfograph.range("H62").value, "0") 'Custo Atual
    ChartData.Workbook.Sheets(1).range("C3").value = Format(wksInfograph.range("H65").value, "0") 'Custo Máximo
    ChartData.Workbook.Sheets(1).range("B4").value = Format(wksInfograph.range("H68").value, "0") 'Custo Consórcio
    ChartData.Workbook.Sheets(1).range("B5").value = Format(wksInfograph.range("H71").value, "0") 'Custo Mínimo
    ChartData.Workbook.Close True
    
    ' Slide 12 - Compostagem
    Set pptSlide = pptPres.Slides(12)
    pptSlide.Shapes("InvestimentoDireto").TextFrame.TextRange.Text = "R$ " & FormatNumber(wksInfograph.range("G97").value, 2) & " Bi"
    pptSlide.Shapes("EmpregosDiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("G42").value, "#,##0")
    pptSlide.Shapes("EmpregosIndiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("G43").value, "#,##0")
    pptSlide.Shapes("RepasseMaterialReciclável").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("G40").value, 0)
    pptSlide.Shapes("RepasseMaterialReciclávelReal").TextFrame.TextRange.Text = "(" & Format(wksInfograph.range("G41").value, "#,##0.0") & " Milhões R$/a)"
    pptSlide.Shapes("ReduçãoEmissões").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("G44").value, 0)
    pptSlide.Shapes("IRR").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("G38").value, 0)
    pptSlide.Shapes("SeparaçãoMateriais").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("G39").value, 1)
    pptSlide.Shapes("EstruturaCapital").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("G45").value, 0)
    pptSlide.Shapes("Wacc").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("G46").value, 1)
    pptSlide.Shapes("VendaMaterialRecicláveis").TextFrame.TextRange.Text = Format(wksInfograph.range("G49").value, "0")
    pptSlide.Shapes("VendaMaterialRecicláveisReal").TextFrame.TextRange.Text = Format(wksInfograph.range("G50").value, "0")
    pptSlide.Shapes("VendaCDR").TextFrame.TextRange.Text = Format(wksInfograph.range("G53").value, "0")
    pptSlide.Shapes("VendaCDRReal").TextFrame.TextRange.Text = Format(wksInfograph.range("G54").value, "0")
    pptSlide.Shapes("VendaComposto").TextFrame.TextRange.Text = Format(wksInfograph.range("G59").value, "0")
    pptSlide.Shapes("VendaCompostoReal").TextFrame.TextRange.Text = Format(wksInfograph.range("G60").value, "0")
    Set Chart = pptPres.Slides(12).Shapes("Gráfico").Chart
    Set ChartData = Chart.ChartData
    ChartData.Activate
    ChartData.Workbook.Application.Windows(1).Visible = False
    ChartData.Workbook.Sheets(1).range("C2").value = Format(wksInfograph.range("G62").value, "0") 'Custo Atual
    ChartData.Workbook.Sheets(1).range("C3").value = Format(wksInfograph.range("G65").value, "0") 'Custo Máximo
    ChartData.Workbook.Sheets(1).range("B4").value = Format(wksInfograph.range("G68").value, "0") 'Custo Consórcio
    ChartData.Workbook.Sheets(1).range("B5").value = Format(wksInfograph.range("G71").value, "0") 'Custo Mínimo
    ChartData.Workbook.Close True
    
    ' Slide 13 - Incineração
    Set pptSlide = pptPres.Slides(13)
    pptSlide.Shapes("InvestimentoDireto").TextFrame.TextRange.Text = "R$ " & FormatNumber(wksInfograph.range("I97").value, 2) & " Bi"
    pptSlide.Shapes("EmpregosDiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("I42").value, "#,##0")
    pptSlide.Shapes("EmpregosIndiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("I43").value, "#,##0")
    pptSlide.Shapes("RepasseMaterialReciclável").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("I40").value, 0)
    pptSlide.Shapes("RepasseMaterialReciclávelReal").TextFrame.TextRange.Text = "(" & Format(wksInfograph.range("I41").value, "#,##0.0") & " Milhões R$/a)"
    pptSlide.Shapes("ReduçãoEmissões").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("I44").value, 0)
    pptSlide.Shapes("IRR").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("I38").value, 0)
    pptSlide.Shapes("SeparaçãoMateriais").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("I39").value, 1)
    pptSlide.Shapes("EstruturaCapital").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("I45").value, 0)
    pptSlide.Shapes("Wacc").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("I46").value, 1)
    pptSlide.Shapes("VendaMaterialRecicláveis").TextFrame.TextRange.Text = Format(wksInfograph.range("I49").value, "0")
    pptSlide.Shapes("VendaMaterialRecicláveisReal").TextFrame.TextRange.Text = Format(wksInfograph.range("I50").value, "0")
    pptSlide.Shapes("VendaEnergia").TextFrame.TextRange.Text = Format(wksInfograph.range("I57").value, "0")
    pptSlide.Shapes("VendaEnergiaReal").TextFrame.TextRange.Text = Format(wksInfograph.range("I58").value, "0")
    Set Chart = pptPres.Slides(13).Shapes("Gráfico").Chart
    Set ChartData = Chart.ChartData
    ChartData.Activate
    ChartData.Workbook.Application.Windows(1).Visible = False
    ChartData.Workbook.Sheets(1).range("C2").value = Format(wksInfograph.range("I62").value, "0") 'Custo Atual
    ChartData.Workbook.Sheets(1).range("C3").value = Format(wksInfograph.range("I65").value, "0") 'Custo Máximo
    ChartData.Workbook.Sheets(1).range("B4").value = Format(wksInfograph.range("I68").value, "0") 'Custo Consórcio
    ChartData.Workbook.Sheets(1).range("B5").value = Format(wksInfograph.range("I71").value, "0") 'Custo Mínimo
    ChartData.Workbook.Close True
    
    ' Slide 14 - Arranjo com dois subarranjos
    Set pptSlide = pptPres.Slides(14)
    pptSlide.Shapes("InvestimentoDireto").TextFrame.TextRange.Text = "R$ " & FormatNumber(wksInfograph.range("K97").value, 2) & " Bi"
    pptSlide.Shapes("UTVRSub1").TextFrame.TextRange.Text = "• " & wksInfograph.range("K20").value & " (" & wksInfograph.range("K33").value & ")"
    pptSlide.Shapes("UTVRSub2").TextFrame.TextRange.Text = "• " & wksInfograph.range("K21").value & " (" & wksInfograph.range("K34").value & ")"
    pptSlide.Shapes("Aterro").TextFrame.TextRange.Text = "• Aterro Sanitário " & wksInfograph.range("K16").value & " e " & wksInfograph.range("K17").value
    pptSlide.Shapes("QuantitativoTotalRSU").TextFrame.TextRange.Text = "• " & wksInfograph.range("E24").value & " t/d | " & Round(wksInfograph.range("E24").value * 312 / 1000, 2) & " Kt/a"
    pptSlide.Shapes("QuantitativoSub1").TextFrame.TextRange.Text = wksInfograph.range("K24").value & " t/d"
    pptSlide.Shapes("TecnologiaValorSub1").TextFrame.TextRange.Text = wksInfograph.range("K33").value & ": " & "R$ " & FormatNumber(wksInfograph.range("K103").value, 2) & " Milhões"
    pptSlide.Shapes("QuantitativoSub2").TextFrame.TextRange.Text = wksInfograph.range("K25").value & " t/d"
    pptSlide.Shapes("TecnologiaValorSub2").TextFrame.TextRange.Text = wksInfograph.range("K34").value & ": " & "R$ " & FormatNumber(wksInfograph.range("K104").value, 2) & " Milhões"
    
    pptSlide.Shapes("EmpregosDiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("K42").value, "#,##0")
    pptSlide.Shapes("EmpregosIndiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("K43").value, "#,##0")
    pptSlide.Shapes("RepasseMaterialReciclável").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("K40").value, 0)
    pptSlide.Shapes("RepasseMaterialReciclávelReal").TextFrame.TextRange.Text = "(" & Format(wksInfograph.range("K41").value, "#,##0.0") & " Milhões R$/a)"
    pptSlide.Shapes("ReduçãoEmissões").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("K44").value, 0)
    pptSlide.Shapes("IRR").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("K38").value, 0)
    pptSlide.Shapes("SeparaçãoMateriais").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("K39").value, 1)
    pptSlide.Shapes("EstruturaCapital").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("K45").value, 0)
    pptSlide.Shapes("Wacc").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("K46").value, 1)
    
    pptSlide.Shapes("Preço Médio Venda Materiais Recicláveis - Mínimo").TextFrame.TextRange.Text = Format(wksInfograph.range("K49").value, "0")
    pptSlide.Shapes("Preço Médio Venda Materiais Recicláveis - Máximo").TextFrame.TextRange.Text = Format(wksInfograph.range("K50").value, "0")
    pptSlide.Shapes("Preço Venda CDR - Mínimo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("K53").value, 0)
    pptSlide.Shapes("Preço Venda CDR - Máximo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("K54").value, 0)
    pptSlide.Shapes("Preço Venda Composto Orgânico - Mínimo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("K59").value, 0)
    pptSlide.Shapes("Preço Venda Composto Orgânico - Máximo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("K60").value, 0)
    pptSlide.Shapes("Preço Venda Biometano - Mínimo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("K51").value, 2)
    pptSlide.Shapes("Preço Venda Biometano - Máximo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("K52").value, 2)
    pptSlide.Shapes("Preço Venda Energia Elétrica Biodigestão - Mínimo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("K55").value, 0)
    pptSlide.Shapes("Preço Venda Energia Elétrica Biodigestão - Máximo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("K56").value, 0)
    pptSlide.Shapes("Preço Venda Energia Elétrica Incineração - Mínimo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("K57").value, 0)
    pptSlide.Shapes("Preço Venda Energia Elétrica Incineração - Máximo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("K58").value, 0)

    Set Chart = pptPres.Slides(14).Shapes("Gráfico").Chart
    Set ChartData = Chart.ChartData
    ChartData.Activate
    ChartData.Workbook.Application.Windows(1).Visible = False
    ChartData.Workbook.Sheets(1).range("C2").value = Format(wksInfograph.range("K62").value, "0") 'Custo Atual
    ChartData.Workbook.Sheets(1).range("C3").value = Format(wksInfograph.range("K65").value, "0") 'Custo Máximo
    ChartData.Workbook.Sheets(1).range("B4").value = Format(wksInfograph.range("K68").value, "0") 'Custo Consórcio
    ChartData.Workbook.Sheets(1).range("B5").value = Format(wksInfograph.range("K71").value, "0") 'Custo Mínimo
    ChartData.Workbook.Close True
    
    ' Slide 15 - Arranjo com três subarranjos
    Set pptSlide = pptPres.Slides(15)
    pptSlide.Shapes("InvestimentoDireto").TextFrame.TextRange.Text = "R$ " & FormatNumber(wksInfograph.range("L97").value, 2) & " Bi"
    pptSlide.Shapes("UTVRSub1").TextFrame.TextRange.Text = "• " & wksInfograph.range("L20").value & " (" & wksInfograph.range("L33").value & ")"
    pptSlide.Shapes("UTVRSub2").TextFrame.TextRange.Text = "• " & wksInfograph.range("L21").value & " (" & wksInfograph.range("L34").value & ")"
    pptSlide.Shapes("UTVRSub3").TextFrame.TextRange.Text = "• " & wksInfograph.range("L22").value & " (" & wksInfograph.range("L35").value & ")"
    pptSlide.Shapes("Aterro").TextFrame.TextRange.Text = "• Aterro Sanitário " & wksInfograph.range("L16").value & ", " & wksInfograph.range("L17").value & " e " & wksInfograph.range("L18").value
    pptSlide.Shapes("QuantitativoTotalRSU").TextFrame.TextRange.Text = "• " & wksInfograph.range("E24").value & " t/d | " & Round(wksInfograph.range("E24").value * 312 / 1000, 2) & " Kt/a"
    pptSlide.Shapes("QuantitativoSub1").TextFrame.TextRange.Text = wksInfograph.range("L24").value & " t/d"
    pptSlide.Shapes("TecnologiaValorSub1").TextFrame.TextRange.Text = wksInfograph.range("L33").value & ": " & "R$ " & FormatNumber(wksInfograph.range("L103").value, 2) & " Milhões"
    pptSlide.Shapes("QuantitativoSub2").TextFrame.TextRange.Text = wksInfograph.range("L25").value & " t/d"
    pptSlide.Shapes("TecnologiaValorSub2").TextFrame.TextRange.Text = wksInfograph.range("L34").value & ": " & "R$ " & FormatNumber(wksInfograph.range("L104").value, 2) & " Milhões"
    pptSlide.Shapes("QuantitativoSub3").TextFrame.TextRange.Text = wksInfograph.range("L26").value & " t/d"
    pptSlide.Shapes("TecnologiaValorSub3").TextFrame.TextRange.Text = wksInfograph.range("L35").value & ": " & "R$ " & FormatNumber(wksInfograph.range("L105").value, 2) & " Milhões"

    pptSlide.Shapes("EmpregosDiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("L42").value, "#,##0")
    pptSlide.Shapes("EmpregosIndiretos").TextFrame.TextRange.Text = Format(wksInfograph.range("L43").value, "#,##0")
    pptSlide.Shapes("RepasseMaterialReciclável").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("L40").value, 0)
    pptSlide.Shapes("RepasseMaterialReciclávelReal").TextFrame.TextRange.Text = "(" & Format(wksInfograph.range("L41").value, "#,##0.0") & " Milhões R$/a)"
    pptSlide.Shapes("ReduçãoEmissões").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("L44").value, 0)
    pptSlide.Shapes("IRR").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("L38").value, 0)
    pptSlide.Shapes("SeparaçãoMateriais").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("L39").value, 1)
    pptSlide.Shapes("EstruturaCapital").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("L45").value, 0)
    pptSlide.Shapes("Wacc").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("L46").value, 1)
    
    pptSlide.Shapes("Preço Médio Venda Materiais Recicláveis - Mínimo").TextFrame.TextRange.Text = Format(wksInfograph.range("L49").value, "0")
    pptSlide.Shapes("Preço Médio Venda Materiais Recicláveis - Máximo").TextFrame.TextRange.Text = Format(wksInfograph.range("L50").value, "0")
    pptSlide.Shapes("Preço Venda CDR - Mínimo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("L53").value, 0)
    pptSlide.Shapes("Preço Venda CDR - Máximo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("L54").value, 0)
    pptSlide.Shapes("Preço Venda Composto Orgânico - Mínimo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("L59").value, 0)
    pptSlide.Shapes("Preço Venda Composto Orgânico - Máximo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("L60").value, 0)
    pptSlide.Shapes("Preço Venda Biometano - Mínimo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("L51").value, 2)
    pptSlide.Shapes("Preço Venda Biometano - Máximo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("L52").value, 2)
    pptSlide.Shapes("Preço Venda Energia Elétrica Biodigestão - Mínimo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("L55").value, 0)
    pptSlide.Shapes("Preço Venda Energia Elétrica Biodigestão - Máximo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("L56").value, 0)
    pptSlide.Shapes("Preço Venda Energia Elétrica Incineração - Mínimo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("L57").value, 0)
    pptSlide.Shapes("Preço Venda Energia Elétrica Incineração - Máximo").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("L58").value, 0)

    Set Chart = pptPres.Slides(15).Shapes("Gráfico").Chart
    Set ChartData = Chart.ChartData
    ChartData.Activate
    ChartData.Workbook.Application.Windows(1).Visible = False
    ChartData.Workbook.Sheets(1).range("C2").value = Format(wksInfograph.range("L62").value, "0") 'Custo Atual
    ChartData.Workbook.Sheets(1).range("C3").value = Format(wksInfograph.range("L65").value, "0") 'Custo Máximo
    ChartData.Workbook.Sheets(1).range("B4").value = Format(wksInfograph.range("L68").value, "0") 'Custo Consórcio
    ChartData.Workbook.Sheets(1).range("B5").value = Format(wksInfograph.range("L71").value, "0") 'Custo Mínimo
    ChartData.Workbook.Close True
    
    ' Slide 16 - Comparativo de Alternativas
    Set pptSlide = pptPres.Slides(16)
    pptSlide.Shapes("TetoCusto").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("E63").value, 0)
    pptSlide.Shapes("CustoAtual").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("E62").value, 0)
    pptSlide.Shapes("População").TextFrame.TextRange.Text = Format(wksInfograph.range("C5").value, "#,##0")
    pptSlide.Shapes("Resíduos").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("E87").value, 1)
    
    pptSlide.Shapes("EficiênciaValorizaçãoBioElétrica").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("E85").value, 0)
    pptSlide.Shapes("EficiênciaValorizaçãoBioBiometano").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("F85").value, 0)
    pptSlide.Shapes("EficiênciaValorizaçãoBiosecagem").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("H85").value, 0)
    pptSlide.Shapes("EficiênciaValorizaçãoCompostagem").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("G85").value, 0)
    pptSlide.Shapes("EficiênciaValorizaçãoIncineração").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("I85").value, 0)
    pptSlide.Shapes("EficiênciaValorizaçãoSub2").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("K85").value, 0)
    pptSlide.Shapes("EficiênciaValorizaçãoSub3").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("L85").value, 0)
    
    pptSlide.Shapes("LabelA2Sub1-" & wksInfograph.range("K33").value).Visible = True
    pptSlide.Shapes("LabelA2Sub2-" & wksInfograph.range("K34").value).Visible = True
    pptSlide.Shapes("LabelA3Sub1-" & wksInfograph.range("L33").value).Visible = True
    pptSlide.Shapes("LabelA3Sub2-" & wksInfograph.range("L34").value).Visible = True
    pptSlide.Shapes("LabelA3Sub3-" & wksInfograph.range("L35").value).Visible = True
    
    Set Chart = pptPres.Slides(16).Shapes("Gráfico").Chart
    Set ChartData = Chart.ChartData
    ChartData.Activate
    ChartData.Workbook.Application.Windows(1).Visible = False
    ChartData.Workbook.Sheets(1).range("E2").value = Format(wksInfograph.range("E62").value, "0") 'Custo Atual
    ChartData.Workbook.Sheets(1).range("D2").value = Format(wksInfograph.range("E63").value, "0") 'Custo Teto
    ChartData.Workbook.Sheets(1).range("B3").value = Format(wksInfograph.range("E71").value, "0")
    ChartData.Workbook.Sheets(1).range("C3").value = Format(wksInfograph.range("E65").value, "0")
    ChartData.Workbook.Sheets(1).range("B4").value = Format(wksInfograph.range("F71").value, "0")
    ChartData.Workbook.Sheets(1).range("C4").value = Format(wksInfograph.range("F65").value, "0")
    ChartData.Workbook.Sheets(1).range("B5").value = Format(wksInfograph.range("H71").value, "0")
    ChartData.Workbook.Sheets(1).range("C5").value = Format(wksInfograph.range("H65").value, "0")
    ChartData.Workbook.Sheets(1).range("B6").value = Format(wksInfograph.range("G71").value, "0")
    ChartData.Workbook.Sheets(1).range("C6").value = Format(wksInfograph.range("G65").value, "0")
    ChartData.Workbook.Sheets(1).range("B7").value = Format(wksInfograph.range("I71").value, "0")
    ChartData.Workbook.Sheets(1).range("C7").value = Format(wksInfograph.range("I65").value, "0")
    ChartData.Workbook.Sheets(1).range("B8").value = Format(wksInfograph.range("K71").value, "0")
    ChartData.Workbook.Sheets(1).range("C8").value = Format(wksInfograph.range("K65").value, "0")
    ChartData.Workbook.Sheets(1).range("B9").value = Format(wksInfograph.range("L71").value, "0")
    ChartData.Workbook.Sheets(1).range("C9").value = Format(wksInfograph.range("L65").value, "0")
    ChartData.Workbook.Close True
    
    ' Slide 17 - Comparativo de Alternativas
    Set pptSlide = pptPres.Slides(17)
    pptSlide.Shapes("TetoCusto").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("E111").value, 1)
    pptSlide.Shapes("CustoAtual").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("E110").value, 1)
    pptSlide.Shapes("População").TextFrame.TextRange.Text = Format(wksInfograph.range("C5").value, "#,##0")
    pptSlide.Shapes("Resíduos").TextFrame.TextRange.Text = FormatNumber(wksInfograph.range("E87").value, 1)
    
    pptSlide.Shapes("EficiênciaValorizaçãoBioElétrica").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("E85").value, 0)
    pptSlide.Shapes("EficiênciaValorizaçãoBioBiometano").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("F85").value, 0)
    pptSlide.Shapes("EficiênciaValorizaçãoBiosecagem").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("H85").value, 0)
    pptSlide.Shapes("EficiênciaValorizaçãoCompostagem").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("G85").value, 0)
    pptSlide.Shapes("EficiênciaValorizaçãoIncineração").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("I85").value, 0)
    pptSlide.Shapes("EficiênciaValorizaçãoSub2").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("K85").value, 0)
    pptSlide.Shapes("EficiênciaValorizaçãoSub3").TextFrame.TextRange.Text = FormatPercent(wksInfograph.range("L85").value, 0)
    
    pptSlide.Shapes("LabelA2Sub1-" & wksInfograph.range("K33").value).Visible = True
    pptSlide.Shapes("LabelA2Sub2-" & wksInfograph.range("K34").value).Visible = True
    pptSlide.Shapes("LabelA3Sub1-" & wksInfograph.range("L33").value).Visible = True
    pptSlide.Shapes("LabelA3Sub2-" & wksInfograph.range("L34").value).Visible = True
    pptSlide.Shapes("LabelA3Sub3-" & wksInfograph.range("L35").value).Visible = True
    
    Set Chart = pptPres.Slides(17).Shapes("Gráfico").Chart
    Set ChartData = Chart.ChartData
    ChartData.Activate
    ChartData.Workbook.Application.Windows(1).Visible = False
    ChartData.Workbook.Sheets(1).range("E2").value = wksInfograph.range("E110").value
    ChartData.Workbook.Sheets(1).range("D2").value = wksInfograph.range("E111").value
    ChartData.Workbook.Sheets(1).range("B3").value = wksInfograph.range("E107").value
    ChartData.Workbook.Sheets(1).range("C3").value = wksInfograph.range("E108").value
    ChartData.Workbook.Sheets(1).range("B4").value = wksInfograph.range("F107").value
    ChartData.Workbook.Sheets(1).range("C4").value = wksInfograph.range("F108").value
    ChartData.Workbook.Sheets(1).range("B5").value = wksInfograph.range("H107").value
    ChartData.Workbook.Sheets(1).range("C5").value = wksInfograph.range("H108").value
    ChartData.Workbook.Sheets(1).range("B6").value = wksInfograph.range("G107").value
    ChartData.Workbook.Sheets(1).range("C6").value = wksInfograph.range("G108").value
    ChartData.Workbook.Sheets(1).range("B7").value = wksInfograph.range("I107").value
    ChartData.Workbook.Sheets(1).range("C7").value = wksInfograph.range("I108").value
    ChartData.Workbook.Sheets(1).range("B8").value = wksInfograph.range("K107").value
    ChartData.Workbook.Sheets(1).range("C8").value = wksInfograph.range("K108").value
    ChartData.Workbook.Sheets(1).range("B9").value = wksInfograph.range("L107").value
    ChartData.Workbook.Sheets(1).range("C9").value = wksInfograph.range("L108").value
    ChartData.Workbook.Close True
    
    
    'Criar o path para salvar o relatório se o mesmo ainda não existir
    prjPath = Database.GetDatabaseValue("ProjectPathFolder", colUserValue)
    prjPath = Util.FolderCreate(prjPath, prjName)
    presentationPath = Util.FolderCreate(prjPath, FOLDERREPORT)
    
    'Salva o relatório final como Power Point para posterior alteração
    pptPres.SaveCopyAs presentationPath & "\Infográficos do Projeto " & prjName & ".pptx"
    
    'Exportar em PDF
    presentationPath = presentationPath & "\Infográficos do Projeto " & prjName & ".pdf"
    pptPres.ExportAsFixedFormat presentationPath, _
            ppFixedFormatTypePDF, Intent:=2, PrintRange:=Nothing
    
    'Finaliza as aplicações
    pptPres.Close
    pptApp.Quit
    
    'Libera a memória
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
    
    generatePresentation = presentationPath
    
End Function

Public Function getTable(s As String, r As Variant) As Table
    Dim tbl As Table
    For Each tbl In r.Tables
    If tbl.title = s Then
       Set getTable = tbl
       Exit Function
    End If
    Next
End Function




