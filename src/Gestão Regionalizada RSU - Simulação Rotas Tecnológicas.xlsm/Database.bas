Attribute VB_Name = "Database"
Option Explicit

Public wksDatabase As Worksheet
Public locProjectName As String
Public locProjectPathFolder As String
Public locGenerationPerCapitaRDO As String
Public locIndexSelectiveColletionRSU As String
Public locAnnualGrowthPopulation As String
Public locAnnualGrowthCollect As String
Public locCOEmission As String
Public locAverageCostTransportation As String
Public locReducingCostMovimentation As String
Public locFoodWaste As String
Public locGreenWaste As String
Public locPaper As String
Public locPlasticFilm As String
Public locHardPlastics As String
Public locGlass As String
Public locFerrousMetals As String
Public locNonFerrousMetals As String
Public locTextiles As String
Public locRubber As String
Public locDiapers As String
Public locWood As String
Public locMineralResidues As String
Public locOthers As String
Public locLandfillDeviationTarget As String
Public locExpectedDeadline As String
Public locMixedRecyclingIndex As String
Public locTargetExpectation As String
Public locValuationEfficiency As String



Sub initializeDB()
    Set wksDatabase = Util.getDatabaseWorksheet
    'Projeto
    locProjectName = "F2"
    locProjectPathFolder = "F3"
    'Definição do Estudo de Caso
    locGenerationPerCapitaRDO = "F4"
    locIndexSelectiveColletionRSU = "F5"
    locAnnualGrowthPopulation = "F6"
    locAnnualGrowthCollect = "F7"
    locCOEmission = "F8"
    locAverageCostTransportation = "F9"
    locReducingCostMovimentation = "F10"
    'Gravimetria do RSU
    locFoodWaste = "F11"
    locGreenWaste = "F12"
    locPaper = "F13"
    locPlasticFilm = "F14"
    locHardPlastics = "F15"
    locGlass = "F16"
    locFerrousMetals = "F17"
    locNonFerrousMetals = "F18"
    locTextiles = "F19"
    locRubber = "F20"
    locDiapers = "F21"
    locWood = "F22"
    locMineralResidues = "F23"
    locOthers = "F24"
    'Simulação do Estudo de Caso
    locLandfillDeviationTarget = "F25"
    locExpectedDeadline = "F26"
    locMixedRecyclingIndex = "F27"
    locTargetExpectation = "F28"
    locValuationEfficiency = "F29"
End Sub

Sub cleanDB()
    setProjectName ("")
    setProjectPathFolder ("")
End Sub

'Projeto
Function setProjectName(projectName As String)
    wksDatabase.Range(locProjectName).value = projectName
End Function
Function setProjectPathFolder(projectPathFolder As String)
    wksDatabase.Range(locProjectPathFolder).value = projectPathFolder
End Function

Function getProjectName() As String
    getProjectName = wksDatabase.Range(locProjectName).value
End Function
Function getProjectPathFolder() As String
    getProjectPathFolder = wksDatabase.Range(locProjectPathFolder).value
End Function

'Definição do Estudo de Caso
Function setGenerationPerCapitaRDO(g As Double)
    wksDatabase.Range(locGenerationPerCapitaRDO).value = g
End Function
Function setIndexSelectiveColletionRSU(i As Double)
    wksDatabase.Range(locIndexSelectiveColletionRSU).value = i
End Function
Function setAnnualGrowthPopulation(a As Double)
    wksDatabase.Range(locAnnualGrowthPopulation).value = a
End Function
Function setAnnualGrowthCollect(a As Double)
    wksDatabase.Range(locAnnualGrowthCollect).value = a
End Function
Function setCOEmission(c As Double)
    wksDatabase.Range(locCOEmission).value = c
End Function
Function setAverageCostTransportation(a As Double)
    wksDatabase.Range(locAverageCostTransportation).value = a
End Function
Function setReducingCostMovimentation(r As Double)
    wksDatabase.Range(locReducingCostMovimentation).value = r
End Function
Function getGenerationPerCapitaRDO() As Double
    getGenerationPerCapitaRDO = wksDatabase.Range(locGenerationPerCapitaRDO).value
End Function
Function getIndexSelectiveColletionRSU() As Double
    getIndexSelectiveColletionRSU = wksDatabase.Range(locIndexSelectiveColletionRSU).value
End Function
Function getAnnualGrowthPopulation() As Double
    getAnnualGrowthPopulation = wksDatabase.Range(locAnnualGrowthPopulation).value
End Function
Function getAnnualGrowthCollect() As Double
    getAnnualGrowthCollect = wksDatabase.Range(locAnnualGrowthCollect).value
End Function
Function getCOEmission() As Double
    getCOEmission = wksDatabase.Range(locCOEmission).value
End Function
Function getAverageCostTransportation() As Double
    getAverageCostTransportation = wksDatabase.Range(locAverageCostTransportation).value
End Function
Function getReducingCostMovimentation() As Double
    getReducingCostMovimentation = wksDatabase.Range(locReducingCostMovimentation).value
End Function

'Gravimetria do RSU
Function setFoodWaste(value As Double)
    wksDatabase.Range(locFoodWaste).value = value
End Function
Function setGreenWaste(value As Double)
    wksDatabase.Range(locGreenWaste).value = value
End Function
Function setPaper(value As Double)
    wksDatabase.Range(locPaper).value = value
End Function
Function setPlasticFilm(value As Double)
    wksDatabase.Range(locPlasticFilm).value = value
End Function
Function setHardPlastics(value As Double)
    wksDatabase.Range(locHardPlastics).value = value
End Function
Function setGlass(value As Double)
    wksDatabase.Range(locGlass).value = value
End Function
Function setFerrousMetals(value As Double)
    wksDatabase.Range(locFerrousMetals).value = value
End Function
Function setNonFerrousMetals(value As Double)
    wksDatabase.Range(locNonFerrousMetals).value = value
End Function
Function setTextiles(value As Double)
    wksDatabase.Range(locTextiles).value = value
End Function
Function setRubber(value As Double)
    wksDatabase.Range(locRubber).value = value
End Function
Function setDiapers(value As Double)
    wksDatabase.Range(locDiapers).value = value
End Function
Function setWood(value As Double)
    wksDatabase.Range(locWood).value = value
End Function
Function setMineralResidues(value As Double)
    wksDatabase.Range(locMineralResidues).value = value
End Function
Function setOthers(value As Double)
    wksDatabase.Range(locOthers).value = value
End Function


Function getFoodWaste(Optional default As Boolean) As Double
    Dim col As Range
    Set col = wksDatabase.Range(locFoodWaste)
    If default = True Then
        Set col = col.Offset(0, 1)
    End If
    getFoodWaste = col.value
End Function
Function getGreenWaste(Optional default As Boolean) As Double
    Dim col As Range
    Set col = wksDatabase.Range(locGreenWaste)
    If default = True Then
        Set col = col.Offset(0, 1)
    End If
    getGreenWaste = col.value
End Function
Function getPaper(Optional default As Boolean) As Double
    Dim col As Range
    Set col = wksDatabase.Range(locPaper)
    If default = True Then
        Set col = col.Offset(0, 1)
    End If
    getPaper = col.value
End Function
Function getPlasticFilm(Optional default As Boolean) As Double
    Dim col As Range
    Set col = wksDatabase.Range(locPlasticFilm)
    If default = True Then
        Set col = col.Offset(0, 1)
    End If
    getPlasticFilm = col.value
End Function
Function getHardPlastics(Optional default As Boolean) As Double
    Dim col As Range
    Set col = wksDatabase.Range(locHardPlastics)
    If default = True Then
        Set col = col.Offset(0, 1)
    End If
    getHardPlastics = col.value
End Function
Function getGlass(Optional default As Boolean) As Double
    Dim col As Range
    Set col = wksDatabase.Range(locGlass)
    If default = True Then
        Set col = col.Offset(0, 1)
    End If
    getGlass = col.value
End Function
Function getFerrousMetals(Optional default As Boolean) As Double
    Dim col As Range
    Set col = wksDatabase.Range(locFerrousMetals)
    If default = True Then
        Set col = col.Offset(0, 1)
    End If
    getFerrousMetals = col.value
End Function
Function getNonFerrousMetals(Optional default As Boolean) As Double
    Dim col As Range
    Set col = wksDatabase.Range(locNonFerrousMetals)
    If default = True Then
        Set col = col.Offset(0, 1)
    End If
    getNonFerrousMetals = col.value
End Function
Function getTextiles(Optional default As Boolean) As Double
    Dim col As Range
    Set col = wksDatabase.Range(locTextiles)
    If default = True Then
        Set col = col.Offset(0, 1)
    End If
    getTextiles = col.value
End Function
Function getRubber(Optional default As Boolean) As Double
    Dim col As Range
    Set col = wksDatabase.Range(locRubber)
    If default = True Then
        Set col = col.Offset(0, 1)
    End If
    getRubber = col.value
End Function
Function getDiapers(Optional default As Boolean) As Double
    Dim col As Range
    Set col = wksDatabase.Range(locDiapers)
    If default = True Then
        Set col = col.Offset(0, 1)
    End If
    getDiapers = col.value
End Function
Function getWood(Optional default As Boolean) As Double
    Dim col As Range
    Set col = wksDatabase.Range(locWood)
    If default = True Then
        Set col = col.Offset(0, 1)
    End If
    getWood = col.value
End Function
Function getMineralResidues(Optional default As Boolean) As Double
    Dim col As Range
    Set col = wksDatabase.Range(locMineralResidues)
    If default = True Then
        Set col = col.Offset(0, 1)
    End If
    getMineralResidues = col.value
End Function
Function getOthers(Optional default As Boolean) As Double
    Dim col As Range
    Set col = wksDatabase.Range(locOthers)
    If default = True Then
        Set col = col.Offset(0, 1)
    End If
    getOthers = col.value
End Function

'Simulação do Estudo de Caso
Function setLandfillDeviationTarget(value As Double)
    wksDatabase.Range(locLandfillDeviationTarget).value = value
End Function

Function setExpectedDeadline(value As Double)
    wksDatabase.Range(locExpectedDeadline).value = value
End Function

Function setMixedRecyclingIndex(value As Double)
    wksDatabase.Range(locMixedRecyclingIndex).value = value
End Function

Function setTargetExpectation(value As Double)
    wksDatabase.Range(locTargetExpectation).value = value
End Function

Function setValuationEfficiency(value As Double)
    wksDatabase.Range(locValuationEfficiency).value = value
End Function

Function getLandfillDeviationTarget() As Double
    getLandfillDeviationTarget = wksDatabase.Range(locLandfillDeviationTarget).value
End Function

Function getExpectedDeadline() As Double
    getExpectedDeadline = wksDatabase.Range(locExpectedDeadline).value
End Function
Function getMixedRecyclingIndex() As Double
    getMixedRecyclingIndex = wksDatabase.Range(locMixedRecyclingIndex).value
End Function

Function getTargetExpectation() As Double
    getTargetExpectation = wksDatabase.Range(locTargetExpectation).value
End Function

Function getValuationEfficiency() As Double
    getValuationEfficiency = wksDatabase.Range(locValuationEfficiency).value
End Function
    
