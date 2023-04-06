Attribute VB_Name = "modCity"
Option Explicit

Public Enum DatabaseCityColumn
    colUF = 1
    colUFCode = 2
    colCityCode = 3
    colIBGECode = 4
    colCityName = 5
    colPoulation = 6
    colLatitude = 7
    colLongitude = 8
End Enum

Public Enum SelectedCityColumn
    colCityName = 1
    colIBGECode = 2
    colLatitude = 3
    colLongitude = 4
    colPoulation = 5
    colTrash = 6
    colConventionalCost = 7
    colTransshipmentCost = 8
    colCostPostTransshipment = 9
    colUTVR = 10
    colExistentLandfill = 11
    colPotentialLandfill = 12
End Enum

Public Function readSelectedCities()
    Dim cities As New Collection
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.GetSelectedCitiesWorksheet
    Dim lastRow As Integer
    Dim r As Integer
    
'On Error GoTo errHandler:
    lastRow = wksDatabase.Cells(Rows.Count, 1).End(xlUp).row
    For r = 2 To lastRow
        Dim c As clsCity
        Set c = New clsCity
        c.vCityName = wksDatabase.Cells(r, 1).value
        c.vIBGECode = wksDatabase.Cells(r, 2).value
        c.vLatitude = wksDatabase.Cells(r, 3).value
        c.vLongitude = wksDatabase.Cells(r, 4).value
        c.vPopulation = wksDatabase.Cells(r, 5).value
        c.vTrash = Round(wksDatabase.Cells(r, 6).value, 2)
        c.vConventionalCost = wksDatabase.Cells(r, 7).value
        c.vTransshipmentCost = wksDatabase.Cells(r, 8).value
        c.vCostPostTransshipment = wksDatabase.Cells(r, 9).value
        If wksDatabase.Cells(r, 10).value = "Sim" Then
            c.vUTVR = True
        Else
            c.vUTVR = False
        End If
        If wksDatabase.Cells(r, 11).value = "Sim" Then
            c.vExistentLandfill = True
        Else
            c.vExistentLandfill = False
        End If
        If wksDatabase.Cells(r, 12).value = "Sim" Then
            c.vPotentialLandfill = True
        Else
            c.vPotentialLandfill = False
        End If
        cities.Add c
    Next r
    Set readSelectedCities = cities
    
'errHandler:
    'Set readSelectedCities = Nothing
    'MsgBox Err.description
    
End Function

Public Function readDatabaseCities()
    Dim cities As New Collection
    Dim wks As Worksheet
    Set wks = Util.GetCitiesWorksheet
    Dim lastRow As Integer
    Dim r As Integer
    lastRow = wks.Cells(Rows.Count, 1).End(xlUp).row
    For r = 2 To lastRow
        Dim c As clsCity
        Set c = New clsCity
        c.vUF = wks.Cells(r, DatabaseCityColumn.colUF).value
        c.vUFCode = wks.Cells(r, DatabaseCityColumn.colUFCode).value
        c.vIBGECode = wks.Cells(r, DatabaseCityColumn.colIBGECode).value
        c.vCityName = wks.Cells(r, DatabaseCityColumn.colCityName).value
        c.vPopulation = wks.Cells(r, DatabaseCityColumn.colPoulation).value
        c.vLatitude = wks.Cells(r, DatabaseCityColumn.colLatitude).value
        c.vLongitude = wks.Cells(r, DatabaseCityColumn.colLongitude).value
        cities.Add c
    Next r
    Set readDatabaseCities = cities
    
End Function

Public Function validateDatabaseCities()
    Dim cities As New Collection
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.GetSelectedCitiesWorksheet
    Dim lastRow, utvr, potentialLandfill, existentLandfill As Integer
    Dim r As Integer
    lastRow = wksDatabase.Cells(Rows.Count, 1).End(xlUp).row
    
    utvr = 0
    potentialLandfill = 0
    existentLandfill = 0
    
    validateDatabaseCities = False
    
    For r = 2 To lastRow
        Dim c As clsCity
        If wksDatabase.Cells(r, 10).value = "Sim" Then
            utvr = utvr + 1
        End If
        If wksDatabase.Cells(r, 11).value = "Sim" Then
            existentLandfill = existentLandfill + 1
        End If
        If wksDatabase.Cells(r, 12).value = "Sim" Then
            potentialLandfill = potentialLandfill + 1
        End If

    Next r
    
    If utvr >= 1 And potentialLandfill >= 1 And existentLandfill >= 1 Then
        validateDatabaseCities = True
    End If
    
End Function



Public Function updateCityValues(ByVal cities As Collection)
    Dim wks As Worksheet
    Set wks = Util.GetSelectedCitiesWorksheet
    
    Dim lastRow As Integer
    Dim r, id As Integer
    Dim c As clsCity
    lastRow = wks.Cells(Rows.Count, 1).End(xlUp).row
    
    r = 2
    For Each c In cities
        wks.Cells(r, SelectedCityColumn.colConventionalCost) = c.vConventionalCost
        wks.Cells(r, SelectedCityColumn.colTransshipmentCost) = c.vTransshipmentCost
        wks.Cells(r, SelectedCityColumn.colCostPostTransshipment) = c.vCostPostTransshipment
        If c.vUTVR Then
            wks.Cells(r, SelectedCityColumn.colUTVR) = "Sim"
        Else
            wks.Cells(r, SelectedCityColumn.colUTVR) = "Não"
        End If
        If c.vExistentLandfill Then
            wks.Cells(r, SelectedCityColumn.colExistentLandfill) = "Sim"
        Else
            wks.Cells(r, SelectedCityColumn.colExistentLandfill) = "Não"
        End If
        If c.vExistentLandfill Then
            wks.Cells(r, SelectedCityColumn.colExistentLandfill) = "Sim"
        Else
            wks.Cells(r, SelectedCityColumn.colExistentLandfill) = "Não"
        End If
        If c.vPotentialLandfill Then
            wks.Cells(r, SelectedCityColumn.colPotentialLandfill) = "Sim"
        Else
            wks.Cells(r, SelectedCityColumn.colPotentialLandfill) = "Não"
        End If
        r = r + 1
    Next c
End Function

