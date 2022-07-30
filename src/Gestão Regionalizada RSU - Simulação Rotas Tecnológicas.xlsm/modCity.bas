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
    lastRow = wksDatabase.Cells(Rows.count, 1).End(xlUp).row
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
End Function

Public Function readDatabaseCities()
    Dim cities As New Collection
    Dim wks As Worksheet
    Set wks = Util.GetCitiesWorksheet
    Dim lastRow As Integer
    Dim r As Integer
    lastRow = wks.Cells(Rows.count, 1).End(xlUp).row
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
    lastRow = wksDatabase.Cells(Rows.count, 1).End(xlUp).row
    
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

Public Sub calculateDistances()
    Dim wksCitiesDistance As Worksheet
    Set wksCitiesDistance = GetCitiesDistanceWorksheet
    wksCitiesDistance.Cells.Clear
    Dim CityRow, CityCol As clsCity
    
    Dim cities As New Collection
    Set cities = readSelectedCities()
    Dim distance As Double
    
    Dim row As Integer
    Dim col As Integer
    row = 1
    col = 1
    For Each CityRow In cities
        For Each CityCol In cities
            distance = modCity.GetDistanceCoord(CityRow.vLatitude, CityRow.vLongitude, CityCol.vLatitude, CityCol.vLongitude, "K")
            'Debug.Print "A distância entre " & CityRow.vCityName & " e " & CityCol.vCityName & " é: " & distance
            wksCitiesDistance.Cells(row, col).value = distance
            col = col + 1
        Next CityCol
        col = 1
        row = row + 1
    Next CityRow

End Sub

Public Function GetDistanceCoord(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double, ByVal unit As String) As Double
    Dim theta As Double: theta = lon1 - lon2
    Dim dist As Double: dist = Math.Sin(deg2rad(lat1)) * Math.Sin(deg2rad(lat2)) + Math.Cos(deg2rad(lat1)) * Math.Cos(deg2rad(lat2)) * Math.Cos(deg2rad(theta))
    dist = WorksheetFunction.Acos(dist)
    dist = rad2deg(dist)
    dist = dist * 60 * 1.1515
    If unit = "K" Then
        dist = dist * 1.609344
    ElseIf unit = "N" Then
        dist = dist * 0.8684
    End If
    GetDistanceCoord = Round(dist, 2)
End Function
 
Function deg2rad(ByVal deg As Double) As Double
    deg2rad = (deg * WorksheetFunction.Pi / 180#)
End Function
 
Function rad2deg(ByVal rad As Double) As Double
    rad2deg = rad / WorksheetFunction.Pi * 180#
End Function

Public Function updateCityValues(ByVal cities As Collection)
    Dim wks As Worksheet
    Set wks = Util.GetSelectedCitiesWorksheet
    
    Dim lastRow As Integer
    Dim r, id As Integer
    Dim c As clsCity
    lastRow = wks.Cells(Rows.count, 1).End(xlUp).row
    
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
