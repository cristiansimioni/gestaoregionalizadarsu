Attribute VB_Name = "modCity"
Public Function readCities()
    Dim cities As New Collection
    Set wksDatabase = Util.GetSelectedCitiesWorksheet
    Dim lastRow As Integer
    Dim r As Integer
    lastRow = wksDatabase.Cells(Rows.Count, 1).End(xlUp).row
    For r = 2 To lastRow
        Dim c As clsCity
        Set c = New clsCity
        c.vCityName = wksDatabase.Cells(r, 1).value
        c.vLatitude = wksDatabase.Cells(r, 2).value
        c.vLongitude = wksDatabase.Cells(r, 3).value
        c.vPopulation = wksDatabase.Cells(r, 4).value
        c.vTrash = CDbl(wksDatabase.Cells(r, 5).value)
        c.vConventionalCost = wksDatabase.Cells(r, 6).value
        c.vTransshipmentCost = wksDatabase.Cells(r, 7).value
        c.vCostPostTransshipment = wksDatabase.Cells(r, 8).value
        If wksDatabase.Cells(r, 9).value = "Sim" Then
            c.vUTVR = True
        Else
            c.vUTVR = False
        End If
        If wksDatabase.Cells(r, 10).value = "Sim" Then
            c.vExistentLandfill = True
        Else
            c.vExistentLandfill = False
        End If
        If wksDatabase.Cells(r, 11).value = "Sim" Then
            c.vPotentialLandfill = True
        Else
            c.vPotentialLandfill = False
        End If
        cities.Add c
    Next r
    Set readCities = cities
End Function

Public Sub calculateDistances()
    Set wksCitiesDistance = GetCitiesDistanceWorksheet
    wksCitiesDistance.Cells.Clear
    
    Dim cities As New Collection
    Set cities = readCities()
    Dim distance As Double
    
    Dim row As Integer
    Dim col As Integer
    row = 1
    col = 1
    For Each CityRow In cities
        For Each CityCol In cities
            distance = modCity.GetDistanceCoord(CityRow.vLatitude, CityRow.vLongitude, CityCol.vLatitude, CityCol.vLongitude, "K")
            Debug.Print "A distância entre " & CityRow.vCityName & " e " & CityCol.vCityName & " é: " & distance
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
    GetDistanceCoord = dist
End Function
 
Function deg2rad(ByVal deg As Double) As Double
    deg2rad = (deg * WorksheetFunction.Pi / 180#)
End Function
 
Function rad2deg(ByVal rad As Double) As Double
    rad2deg = rad / WorksheetFunction.Pi * 180#
End Function
