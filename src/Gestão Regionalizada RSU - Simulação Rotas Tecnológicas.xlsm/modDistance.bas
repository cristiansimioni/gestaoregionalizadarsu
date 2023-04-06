Attribute VB_Name = "modDistance"
Option Explicit

Public Enum DistanceMethod
    euclidean
    bing
    googlemaps
End Enum

Public Function calculateDistance(method As DistanceMethod, cities As Collection, form As UserForm, Optional key As String)
    Dim wksCitiesDistance As Worksheet
    Dim CityRow, CityCol As clsCity
    Dim result As Boolean
    
    Set wksCitiesDistance = GetCitiesDistanceWorksheet
    result = True
    
    Dim row As Integer
    Dim col As Integer
    Dim distance As Double
    
    Dim total As Long
    Dim processed As Long
    Dim width As Long
    Dim percent As Double
    
    total = cities.Count * cities.Count
    width = form.lblProgress.width
    form.lblProgress.width = 0
    form.lblProgress.BackColor = ApplicationColors.bgColorLevel2
    form.txtPercent.Visible = True
    form.Repaint
    
    processed = 1
    
    row = 3
    col = 2
    For Each CityRow In cities
        For Each CityCol In cities
            If CityRow.vCityName = CityCol.vCityName Then
                distance = 0
            Else
                If method = euclidean Then
                    distance = GetDistanceCoord(CityRow.vLatitude, CityRow.vLongitude, CityCol.vLatitude, CityCol.vLongitude, "K")
                Else
                    distance = GetDistanceBing(CityRow, CityCol, key)
                End If
            End If
            wksCitiesDistance.Cells(row, col).value = distance
            col = col + 1
            
            percent = processed / total
            form.lblProgress.width = percent * width
            form.txtPercent.Text = Round(percent * 100, 0) & "%"
            processed = processed + 1
            form.Repaint
            
        Next CityCol
        col = 2
        row = row + 1
        
    Next CityRow
    
    
    calculateDistance = result

End Function

Public Function validateBingKey(key As String)
    Dim Url As String
    Dim objHTTP As Variant
    
    If Len(key) < 8 Then
        validateBingKey = False
    Else
        Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
        Url = "https://dev.virtualearth.net/REST/v1/Routes/DistanceMatrix?origins=-25.0144,-47.9341&destinations=-22.2904,51.9084&travelMode=driving&o=xml&key=" & key
        objHTTP.Open "GET", Url, False
        objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
        objHTTP.send ("")
        validateBingKey = WorksheetFunction.FilterXML(objHTTP.responseText, "//AuthenticationResultCode") = "ValidCredentials"
    End If
    
End Function

Public Function GetDistanceBing(cityA, cityB, key As String)

    Dim firstVal As String, secondVal As String, lastVal As String, start As String, dest As String, Url As String
    Dim objHTTP As Variant
    Dim resultcode As Boolean

    firstVal = "https://dev.virtualearth.net/REST/v1/Routes/DistanceMatrix?origins="
    secondVal = "&destinations="
    lastVal = "&travelMode=driving&o=xml&key=" & key

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

    start = Replace(cityA.vLatitude, ",", ".") & "," & Replace(cityA.vLongitude, ",", ".")
    dest = Replace(cityB.vLatitude, ",", ".") & "," & Replace(cityB.vLongitude, ",", ".")
    
    Url = firstVal & start & secondVal & dest & lastVal
    objHTTP.Open "GET", Url, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")
    If WorksheetFunction.FilterXML(objHTTP.responseText, "//AuthenticationResultCode") = "ValidCredentials" Then
        GetDistanceBing = Round(WorksheetFunction.FilterXML(objHTTP.responseText, "//TravelDistance"), 2)
    Else
        resultcode = False
    End If

End Function


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

Public Sub cleanDistances()
    Dim answer As Integer
    answer = MsgBox("Você tem certeza que quer apagar as distâncias já inseridas?", vbExclamation + vbYesNo + vbDefaultButton2, MSG_ATTENTION)
    If answer = vbYes Then
        Dim wksCitiesDistance As Worksheet
        Set wksCitiesDistance = Util.GetCitiesDistanceWorksheet
        wksCitiesDistance.range("B3:XFD500").ClearContents
    End If
End Sub

'Verifica se a distância entre os municípios foi preenchido de maneira adequada,
'caso contrário retorna False e a mensagem de erro (errMsg) ao encontrar o primeiro problema
Public Function checkDistances(ByRef errMsg As String)
    Dim wks As Worksheet
    Dim result As Boolean
    Dim row, column As Integer
    
    Set wks = Util.GetCitiesDistanceWorksheet
    result = True
    
    Dim cities As New Collection
    Set cities = readSelectedCities()
    
    For row = 3 To cities.Count + 2
        For column = 2 To cities.Count + 1
            Dim value As Variant
            value = wks.Cells(row, column).value
            If IsEmpty(value) = True Then
                result = False
                errMsg = "A Distância entre " & cities(row - 2).vCityName & " e " & cities(column - 1).vCityName & " não está preenchida."
                Exit Function
            End If
            If IsNumeric(value) = False Then
                result = False
                errMsg = "A Distância entre " & cities(row - 2).vCityName & " e " & cities(column - 1).vCityName & " não é um valor numérico."
                Exit Function
            End If
            If value < 0 Then
                result = False
                errMsg = "A Distância entre " & cities(row - 2).vCityName & " e " & cities(column - 1).vCityName & " é menor que 0."
                Exit Function
            End If
            Debug.Print "A Distância entre " & cities(row - 2).vCityName & " e " & cities(column - 1).vCityName & " é de " & wks.Cells(row, column).value & "km"
        Next column
    Next row
    
    checkDistances = result
End Function

