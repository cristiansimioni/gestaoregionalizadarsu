Attribute VB_Name = "Util"
Option Explicit

Public xApplicationName As String
Public xApplicationVersion As String
Public xApplicationLastUpdate As String

Public xColorRed
Public xColorGreen
Public xColorLevel1
Public xColorLevel2
Public xColorLevel3
Public xColorLevel4

Sub initializeDefinitions()
   xApplicationName = "Gestão Regionalizada RSU - Simulação Rotas Tecnológicas: Tratamento/Disposição"
   xApplicationVersion = "1.0.0"
   xApplicationLastUpdate = "19.05.2022"
   xColorRed = RGB(255, 89, 89)
   xColorGreen = RGB(73, 179, 182)
   xColorLevel1 = RGB(255, 242, 204)
   xColorLevel2 = RGB(255, 217, 102)
   xColorLevel3 = RGB(191, 144, 0)
   xColorLevel4 = RGB(127, 96, 0)
End Sub

Function getDatabaseWorksheet() As Worksheet
    Set getDatabaseWorksheet = ThisWorkbook.Worksheets("Banco de Dados")
End Function

Function getCitiesWorksheet() As Worksheet
    Set getCitiesWorksheet = ThisWorkbook.Worksheets("Municípios")
End Function

Function getSelectedCitiesWorksheet() As Worksheet
    Set getSelectedCitiesWorksheet = ThisWorkbook.Worksheets("Municípios Selecionados")
End Function

Function validateRange(ByVal value As String, ByVal down, ByVal up, ByRef message As String) As Boolean
    validateRange = True
    If IsNumeric(value) Then
        Dim number As Double
        number = CDbl(value)
        If number >= down And number <= up Then
            message = ""
        Else
            validateRange = False
            message = "O valor deve ser maior que " & down & " e menor que " & up
        End If
    Else
        validateRange = False
        message = "O valor deve ser numérico entre " & down & " e " & up
    End If
End Function
