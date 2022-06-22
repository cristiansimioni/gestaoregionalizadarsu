Attribute VB_Name = "Util"
Option Explicit

Public Const APPNAME        As String = "Gestão Regionalizada RSU - Simulação Rotas Tecnológicas: Tratamento/Disposição"
Public Const APPVERSION     As String = "1.0.0"
Public Const APPLASTUPDATED As String = "26.05.2021"
Public Const APPDEVELOPER   As String = "Cristian Simioni Milani"

Public Enum ApplicationColors
    'Form
    '#Background
    bgColorLevel1 = 16777215    'RGB(255, 255, 255)
    bgColorLevel2 = 16777215    'RGB(255, 255, 255)
    bgColorLevel3 = 16777215    'RGB(255, 255, 255)
    bgColorLevel4 = 16777215    'RGB(255, 255, 255)
    '#Button
    btColorLevel1 = 14855222 '14602886 '809194      'RGB(234, 88, 12)
    btColorLevel2 = 1536493     'RGB(237, 113, 23)
    btColorLevel3 = 2461170     'RGB(242, 141, 37)
    'Text Box
    bgColorValidTextBox = 11973449 'RGB(73, 179, 182)
    bgColorInvalidTextBox = 5855743 'RGB(255, 89, 89)
End Enum

Function GetDatabaseWorksheet() As Worksheet
    Set GetDatabaseWorksheet = ThisWorkbook.Worksheets("Banco de Dados")
End Function

Function GetCitiesWorksheet() As Worksheet
    Set GetCitiesWorksheet = ThisWorkbook.Worksheets("Municípios")
End Function

Function GetSelectedCitiesWorksheet() As Worksheet
    Set GetSelectedCitiesWorksheet = ThisWorkbook.Worksheets("Municípios Selecionados")
End Function

Function GetCitiesDistanceWorksheet() As Worksheet
    Set GetCitiesDistanceWorksheet = ThisWorkbook.Worksheets("Distancias entre Municípios")
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

Sub saveAsCSV(projectName As String, directory As String)
    Dim sFileName As String
    Dim WB As Workbook

    Application.DisplayAlerts = False

    sFileName = "cities-" & projectName & ".csv"
    'Copy the contents of required sheet ready to paste into the new CSV
    Sheets("Municípios Selecionados").Range("A1:J41").Copy

    'Open a new XLS workbook, save it as the file name
    Set WB = Workbooks.Add
    With WB
        .Title = "Cidades"
        .Subject = projectName
        .Sheets(1).Select
        ActiveSheet.Paste
        .SaveAs directory & "\" & sFileName, xlCSV
        .Close
    End With

    Application.DisplayAlerts = True
End Sub


Sub RunPythonScript()

'Declare Variables
Dim objShell As Object
Dim PythonExe, PythonScript As String

'Create a new Object shell.
Set objShell = VBA.CreateObject("Wscript.Shell")

'Provide file path to Python.exe
'USE TRIPLE QUOTES WHEN FILE PATH CONTAINS SPACES.
PythonExe = """C:\Users\cristiansimioni\AppData\Local\Programs\Python\Python310\python.exe"""
PythonExe = "C:\Users\cristiansimioni\AppData\Local\Microsoft\WindowsApps\python3.exe"
PythonScript = "C:\Users\cristiansimioni\Documents\Projetos\gestaoregionalizadarsu\src\combinations\combinations.py"

'Run the Python Script
'objShell.Run PythonExe & PythonScript

End Sub
