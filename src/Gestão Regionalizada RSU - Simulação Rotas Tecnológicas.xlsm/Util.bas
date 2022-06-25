Attribute VB_Name = "Util"
Option Explicit

Public Const APPNAME        As String = "Gestão Regionalizada RSU - Simulação Rotas Tecnológicas: Tratamento/Disposição"
Public Const APPVERSION     As String = "1.0.0"
Public Const APPLASTUPDATED As String = "25.06.2021"
Public Const APPDEVELOPER   As String = "Cristian Simioni Milani"

Public Const FOLDERALGORITHM As String = "Algoritmo"


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
    btColorLevel4 = 2461170     'RGB(242, 141, 37)
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

Sub saveAsCSV(projectName As String, directory As String, sheet As String)
    Dim sFileName As String
    Dim WB As Workbook
    Dim wks As Worksheet

    Application.DisplayAlerts = False

    
    'Copy the contents of required sheet ready to paste into the new CSV
    If sheet = "city" Then
        sFileName = "cities-" & projectName & ".csv"
        Set wks = Util.GetSelectedCitiesWorksheet
    Else
        sFileName = "distance-" & projectName & ".csv"
        Set wks = Util.GetCitiesDistanceWorksheet
    End If
    
    Dim lRow As Long
    Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = wks.Cells(Rows.Count, 1).End(xlUp).row
    
    'Find the last non-blank cell in row 1
    lCol = wks.Cells(1, Columns.Count).End(xlToLeft).column
    wks.Range(wks.Cells(1, 1), wks.Cells(lRow, lCol)).Copy

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
Dim PythonExe, PythonScript, Params, cmd As String
Dim wsh As Object
Set wsh = VBA.CreateObject("WScript.Shell")
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1
Dim errorCode As Integer

'Provide file path to Python.exe
'USE TRIPLE QUOTES WHEN FILE PATH CONTAINS SPACES.
PythonExe = """C:\Users\cristiansimioni\AppData\Local\Programs\Python\Python310\python.exe"""
PythonExe = """C:\Users\cristiansimioni\AppData\Local\Microsoft\WindowsApps\python3.exe"""
PythonScript = """C:\Users\cristiansimioni\OneDrive\Área de Trabalho\gestaoregionalizadarsu\src\combinations\combinations.py"""

Params = """C:\Users\cristiansimioni\OneDrive\Área de Trabalho\Teste\Cristian\Algoritmo\cities-Cristian.csv""" & _
         " " & _
         """C:\Users\cristiansimioni\OneDrive\Área de Trabalho\Teste\Cristian\Algoritmo\distance-Cristian.csv""" & _
         " " & _
         "10 50" & _
         " " & _
         """C:\Users\cristiansimioni\OneDrive\Área de Trabalho\Teste\Cristian\Algoritmo\alg-report.txt""" & _
         " " & _
         """C:\Users\cristiansimioni\OneDrive\Área de Trabalho\Teste\Cristian\Algoritmo\alg-out.csv"""

cmd = "%comspec% /c " & Chr(34) & PythonExe & " " & PythonScript & " " & Params & Chr(34)
'Run the Python Script
errorCode = wsh.Run(cmd, windowStyle, waitOnReturn)

If errorCode = 0 Then
    'Insert your code here
    MsgBox "Program finished successfully."
Else
    MsgBox "Program exited with error code " & errorCode & "."
End If



End Sub

Public Function FolderCreate(ByVal strPathToFolder As String, ByVal strFolder As String) As Variant
    'The function FolderCreate attemps to create the folder strFolder on the path strPathToFolder _
    ' and returns an array where the first element is a boolean indicating if the folder was created/already exists
    ' True meaning that the folder already exists or was successfully created, and False meaning that the folder _
    ' wans't created and doesn't exists
    '
    'The second element of the returned array is the Full Folder Path , meaning ex: "C:\MyExamplePath\MyCreatedFolder"
    
    Dim fso As Object
    'Dim fso As New FileSystemObject
    Dim FullDirPath As String
    Dim Length As Long
    
    'Check if the path to folder string finishes by the path separator (ex: \) ,and if not add it
    If Right(strPathToFolder, 1) <> Application.PathSeparator Then
        strPathToFolder = strPathToFolder & Application.PathSeparator
    End If
    
    'Check if the folder string starts by the path separator (ex: \) , and if it does remove it
    If Left(strFolder, 1) = Application.PathSeparator Then
        Length = Len(strFolder) - 1
        strFolder = Right(strFolder, Length)
    End If
    
    FullDirPath = strPathToFolder & strFolder
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(FullDirPath) Then
        FolderCreate = FullDirPath
    Else
        On Error GoTo ErrorHandler
        fso.CreateFolder Path:=FullDirPath
        FolderCreate = FullDirPath
        On Error GoTo 0
    End If
    
SafeExit:
        Exit Function
    
ErrorHandler:
        MsgBox prompt:="A folder could not be created for the following path: " & FullDirPath & vbCrLf & _
                "Check the path name and try again."
        FolderCreate = ""
End Function
