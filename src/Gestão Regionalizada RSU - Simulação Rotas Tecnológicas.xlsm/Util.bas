Attribute VB_Name = "Util"
Option Explicit

'Application
Public Const APPNAME                As String = "Gestão Regionalizada RSU - Simulação Rotas Tecnológicas: Tratamento/Disposição"
Public Const APPSHORTNAME           As String = "Gestão Regionalizada RSU"
Public Const APPSUBNAME             As String = "Simulação Rotas Tecnológicas: Tratamento/Disposição"
Public Const APPVERSION             As String = "1.0.0"
Public Const APPLASTUPDATED         As String = "14.07.2022"
Public Const APPDEVELOPERNAME       As String = "Cristian Simioni Milani"
Public Const APPDEVELOPEREMAIL      As String = "cristiansimionimilani@gmail.com"

'Folder
Public Const FOLDERASSETS           As String = "assets"
Public Const FOLDERICONS            As String = "assets\icons"
Public Const FOLDERMANUAL           As String = "assets\manual"
Public Const FOLDERTEMPLATES        As String = "templates"
Public Const FOLDERALGORITHM        As String = "Algoritmo"
Public Const FOLDERBASEMARKET       As String = "Mercado Base"
Public Const FOLDEROPTIMIZEDMARKET  As String = "Mercado Otimizado"
Public Const FOLDERLANDFILLMARKET   As String = "Mercado Aterro Existentes"

'Icons
Public Const ICONCHECK              As String = "check-icon.jpg"
Public Const ICONWARNING            As String = "error-icon.jpg"

'Messages
Public Const MSG_ATTENTION                          As String = "Atenção"
Public Const MSG_CLEAN_DATABASE                     As String = "Tem certeza que você deseja apagar tudo? Todos os dados inseridos serão apagados."
Public Const MSG_CHANGED_NOT_SAVED_TITLE            As String = "Salvar Alterações"
Public Const MSG_CHANGED_NOT_SAVED                  As String = "Você realizou alterações no formulário. Gostaria de salvar?"
Public Const MSG_INVALID_DATA_TITLE                 As String = "Dados Inválidos"
Public Const MSG_INVALID_DATA                       As String = "Um ou mais dados estão preechidos de maneira incorreta. Favor verificar!"
Public Const MSG_ALGORITHM_COMPLETE_SUCCESSFULLY    As String = "A execução do algoritmo terminou com sucesso."
Public Const MSG_ALGORITHM_COMPLETE_FAILED          As String = "A execução do algoritmo falhou."

'Look and feel
Public Enum ApplicationColors
    'Form
    '#Background
    frmBgColorLevel1 = 16777215    'RGB(255, 255, 255)
    frmBgColorLevel2 = 16777215    'RGB(255, 255, 255)
    frmBgColorLevel3 = 16777215    'RGB(255, 255, 255)
    frmBgColorLevel4 = 16777215    'RGB(255, 255, 255)
    '#Button
    bgColorLevel1 = 14602886    'RGB(134, 210, 222)
    bgColorLevel2 = 14855222    'RGB(54,  172, 226)
    bgColorLevel3 = 7220525     'RGB(45,   45, 110)
    bgColorLevel4 = 2461170     'RGB(242, 141,  37)
    fgColorLevel1 = 0           'RGB(0, 0, 0)
    fgColorLevel2 = 16777215    'RGB(255, 255, 255)
    fgColorLevel3 = 16777215    'RGB(255, 255, 255)
    fgColorLevel4 = 16777215    'RGB(255, 255, 255)
    'Text Box
    bgColorValidTextBox = 11973449 'RGB(73, 179, 182)
    bgColorInvalidTextBox = 5855743 'RGB(255, 89, 89)
    txtFgColorLevel1 = 0           'RGB(0, 0, 0)
    txtFgColorLevel2 = 0           'RGB(0, 0, 0)
    txtFgColorLevel3 = 16777215    'RGB(255, 255, 255)
    txtFgColorLevel4 = 16777215    'RGB(255, 255, 255)
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
    Set GetCitiesDistanceWorksheet = ThisWorkbook.Worksheets("Distâncias entre Municípios")
End Function

Function GetArraysWorksheet() As Worksheet
    Set GetArraysWorksheet = ThisWorkbook.Worksheets("Arranjos")
End Function

Function GetDefinedArraysWorksheet() As Worksheet
    Set GetDefinedArraysWorksheet = ThisWorkbook.Worksheets("Arranjos Consolidados")
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
        sFileName = "cidades-" & projectName & ".csv"
        Set wks = Util.GetSelectedCitiesWorksheet
    Else
        sFileName = "distancias-" & projectName & ".csv"
        Set wks = Util.GetCitiesDistanceWorksheet
    End If
    
    Dim lRow As Long
    Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = wks.Cells(Rows.Count, 1).End(xlUp).row
    
    'Find the last non-blank cell in row 1
    lCol = wks.Cells(1, Columns.Count).End(xlToLeft).column
    wks.range(wks.Cells(1, 1), wks.Cells(lRow, lCol)).Copy

    'Open a new XLS workbook, save it as the file name
    Set WB = Workbooks.Add
    With WB
        .title = "Cidades"
        .Subject = projectName
        .Sheets(1).Select
        ActiveSheet.Paste
        .SaveAs directory & "\" & sFileName, xlCSV
        .Close
    End With

    Application.DisplayAlerts = True
End Sub


Sub RunPythonScript(ByVal algPath As String, ByVal prjName As String)

'Declare Variables
Dim PythonExe, PythonScript, Params, cmd As String
Dim wsh As Object
Set wsh = VBA.CreateObject("WScript.Shell")
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1
Dim errorCode As Integer

'Provide file path to Python.exe
PythonExe = Database.GetDatabaseValue("PythonPath", colUserValue)
PythonScript = Application.ActiveWorkbook.Path & "\src\combinations\combinations.py"

Dim maxCluster, trashThreshold As Double
maxCluster = Database.GetDatabaseValue("MaxClusters", colUserValue)
trashThreshold = Database.GetDatabaseValue("TrashThreshold", colUserValue)

Params = algPath & "\cidades-" & prjName & ".csv" & _
         " " & _
         algPath & "\distancias-" & prjName & ".csv" & _
         " " & _
         maxCluster & " " & trashThreshold & _
         " " & _
         algPath & "\relatório-" & prjName & ".txt" & _
         " " & _
         algPath & "\output-" & prjName & ".csv"

cmd = "%comspec% /c " & Chr(34) & PythonExe & " " & PythonScript & " " & Params & Chr(34)
'Run the Python Script
errorCode = wsh.Run(cmd, windowStyle, waitOnReturn)

If errorCode = 0 Then
    'Insert your code here
    MsgBox "Algoritmo finalizado com sucesso!"
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

Public Function CSVImport(ByVal algPath As String, ByVal prjName As String)
    Dim ws As Worksheet, strFile As String, sPath As String

    Set ws = ActiveWorkbook.Sheets("Arranjos") 'set to current worksheet name
    ws.Rows("2:" & Rows.Count).ClearContents

    Dim line As String
    Dim arrayOfElements
    Dim element As Variant
    Dim filePath As String
    Dim ImportToRow, StartColumn, ArrayId, SubArrayId As Integer
    
    ArrayId = 0
    SubArrayId = 0
    ImportToRow = 1
    
    filePath = algPath & "\output-" & prjName & ".csv"
    Open filePath For Input As #1 ' Open file for input
        Do While Not EOF(1) ' Loop until end of file
            ImportToRow = ImportToRow + 1
            Line Input #1, line
            arrayOfElements = Split(line, ";") 'Split the line into the array.
    
            If arrayOfElements(1) = "Sumário" Then
                ArrayId = ArrayId + 1
                SubArrayId = 0
            End If
            ws.Cells(ImportToRow, 1) = ArrayId
            
            If ArrayId = 1 Then 'Centralized array
                ws.Cells(ImportToRow, 2) = "Sim"
            Else
                ws.Cells(ImportToRow, 2) = "Não"
            End If
            
            If arrayOfElements(1) = "Sumário" Then
                ws.Cells(ImportToRow, 3) = "A" & ArrayId
            Else
                ws.Cells(ImportToRow, 3) = "A" & ArrayId & "SA" & SubArrayId
            End If
            'Loop thorugh every element in the array and print to Excelfile
            StartColumn = 4
            For Each element In arrayOfElements
                ws.Cells(ImportToRow, StartColumn).value = element
                StartColumn = StartColumn + 1
                Debug.Print element
            Next
            
            SubArrayId = SubArrayId + 1
        Loop
    Close #1 ' Close file.
    
End Function


Public Function GetMarketCode(ByVal market As String)
    If market = FOLDERBASEMARKET Then
        GetMarketCode = "M1"
    ElseIf market = FOLDEROPTIMIZEDMARKET Then
        GetMarketCode = "M2"
    Else
        GetMarketCode = "M3"
    End If
End Function
