Attribute VB_Name = "modTool"
Sub CopiarDistanciasTeste()
Attribute CopiarDistanciasTeste.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CopiarDistanciasTeste Macro
'

'
    Windows( _
        "Gest�o Regionalizada RSU - Simula��o Rotas Tecnol�gicas - GRANFPOLIS - 3.2.2.xlsm" _
        ).Activate
    Sheets("Dist�ncias entre Munic�pios").Select
    range("A1").Select
    range(Selection, Selection.End(xlToRight)).Select
    range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("Gest�o Regionalizada RSU - Simula��o Rotas Tecnol�gicas.xlsm"). _
        Activate
    Sheets("Dist�ncias entre Munic�pios").Select
    range("B3").Select
    ActiveSheet.Paste
    range("I8").Select
    Application.CutCopyMode = False
End Sub
