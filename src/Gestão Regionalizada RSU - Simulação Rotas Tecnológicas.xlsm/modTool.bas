Attribute VB_Name = "modTool"
Sub CopiarDistanciasTeste()
Attribute CopiarDistanciasTeste.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CopiarDistanciasTeste Macro
'

'
    Windows( _
        "Gestão Regionalizada RSU - Simulação Rotas Tecnológicas - GRANFPOLIS - 3.2.2.xlsm" _
        ).Activate
    Sheets("Distâncias entre Municípios").Select
    range("A1").Select
    range(Selection, Selection.End(xlToRight)).Select
    range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("Gestão Regionalizada RSU - Simulação Rotas Tecnológicas.xlsm"). _
        Activate
    Sheets("Distâncias entre Municípios").Select
    range("B3").Select
    ActiveSheet.Paste
    range("I8").Select
    Application.CutCopyMode = False
End Sub
