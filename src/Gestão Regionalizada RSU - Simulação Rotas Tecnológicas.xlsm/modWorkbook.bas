Attribute VB_Name = "modWorkbook"
Option Explicit

Public Sub EditRouteToolData(ByVal filename, ByVal arr, ByVal market As String)
    Application.DisplayAlerts = False
    
    Dim value As Integer
    value = 200
    
    Workbooks.Open filename
    
    ' Valores sub-arranjo
    ActiveWorkbook.Sheets("R-Entrada").Range("E10") = arr.vTrash
    ActiveWorkbook.Sheets("R&C-Painel de Controle").Range("D84") = arr.vInbound
    ActiveWorkbook.Sheets("R&C-Painel de Controle").Range("D88") = arr.vOutbound
    
    If market = FOLDERLANDFILLMARKET Then
        ActiveWorkbook.Sheets("R-Definição").Range("E121") = "Existente"
    End If
    
    'Valores da ferramenta
    
    
    ActiveWorkbook.Save
    ActiveWindow.Close
    
    Application.DisplayAlerts = True
End Sub
