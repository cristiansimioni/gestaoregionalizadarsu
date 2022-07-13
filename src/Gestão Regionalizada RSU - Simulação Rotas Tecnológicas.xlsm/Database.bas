Attribute VB_Name = "Database"
Option Explicit

Public Enum DatabaseColumn
    colGroup = 1
    colStep = 2
    colDescription = 3
    colName = 4
    colType = 5
    colFormula = 6
    colUnit = 7
    colUserValue = 8
    colDefaultValue = 9
    colWorkbook = 10
    colSheet = 11
    colCell = 12
    colValid = 13
    colMinValue = 14
    colMaxValue = 15
End Enum

Function LocateVariableRow(ByVal name As String)
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.GetDatabaseWorksheet
    Dim lastRow As Integer
    Dim r As Integer
    lastRow = wksDatabase.Cells(Rows.Count, DatabaseColumn.colName).End(xlUp).row
    For r = 2 To lastRow
        If wksDatabase.Cells(r, DatabaseColumn.colName).value = name Then
            LocateVariableRow = r
        End If
    Next r
End Function

Function GetDatabaseValue(ByVal name As String, ByVal column As DatabaseColumn)
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.GetDatabaseWorksheet
    Dim row As Integer
    row = LocateVariableRow(name)
    GetDatabaseValue = wksDatabase.Cells(row, column).value
End Function

Sub SetDatabaseValue(ByVal name As String, ByVal column As DatabaseColumn, ByVal v)
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.GetDatabaseWorksheet
    Dim row As Integer
    row = LocateVariableRow(name)
    wksDatabase.Cells(row, column).value = v
End Sub

Function Validate(ByVal name As String, ByVal value As String, Optional ByRef message As String) As Boolean
    Validate = True
    Dim varType As String
    Dim varMinValue As Double
    Dim varMaxValue As Double
    
    varType = GetDatabaseValue(name, DatabaseColumn.colType)
    varMinValue = GetDatabaseValue(name, DatabaseColumn.colMinValue)
    varMaxValue = GetDatabaseValue(name, DatabaseColumn.colMaxValue)
    
    If varType = "Double" Then
        If IsNumeric(value) Then
            Dim number As Double
            number = CDbl(value)
            If number >= varMinValue And number <= varMaxValue Then
                message = ""
            Else
                Validate = False
                message = "O valor deve ser maior ou igual a " & varMinValue & " e menor ou igual a " & varMaxValue & "."
            End If
        Else
            Validate = False
            message = "O valor deve ser numérico entre " & varMinValue & " e " & varMaxValue
        End If
    End If
    
End Function

Function checkStepStatus(ByVal step As String)
    checkStepStatus = True
End Function

Sub CleanDatabase()

End Sub
