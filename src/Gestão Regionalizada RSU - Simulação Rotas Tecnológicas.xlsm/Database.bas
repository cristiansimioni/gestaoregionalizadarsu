Attribute VB_Name = "Database"
Option Explicit

Public Enum DatabaseColumn
    colGroup = 1
    colDescription = 2
    colName = 3
    colType
    colFormula
    colUnit
    colUserValue
    colDefaultValue
    colWorkbook
    colSheet
    colCell
    colValid
    colMinValue
    colMaxValue
End Enum

Public Enum Project
    projectName
End Enum

Function LocateVariableRow(ByVal name As String)
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.getDatabaseWorksheet
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
    Set wksDatabase = Util.getDatabaseWorksheet
    Dim row As Integer
    row = LocateVariableRow(name)
    GetDatabaseValue = wksDatabase.Cells(row, column)
End Function

Sub SetDatabaseValue(ByVal name As String, ByVal column As DatabaseColumn, ByVal v)
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.getDatabaseWorksheet
    Dim row As Integer
    row = LocateVariableRow(name)
    wksDatabase.Cells(row, column).value = v
End Sub

Sub CleanDatabase()
    setProjectName ("")
    setProjectPathFolder ("")
End Sub
