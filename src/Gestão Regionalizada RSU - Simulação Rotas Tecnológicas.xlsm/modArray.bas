Attribute VB_Name = "modArray"
Option Explicit

Public Enum DatabaseArrayColumn
    colId = 1
    colSelected = 2
    colArrayRaw = 3
    colSubRaw = 4
    colTotal = 5
    colTrash = 6
    colInbound = 7
    colOutbound = 8
End Enum

Public Function readArrays()
    Dim arrays As New Collection
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.GetArraysWorksheet
    Dim lastRow As Integer
    Dim r As Integer
    lastRow = wksDatabase.Cells(Rows.Count, 1).End(xlUp).row
    
    
    Dim arr As clsArray
    Dim subArr As clsArray
    Dim subArrCollection As New Collection
    For r = 2 To lastRow
        If wksDatabase.Cells(r, DatabaseArrayColumn.colSubRaw).value = "Sumário" Then
            If r <> 2 Then
                arrays.Add arr
            End If
            Set arr = New clsArray
            arr.vArrayRaw = wksDatabase.Cells(r, DatabaseArrayColumn.colArrayRaw).value
            arr.vTotal = wksDatabase.Cells(r, DatabaseArrayColumn.colTotal).value
            arr.vTrash = wksDatabase.Cells(r, DatabaseArrayColumn.colTrash).value
            arr.vInbound = wksDatabase.Cells(r, DatabaseArrayColumn.colInbound).value
            arr.vOutbound = wksDatabase.Cells(r, DatabaseArrayColumn.colOutbound).value
            Set arr.vSubArray = New Collection
        Else
            Set subArr = New clsArray
            subArr.vArrayRaw = wksDatabase.Cells(r, DatabaseArrayColumn.colSubRaw).value
            subArr.vTotal = wksDatabase.Cells(r, DatabaseArrayColumn.colTotal).value
            subArr.vTrash = wksDatabase.Cells(r, DatabaseArrayColumn.colTrash).value
            subArr.vInbound = wksDatabase.Cells(r, DatabaseArrayColumn.colInbound).value
            subArr.vOutbound = wksDatabase.Cells(r, DatabaseArrayColumn.colOutbound).value
            arr.vSubArray.Add subArr
        End If
    Next r
    Set readArrays = arrays
End Function
