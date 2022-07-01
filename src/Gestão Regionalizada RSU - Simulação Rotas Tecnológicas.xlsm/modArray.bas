Attribute VB_Name = "modArray"
Option Explicit

Public Enum DatabaseArrayColumn
    colId = 1
    colSelected = 2
    colArrayRaw = 3
    colSubRaw = 4
    colLandfill = 5
    colUTVR = 6
    colTotal = 7
    colTrash = 8
    colTechnology = 9
    colInbound = 10
    colOutbound = 11
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
            If wksDatabase.Cells(r, DatabaseArrayColumn.colSelected).value = "Sim" Then
                arr.vSelected = True
            Else
                arr.vSelected = False
            End If
            arr.vLandfill = wksDatabase.Cells(r, DatabaseArrayColumn.colLandfill).value
            arr.vUTVR = wksDatabase.Cells(r, DatabaseArrayColumn.colUTVR).value
            arr.vTotal = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colTotal).value, 2)
            arr.vTrash = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colTrash).value, 2)
            arr.vTechnology = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colTechnology).value, 2)
            arr.vInbound = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colInbound).value, 2)
            arr.vOutbound = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colOutbound).value, 2)
            Set arr.vSubArray = New Collection
        Else
            Set subArr = New clsArray
            subArr.vArrayRaw = wksDatabase.Cells(r, DatabaseArrayColumn.colSubRaw).value
            subArr.vArrayRaw = Replace(subArr.vArrayRaw, "[", "")
            subArr.vArrayRaw = Replace(subArr.vArrayRaw, "]", "")
            subArr.vArrayRaw = Replace(subArr.vArrayRaw, "'", "")
            subArr.vLandfill = wksDatabase.Cells(r, DatabaseArrayColumn.colLandfill).value
            subArr.vUTVR = wksDatabase.Cells(r, DatabaseArrayColumn.colUTVR).value
            subArr.vTotal = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colTotal).value, 2)
            subArr.vTrash = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colTrash).value, 2)
            subArr.vTechnology = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colTechnology).value, 2)
            subArr.vInbound = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colInbound).value, 2)
            subArr.vOutbound = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colOutbound).value, 2)
            arr.vSubArray.Add subArr
        End If
    Next r
    arrays.Add arr
    Set readArrays = arrays
End Function

Public Function updateValues(ByVal arrays As Collection)
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.GetArraysWorksheet
    Dim lastRow As Integer
    Dim r, id As Integer
    lastRow = wksDatabase.Cells(Rows.Count, 1).End(xlUp).row
    
    For r = 2 To lastRow
        id = wksDatabase.Cells(r, DatabaseArrayColumn.colId).value
        If arrays(id).vSelected Then
            wksDatabase.Cells(r, DatabaseArrayColumn.colSelected) = "Sim"
        Else
            wksDatabase.Cells(r, DatabaseArrayColumn.colSelected) = "Não"
        End If
    Next r
End Function
