Attribute VB_Name = "modArray"
Option Explicit

Public Enum DatabaseArrayColumn
    colId = 1
    colSelected = 2
    colCode = 3
    colArrayRaw = 4
    colSubRaw = 5
    colLandfill = 6
    colExistentLandfill = 7
    colUTVR = 8
    colPopulation = 9
    colTotal = 10
    colTrash = 11
    colTechnology = 12
    colInbound = 13
    colOutbound = 14
    colOutboundExistentLandfill = 15
End Enum


Public Function countSelectedArrays()
    Dim arrays As New Collection
    Dim count As Integer
    Dim e As Variant
    count = 0
    Set arrays = readArrays
    If arrays.count <> 0 Then
        For Each e In arrays
            If e.vSelected Then
                count = count + 1
            End If
        Next e
    End If
    
    countSelectedArrays = count
End Function

Public Function readArrays()
    Dim arrays As New Collection
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.GetArraysWorksheet
    Dim lastRow As Integer
    Dim r As Integer
    lastRow = wksDatabase.Cells(Rows.count, 1).End(xlUp).row
    
    
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
            arr.vCode = wksDatabase.Cells(r, DatabaseArrayColumn.colCode).value
            arr.vLandfill = wksDatabase.Cells(r, DatabaseArrayColumn.colLandfill).value
            arr.vExistentLandfill = wksDatabase.Cells(r, DatabaseArrayColumn.colExistentLandfill).value
            arr.vUTVR = wksDatabase.Cells(r, DatabaseArrayColumn.colUTVR).value
            arr.vPopulation = wksDatabase.Cells(r, DatabaseArrayColumn.colPopulation).value
            arr.vTotal = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colTotal).value, 3)
            arr.vTrash = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colTrash).value, 3)
            arr.vTechnology = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colTechnology).value, 3)
            arr.vInbound = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colInbound).value, 3)
            arr.vOutbound = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colOutbound).value, 3)
            arr.vOutboundExistentLandfill = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colOutboundExistentLandfill).value, 3)
            Set arr.vSubArray = New Collection
        Else
            Set subArr = New clsArray
            subArr.vCode = wksDatabase.Cells(r, DatabaseArrayColumn.colCode).value
            subArr.vArrayRaw = wksDatabase.Cells(r, DatabaseArrayColumn.colSubRaw).value
            subArr.vArrayRaw = Replace(subArr.vArrayRaw, "[", "")
            subArr.vArrayRaw = Replace(subArr.vArrayRaw, "]", "")
            subArr.vArrayRaw = Replace(subArr.vArrayRaw, "'", "")
            subArr.vLandfill = wksDatabase.Cells(r, DatabaseArrayColumn.colLandfill).value
            subArr.vLandfill = Replace(subArr.vLandfill, "'", "")
            subArr.vExistentLandfill = wksDatabase.Cells(r, DatabaseArrayColumn.colExistentLandfill).value
            subArr.vExistentLandfill = Replace(subArr.vExistentLandfill, "'", "")
            subArr.vUTVR = wksDatabase.Cells(r, DatabaseArrayColumn.colUTVR).value
            subArr.vUTVR = Replace(subArr.vUTVR, "'", "")
            subArr.vPopulation = wksDatabase.Cells(r, DatabaseArrayColumn.colPopulation).value
            subArr.vTotal = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colTotal).value, 3)
            subArr.vTrash = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colTrash).value, 3)
            subArr.vTechnology = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colTechnology).value, 3)
            subArr.vInbound = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colInbound).value, 3)
            subArr.vOutbound = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colOutbound).value, 3)
            subArr.vOutboundExistentLandfill = Round(wksDatabase.Cells(r, DatabaseArrayColumn.colOutboundExistentLandfill).value, 3)
            arr.vSubArray.Add subArr
        End If
    Next r
    If Not arr Is Nothing Then
        arrays.Add arr
    End If
    Set readArrays = arrays
End Function

Public Function updateValues(ByVal arrays As Collection)
    Dim wksDatabase As Worksheet
    Set wksDatabase = Util.GetArraysWorksheet
    Dim lastRow As Integer
    Dim r, id As Integer
    lastRow = wksDatabase.Cells(Rows.count, 1).End(xlUp).row
    
    For r = 2 To lastRow
        id = wksDatabase.Cells(r, DatabaseArrayColumn.colId).value
        If arrays(id).vSelected Then
            wksDatabase.Cells(r, DatabaseArrayColumn.colSelected) = "Sim"
        Else
            wksDatabase.Cells(r, DatabaseArrayColumn.colSelected) = "Não"
        End If
    Next r
End Function
