VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditCities 
   Caption         =   "Editar Cidades"
   ClientHeight    =   7815
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   15555
   OleObjectBlob   =   "frmEditCities.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditCities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cities As New Collection

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()

End Sub


Private Sub txtConventionalCost1_Change()
    Dim index As Integer
    index = 0
    If txtConventionalCost1.Text <> "" Then
        index = index + vScrollBar.value
        cities.Item(index).vConventionalCost = CDbl(txtConventionalCost1.Text)
    End If
End Sub

Private Sub txtConventionalCost2_Change()
    Dim index As Integer
    index = 1
    If vScrollBar.value <> 1 Then
        index = index + vScrollBar.value
    End If
    cities.Item(index).vConventionalCost = CDbl(txtConventionalCost2.Text)
End Sub

Private Sub txtConventionalCost3_Change()
    Dim index As Integer
    index = 1
    If vScrollBar.value > 1 Then
        index = index + vScrollBar.value
    End If
    cities.Item(index).vConventionalCost = CDbl(txtConventionalCost1.Text)
End Sub

Private Sub txtConventionalCost4_Change()

End Sub

Private Sub txtConventionalCost5_Change()

End Sub

Private Sub txtConventionalCost6_Change()

End Sub
Private Sub txtConventionalCost7_Change()

End Sub
Private Sub txtConventionalCost8_Change()

End Sub
Private Sub txtConventionalCost9_Change()

End Sub

Private Sub txtConventionalCost10_Change()

End Sub

Private Sub UserForm_Initialize()

    Set wksDatabase = Util.GetSelectedCitiesWorksheet
    Dim lastRow As Integer
    Dim r As Integer
    lastRow = wksDatabase.Cells(Rows.Count, 1).End(xlUp).row
    For r = 2 To lastRow
        Dim c As clsCity
        Set c = New clsCity
        c.vCityName = wksDatabase.Cells(r, 1).value
        c.vPopulation = wksDatabase.Cells(r, 2).value
        c.vTrash = CDbl(wksDatabase.Cells(r, 4).value)
        c.vConventionalCost = wksDatabase.Cells(r, 5).value
        c.vTransshipmentCost = wksDatabase.Cells(r, 6).value
        c.vCostPostTransshipment = wksDatabase.Cells(r, 7).value
        If wksDatabase.Cells(r, 8).value = "Sim" Then
            c.vUTVR = True
        Else
            c.vUTVR = False
        End If
        If wksDatabase.Cells(r, 9).value = "Sim" Then
            c.vExistentLandfill = True
        Else
            c.vExistentLandfill = False
        End If
        If wksDatabase.Cells(r, 10).value = "Sim" Then
            c.vPotentialLandfill = True
        Else
            c.vPotentialLandfill = False
        End If
        cities.Add c
    Next r
    
    vScrollBar.Min = 1
    If cities.Count >= 10 Then
        vScrollBar.Max = cities.Count - 9
    Else
        vScrollBar.Enabled = False
    End If
    
    GetRangeToDisplay vScrollBar.value
End Sub

Private Sub vScrollBar_Change()
    GetRangeToDisplay vScrollBar.value
End Sub

Sub GetRangeToDisplay(currentValue As Integer)
    Debug.Print currentValue
    Dim i As Integer
    i = 1
    If cities.Count > 10 Then
        i = currentValue
    End If
    
    t = 1
    While t <= 10
        Me.Controls("txtCity" & t).value = cities.Item(i).vCityName
        Me.Controls("txtPopulation" & t).value = cities.Item(i).vPopulation
        Me.Controls("txtTrash" & t).value = cities.Item(i).vTrash
        Me.Controls("txtConventionalCost" & t).value = cities.Item(i).vConventionalCost
        Me.Controls("txtTransshipmentCost" & t).value = cities.Item(i).vTransshipmentCost
        Me.Controls("txtCostPostTranshipment" & t).value = cities.Item(i).vCostPostTransshipment
        Me.Controls("cbUTVR" & t).value = cities.Item(i).vUTVR
        Me.Controls("cbExistentLandfill" & t).value = cities.Item(i).vExistentLandfill
        Me.Controls("cbPotentialLandfill" & t).value = cities.Item(i).vPotentialLandfill
        i = i + 1
        t = t + 1
    Wend
    
End Sub
