VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditCities 
   Caption         =   "Editar Cidades"
   ClientHeight    =   7860
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
    updateCityValues cities
    Unload Me
End Sub

Private Sub txtConventionalCost1_Change()
    Call updateConventionalCost(txtConventionalCost1, 1)
End Sub

Private Sub txtConventionalCost2_Change()
    Call updateConventionalCost(txtConventionalCost3, 2)
End Sub

Private Sub txtConventionalCost3_Change()
    Call updateConventionalCost(txtConventionalCost3, 3)
End Sub

Private Sub txtConventionalCost4_Change()
    Call updateConventionalCost(txtConventionalCost4, 4)
End Sub

Private Sub txtConventionalCost5_Change()
    Call updateConventionalCost(txtConventionalCost5, 5)
End Sub

Private Sub txtConventionalCost6_Change()
    Call updateConventionalCost(txtConventionalCost6, 6)
End Sub
Private Sub txtConventionalCost7_Change()
    Call updateConventionalCost(txtConventionalCost7, 7)
End Sub
Private Sub txtConventionalCost8_Change()
    Call updateConventionalCost(txtConventionalCost8, 8)
End Sub
Private Sub txtConventionalCost9_Change()
    Call updateConventionalCost(txtConventionalCost9, 9)
End Sub

Private Sub txtConventionalCost10_Change()
    Call updateConventionalCost(txtConventionalCost10, 10)
End Sub

Private Sub updateConventionalCost(ByRef txtBox, ByVal index As Integer)
    Dim i As Integer
    i = index
    If vScrollBar.value > 1 Then
        i = i + vScrollBar.value
    End If
    If IsNumeric(txtBox.Text) Then
        cities.Item(i).vConventionalCost = CDbl(txtBox.Text)
    End If
End Sub

Private Sub UserForm_Initialize()

    Set cities = readSelectedCities
    
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
