VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditCities 
   Caption         =   "Editar Cidades"
   ClientHeight    =   7995
   ClientLeft      =   360
   ClientTop       =   1395
   ClientWidth     =   15600
   OleObjectBlob   =   "frmEditCities.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditCities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cities As New Collection
Public changeValues As Boolean

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    updateCityValues cities
    Unload Me
End Sub

Private Sub cbExistentLandfill1_Click()
    Call updateExistentLandfill(cbExistentLandfill1, 1)
End Sub

Private Sub cbExistentLandfill10_Click()
    Call updateExistentLandfill(cbExistentLandfill10, 10)
End Sub

Private Sub cbExistentLandfill2_Click()
    Call updateExistentLandfill(cbExistentLandfill2, 2)
End Sub

Private Sub cbExistentLandfill3_Click()
    Call updateExistentLandfill(cbExistentLandfill3, 3)
End Sub

Private Sub cbExistentLandfill4_Click()
    Call updateExistentLandfill(cbExistentLandfill4, 4)
End Sub

Private Sub cbExistentLandfill5_Click()
    Call updateExistentLandfill(cbExistentLandfill5, 5)
End Sub

Private Sub cbExistentLandfill6_Click()
    Call updateExistentLandfill(cbExistentLandfill6, 6)
End Sub

Private Sub cbExistentLandfill7_Click()
    Call updateExistentLandfill(cbExistentLandfill7, 7)
End Sub

Private Sub cbExistentLandfill8_Click()
    Call updateExistentLandfill(cbExistentLandfill8, 8)
End Sub

Private Sub cbExistentLandfill9_Click()
    Call updateExistentLandfill(cbExistentLandfill9, 9)
End Sub

Private Sub updateExistentLandfill(ByRef chkBox, ByVal index As Integer)
    Dim i As Integer
    i = index
    If vScrollBar.value > 1 Then
        i = i + vScrollBar.value - 1
    End If
    
    If changeValues Then
        cities.Item(i).vExistentLandfill = chkBox.value
    End If
End Sub

Private Sub cbPotentialLandfill1_Click()
    Call updatePotentialLandfill(cbPotentialLandfill1, 1)
End Sub

Private Sub cbPotentialLandfill10_Click()
    Call updatePotentialLandfill(cbPotentialLandfill10, 10)
End Sub

Private Sub cbPotentialLandfill2_Click()
    Call updatePotentialLandfill(cbPotentialLandfill2, 2)
End Sub

Private Sub cbPotentialLandfill3_Click()
    Call updatePotentialLandfill(cbPotentialLandfill3, 3)
End Sub

Private Sub cbPotentialLandfill4_Click()
    Call updatePotentialLandfill(cbPotentialLandfill4, 4)
End Sub

Private Sub cbPotentialLandfill5_Click()
    Call updatePotentialLandfill(cbPotentialLandfill5, 5)
End Sub

Private Sub cbPotentialLandfill6_Click()
    Call updatePotentialLandfill(cbPotentialLandfill6, 6)
End Sub

Private Sub cbPotentialLandfill7_Click()
    Call updatePotentialLandfill(cbPotentialLandfill7, 7)
End Sub

Private Sub cbPotentialLandfill8_Click()
    Call updatePotentialLandfill(cbPotentialLandfill8, 8)
End Sub

Private Sub cbPotentialLandfill9_Click()
    Call updatePotentialLandfill(cbPotentialLandfill9, 9)
End Sub

Private Sub updatePotentialLandfill(ByRef chkBox, ByVal index As Integer)
    Dim i As Integer
    i = index
    If vScrollBar.value > 1 Then
        i = i + vScrollBar.value - 1
    End If
    
    If changeValues Then
        cities.Item(i).vPotentialLandfill = chkBox.value
    End If
End Sub

Private Sub cbUTVR1_Click()
    Call updateUTVR(cbUTVR1, 1)
End Sub

Private Sub cbUTVR10_Click()
    Call updateUTVR(cbUTVR10, 10)
End Sub

Private Sub cbUTVR2_Click()
    Call updateUTVR(cbUTVR2, 2)
End Sub

Private Sub cbUTVR3_Click()
    Call updateUTVR(cbUTVR3, 3)
End Sub

Private Sub cbUTVR4_Click()
    Call updateUTVR(cbUTVR4, 4)
End Sub

Private Sub cbUTVR5_Click()
    Call updateUTVR(cbUTVR5, 5)
End Sub

Private Sub cbUTVR6_Click()
    Call updateUTVR(cbUTVR6, 6)
End Sub

Private Sub cbUTVR7_Click()
    Call updateUTVR(cbUTVR7, 7)
End Sub

Private Sub cbUTVR8_Click()
    Call updateUTVR(cbUTVR8, 8)
End Sub

Private Sub cbUTVR9_Click()
    Call updateUTVR(cbUTVR9, 9)
End Sub

Private Sub updateUTVR(ByRef chkBox, ByVal index As Integer)
    Dim i As Integer
    i = index
    If vScrollBar.value > 1 Then
        i = i + vScrollBar.value - 1
    End If
    
    If changeValues Then
        cities.Item(i).vUTVR = chkBox.value
    End If
End Sub

Private Sub txtConventionalCost1_Change()
    Call updateConventionalCost(txtConventionalCost1, 1)
End Sub

Private Sub txtConventionalCost2_Change()
    Call updateConventionalCost(txtConventionalCost2, 2)
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
    Dim errorMsg As String
    
    i = index
    If vScrollBar.value > 1 Then
        i = i + vScrollBar.value - 1
    End If
    
    If Util.validateRange(txtBox.Text, 0#, 500#, errorMsg) Then
        txtBox.BackColor = ApplicationColors.bgColorValidTextBox
        txtBox.ControlTipText = errorMsg
    Else
        txtBox.BackColor = ApplicationColors.bgColorInvalidTextBox
        txtBox.ControlTipText = errorMsg
    End If
    
    If IsNumeric(txtBox.Text) And changeValues Then
        cities.Item(i).vConventionalCost = CDbl(txtBox.Text)
    End If
End Sub

Private Sub updateTransshipmentCost(ByRef txtBox, ByVal index As Integer)
    Dim i As Integer
    Dim errorMsg As String
    
    i = index
    If vScrollBar.value > 1 Then
        i = i + vScrollBar.value - 1
    End If
    
    If Util.validateRange(txtBox.Text, 0#, 1500#, errorMsg) Then
        txtBox.BackColor = ApplicationColors.bgColorValidTextBox
        txtBox.ControlTipText = errorMsg
    Else
        txtBox.BackColor = ApplicationColors.bgColorInvalidTextBox
        txtBox.ControlTipText = errorMsg
    End If
    
    If IsNumeric(txtBox.Text) And changeValues Then
        cities.Item(i).vTransshipmentCost = CDbl(txtBox.Text)
    End If
End Sub

Private Sub updateCostPostTranshipment(ByRef txtBox, ByVal index As Integer)
    Dim i As Integer
    Dim errorMsg As String
    
    i = index
    If vScrollBar.value > 1 Then
        i = i + vScrollBar.value - 1
    End If
    
    If Util.validateRange(txtBox.Text, 0#, 10#, errorMsg) Then
        txtBox.BackColor = ApplicationColors.bgColorValidTextBox
        txtBox.ControlTipText = errorMsg
    Else
        txtBox.BackColor = ApplicationColors.bgColorInvalidTextBox
        txtBox.ControlTipText = errorMsg
    End If
    
    If IsNumeric(txtBox.Text) And changeValues Then
        cities.Item(i).vCostPostTranshipment = CDbl(txtBox.Text)
    End If
End Sub

Private Sub txtCostPostTranshipment1_Change()
    Call updateCostPostTranshipment(txtCostPostTranshipment1, 1)
End Sub

Private Sub txtCostPostTranshipment10_Change()
    Call updateCostPostTranshipment(txtCostPostTranshipment10, 10)
End Sub

Private Sub txtCostPostTranshipment2_Change()
    Call updateCostPostTranshipment(txtCostPostTranshipment2, 2)
End Sub

Private Sub txtCostPostTranshipment3_Change()
    Call updateCostPostTranshipment(txtCostPostTranshipment3, 3)
End Sub

Private Sub txtCostPostTranshipment4_Change()
    Call updateCostPostTranshipment(txtCostPostTranshipment4, 4)
End Sub

Private Sub txtCostPostTranshipment5_Change()
    Call updateCostPostTranshipment(txtCostPostTranshipment5, 5)
End Sub

Private Sub txtCostPostTranshipment6_Change()
    Call updateCostPostTranshipment(txtCostPostTranshipment6, 6)
End Sub

Private Sub txtCostPostTranshipment7_Change()
    Call updateCostPostTranshipment(txtCostPostTranshipment7, 7)
End Sub

Private Sub txtCostPostTranshipment8_Change()
    Call updateCostPostTranshipment(txtCostPostTranshipment8, 9)
End Sub

Private Sub txtCostPostTranshipment9_Change()
    Call updateCostPostTranshipment(txtCostPostTranshipment9, 9)
End Sub

Private Sub txtTransshipmentCost1_Change()
    Call updateTransshipmentCost(txtTransshipmentCost1, 1)
End Sub

Private Sub txtTransshipmentCost10_Change()
    Call updateTransshipmentCost(txtTransshipmentCost10, 10)
End Sub

Private Sub txtTransshipmentCost2_Change()
    Call updateTransshipmentCost(txtTransshipmentCost2, 2)
End Sub

Private Sub txtTransshipmentCost3_Change()
    Call updateTransshipmentCost(txtTransshipmentCost3, 3)
End Sub

Private Sub txtTransshipmentCost4_Change()
    Call updateTransshipmentCost(txtTransshipmentCost4, 4)
End Sub

Private Sub txtTransshipmentCost5_Change()
    Call updateTransshipmentCost(txtTransshipmentCost5, 5)
End Sub

Private Sub txtTransshipmentCost6_Change()
    Call updateTransshipmentCost(txtTransshipmentCost6, 6)
End Sub

Private Sub txtTransshipmentCost7_Change()
    Call updateTransshipmentCost(txtTransshipmentCost7, 7)
End Sub

Private Sub txtTransshipmentCost8_Change()
    Call updateTransshipmentCost(txtTransshipmentCost8, 8)
End Sub

Private Sub txtTransshipmentCost9_Change()
    Call updateTransshipmentCost(txtTransshipmentCost9, 9)
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Editar Cidades", True)
    
    Set cities = readSelectedCities
    
    vScrollBar.Min = 1
    If cities.Count >= 10 Then
        vScrollBar.Max = cities.Count - 9
    Else
        vScrollBar.Enabled = False
    End If
    
    changeValues = True
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
    
    changeValues = False
    
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
    
    changeValues = True
    
End Sub
