VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectArrays 
   Caption         =   "UserForm1"
   ClientHeight    =   11010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22095
   OleObjectBlob   =   "frmSelectArrays.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arrays As Collection

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    updateValues arrays
    Unload Me
End Sub

Private Sub txtArraySelected_Click()
    Dim currentValue As Integer
    currentValue = vScrollBar.value
    currentValue = currentValue + 1
    
    arrays.Item(currentValue).vSelected = txtArraySelected.value
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Definir Arranjos Centralizados"
    Me.BackColor = ApplicationColors.frmBgColorLevel2
    
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.bgColorLevel3
         End If
    Next Ctrl
    
    Set arrays = readArrays
    
    'Centralized Array is always the first one
    txtCentralizedArray.Text = arrays(1).vSubArray(1).vArrayRaw
    txtCentralizedLandfill.Text = arrays(1).vSubArray(1).vLandfill
    txtCentralizedUTVR.Text = arrays(1).vSubArray(1).vUTVR
    txtCentralizedTotal.Text = arrays(1).vSubArray(1).vTotal
    txtCentralizedTrash.Text = arrays(1).vSubArray(1).vTrash
    txtCentralizedTechnology = arrays(1).vSubArray(1).vTechnology
    txtCentralizedInbound.Text = arrays(1).vSubArray(1).vInbound
    txtCentralizedOutbound.Text = arrays(1).vSubArray(1).vOutbound
    
    'txtArray2.Text = arrays(2).vArrayRaw
    vScrollBar.Min = 1
    vScrollBar.Max = arrays.Count - 1
End Sub

Private Sub vScrollBar_Change()
    Dim currentValue As Integer
    currentValue = vScrollBar.value
    currentValue = currentValue + 1
    
    'Clear
    t = 1
    While t <= 6
        Me.Controls("txtSubArray" & t).value = ""
        Me.Controls("txtSubArrayLandfill" & t).value = ""
        Me.Controls("txtSubArrayUTVR" & t).value = ""
        Me.Controls("txtSubArrayTotal" & t).value = ""
        Me.Controls("txtSubArrayTrash" & t).value = ""
        Me.Controls("txtSubArrayTechnology" & t).value = ""
        Me.Controls("txtSubArrayInbound" & t).value = ""
        Me.Controls("txtSubArrayOutbound" & t).value = ""
        t = t + 1
    Wend
    
    'Fill sub array
    t = 1
    While t <= arrays(currentValue).vSubArray.Count
        Me.Controls("txtSubArray" & t).value = arrays.Item(currentValue).vSubArray(t).vArrayRaw
        Me.Controls("txtSubArrayLandfill" & t).value = arrays.Item(currentValue).vSubArray(t).vLandfill
        Me.Controls("txtSubArrayUTVR" & t).value = arrays.Item(currentValue).vSubArray(t).vUTVR
        Me.Controls("txtSubArrayTotal" & t).value = arrays.Item(currentValue).vSubArray(t).vTotal
        Me.Controls("txtSubArrayTrash" & t).value = arrays.Item(currentValue).vSubArray(t).vTrash
        Me.Controls("txtSubArrayTechnology" & t).value = arrays.Item(currentValue).vSubArray(t).vTechnology
        Me.Controls("txtSubArrayInbound" & t).value = arrays.Item(currentValue).vSubArray(t).vInbound
        Me.Controls("txtSubArrayOutbound" & t).value = arrays.Item(currentValue).vSubArray(t).vOutbound
        t = t + 1
    Wend
    
    'Fill array
    txtArrayTotal.Text = arrays.Item(currentValue).vTotal
    txtArrayTrash.Text = arrays.Item(currentValue).vTrash
    txtArrayTechnology.Text = arrays.Item(currentValue).vTechnology
    txtArrayInbound.Text = arrays.Item(currentValue).vInbound
    txtArrayOutbound.Text = arrays.Item(currentValue).vOutbound
    txtArraySelected.value = arrays.Item(currentValue).vSelected
End Sub
