VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectArrays 
   Caption         =   "UserForm1"
   ClientHeight    =   11595
   ClientLeft      =   300
   ClientTop       =   1116
   ClientWidth     =   24240
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
    Dim count As Integer
    Dim e As Variant
    count = 0
    For Each e In arrays
        If e.vSelected Then
            count = count + 1
        End If
    Next e
    
    If count = 4 Then
        updateValues arrays
        frmStepTwo.updateForm
        Unload Me
        ThisWorkbook.Save
    Else
        Call MsgBox(MSG_WRONG_NUMBER_ARRAYS, vbCritical, MSG_WRONG_NUMBER_ARRAYS_TITLE)
    End If
End Sub


Private Sub txtArraySelected_Click()
    Dim currentValue As Integer
    currentValue = vScrollBar.value
    currentValue = currentValue + 1
    
    arrays.Item(currentValue).vSelected = txtArraySelected.value
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Definir Arranjos Consolidados", True)
    
    Set arrays = readArrays
    
    'Centralized Array is always the first one
    lblCentralizedCode.Caption = arrays(1).vCode
    txtCentralizedArray.Text = arrays(1).vSubArray(1).vArrayRaw
    txtCentralizedLandfill.Text = arrays(1).vSubArray(1).vLandfill
    txtCentralizedExistentLandfill.Text = arrays(1).vSubArray(1).vExistentLandfill
    txtCentralizedUTVR.Text = arrays(1).vSubArray(1).vUTVR
    txtCentralizedTotal.Text = arrays(1).vTotal
    txtCentralizedTrash.Text = arrays(1).vTrash
    txtCentralizedTechnology = arrays(1).vTechnology
    txtCentralizedInbound.Text = arrays(1).vInbound
    txtCentralizedOutbound.Text = arrays(1).vOutbound
    txtCentralizedOutboundExistent.Text = arrays(1).vOutboundExistentLandfill
    
    vScrollBar.Min = 1
    vScrollBar.Max = arrays.count - 1
    
    frmSelectArrays.Height = 609
    frmSelectArrays.width = 1225
    
End Sub

Private Sub vScrollBar_Change()
    Dim currentValue As Integer
    currentValue = vScrollBar.value
    currentValue = currentValue + 1
    
    'Clear
    t = 1
    While t <= 3
        Me.Controls("txtSubArray" & t).value = ""
        Me.Controls("txtSubArrayLandfill" & t).value = ""
        Me.Controls("txtSubArrayExistentLandfill" & t).value = ""
        Me.Controls("txtSubArrayUTVR" & t).value = ""
        Me.Controls("txtSubArrayTotal" & t).value = ""
        Me.Controls("txtSubArrayTrash" & t).value = ""
        Me.Controls("txtSubArrayTechnology" & t).value = ""
        Me.Controls("txtSubArrayInbound" & t).value = ""
        Me.Controls("txtSubArrayOutbound" & t).value = ""
        Me.Controls("txtSubArrayOutboundExistent" & t).value = ""
        t = t + 1
    Wend
    
    'Fill sub array
    t = 1
    While t <= arrays(currentValue).vSubArray.count
        Me.Controls("txtSubArray" & t).value = arrays.Item(currentValue).vSubArray(t).vArrayRaw
        Me.Controls("txtSubArrayLandfill" & t).value = arrays.Item(currentValue).vSubArray(t).vLandfill
        Me.Controls("txtSubArrayExistentLandfill" & t).value = arrays.Item(currentValue).vSubArray(t).vExistentLandfill
        Me.Controls("txtSubArrayUTVR" & t).value = arrays.Item(currentValue).vSubArray(t).vUTVR
        Me.Controls("txtSubArrayTotal" & t).value = arrays.Item(currentValue).vSubArray(t).vTotal
        Me.Controls("txtSubArrayTrash" & t).value = arrays.Item(currentValue).vSubArray(t).vTrash
        Me.Controls("txtSubArrayTechnology" & t).value = arrays.Item(currentValue).vSubArray(t).vTechnology
        Me.Controls("txtSubArrayInbound" & t).value = arrays.Item(currentValue).vSubArray(t).vInbound
        Me.Controls("txtSubArrayOutbound" & t).value = arrays.Item(currentValue).vSubArray(t).vOutbound
        Me.Controls("txtSubArrayOutboundExistent" & t).value = arrays.Item(currentValue).vSubArray(t).vOutboundExistentLandfill
        t = t + 1
    Wend
    
    'Fill array
    lblArrayCode.Caption = arrays.Item(currentValue).vCode
    txtArrayTotal.Text = arrays.Item(currentValue).vTotal
    txtArrayTrash.Text = arrays.Item(currentValue).vTrash
    txtArrayTechnology.Text = arrays.Item(currentValue).vTechnology
    txtArrayInbound.Text = arrays.Item(currentValue).vInbound
    txtArrayOutbound.Text = arrays.Item(currentValue).vOutbound
    txtArrayOutboundExistent.Text = arrays.Item(currentValue).vOutboundExistentLandfill
    txtArraySelected.value = arrays.Item(currentValue).vSelected
End Sub
