VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectArrays 
   Caption         =   "UserForm1"
   ClientHeight    =   7584
   ClientLeft      =   75
   ClientTop       =   120
   ClientWidth     =   16605
   OleObjectBlob   =   "frmSelectArrays.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arrays2 As New Collection
Dim arrays3 As New Collection
Dim arrays As Collection
Dim arraySave As Collection
Dim subArraySize As Integer

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    Dim count As Integer
    Dim e As Variant
    count = 0
    For Each e In arraySave
        If e.vSelected Then
            count = count + 1
        End If
    Next e
    
    If count = 4 Then
        modArray.updateValues arraySave
        frmStepTwo.updateForm
        Unload Me
        ThisWorkbook.Save
    Else
        Call MsgBox(MSG_WRONG_NUMBER_ARRAYS, vbCritical, MSG_WRONG_NUMBER_ARRAYS_TITLE)
    End If
End Sub


Private Sub subarrayTab_Change()
    If subarrayTab.SelectedItem.index = 0 Then
        subArraySize = 2
    Else
        subArraySize = 3
    End If
    
    Set arrays = arraySave
    Set arrays2 = New Collection
    Set arrays3 = New Collection
    For Each A In arrays
        If A.vSubArray.count = 2 Then
            arrays2.Add A
        ElseIf A.vSubArray.count = 3 Then
            arrays3.Add A
        End If
    Next A
    
    If subArraySize = 2 Then
        Set arrays = arrays2
    Else
        Set arrays = arrays3
    End If
    
    vScrollBar.Min = 1
    vScrollBar.Max = arrays.count
    
    vScrollBar.value = 1
    vScrollBar_Change
    
End Sub

Private Sub txtArraySelected_Click()
    Dim currentValue As Integer
    currentValue = vScrollBar.value
    'currentValue = currentValue + 1
    
    For Each A In arraySave
        If arrays.Item(currentValue).vCode = A.vCode Then
            A.vSelected = txtArraySelected.value
            Exit For
        End If
    Next A
    
    Set arrays = arraySave
    Set arrays2 = New Collection
    Set arrays3 = New Collection
    For Each A In arrays
        If A.vSubArray.count = 2 Then
            arrays2.Add A
        ElseIf A.vSubArray.count = 3 Then
            arrays3.Add A
        End If
    Next A
    
    If subArraySize = 2 Then
        Set arrays = arrays2
    Else
        Set arrays = arrays3
    End If
    
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Definir Arranjos Consolidados", True)
    
    
    Set arraySave = readArrays
    
    'Centralized Array is always the first one
    lblCentralizedCode.Caption = arraySave(1).vCode
    txtCentralizedArray.Text = arraySave(1).vSubArray(1).vArrayRaw
    txtCentralizedLandfill.Text = arraySave(1).vSubArray(1).vLandfill
    txtCentralizedExistentLandfill.Text = arraySave(1).vSubArray(1).vExistentLandfill
    txtCentralizedUTVR.Text = arraySave(1).vSubArray(1).vUTVR
    txtCentralizedTotal.Text = arraySave(1).vTotal
    txtCentralizedTrash.Text = arraySave(1).vTrash
    txtCentralizedTechnology = arraySave(1).vTechnology
    txtCentralizedInbound.Text = arraySave(1).vInbound
    txtCentralizedOutbound.Text = arraySave(1).vOutbound
    txtCentralizedOutboundExistent.Text = arraySave(1).vOutboundExistentLandfill
    
    subArraySize = 2
    
    Set arrays = arraySave
    Set arrays2 = New Collection
    Set arrays3 = New Collection
    For Each A In arrays
        If A.vSubArray.count = 2 Then
            arrays2.Add A
        ElseIf A.vSubArray.count = 3 Then
            arrays3.Add A
        End If
    Next A
    
    If subArraySize = 2 Then
        Set arrays = arrays2
    Else
        Set arrays = arrays3
    End If
    
    vScrollBar.Min = 1
    vScrollBar.Max = arrays.count
    
    subarrayTab.Tabs(0).Caption = "Subarranjos de Tamanho 2"
    subarrayTab.Tabs(1).Caption = "Subarranjos de Tamanho 3"

    If Database.GetDatabaseValue("MaxSubarrays", colUserValue) < 3 Then
        subarrayTab.Tabs(1).Visible = False
    End If
    
    frmSelectArrays.Height = 621
    frmSelectArrays.width = 1225
    
End Sub

Private Sub vScrollBar_Change()
    Dim currentValue As Integer
    currentValue = vScrollBar.value
    'currentValue = currentValue + 1
    
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
