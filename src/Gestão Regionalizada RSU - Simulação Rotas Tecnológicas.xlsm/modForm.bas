Attribute VB_Name = "modForm"
Option Explicit

Public Sub textBoxChange(ByRef txtBox, ByVal varName As String, ByRef FormChanged As Boolean)
    Dim errorMsg As String
    If Database.Validate(varName, txtBox.Text, errorMsg) Then
        txtBox.BackColor = ApplicationColors.bgColorValidTextBox
        txtBox.ControlTipText = errorMsg
    Else
        txtBox.BackColor = ApplicationColors.bgColorInvalidTextBox
        txtBox.ControlTipText = errorMsg
    End If
    FormChanged = True
End Sub

Public Function validateForm()
    validateForm = True
End Function
