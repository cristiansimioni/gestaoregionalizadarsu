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

Public Sub applyLookAndFeel(ByVal form As Variant, ByVal level As Integer, ByVal title As String)
    Dim frmBackColor As Double
    Dim btnBackColor, btnForeColor As Double
    
    'Set values depending on level
    Select Case level
        Case 1
            form.Caption = APPNAME & " - " & title
            frmBackColor = ApplicationColors.frmBgColorLevel1
            btnBackColor = ApplicationColors.bgColorLevel1
            btnForeColor = ApplicationColors.fgColorLevel1
        Case 2
            form.Caption = APPSHORTNAME & " - " & title
            frmBackColor = ApplicationColors.frmBgColorLevel2
            btnBackColor = ApplicationColors.bgColorLevel2
            btnForeColor = ApplicationColors.fgColorLevel2
        Case 3
            form.Caption = APPSHORTNAME & " - " & title
            frmBackColor = ApplicationColors.frmBgColorLevel3
            btnBackColor = ApplicationColors.bgColorLevel3
            btnForeColor = ApplicationColors.fgColorLevel3
        Case Else
            form.Caption = APPSHORTNAME & " - " & title
            frmBackColor = ApplicationColors.frmBgColorLevel1
            btnBackColor = ApplicationColors.bgColorLevel1
            btnForeColor = ApplicationColors.fgColorLevel1
    End Select

    'Apply the look and feel
    form.BackColor = frmBackColor
    
    Dim Ctrl As Control
    For Each Ctrl In form.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = btnBackColor
            Ctrl.ForeColor = btnForeColor
         End If
    Next Ctrl
    
    'Repaint form
    form.Repaint
End Sub
