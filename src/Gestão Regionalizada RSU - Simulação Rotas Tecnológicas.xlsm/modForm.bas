Attribute VB_Name = "modForm"
Option Explicit

'Função que retornará o nome da classe e o nome do UserForm
Private Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Função que recupera as informações sobre o nome da classe e o estilo da janela do UserForm
Private Declare PtrSafe Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Função que altera o estilo da janela do UserForm
Private Declare PtrSafe Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Abre o formulário principal da ferramenta
Sub OpenTool()
    frmTool.Show
End Sub

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

Public Sub applyLookAndFeel(ByVal form As Variant, ByVal level As Integer, ByVal title As String, Optional ByVal bgWhite As Boolean)
    
    Dim frmBackColor As Double
    Dim btnBackColor, btnForeColor As Double
    Dim txtBackColor, txtForeColor, txtAlign As Double
    
    'Set values depending on level
    Select Case level
        Case 1
            form.Caption = APPNAME & " - " & title
            frmBackColor = ApplicationColors.frmBgColorLevel1
            btnBackColor = ApplicationColors.bgColorLevel1
            btnForeColor = ApplicationColors.fgColorLevel1
            txtForeColor = ApplicationColors.txtFgColorLevel1
            txtAlign = 1
        Case 2
            form.Caption = APPSHORTNAME & " > " & title
            frmBackColor = ApplicationColors.frmBgColorLevel2
            btnBackColor = ApplicationColors.bgColorLevel2
            btnForeColor = ApplicationColors.fgColorLevel2
            txtForeColor = ApplicationColors.txtFgColorLevel2
            txtAlign = 1
        Case 3
            form.Caption = APPSHORTNAME & " > " & title
            frmBackColor = ApplicationColors.frmBgColorLevel3
            btnBackColor = ApplicationColors.bgColorLevel3
            btnForeColor = ApplicationColors.fgColorLevel3
            txtForeColor = ApplicationColors.txtFgColorLevel3
            txtAlign = 2
        Case Else
            form.Caption = APPSHORTNAME & " > " & title
            frmBackColor = ApplicationColors.frmBgColorLevel1
            btnBackColor = ApplicationColors.bgColorLevel1
            btnForeColor = ApplicationColors.fgColorLevel1
            txtForeColor = ApplicationColors.txtFgColorLevel1
            txtAlign = 2
    End Select

    'Apply the look and feel
    form.BackColor = frmBackColor
    
    Dim Ctrl As Control
    For Each Ctrl In form.Controls
        If InStr(Ctrl.Tag, "DO-NOT-APPLY-UI") = 0 Then
            If TypeName(Ctrl) = "CommandButton" And Ctrl.name <> "btnAbout" And Ctrl.name <> "btnHelp" And Ctrl.name <> "btnClean" Then
                Ctrl.BackColor = btnBackColor
                Ctrl.ForeColor = btnForeColor
                Ctrl.Font.Size = 9
                Ctrl.FontName = "Open Sans"
                Ctrl.FontBold = False
            ElseIf TypeName(Ctrl) = "TextBox" Then
                If bgWhite Then
                    Ctrl.ForeColor = RGB(0, 0, 0)
                Else
                    Ctrl.ForeColor = txtForeColor
                End If
                Ctrl.TextAlign = txtAlign
                Ctrl.SpecialEffect = 0
                Ctrl.BorderStyle = 1
                Ctrl.Font.Size = 9
                Ctrl.FontBold = False
                Ctrl.FontName = "Open Sans"
            End If
            
            If Ctrl.name = "imgLogo" Then
                Ctrl.Picture = LoadPicture(Application.ThisWorkbook.Path & "\" & FOLDERASSETS & "\" & IMAGELOGOEXTRASMALL)
                Ctrl.width = 110
                Ctrl.Height = 40
                Ctrl.Left = 10
                Ctrl.Top = 10
                Ctrl.BackColor = RGB(240, 240, 240)
                Ctrl.BorderStyle = 0
            End If
            
            If Ctrl.name = "lblTitle" Then
                Ctrl.Left = 130
                Ctrl.Top = 20
                Ctrl.BackColor = RGB(240, 240, 240)
            End If
        End If
    Next Ctrl
    
    'Repaint form
    form.Repaint
End Sub


'Sub que irá obter o nome do UserForm (ObjForm)
Sub EnableMinimizeMaximize(ObjForm As Object)

    'Código que atribui os botões minimizar e maximizar e possibilita redimensionar o UserForm
    SetWindowLong FindWindow("ThunderDFrame", ObjForm.Caption), -16, GetWindowLong(FindWindow("ThunderDFrame", ObjForm.Caption), -16) Or &H10000 Or &H20000 Or &H40000

End Sub
