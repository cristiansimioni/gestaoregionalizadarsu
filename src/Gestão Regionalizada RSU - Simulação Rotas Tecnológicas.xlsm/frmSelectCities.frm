VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectCities 
   Caption         =   "Selecionar Cidades"
   ClientHeight    =   5655
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7620
   OleObjectBlob   =   "frmSelectCities.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectCities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnAdd_Click()
    lstSelected.AddItem (lstAvailable.List(lstAvailable.ListIndex))
End Sub

Private Sub btnRemove_Click()
    If lstSelected.ListIndex > -1 Then
        lstSelected.RemoveItem (lstSelected.ListIndex)
    End If
End Sub

Private Sub btnSave_Click()
    Set wksSelectedCities = Util.GetSelectedCitiesWorksheet
    
    x = 2
    For Index = 0 To lstSelected.ListCount - 1
        wksSelectedCities.Cells(x, "A").value = CStr(lstSelected.List(Index))
        x = x + 1
    Next Index
    
End Sub

Private Sub cbxUF_Change()
    Set wksCities = Util.GetCitiesWorksheet
    lastRow = wksCities.Cells(Rows.Count, 1).End(xlUp).row
    lstAvailable.Clear
    currentUF = cbxUF
    For x = 2 To lastRow
        uf = wksCities.Cells(x, "A")
        city = wksCities.Cells(x, "D")
        If uf = cbxUF Then
            lstAvailable.AddItem (city)
        End If
    Next x
End Sub

Private Sub ToggleButton1_Click()
    Unload Me
End Sub



Private Sub UserForm_Initialize()
    'Form Appearance
    Me.Caption = APPNAME & " - Selectionar Cidades"
    Me.BackColor = ApplicationColors.bgColorLevel3
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.btColorLevel3
         End If
    Next Ctrl
    
    'Load UF
    Set wksCities = Util.GetCitiesWorksheet
    lastRow = wksCities.Cells(Rows.Count, 1).End(xlUp).row
    For x = 2 To lastRow
        uf = wksCities.Cells(x, "A")
        inList = False
        For Index = 0 To cbxUF.ListCount - 1
            If uf = CStr(cbxUF.List(Index)) Then
                inList = True
                Exit For
            End If
        Next Index
        If inList = False Then cbxUF.AddItem (uf)
    Next x
    
End Sub
