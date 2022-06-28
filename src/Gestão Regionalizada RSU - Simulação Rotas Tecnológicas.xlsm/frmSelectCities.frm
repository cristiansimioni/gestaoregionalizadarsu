VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectCities 
   Caption         =   "Selecionar Cidades"
   ClientHeight    =   9045.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10995
   OleObjectBlob   =   "frmSelectCities.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectCities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public databaseCities As New Collection
Public selectedCities As New Collection

Private Sub btnAdd_Click()
    If cbxUF.ListIndex > -1 And lstAvailable.ListIndex > -1 Then
        'Find city base on IBGE code
        Dim city As clsCity
        Set city = FindInCollection(databaseCities, lstAvailable.List(lstAvailable.ListIndex, 1))
        lstSelected.AddItem
        lstSelected.List(lstSelected.ListCount - 1, 0) = city.vCityName
        lstSelected.List(lstSelected.ListCount - 1, 1) = city.vIBGECode
        lstAvailable.RemoveItem (lstAvailable.ListIndex)
        city.vUTVR = True
        selectedCities.Add city
    End If
End Sub

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnRemove_Click()
    If lstSelected.ListIndex > -1 Then
        selectedCities.Remove (lstSelected.ListIndex + 1)
        lstSelected.RemoveItem (lstSelected.ListIndex)
        Call cbxUF_Change
    End If
End Sub

Private Sub btnSave_Click()
    Set wksSelectedCities = Util.GetSelectedCitiesWorksheet
    
    'Clear currect selected cities worksheet
    wksSelectedCities.Range("A2:B100").Clear
    wksSelectedCities.Range("G2:L100").Clear
    
    'Fill with current values
    row = 2
    For Each city In selectedCities
        wksSelectedCities.Cells(row, SelectedCityColumn.colCityName) = city.vCityName
        wksSelectedCities.Cells(row, SelectedCityColumn.colIBGECode) = city.vIBGECode
        wksSelectedCities.Cells(row, SelectedCityColumn.colConventionalCost) = city.vConventionalCost
        wksSelectedCities.Cells(row, SelectedCityColumn.colTransshipmentCost) = city.vTransshipmentCost
        wksSelectedCities.Cells(row, SelectedCityColumn.colCostPostTransshipment) = city.vCostPostTransshipment
        If city.vUTVR Then
            wksSelectedCities.Cells(row, SelectedCityColumn.colUTVR) = "Sim"
        Else
            wksSelectedCities.Cells(row, SelectedCityColumn.colUTVR) = "N�o"
        End If
        If city.vExistentLandfill Then
            wksSelectedCities.Cells(row, SelectedCityColumn.colExistentLandfill) = "Sim"
        Else
            wksSelectedCities.Cells(row, SelectedCityColumn.colExistentLandfill) = "N�o"
        End If
        If city.vPotentialLandfill Then
            wksSelectedCities.Cells(row, SelectedCityColumn.colPotentialLandfill) = "Sim"
        Else
            wksSelectedCities.Cells(row, SelectedCityColumn.colPotentialLandfill) = "N�o"
        End If
        row = row + 1
    Next city
    
End Sub

Private Sub cbxUF_Change()
    Set wksCities = Util.GetCitiesWorksheet
    lastRow = wksCities.Cells(Rows.Count, 1).End(xlUp).row
    lstAvailable.Clear
    currentUF = cbxUF
    
    For Each city In databaseCities
        If city.vUF = cbxUF Then
            If Not IsInCollection(selectedCities, city.vIBGECode) Then
                lstAvailable.AddItem
                lstAvailable.List(lstAvailable.ListCount - 1, 0) = city.vCityName
                lstAvailable.List(lstAvailable.ListCount - 1, 1) = city.vIBGECode
            End If
        End If
    Next city
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Me.Caption = APPNAME & " - Selectionar Cidades"
    Me.BackColor = ApplicationColors.bgColorLevel3
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.btColorLevel3
            Ctrl.ForeColor = ApplicationColors.fgColorLevel3
         End If
    Next Ctrl
    
    lstAvailable.ColumnWidths = "130,10"
    lstSelected.ColumnWidths = "130,10"
    
    
    'Load database cities
    Set databaseCities = readDatabaseCities
    
    'Load alrady selected cities if available
    Set selectedCities = readSelectedCities
    
    'Load UF
    For Each city In databaseCities
        inList = False
        For index = 0 To cbxUF.ListCount - 1
            If city.vUF = CStr(cbxUF.List(index)) Then
                inList = True
                Exit For
            End If
        Next index
        If inList = False Then cbxUF.AddItem (city.vUF)
    Next city
    
    'Show current selected cities
    lstSelected.Clear
    For Each city In selectedCities
        lstSelected.AddItem
        lstSelected.List(lstSelected.ListCount - 1, 0) = city.vCityName
        lstSelected.List(lstSelected.ListCount - 1, 1) = city.vIBGECode
    Next city
    
End Sub

Function IsInCollection(ByVal oCollection As Collection, ByVal sItem As Double) As Boolean
    Dim vItem As Variant
    For Each vItem In oCollection
        If vItem.vIBGECode = sItem Then
            IsInCollection = True
            Exit Function
        End If
    Next vItem
    IsInCollection = False
End Function

Function FindInCollection(ByVal oCollection As Collection, ByVal sItem As Double) As clsCity
    Dim vItem As clsCity
    For Each vItem In oCollection
        If vItem.vIBGECode = sItem Then
            Set FindInCollection = vItem
            Exit Function
        End If
    Next vItem
    Set FindInCollection = Nothing
End Function
