VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectCities 
   Caption         =   "UserForm1"
   ClientHeight    =   8028
   ClientLeft      =   -285
   ClientTop       =   -1185
   ClientWidth     =   8835.001
   OleObjectBlob   =   "frmSelectCities.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectCities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim databaseCities As New Collection
Dim selectedCities As New Collection
Dim FormChanged As Boolean

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
    If FormChanged Then
        answer = MsgBox(MSG_CHANGED_NOT_SAVED, vbQuestion + vbYesNo + vbDefaultButton2, MSG_CHANGED_NOT_SAVED_TITLE)
        If answer = vbYes Then
          Call btnSave_Click
        Else
          Unload Me
        End If
    Else
        Unload Me
    End If
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
    Set wksCitiesDistance = Util.GetCitiesDistanceWorksheet
    
    If selectedCities.count >= 2 Then
        'Clear currect selected cities worksheet
        wksSelectedCities.range("A2:B100").ClearContents
        wksSelectedCities.range("G2:L100").ClearContents
        
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
                wksSelectedCities.Cells(row, SelectedCityColumn.colUTVR) = "Não"
            End If
            If city.vExistentLandfill Then
                wksSelectedCities.Cells(row, SelectedCityColumn.colExistentLandfill) = "Sim"
            Else
                wksSelectedCities.Cells(row, SelectedCityColumn.colExistentLandfill) = "Não"
            End If
            If city.vPotentialLandfill Then
                wksSelectedCities.Cells(row, SelectedCityColumn.colPotentialLandfill) = "Sim"
            Else
                wksSelectedCities.Cells(row, SelectedCityColumn.colPotentialLandfill) = "Não"
            End If
            row = row + 1
        Next city
        
        'Salva o nome dos municípios na aba de distâncias para facilitar a entrada de dados e visualização
        row = 3
        column = 2
        wksCitiesDistance.range("A2:A100").ClearContents
        wksCitiesDistance.range("B2:DQ2").ClearContents
        For Each city In selectedCities
            wksCitiesDistance.Cells(row, 1) = city.vCityName
            wksCitiesDistance.Cells(2, column) = city.vCityName
            row = row + 1
            column = column + 1
        Next city
        
        Unload Me
        frmStepOne.updateForm
        ThisWorkbook.Save
    Else
        answer = MsgBox(MSG_WRONG_NUMBER_CITIES, vbInformation, MSG_WRONG_NUMBER_CITIES_TITLE)
    End If
    
End Sub

Private Sub cbxUF_Change()
    lstAvailable.Clear
    currentUF = cbxUF
    For Each city In databaseCities
        If city.vUF = cbxUF Then
            If Not IsInCollection(selectedCities, city.vIBGECode) And InStr(LCase(city.vCityName), LCase(txtAvailableSearch.Text)) <> 0 Then
                lstAvailable.AddItem
                lstAvailable.List(lstAvailable.ListCount - 1, 0) = city.vCityName
                lstAvailable.List(lstAvailable.ListCount - 1, 1) = city.vIBGECode
            End If
        End If
    Next city
    
    If cbxUF <> "" Then
        lblSearch.Visible = True
        txtAvailableSearch.Visible = True
    Else
        lblSearch.Visible = False
        txtAvailableSearch.Visible = False
    End If
End Sub

Private Sub txtAvailableSearch_Change()
    lstAvailable.Clear
    currentUF = cbxUF
    For Each city In databaseCities
        If city.vUF = cbxUF Then
            If Not IsInCollection(selectedCities, city.vIBGECode) And InStr(LCase(city.vCityName), LCase(txtAvailableSearch.Text)) <> 0 Then
                lstAvailable.AddItem
                lstAvailable.List(lstAvailable.ListCount - 1, 0) = city.vCityName
                lstAvailable.List(lstAvailable.ListCount - 1, 1) = city.vIBGECode
            End If
        End If
    Next city
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Selecionar Municípios")
    
    'Load database cities
    Set databaseCities = readDatabaseCities
    
    'Load alrady selected cities if available
    Set selectedCities = readSelectedCities
    
    lblSearch.Visible = False
    txtAvailableSearch.Visible = False
    txtAvailableSearch.Text = ""
    
    'Load UF
    For Each city In databaseCities
        inList = False
        For index = 0 To cbxUF.ListCount - 1
            If city.vUF = CStr(cbxUF.List(index)) Then
                inList = True
                Exit For
            End If
        Next index
        If inList = False Then
            Dim added As Boolean
            added = False
            For index = 0 To cbxUF.ListCount - 1
                UF = cbxUF.List(index)
                If city.vUF < UF Then
                    cbxUF.AddItem (city.vUF), index
                    added = True
                    Exit For
                End If
            Next index
            If added = False Then
                cbxUF.AddItem (city.vUF)
            End If
        End If
    Next city
    
    'Show current selected cities
    lstSelected.Clear
    For Each city In selectedCities
        lstSelected.AddItem
        lstSelected.List(lstSelected.ListCount - 1, 0) = city.vCityName
        lstSelected.List(lstSelected.ListCount - 1, 1) = city.vIBGECode
    Next city
    
    With frmSelectCities
        Height = 531
        width = 564
    End With
    
    FormChanged = False
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

