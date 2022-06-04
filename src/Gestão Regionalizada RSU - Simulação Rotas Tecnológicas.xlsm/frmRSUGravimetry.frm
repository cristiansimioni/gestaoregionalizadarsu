VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRSUGravimetry 
   Caption         =   "Dados de Gravimetria do RSU"
   ClientHeight    =   8265.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7575
   OleObjectBlob   =   "frmRSUGravimetry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRSUGravimetry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FoodWaste As Double
Dim GreenWaste As Double
Dim Paper As Double
Dim PlasticFilm As Double
Dim HardPlastics As Double
Dim Glass As Double
Dim FerrousMetals As Double
Dim NonFerrousMetals As Double
Dim Textiles As Double
Dim Rubber As Double
Dim Diapers As Double
Dim Wood As Double
Dim MineralResidues As Double
Dim Others As Double
Dim FormChanged As Boolean

Function validateForm() As Boolean
    validateForm = True
End Function

Private Sub btnBack_Click()
    If FormChanged Then
        answer = MsgBox("Você realizou alterações, gostaria de salvar?", vbQuestion + vbYesNo + vbDefaultButton2, "Salvar Alterações")
        If answer = vbYes Then
          Call btnSave_Click
        Else
          Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub btnDefaultValues_Click()
    'Set default values
    txtFoodWaste.Text = Database.GetDatabaseValue("FoodWaste", colDefaultValue)
    txtGreenWaste.Text = Database.GetDatabaseValue("GreenWaste", colDefaultValue)
    txtPaper.Text = Database.GetDatabaseValue("Paper", colDefaultValue)
    txtPlasticFilm.Text = Database.GetDatabaseValue("PlasticFilm", colDefaultValue)
    txtHardPlastics.Text = Database.GetDatabaseValue("HardPlastics", colDefaultValue)
    txtGlass.Text = Database.GetDatabaseValue("Glass", colDefaultValue)
    txtFerrousMetals.Text = Database.GetDatabaseValue("FerrousMetals", colDefaultValue)
    txtNonFerrousMetals.Text = Database.GetDatabaseValue("NonFerrousMetals", colDefaultValue)
    txtTextiles.Text = Database.GetDatabaseValue("Textiles", colDefaultValue)
    txtRubber.Text = Database.GetDatabaseValue("Rubber", colDefaultValue)
    txtDiapers.Text = Database.GetDatabaseValue("Diapers", colDefaultValue)
    txtWood.Text = Database.GetDatabaseValue("Wood", colDefaultValue)
    txtMineralResidues.Text = Database.GetDatabaseValue("MineralResidues", colDefaultValue)
    txtOthers.Text = Database.GetDatabaseValue("Others", colDefaultValue)
    
    FoodWaste = CDbl(txtFoodWaste.Text)
    GreenWaste = CDbl(txtGreenWaste.Text)
    Paper = CDbl(txtPaper.Text)
    PlasticFilm = CDbl(txtPlasticFilm.Text)
    HardPlastics = CDbl(txtHardPlastics.Text)
    Glass = CDbl(txtGlass.Text)
    FerrousMetals = CDbl(txtFerrousMetals.Text)
    NonFerrousMetals = CDbl(txtNonFerrousMetals.Text)
    Textiles = CDbl(txtTextiles.Text)
    Rubber = CDbl(txtRubber.Text)
    Diapers = CDbl(txtDiapers.Text)
    Wood = CDbl(txtWood.Text)
    MineralResidues = CDbl(txtMineralResidues.Text)
    Others = CDbl(txtOthers.Text)
    
    Call calculateTotal
End Sub

Private Sub btnSave_Click()
    If validateForm() Then
        Call Database.SetDatabaseValue("FoodWaste", colUserValue, FoodWaste)
        Call Database.SetDatabaseValue("GreenWaste", colUserValue, GreenWaste)
        Call Database.SetDatabaseValue("Paper", colUserValue, Paper)
        Call Database.SetDatabaseValue("PlasticFilm", colUserValue, PlasticFilm)
        Call Database.SetDatabaseValue("HardPlastics", colUserValue, HardPlastics)
        Call Database.SetDatabaseValue("Glass", colUserValue, Glass)
        Call Database.SetDatabaseValue("FerrousMetals", colUserValue, FerrousMetals)
        Call Database.SetDatabaseValue("NonFerrousMetals", colUserValue, NonFerrousMetals)
        Call Database.SetDatabaseValue("Textiles", colUserValue, Textiles)
        Call Database.SetDatabaseValue("Rubber", colUserValue, Rubber)
        Call Database.SetDatabaseValue("Diapers", colUserValue, Diapers)
        Call Database.SetDatabaseValue("Wood", colUserValue, Wood)
        Call Database.SetDatabaseValue("MineralResidues", colUserValue, MineralResidues)
        Call Database.SetDatabaseValue("Others", colUserValue, Others)
        FormChanged = False
        Unload Me
    Else
        answer = MsgBox("Valores inválidos. Favor verificar!", vbExclamation, "Dados inválidos")
    End If
End Sub

Private Sub txtDiapers_Change()
    Call validateEntry(txtDiapers, Diapers)
End Sub

Private Sub txtFerrousMetals_Change()
    Call validateEntry(txtFerrousMetals, FerrousMetals)
End Sub

Private Sub txtFoodWaste_Change()
    Call validateEntry(txtFoodWaste, FoodWaste)
End Sub

Private Sub txtGlass_Change()
    Call validateEntry(txtGlass, Glass)
End Sub

Private Sub txtGreenWaste_Change()
    Call validateEntry(txtGreenWaste, GreenWaste)
End Sub

Private Sub txtHardPlastics_Change()
    Call validateEntry(txtHardPlastics, HardPlastics)
End Sub

Private Sub txtMineralResidues_Change()
    Call validateEntry(txtMineralResidues, MineralResidues)
End Sub

Private Sub txtNonFerrousMetals_Change()
    Call validateEntry(txtNonFerrousMetals, NonFerrousMetals)
End Sub

Private Sub txtOthers_Change()
    Call validateEntry(txtOthers, Others)
End Sub

Private Sub txtPaper_Change()
    Call validateEntry(txtPaper, Paper)
End Sub



Private Sub validateEntry(ByRef txt, ByRef value As Double)
    Dim errorMsg As String
    If Util.validateRange(txt.Text, 0#, 100#, errorMsg) Then
        txt.BackColor = ApplicationColors.bgColorValidTextBox
        txt.ControlTipText = errorMsg
        value = (CDbl(txt.Text))
        Call calculateTotal
    Else
        If IsNumeric(txt.Text) Then
            value = (CDbl(txt.Text))
            Call calculateTotal
        End If
        txt.BackColor = ApplicationColors.bgColorInvalidTextBox
        txt.ControlTipText = errorMsg
    End If
End Sub

Private Sub txtPlasticFilm_Change()
    Call validateEntry(txtPlasticFilm, PlasticFilm)
End Sub

Private Sub txtRubber_Change()
    Call validateEntry(txtRubber, Rubber)
End Sub

Private Sub txtTextiles_Change()
    Call validateEntry(txtTextiles, Textiles)
End Sub

Private Sub txtWood_Change()
    Call validateEntry(txtWood, Wood)
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Me.Caption = APPNAME & " - Gravimetria do RSU"
    Me.BackColor = ApplicationColors.bgColorLevel3
    
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.btColorLevel3
         End If
    Next Ctrl
    
    'Read database values (default)
    lblFoodWaste.Caption = Database.GetDatabaseValue("FoodWaste", colDefaultValue)
    lblGreenWaste.Caption = Database.GetDatabaseValue("GreenWaste", colDefaultValue)
    lblPaper.Caption = Database.GetDatabaseValue("Paper", colDefaultValue)
    lblPlasticFilm.Caption = Database.GetDatabaseValue("PlasticFilm", colDefaultValue)
    lblHardPlastics.Caption = Database.GetDatabaseValue("HardPlastics", colDefaultValue)
    lblGlass.Caption = Database.GetDatabaseValue("Glass", colDefaultValue)
    lblFerrousMetals.Caption = Database.GetDatabaseValue("FerrousMetals", colDefaultValue)
    lblNonFerrousMetals.Caption = Database.GetDatabaseValue("NonFerrousMetals", colDefaultValue)
    lblTextiles.Caption = Database.GetDatabaseValue("Textiles", colDefaultValue)
    lblRubber.Caption = Database.GetDatabaseValue("Rubber", colDefaultValue)
    lblDiapers.Caption = Database.GetDatabaseValue("Diapers", colDefaultValue)
    lblWood.Caption = Database.GetDatabaseValue("Wood", colDefaultValue)
    lblMineralResidues.Caption = Database.GetDatabaseValue("MineralResidues", colDefaultValue)
    lblOthers.Caption = Database.GetDatabaseValue("Others", colDefaultValue)
    
    'Read database values (user)
    FoodWaste = Database.GetDatabaseValue("FoodWaste", colUserValue)
    GreenWaste = Database.GetDatabaseValue("GreenWaste", colUserValue)
    Paper = Database.GetDatabaseValue("Paper", colUserValue)
    PlasticFilm = Database.GetDatabaseValue("PlasticFilm", colUserValue)
    HardPlastics = Database.GetDatabaseValue("HardPlastics", colUserValue)
    Glass = Database.GetDatabaseValue("Glass", colUserValue)
    FerrousMetals = Database.GetDatabaseValue("FerrousMetals", colUserValue)
    NonFerrousMetals = Database.GetDatabaseValue("NonFerrousMetals", colUserValue)
    Textiles = Database.GetDatabaseValue("Textiles", colUserValue)
    Rubber = Database.GetDatabaseValue("Rubber", colUserValue)
    Diapers = Database.GetDatabaseValue("Diapers", colUserValue)
    Wood = Database.GetDatabaseValue("Wood", colUserValue)
    MineralResidues = Database.GetDatabaseValue("MineralResidues", colUserValue)
    Others = Database.GetDatabaseValue("Others", colUserValue)
    
    'Only show the data if it's available
    If Others <> 0 Then
        txtFoodWaste.Text = FoodWaste
        txtGreenWaste.Text = GreenWaste
        txtPaper.Text = Paper
        txtPlasticFilm.Text = PlasticFilm
        txtHardPlastics.Text = HardPlastics
        txtGlass.Text = Glass
        txtFerrousMetals.Text = FerrousMetals
        txtNonFerrousMetals.Text = NonFerrousMetals
        txtTextiles.Text = Textiles
        txtRubber.Text = Rubber
        txtDiapers.Text = Diapers
        txtWood.Text = Wood
        txtMineralResidues.Text = MineralResidues
        txtOthers.Text = Others
    End If
    
    Call calculateTotal
    
    FormChanged = False
End Sub

Sub calculateTotal()
    Dim total As Double
    total = FoodWaste + GreenWaste + Paper + PlasticFilm + HardPlastics + _
            Glass + FerrousMetals + NonFerrousMetals + Textiles + Rubber + _
            Diapers + Wood + MineralResidues + Others
    lblTotal.Caption = total
    If total > 100 Then
        lblTotal.BackColor = ApplicationColors.bgColorInvalidTextBox
    Else
        lblTotal.BackColor = ApplicationColors.bgColorValidTextBox
    End If
End Sub


