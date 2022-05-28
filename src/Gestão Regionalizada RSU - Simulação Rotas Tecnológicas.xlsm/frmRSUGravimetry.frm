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

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnDefaultValues_Click()
    txtFoodWaste.Caption = Database.GetDatabaseValue("FoodWaste", colDefaultValue)
    txtGreenWaste.Caption = Database.GetDatabaseValue("GreenWaste", colDefaultValue)
    txtPaper.Caption = Database.GetDatabaseValue("Paper", colDefaultValue)
    txtPlasticFilm.Caption = Database.GetDatabaseValue("PlasticFilm", colDefaultValue)
    txtHardPlastics.Caption = Database.GetDatabaseValue("HardPlastics", colDefaultValue)
    txtGlass.Caption = Database.GetDatabaseValue("Glass", colDefaultValue)
    txtFerrousMetals.Caption = Database.GetDatabaseValue("FerrousMetals", colDefaultValue)
    txtNonFerrousMetals.Caption = Database.GetDatabaseValue("NonFerrousMetals", colDefaultValue)
    txtTextiles.Caption = Database.GetDatabaseValue("Textiles", colDefaultValue)
    txtRubber.Caption = Database.GetDatabaseValue("Rubber", colDefaultValue)
    txtDiapers.Caption = Database.GetDatabaseValue("Diapers", colDefaultValue)
    txtWood.Caption = Database.GetDatabaseValue("Wood", colDefaultValue)
    txtMineralResidues.Caption = Database.GetDatabaseValue("MineralResidues", colDefaultValue)
    txtOthers.Caption = Database.GetDatabaseValue("Others", colDefaultValue)
    
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
    Database.setFoodWaste (FoodWaste)
    Database.setGreenWaste (GreenWaste)
    Database.setPaper (Paper)
    Database.setPlasticFilm (PlasticFilm)
    Database.setHardPlastics (HardPlastics)
    Database.setGlass (Glass)
    Database.setFerrousMetals (FerrousMetals)
    Database.setNonFerrousMetals (NonFerrousMetals)
    Database.setTextiles (Textiles)
    Database.setRubber (Rubber)
    Database.setDiapers (Diapers)
    Database.setWood (Wood)
    Database.setMineralResidues (MineralResidues)
    Database.setOthers (Others)
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
    Me.Caption = APPNAME & " - xxxx"
    Me.BackColor = ApplicationColors.bgColorLevel3
    
    Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Or TypeName(Ctrl) = "CommandButton" Then
            Ctrl.BackColor = ApplicationColors.btColorLevel3
         End If
    Next Ctrl
    
    'Preenche valores padrão nas labels
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
    
End Sub

Sub calculateTotal()
    Dim total As Double
    total = FoodWaste + GreenWaste + Paper + PlasticFilm + HardPlastics + Glass + FerrousMetals + NonFerrousMetals + Textiles + Rubber + Diapers + Wood + MineralResidues + Others
    lblTotal.Caption = total
    If total > 100 Then
        lblTotal.BackColor = ApplicationColors.bgColorInvalidTextBox
    Else
        lblTotal.BackColor = ApplicationColors.bgColorValidTextBox
    End If
End Sub


