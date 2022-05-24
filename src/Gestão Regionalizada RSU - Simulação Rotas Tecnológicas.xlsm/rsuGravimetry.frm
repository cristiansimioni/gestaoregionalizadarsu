VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} rsuGravimetry 
   Caption         =   "Dados de Gravimetria do RSU"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235.001
   OleObjectBlob   =   "rsuGravimetry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "rsuGravimetry"
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
    txtFoodWaste.Text = Database.getFoodWaste(True)
    txtGreenWaste.Text = Database.getGreenWaste(True)
    txtPaper.Text = Database.getPaper(True)
    txtPlasticFilm.Text = Database.getPlasticFilm(True)
    txtHardPlastics.Text = Database.getHardPlastics(True)
    txtGlass.Text = Database.getGlass(True)
    txtFerrousMetals.Text = Database.getFerrousMetals(True)
    txtNonFerrousMetals.Text = Database.getNonFerrousMetals(True)
    txtTextiles.Text = Database.getTextiles(True)
    txtRubber.Text = Database.getRubber(True)
    txtDiapers.Text = Database.getDiapers(True)
    txtWood.Text = Database.getWood(True)
    txtMineralResidues.Text = Database.getMineralResidues(True)
    txtOthers.Text = Database.getOthers(True)
    
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
        txt.BackColor = Util.xColorGreen
        txt.ControlTipText = errorMsg
        value = (CDbl(txt.Text))
        Call calculateTotal
    Else
        If IsNumeric(txt.Text) Then
            value = (CDbl(txt.Text))
            Call calculateTotal
        End If
        txt.BackColor = Util.xColorRed
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
    
    'Preenche valores padrão nas labels
    lblFoodWaste.Caption = Database.getFoodWaste(True)
    lblGreenWaste.Caption = Database.getGreenWaste(True)
    lblPaper.Caption = Database.getPaper(True)
    lblPlasticFilm.Caption = Database.getPlasticFilm(True)
    lblHardPlastics.Caption = Database.getHardPlastics(True)
    lblGlass.Caption = Database.getGlass(True)
    lblFerrousMetals.Caption = Database.getFerrousMetals(True)
    lblNonFerrousMetals.Caption = Database.getNonFerrousMetals(True)
    lblTextiles.Caption = Database.getTextiles(True)
    lblRubber.Caption = Database.getRubber(True)
    lblDiapers.Caption = Database.getDiapers(True)
    lblWood.Caption = Database.getWood(True)
    lblMineralResidues.Caption = Database.getMineralResidues(True)
    lblOthers.Caption = Database.getOthers(True)
    
    FoodWaste = Database.getFoodWaste
    GreenWaste = Database.getGreenWaste
    Paper = Database.getPaper
    PlasticFilm = Database.getPlasticFilm
    HardPlastics = Database.getHardPlastics
    Glass = Database.getGlass
    FerrousMetals = Database.getFerrousMetals
    NonFerrousMetals = Database.getNonFerrousMetals
    Textiles = Database.getTextiles
    Rubber = Database.getRubber
    Diapers = Database.getDiapers
    Wood = Database.getWood
    MineralResidues = Database.getMineralResidues
    Others = Database.getOthers
    
    If Others <> 0 Then
        txtFoodWaste.Text = Database.getFoodWaste
        txtGreenWaste.Text = Database.getGreenWaste
        txtPaper.Text = Database.getPaper
        txtPlasticFilm.Text = Database.getPlasticFilm
        txtHardPlastics.Text = Database.getHardPlastics
        txtGlass.Text = Database.getGlass
        txtFerrousMetals.Text = Database.getFerrousMetals
        txtNonFerrousMetals.Text = Database.getNonFerrousMetals
        txtTextiles.Text = Database.getTextiles
        txtRubber.Text = Database.getRubber
        txtDiapers.Text = Database.getDiapers
        txtWood.Text = Database.getWood
        txtMineralResidues.Text = Database.getMineralResidues
        txtOthers.Text = Database.getOthers
    End If
    
    Call calculateTotal
    
End Sub

Sub calculateTotal()
    Dim total As Double
    total = FoodWaste + GreenWaste + Paper + PlasticFilm + HardPlastics + Glass + FerrousMetals + NonFerrousMetals + Textiles + Rubber + Diapers + Wood + MineralResidues + Others
    lblTotal.Caption = total
    If total > 100 Then
        lblTotal.BackColor = xColorRed
    Else
        lblTotal.BackColor = xColorGreen
    End If
End Sub


