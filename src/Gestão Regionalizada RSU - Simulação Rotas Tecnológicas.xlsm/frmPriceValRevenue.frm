VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPriceValRevenue 
   Caption         =   "UserForm1"
   ClientHeight    =   1728
   ClientLeft      =   84
   ClientTop       =   288
   ClientWidth     =   8340.001
   OleObjectBlob   =   "frmPriceValRevenue.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPriceValRevenue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim FormChanged As Boolean

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


Private Sub txtExtraordinaryTariffAffordability_Change()
    Call modForm.textBoxChange(txtExtraordinaryTariffAffordability, "ExtraordinaryTariffAffordability", FormChanged)
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Receitas Extraordinárias p/ Modicidade Tarifária")
    
    txtExtraordinaryTariffAffordability = Database.GetDatabaseValue("ExtraordinaryTariffAffordability", colUserValue)
    
    FormChanged = False
    
    Me.Height = 164
    Me.width = 532
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrorHandler
        Call Database.SetDatabaseValue("ExtraordinaryTariffAffordability", colUserValue, CDbl(txtExtraordinaryTariffAffordability.Text))
        FormChanged = False
        frmStepFour.updateForm
        Unload Me
        ThisWorkbook.Save
    Exit Sub
    
ErrorHandler:
    Call MsgBox(MSG_INVALID_DATA, vbCritical, MSG_INVALID_DATA_TITLE)
End Sub

Private Sub btnDefault_Click()
    txtExtraordinaryTariffAffordability = Database.GetDatabaseValue("ExtraordinaryTariffAffordability", colDefaultValue)
End Sub

