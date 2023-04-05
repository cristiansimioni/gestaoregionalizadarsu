VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTaxes 
   Caption         =   "UserForm1"
   ClientHeight    =   3555
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   8400.001
   OleObjectBlob   =   "frmTaxes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTaxes"
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

Private Sub txtISSTax_Change()
Call modForm.textBoxChange(txtISSTax, "ISSTax", FormChanged)
End Sub
Private Sub txtICMSTax_Change()
Call modForm.textBoxChange(txtICMSTax, "ICMSTax", FormChanged)
End Sub
Private Sub txtCSLLTax_Change()
Call modForm.textBoxChange(txtCSLLTax, "CSLLTax", FormChanged)
End Sub
Private Sub txtIRPJTax_Change()
Call modForm.textBoxChange(txtIRPJTax, "IRPJTax", FormChanged)
End Sub

Private Sub UserForm_Initialize()
    'Form Appearance
    Call modForm.applyLookAndFeel(Me, 3, "Impostos")
    
    txtISSTax = Database.GetDatabaseValue("ISSTax", colUserValue)
    txtICMSTax = Database.GetDatabaseValue("ICMSTax", colUserValue)
    txtCSLLTax = Database.GetDatabaseValue("CSLLTax", colUserValue)
    txtIRPJTax = Database.GetDatabaseValue("IRPJTax", colUserValue)

    FormChanged = False
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrorHandler
        Call Database.SetDatabaseValue("ISSTax", colUserValue, CDbl(txtISSTax.Text))
        Call Database.SetDatabaseValue("ICMSTax", colUserValue, CDbl(txtICMSTax.Text))
        Call Database.SetDatabaseValue("CSLLTax", colUserValue, CDbl(txtCSLLTax.Text))
        Call Database.SetDatabaseValue("IRPJTax", colUserValue, CDbl(txtIRPJTax.Text))

        FormChanged = False
        frmStepThree.updateForm
        Unload Me
        ThisWorkbook.Save
    Exit Sub
    
ErrorHandler:
    Call MsgBox(MSG_INVALID_DATA, vbCritical, MSG_INVALID_DATA_TITLE)
End Sub

Private Sub btnDefault_Click()
    txtISSTax = Database.GetDatabaseValue("ISSTax", colDefaultValue)
    txtICMSTax = Database.GetDatabaseValue("ICMSTax", colDefaultValue)
    txtCSLLTax = Database.GetDatabaseValue("CSLLTax", colDefaultValue)
    txtIRPJTax = Database.GetDatabaseValue("IRPJTax", colDefaultValue)
End Sub

