VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectArrays 
   Caption         =   "UserForm1"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16320
   OleObjectBlob   =   "frmSelectArrays.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label33_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim arrays As Collection
    Set arrays = readArrays
    
    txtArray1.Text = arrays(1).vArrayRaw
    txtArray2.Text = arrays(2).vArrayRaw
End Sub
