VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   ClientHeight    =   1872
   ClientLeft      =   168
   ClientTop       =   576
   ClientWidth     =   2448
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Inicializa o formulário
Private Sub UserForm_Initialize()
    lblAppName = APPNAME
    lblAppSubname = APPSUBNAME
    lblVersion = "Versão: " & APPVERSION
    lblReleaseDate = "Última Atualização: " & APPLASTUPDATED
    
    Me.Height = 316
    Me.width = 310
End Sub
