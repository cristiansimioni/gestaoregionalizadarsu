VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   ClientHeight    =   5625
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   5925
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    lblAppName = APPNAME
    lblAppSubname = APPSUBNAME
    lblVersion = "Versão: " & APPVERSION
    lblReleaseDate = "Última Atualização: " & APPLASTUPDATED
End Sub
