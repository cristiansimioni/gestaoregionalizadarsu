VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDistance 
   Caption         =   "UserForm1"
   ClientHeight    =   5124
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8424.001
   OleObjectBlob   =   "frmDistance.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDistance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnCalculate_Click()
    Dim result As Boolean
    Dim cities As New Collection
    Set cities = readSelectedCities()
    
    If cbxDistanceMethod.value = "Bing" Then
        result = modDistance.calculateDistance(DistanceMethod.bing, cities, Me.lblProgress, txtAPIKey.Text)
    ElseIf cbxDistanceMethod.value = "Euclidiana" Then
        result = modDistance.calculateDistance(DistanceMethod.euclidean, cities, Me.lblProgress)
    End If
    
    If result Then
        MsgBox "OK"
    Else
        MsgBox "Deu ruim!"
    End If
    
End Sub

Private Sub cbxDistanceMethod_Change()
    If cbxDistanceMethod.value = "Bing" Then
        txtAPIKey.Enabled = True
        txtAPIKey.BorderColor = vbBlack
        lblAPIKey.Enabled = True
        lblDescription.Caption = "Para utilizar esse método é necessário a geração de uma chave API (API Key) conforme descrito no manual " & _
                                 "do usuário. Verifique a quantidade de requests que serão gerados, pois o máximo permitido por dia são " & _
                                 "3000 mil requests."
    Else
        txtAPIKey.Enabled = False
        txtAPIKey.BorderColor = vbScrollBars
        lblAPIKey.Enabled = False
        lblDescription.Caption = "Atenção: esse método serve apenas para simular de maneira mais rápida um cenário. A distância euclidinada " & _
                                 "calcula a distância em linha reta entre dois municípios e o uso desse método irá gerar distorções no " & _
                                 "resultado final. Para simulções precisas, o método recomendado é o do Bing ou inserção manual."
    End If
End Sub

Private Sub UserForm_Initialize()
    Call modForm.applyLookAndFeel(Me, 2, "Calcular Distâncias")
    
    cbxDistanceMethod.AddItem "Bing"
    cbxDistanceMethod.AddItem "Google"
    cbxDistanceMethod.AddItem "Euclidiana"
    
    lblProgress.BackColor = ApplicationColors.bgColorLevel2

End Sub
