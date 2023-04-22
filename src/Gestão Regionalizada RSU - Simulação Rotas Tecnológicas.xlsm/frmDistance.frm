VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDistance 
   Caption         =   "UserForm1"
   ClientHeight    =   3192
   ClientLeft      =   -15
   ClientTop       =   -120
   ClientWidth     =   4365
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
        If modDistance.validateBingKey(txtAPIKey.Text) Then
            result = modDistance.calculateDistance(DistanceMethod.bing, cities, Me, txtAPIKey.Text)
        Else
            MsgBox "A chave " & txtAPIKey.Text & " não é uma chave válida. Favor verificar!", vbCritical, "Erro"
            Exit Sub
        End If
    ElseIf cbxDistanceMethod.value = "Euclidiana" Then
        result = modDistance.calculateDistance(DistanceMethod.euclidean, cities, Me)
    End If
    
    If result Then
        MsgBox "Distâncias calculadas com sucesso!", vbInformation, "Sucesso"
        Unload Me
    Else
        MsgBox "Algo deu errado ao calcular as distâncias.", vbCritical, "Erro"
    End If
    
End Sub

Private Sub cbxDistanceMethod_Change()
    If cbxDistanceMethod.value = "Bing" Then
        txtAPIKey.Enabled = True
        txtAPIKey.BorderColor = vbBlack
        lblAPIKey.Enabled = True
        lblDescription.Caption = "Para utilizar esse método é necessário a geração de uma chave API (API Key) conforme descrito no manual " & _
                                 "do usuário. Verifique a quantidade de requests que serão gerados antes de continuar. A quantidade máxima de requests " & _
                                 "gratuitos por ano é de 125.000 mil requests. Em um consórcio de 50 municípios, por exemplo, serão necessários 2500 requests (50 municípios x 50 municípios) para cada vez " & _
                                 "que o botão calcular for acionado. Ao continuar você aceita os termos da plataforma Bing."
    Else
        txtAPIKey.Enabled = False
        txtAPIKey.BorderColor = vbScrollBars
        lblAPIKey.Enabled = False
        lblDescription.Caption = "Atenção: esse método serve apenas para simular de maneira mais rápida um cenário. A distância euclidinada " & _
                                 "calcula a distância em linha reta entre dois municípios e o uso desse método irá gerar distorções no " & _
                                 "resultado final. Portanto seu uso NÃO é recomendado para análises finais. Para simulções precisas, o método recomendado " & _
                                 "é o do Bing, que calcula a melhor rota entre dois municípios ou a inserção manual das distâncias."
    End If
End Sub

Private Sub UserForm_Initialize()
    Call modForm.applyLookAndFeel(Me, 2, "Calcular Distâncias")
    
    cbxDistanceMethod.AddItem "Bing"
    cbxDistanceMethod.AddItem "Euclidiana"
    
    Me.Height = 278
    Me.width = 437
End Sub
