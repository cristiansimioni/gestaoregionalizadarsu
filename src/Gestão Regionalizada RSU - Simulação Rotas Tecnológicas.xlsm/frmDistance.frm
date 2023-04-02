VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDistance 
   Caption         =   "UserForm1"
   ClientHeight    =   4755
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8355.001
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
    'cbxDistanceMethod.AddItem "Google"
    cbxDistanceMethod.AddItem "Euclidiana"

End Sub
