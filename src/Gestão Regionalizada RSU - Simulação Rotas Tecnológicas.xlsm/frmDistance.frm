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
            MsgBox "A chave " & txtAPIKey.Text & " n�o � uma chave v�lida. Favor verificar!", vbCritical, "Erro"
            Exit Sub
        End If
    ElseIf cbxDistanceMethod.value = "Euclidiana" Then
        result = modDistance.calculateDistance(DistanceMethod.euclidean, cities, Me)
    End If
    
    If result Then
        MsgBox "Dist�ncias calculadas com sucesso!", vbInformation, "Sucesso"
        Unload Me
    Else
        MsgBox "Algo deu errado ao calcular as dist�ncias.", vbCritical, "Erro"
    End If
    
End Sub

Private Sub cbxDistanceMethod_Change()
    If cbxDistanceMethod.value = "Bing" Then
        txtAPIKey.Enabled = True
        txtAPIKey.BorderColor = vbBlack
        lblAPIKey.Enabled = True
        lblDescription.Caption = "Para utilizar esse m�todo � necess�rio a gera��o de uma chave API (API Key) conforme descrito no manual " & _
                                 "do usu�rio. Verifique a quantidade de requests que ser�o gerados, pois o m�ximo permitido por dia s�o " & _
                                 "3000 mil requests."
    Else
        txtAPIKey.Enabled = False
        txtAPIKey.BorderColor = vbScrollBars
        lblAPIKey.Enabled = False
        lblDescription.Caption = "Aten��o: esse m�todo serve apenas para simular de maneira mais r�pida um cen�rio. A dist�ncia euclidinada " & _
                                 "calcula a dist�ncia em linha reta entre dois munic�pios e o uso desse m�todo ir� gerar distor��es no " & _
                                 "resultado final. Para simul��es precisas, o m�todo recomendado � o do Bing ou inser��o manual."
    End If
End Sub

Private Sub UserForm_Initialize()
    Call modForm.applyLookAndFeel(Me, 2, "Calcular Dist�ncias")
    
    cbxDistanceMethod.AddItem "Bing"
    'cbxDistanceMethod.AddItem "Google"
    cbxDistanceMethod.AddItem "Euclidiana"

End Sub
