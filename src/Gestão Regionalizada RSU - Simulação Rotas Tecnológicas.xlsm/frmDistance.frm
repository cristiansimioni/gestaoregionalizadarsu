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
    cbxDistanceMethod.AddItem "Google"
    cbxDistanceMethod.AddItem "Euclidiana"
    
    lblProgress.BackColor = ApplicationColors.bgColorLevel2

End Sub
