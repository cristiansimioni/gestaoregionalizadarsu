Attribute VB_Name = "GoogleAPI"
Option Explicit

Function Km_Distancia(Origin As String, Destination As String) As Double
    'Requer referência ao: 'Microsoft XML, v6.0'

    Dim Solicitacao As XMLHTTP60
    Dim Doc As DOMDocument60
    Dim Distancia_Pontos As IXMLDOMNode
    Dim Url As String
    

    Let Km_Distancia = 0

    'Checa e limpa as entradas
    On Error GoTo Sair

    Let Origin = Replace(Origin, " ", "+")
    Let Destination = Replace(Destination, " ", "+")

    ' Le os dados XML da API do Google Maps.
    Set Solicitacao = New XMLHTTP60
    
    Url = "https://maps.googleapis.com/maps/api/directions/xml?origin=" _
        & Origin & "&destination=" & Destination _
        & "&key=<KEY>"
        
    Solicitacao.Open "GET", Url, False
    Solicitacao.send
    
    'https://maps.googleapis.com/maps/api/directions/json?origin=Disneyland&destination=Universal+Studios+Hollywood&key=YOUR_API_KEY

    ' Tornando o XML legível por usar o XPath
    
    Set Doc = New DOMDocument60

    Doc.LoadXML Solicitacao.responseText

    ' Obtendo o valor da distância entre os nós.
    Set Distancia_Pontos = Doc.SelectSingleNode("//leg/distance/value")
    If Not Distancia_Pontos Is Nothing Then Km_Distancia = Distancia_Pontos.Text / 1000

Sair:
    ' Tidy up
    Set Distancia_Pontos = Nothing
    Set Doc = Nothing
    Set Solicitacao = Nothing
End Function

