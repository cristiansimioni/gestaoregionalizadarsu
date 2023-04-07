VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public vCityName As String
Public vIBGECode As Double
Public vUF As String
Public vUFCode
Public vLatitude As Double
Public vLongitude As Double
Public vPopulation As Double
Public vTrash As Double
Public vConventionalCost As Double
Public vTransshipmentCost As Double
Public vCostPostTransshipment As Double
Public vUTVR As Boolean
Public vExistentLandfill As Boolean
Public vPotentialLandfill As Boolean

Public Function vUTVRAsString()
    If Me.vUTVR Then
        vUTVRAsString = "Sim"
    Else
        vUTVRAsString = "Não"
    End If
End Function

Public Function vExistentLandfillAsString()
    If Me.vExistentLandfill Then
        vExistentLandfillAsString = "Sim"
    Else
        vExistentLandfillAsString = "Não"
    End If
End Function

Public Function vPotentialLandfillAsString()
    If Me.vPotentialLandfill Then
        vPotentialLandfillAsString = "Sim"
    Else
        vPotentialLandfillAsString = "Não"
    End If
End Function



