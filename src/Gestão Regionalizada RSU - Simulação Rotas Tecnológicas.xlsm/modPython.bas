Attribute VB_Name = "modPython"
Option Explicit

Public Function RunPythonScript(ByVal algPath As String, ByVal prjName As String)

'Declare Variables
Dim PythonExe, PythonScript, Params, cmd As String
Dim wsh As Object
Set wsh = VBA.CreateObject("WScript.Shell")
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1
Dim errorCode As Integer

'Provide file path to Python.exe
PythonExe = Database.GetDatabaseValue("PythonPath", colUserValue)
PythonScript = Chr(34) & Application.ThisWorkbook.Path & "\src\combinations\combinations.py" & Chr(34)

Dim maxCluster, maxSubarrays, trashThreshold, capexInbound, capexOutbound, paymentPeriod, movimentationCost, landfillDeviation As Double
maxCluster = Database.GetDatabaseValue("MaxClusters", colUserValue)
maxSubarrays = Database.GetDatabaseValue("MaxSubarrays", colUserValue)
trashThreshold = Database.GetDatabaseValue("TrashThreshold", colUserValue)
capexInbound = Database.GetDatabaseValue("CapexInbound", colUserValue)
capexOutbound = Database.GetDatabaseValue("CapexOutbound", colUserValue)
paymentPeriod = Database.GetDatabaseValue("ExpectedDeadline", colUserValue)
movimentationCost = (100 - Database.GetDatabaseValue("ReducingCostMovimentation", colUserValue)) / 100#
landfillDeviation = (100 - Database.GetDatabaseValue("LandfillDeviationTarget", colUserValue)) / 100#

Params = Chr(34) & algPath & "\cidades-" & prjName & ".csv" & Chr(34) & _
         " " & _
         Chr(34) & algPath & "\distancias-" & prjName & ".csv" & Chr(34) & _
         " " & _
         maxCluster & " " & maxSubarrays & " " & trashThreshold & " " & capexInbound & " " & capexOutbound & _
         " " & paymentPeriod & " " & Replace(CStr(movimentationCost), ",", ".") & _
         " " & Replace(CStr(landfillDeviation), ",", ".") & _
         " " & _
         Chr(34) & algPath & "\relatório-" & prjName & ".txt" & Chr(34) & _
         " " & _
         Chr(34) & algPath & "\output-" & prjName & ".csv" & Chr(34)

cmd = "%comspec% /c " & Chr(34) & Chr(34) & PythonExe & Chr(34) & " " & PythonScript & " " & Params & Chr(34)
'Run the Python Script
errorCode = wsh.Run(cmd, windowStyle, waitOnReturn)

If errorCode = 0 Then
    RunPythonScript = True
Else
    Debug.Print cmd
    RunPythonScript = False
End If

End Function

