Attribute VB_Name = "modFolder"
Public Function FolderExists(ByVal FolderPath As String) As Boolean

    Dim Fso As Scripting.FileSystemObject
    Set Fso = New Scripting.FileSystemObject
    If Fso.FolderExists(FolderPath) Then
        FolderExists = True
    End If
    
End Function

Public Function HasWriteAccessToFolder(ByVal FolderPath As String) As Boolean

    If Not FolderExists(FolderPath) Then
        Exit Function
    End If
    
    Dim Fso As Scripting.FileSystemObject
    Set Fso = New Scripting.FileSystemObject

    'GET UNIQUE TEMP FilePath, DON'T WANT TO OVERWRITE SOMETHING THAT ALREADY EXISTS
    Do
        Dim count As Integer
        Dim FilePath As String

        FilePath = Fso.BuildPath(FolderPath, "TestWriteAccess" & count & ".tmp")
        count = count + 1
    Loop Until Not Fso.FileExists(FilePath)

    'ATTEMPT TO CREATE THE TMP FILE, ERROR RETURNS FALSE
    On Error GoTo Catch
    Fso.CreateTextFile(FilePath).Write ("Test Folder Access")
    Kill FilePath

    'NO ERROR, ABLE TO WRITE TO FILE; RETURN TRUE!
    HasWriteAccessToFolder = True

Catch:

End Function
