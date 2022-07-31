Attribute VB_Name = "Folder"
Option Explicit
Option Compare Database

Public Function DoesFolderExist(ByRef folderPath As String) As Boolean

    DoesFolderExist = Dir(folderPath, vbDirectory) <> ""
      
End Function

Public Function CreateFolder(ByRef folderPath As String) As Boolean

    If Not DoesFolderExist(folderPath) Then
        MkDir folderPath
    End If
      CreateFolder = DoesFolderExist(folderPath)
End Function

