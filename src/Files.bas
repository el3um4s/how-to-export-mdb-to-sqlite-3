Attribute VB_Name = "Files"
Option Explicit
Option Compare Database

Public Function DoesFileExist(ByRef filePath) As Boolean

    DoesFileExist = Dir(filePath) <> ""
      
End Function

Public Function CopyAFile(ByRef filePath As String, ByRef destinationFile As String) As Boolean

    If DoesFileExist(filePath) Then
        FileCopy filePath, destinationFile
    End If
    
    CopyAFile = Dir(destinationFile) <> ""
End Function

Public Function RenameAFile(ByRef currentName As String, ByRef newName As String) As Boolean

    If DoesFileExist(currentName) Then
        Name currentName As newName
    End If
    
    RenameAFile = Dir(newName) <> ""
End Function


Public Function MoveAFile(ByRef filePath As String, ByRef destinationFile As String) As Boolean

    If DoesFileExist(filePath) Then
        Name filePath As destinationFile
    End If
    
    MoveAFile = Dir(destinationFile) <> ""
End Function


Public Function CopyAFileDeletingOld(ByRef filePath As String, ByRef destinationFile As String) As Boolean
    If DoesFileExist(filePath) Then
        If DoesFileExist(destinationFile) Then
            Kill destinationFile
        End If
        FileCopy filePath, destinationFile
    End If
    
    CopyAFileDeletingOld = Dir(destinationFile) <> ""
End Function
