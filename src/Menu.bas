Attribute VB_Name = "Menu"
Option Explicit
Option Compare Database

Public Function SelectDatabase() As String
    Dim path As String
    path = GetOpenFile()
    
    Forms!Menu!pathDatabase = path
    
    SelectDatabase = path
End Function


Public Function SelectDestinationFolder() As String
    Dim path As String
    path = BrowseFolder("Select destination folder")
    
    Forms!Menu!destinationFolder = path
    
    SelectDestinationFolder = path
End Function

Public Function ShowListTable() As Boolean
    Forms!Menu!tableList.RowSource = ""
    Dim database As String
    database = Forms!Menu!pathDatabase

    Dim db As DAO.database
    Dim tdf As DAO.TableDef
    
    Set db = OpenDatabase(database, False)
    
    For Each tdf In db.TableDefs
        If Not (tdf.name Like "MSys*" Or tdf.name Like "~*") Then
            If Forms!Menu!tableList.RowSource = "" Then
                Forms!Menu!tableList.RowSource = tdf.name
            Else
                ' https://www.599cd.com/tips/access/listbox-additem-2000/
                Forms!Menu!tableList.RowSource = Forms!Menu!tableList.RowSource & ";" & tdf.name
            End If

        End If
    
    Next
    ShowListTable = True
End Function

Public Function SelectAllTables() As Boolean
    ListBoxSelectAll Forms!Menu!tableList
    SelectAllTables = ListBoxSelectAll(Forms!Menu!tableList)
End Function


Public Function nameNewDatabaseFromOriginalPath() As String
    Dim strFullPath As String
    strFullPath = Forms!Menu!pathDatabase
    
    Dim nameWithExtension As String
    nameWithExtension = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))

    Dim name As String
    name = Left(nameWithExtension, Len(nameWithExtension) - 3) & "db"
    
    Forms!Menu!nameNewDatabase = name
    
    nameNewDatabaseFromOriginalPath = name
End Function

Public Function pathOriginal() As String
    Dim result As String
    result = CurrentProject.path & "\NewSQLiteDB.db"
    
    Forms!Menu!pathOriginalSQLiteDB = result
    
    pathOriginal = result

End Function


Public Function createNewDatabase() As Boolean

    Dim originalDB As String
    originalDB = Forms!Menu!pathOriginalSQLiteDB
    
    Dim newNameDB As String
    newNameDB = Forms!Menu!nameNewDatabase
    
    Dim destinationFolder As String
    destinationFolder = Forms!Menu!destinationFolder
    
    CreateFolder destinationFolder
    
    Dim destinationFile As String
    destinationFile = destinationFolder & "\" & newNameDB
       
    CopyAFileDeletingOld originalDB, destinationFile
    
    createNewDatabase = DoesFileExist(destinationFile)
End Function

Public Function exportSelectedTables() As Boolean

    Dim dbAccess As String
    dbAccess = Forms!Menu!pathDatabase
    
    Dim newNameDB As String
    newNameDB = Forms!Menu!nameNewDatabase
    
    Dim destinationFolder As String
    destinationFolder = Forms!Menu!destinationFolder
    
    Dim destinationFile As String
    destinationFile = destinationFolder & "\" & newNameDB

    updateMessage "START"
    
    Dim t As Variant
    For Each t In Forms!Menu!tableList.ItemsSelected()
        Dim nameTable As String
        nameTable = Forms!Menu!tableList.Column(0, t)
        
        Dim message As String
        message = Forms!Menu!logExport
        
        updateMessage nameTable & ": EXPORT" & vbCrLf & message
        
        ExportFromOtherDatabaseToSQLite dbAccess, nameTable, destinationFile

        updateMessage nameTable & ": OK" & vbCrLf & message
        
    Next

    exportSelectedTables = True
End Function


Public Function updateMessage(ByVal message As String) As String
    Application.Echo False
    
    Forms!Menu!logExport = message
    Forms!Menu!logExport.Requery
      
    Application.Echo True
    
    updateMessage = message
End Function
