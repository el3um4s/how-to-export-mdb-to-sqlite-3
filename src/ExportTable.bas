Attribute VB_Name = "ExportTable"
Option Explicit
Option Compare Database

Public Function ExportToSQLite(ByVal table As String, ByVal database As String) As Boolean
    
    DoCmd.TransferDatabase acExport, "ODBC", "ODBC;DSN=SQLite3 Datasource;Database=" & database & ";StepAPI=0;SyncPragma=NORMAL;NoTXN=0;Timeout=100000;ShortNames=0;LongNames=0;NoCreat=0;NoWCHAR=0;FKSupport=0;JournalMode=;OEMCP=0;LoadExt=;BigInt=0;JDConv=0;", acTable, table, table, False
          
    
    ExportToSQLite = True
End Function


Public Function ExportFromOtherDatabaseToSQLite(ByVal dbAccess As String, ByVal table As String, ByVal dbSQLite As String) As Boolean

    Dim db As DAO.database
    
    Set db = OpenDatabase(dbAccess, False)
            
    DoCmd.TransferDatabase acImport, "Microsoft Access", dbAccess, acTable, table, table, False
  
    ExportToSQLite table, dbSQLite
    
    DoCmd.DeleteObject acTable, table

    db.Close
    ExportFromOtherDatabaseToSQLite = True
End Function

'Function ExportFromOtherDatabaseToSQLite(database As String) As Boolean
'
'    Dim db As DAO.database
'    Dim tdf As DAO.TableDef
'
'    Set db = OpenDatabase(database, False)
'
'    For Each tdf In db.TableDefs
'        If Not (tdf.name Like "MSys*" Or tdf.name Like "~*") Then
'
'            DoCmd.TransferDatabase acImport, "Microsoft Access", database, acTable, tdf.name, tdf.name, False
'            Debug.Print tdf.name
'            ExportToSQLite tdf.name, "D:\VARIE\SQLite ODBC Driver\ConvertMDBtoSQLite\datico.db"
'            DoCmd.DeleteObject acTable, tdf.name
'        End If
'
'    Next
'    ExportFromOtherDatabaseToSQLite = True
'End Function
