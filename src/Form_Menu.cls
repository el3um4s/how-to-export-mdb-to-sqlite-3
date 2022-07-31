VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Database

Private Sub btnChooseDatabase_Click()
    SelectDatabase
    nameNewDatabaseFromOriginalPath
    ShowListTable
    SelectAllTables
End Sub

Private Sub btnDestinationFolder_Click()
    SelectDestinationFolder
End Sub

Private Sub btnExportTo_Click()
     createNewDatabase
     exportSelectedTables
     
     MsgBox "COMPLETED"
End Sub

Private Sub btnSelectAllTables_Click()
    SelectAllTables
End Sub

Private Sub Form_Open(Cancel As Integer)
    pathOriginal
    Forms!Menu!tableList.RowSource = ""
    Forms!Menu!logExport = ""
End Sub

Private Sub pathDatabase_AfterUpdate()
    nameNewDatabaseFromOriginalPath
    ShowListTable
    SelectAllTables
End Sub

