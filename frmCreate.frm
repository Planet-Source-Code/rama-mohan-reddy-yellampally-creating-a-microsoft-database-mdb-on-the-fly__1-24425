VERSION 5.00
Begin VB.Form frmCreate 
   Caption         =   "Form1"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2115
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create MDB On Fly"
      Height          =   465
      Left            =   675
      TabIndex        =   0
      Top             =   900
      Width           =   2760
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CreateAccessDatabase()
    Dim CatDB As ADOX.Catalog
    Set CatDB = New ADOX.Catalog
    CatDB.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=myMDB.mdb"
    Set CatDB = Nothing
End Sub

Sub CreateTable()
    Dim CatDB As ADOX.Catalog
    Dim TabDB As ADOX.Table
    
    Set CatDB = New ADOX.Catalog
    Set TabDB = New ADOX.Table
    'open the database
    CatDB.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=myMDB.mdb"
    'create new table object
    With TabDB
        .Name = "MyTable" 'set name
        'add fields and specify datatype
        .Columns.Append "First Name", adVarWChar
        .Columns.Append "Last Name", adVarWChar
        .Columns.Append "Age", adInteger
    End With
    'add the table to database
    CatDB.Tables.Append TabDB
    Set CatDB = Nothing
    Set TabDB = Nothing
    MsgBox "Database and Table created..."
End Sub

Private Sub cmdCreate_Click()
    CreateAccessDatabase
    CreateTable
End Sub


