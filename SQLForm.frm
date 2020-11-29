VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SQLForm 
   Caption         =   "SQLEXE"
   ClientHeight    =   5028
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6792
   LinkTopic       =   "SQLForm"
   MDIChild        =   -1  'True
   ScaleHeight     =   8496
   ScaleWidth      =   12192
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "Chart"
      Height          =   252
      Left            =   3600
      TabIndex        =   19
      Top             =   2640
      Width           =   852
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Reset"
      Height          =   252
      Left            =   1800
      TabIndex        =   18
      Top             =   2640
      Width           =   852
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   252
      Left            =   6360
      TabIndex        =   17
      Top             =   1440
      Width           =   732
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   252
      Left            =   1920
      TabIndex        =   16
      Top             =   720
      Width           =   732
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   252
      Left            =   7440
      TabIndex        =   15
      Top             =   1440
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   255
      Left            =   3720
      TabIndex        =   14
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox textSQL 
      Height          =   1080
      Left            =   8400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   1680
      Width           =   3330
   End
   Begin VB.ListBox QryList 
      Height          =   1008
      ItemData        =   "SQLForm.frx":0000
      Left            =   4920
      List            =   "SQLForm.frx":0002
      TabIndex        =   9
      Top             =   1680
      Width           =   3375
   End
   Begin VB.ListBox FldList 
      Height          =   1008
      Left            =   8400
      TabIndex        =   7
      Top             =   240
      Width           =   3375
   End
   Begin VB.ListBox TblList 
      Height          =   1008
      Left            =   4920
      TabIndex        =   6
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox txtSQL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   975
      Width           =   4620
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7560
      Width           =   11655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "SQLForm.frx":0004
      Height          =   4680
      Left            =   135
      OleObjectBlob   =   "SQLForm.frx":0018
      TabIndex        =   1
      Top             =   2895
      Width           =   11625
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   1,17491e-38
   End
   Begin VB.Label Label8 
      Caption         =   "Query Definition"
      Height          =   255
      Left            =   8400
      TabIndex        =   12
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Fields"
      Height          =   255
      Left            =   8520
      TabIndex        =   11
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Queries"
      Height          =   252
      Left            =   5040
      TabIndex        =   10
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label5 
      Caption         =   "Tables"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Database Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   135
      TabIndex        =   5
      Top             =   75
      Width           =   1725
   End
   Begin VB.Label Label3 
      Caption         =   "SQL Statement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   4
      Top             =   720
      Width           =   1680
   End
   Begin VB.Label Label2 
      Caption         =   "Query Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   3
      Top             =   2640
      Width           =   1590
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   135
      TabIndex        =   0
      Top             =   360
      Width           =   4605
   End
   Begin VB.Menu OpenDB 
      Caption         =   "Open Database"
   End
   Begin VB.Menu ExecSQL 
      Caption         =   "Execute SQL"
   End
   Begin VB.Menu End 
      Caption         =   "End"
   End
End
Attribute VB_Name = "SQLForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DB As Database
Dim tbl As TableDef
Dim TName As String
Dim idx As Index
Dim qry As QueryDef
Dim gsconnect, gsdatabase, gsrecordsource As String

Private Sub Command1_Click() ' SAVE QUERY
Dim qdfNew As QueryDef
Dim qdfname As String

On Error GoTo SQLError

qdfname = InputBox("Give a name for the query:", , qdfname)
Set qdfNew = DB.CreateQueryDef(qdfname, _
            txtSQL.Text)
QryList.AddItem qdfNew.Name
QryList.Refresh

SQLError:
    MsgBox Err.Description
End Sub

Private Sub Command2_Click() 'REMOVE QUERY FROM DATABASE
Dim qdfrem As QueryDef
Dim qdfname As String
Dim msg, style, title, response
If QryList.ListIndex = -1 Then
   MsgBox ("Please select a query from the list below")
Else
  qdfname = DB.QueryDefs(QryList.ListIndex).Name
  msg = "Query to be deleted: "
  style = vbOKCancel
  title = "QUERY DELETE"
  response = MsgBox(msg & qdfname, style, title)
  If response = vbOK Then   ' User chose OK
     QryList.RemoveItem (QryList.ListIndex)
     DB.QueryDefs.Delete (qdfname)
     QryList.Refresh
  Else
  End If
End If
End Sub

Private Sub Command3_Click() ' Clear txtSQL in box
txtSQL.Text = ""
End Sub

Private Sub Command4_Click()
MsgBox ("Sorry, information is not yet available for queries")
End Sub

Private Sub Command5_Click() 'RESET
    
End Sub

Private Sub Command6_Click()
Load ChartForm
ChartForm.Show
End Sub

Private Sub data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'This will display the current record position for dynasets and snapshots
  Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
End Sub

Private Sub End_Click()
Unload Me
End Sub

Private Sub ExecSQL_Click()
On Error GoTo SQLError

    Data1.RecordSource = txtSQL
    Data1.Refresh
    Exit Sub
    
SQLError:
    MsgBox Err.Description
End Sub


Private Sub Form_Load()

  'gsconnect = OpenFile.globalConnect
  Set DB = OpenDatabase(OpenFile.globalDataBase)
  'Set db = OpenFile.globalDataBase
  gsdatabase = OpenFile.globalDataBase
  Label1.Caption = gsdatabase
  'gsrecordsource = OpenFile.globalRecordsource
  FldList.Clear
  TblList.Clear
   
Debug.Print "There are " & DB.TableDefs.Count & " tables in the database"
' Process each table
    For Each tbl In DB.TableDefs
        ' EXCLUDE SYSTEM TABLES
        If Left(tbl.Name, 4) <> "MSys" And Left(tbl.Name, 4) <> "USys" Then
            TblList.AddItem tbl.Name
' For each table, process the table's indices
            'For Each idx In tbl.Indexes
               ' TblList.AddItem "  " & idx.Name
            'Next
        End If
    Next
    
Debug.Print "There are " & DB.QueryDefs.Count & " queries in the database"
' Process each stored query
    For Each qry In DB.QueryDefs
        QryList.AddItem qry.Name
    Next
    

    'Data1.Connect = gsconnect
    Data1.DatabaseName = gsdatabase
    'Data1.RecordSource = gsrecordsource
    Data1.RecordsetType = 1     'dynaset
    Data1.Options = 0
    Data1.Refresh
    
NoDatabase:
End Sub

Private Sub QryList_Click()
Dim qry As QueryDef
    
    textSQL.Text = DB.QueryDefs(QryList.ListIndex).SQL
        
End Sub

Private Sub QryList_DblClick()
txtSQL.Text = textSQL.Text
End Sub

Private Sub TblList_Click()
Dim fld As Field
Dim idx As Index
    
    If Left(TblList.Text, 2) = "  " Then Exit Sub
    FldList.Clear
    For Each fld In DB.TableDefs(TblList.Text).Fields
        FldList.AddItem fld.Name
    Next
    
End Sub
Private Sub OpenDB_Click()
On Error GoTo NoDatabase
CommonDialog1.CancelError = True
CommonDialog1.Filter = "Databases|*.MDB"
CommonDialog1.ShowOpen

Data1.DatabaseName = CommonDialog1.FileName
Data1.Refresh
' Open the database
    If CommonDialog1.FileName <> "" Then
        Set DB = OpenDatabase(CommonDialog1.FileName)
        Label1.Caption = CommonDialog1.FileName
    End If
' Clear the ListBox controls
TblList.Clear
FldList.Clear
QryList.Clear
txtSQL.Text = ""
textSQL.Text = ""
   
'Debug.Print "There are " & DB.TableDefs.Count & " tables in the database"
' Process each table
    For Each tbl In DB.TableDefs
        ' EXCLUDE SYSTEM TABLES
        If Left(tbl.Name, 4) <> "MSys" And Left(tbl.Name, 4) <> "USys" Then
            TblList.AddItem tbl.Name
' For each table, process the table's indices
            'For Each idx In tbl.Indexes
               ' TblList.AddItem "  " & idx.Name
            'Next
        End If
    Next
    
Debug.Print "There are " & DB.QueryDefs.Count & " queries in the database"
' Process each stored query
    For Each qry In DB.QueryDefs
        QryList.AddItem qry.Name
    Next
    
If Err = 0 Then
        Label1.Caption = CommonDialog1.FileName
    Else
        MsgBox Err.Description
End If

NoDatabase:
    On Error GoTo 0
End Sub
