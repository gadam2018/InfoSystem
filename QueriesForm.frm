VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form QueriesForm 
   Caption         =   "QueriesForm"
   ClientHeight    =   6720
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   1572
   End
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "QueriesForm.frx":0000
      Height          =   1212
      Left            =   2640
      OleObjectBlob   =   "QueriesForm.frx":0014
      TabIndex        =   18
      Top             =   2400
      Width           =   9372
   End
   Begin VB.TextBox Text6 
      Height          =   372
      Left            =   1320
      TabIndex        =   17
      Top             =   3240
      Width           =   1212
   End
   Begin VB.TextBox Text5 
      Height          =   372
      Left            =   1320
      TabIndex        =   16
      Top             =   2760
      Width           =   1212
   End
   Begin VB.TextBox Text4 
      Height          =   372
      Left            =   1320
      TabIndex        =   15
      Top             =   2280
      Width           =   1212
   End
   Begin VB.TextBox Text3 
      Height          =   372
      Left            =   1320
      TabIndex        =   14
      Top             =   1800
      Width           =   1212
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   1320
      TabIndex        =   13
      Top             =   1320
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   1320
      TabIndex        =   12
      Top             =   840
      Width           =   1212
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SEARCH"
      Height          =   372
      Left            =   1560
      TabIndex        =   11
      Top             =   4560
      Width           =   1212
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "QueriesForm.frx":09E7
      Height          =   1932
      Left            =   3480
      OleObjectBlob   =   "QueriesForm.frx":09FB
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   8772
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   492
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.CommandButton Command3 
      Caption         =   "search2"
      Height          =   492
      Left            =   5760
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   1572
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "QueriesForm.frx":13CE
      Height          =   1572
      Left            =   2640
      OleObjectBlob   =   "QueriesForm.frx":13E2
      TabIndex        =   8
      Top             =   720
      Width           =   9372
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   492
      Left            =   1560
      TabIndex        =   7
      Top             =   5760
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Return"
      Height          =   492
      Left            =   1560
      TabIndex        =   6
      Top             =   5040
      Width           =   1332
   End
   Begin VB.Data Data1 
      Caption         =   "Database"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.Label Label6 
      Caption         =   "MPPD5"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label Label5 
      Caption         =   "MPPD3"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "SCLA"
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "DCLA"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "HCLA"
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "HSIZE"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   972
   End
End
Attribute VB_Name = "QueriesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gsdatabase
Dim dbs As Database


Private Sub Command1_Click() '*** RETURN ***
Unload Me
'Load Master
'Master.Show
' dbs.Close
End Sub

Private Sub Command2_Click()
Dim recRecordset1 As Recordset, recRecordset2 As Recordset
   ' Dim hsize As String
    
    Set recRecordset1 = Data1.Recordset
    
    ' hsize$ = Text1.Text
    'hcla$ = Text2.Text
    'dcla$ = Text3.Text
    'scla$ = Text4.Text
    'mppd3$ = Text5.Text
    'mppd5$ = Text6.Text
    
    ' If Len(hsize$) = 0 Then Exit Sub

    ' Screen.MousePointer = vbHourglass
      recRecordset1.Filter = hsize$
      Set recRecordset2 = recRecordset1.OpenRecordset(recRecordset1.Type) 'establish the filter
      Set Data1.Recordset = recRecordset2
         
    
    
    'Data1.Recordset.Index = "CID1CID2"      'χρήση του code
    'Data1.Recordset.Seek "=", hsize$ 'και αναζήτηση.
    'If Data1.Recordset.NoMatch Then      'αν δε βρεθεί
   ' Data1.Recordset.MoveFirst         'μετάβαση στην πρώτη εγγραφή
    'End If
        
End Sub

Private Sub Command3_Click()
     
   ' hsize$ = Text1.Text
   hcla$ = Text2.Text
    
    Data2.Recordset.Index = "HSIZE"      'χρήση του code
    Data2.Recordset.Seek "=", hsize$ And hcla$         ' και αναζήτηση.
    If Data2.Recordset.NoMatch Then      'αν δε βρεθεί
    Data2.Recordset.MoveFirst         'μετάβαση στην πρώτη εγγραφή
    End If
End Sub
Private Sub Command5_Click()
SelectNX1
MsgBox ("Query is executed")
End Sub

Sub SelectNX1()
    
    Dim qdf, qdf1, qdf2 As QueryDef
    Dim strSql, strSql2 As String, strParm As String
      
    'Set Data1.Recordset = dbs.OpenRecordset("SELECT * FROM charolais where HSIZE = hs;")
            
    ' Define the parameters clause.
    strParm = "PARAMETERS [hs] DOUBLE, [hc] DOUBLE, [dc] DOUBLE, [sc] DOUBLE;" ',[mp3] DOUBLE,[mp5] DOUBLE; "

    ' Define an SQL statement with the parameters clause.
    strSql = strParm & "SELECT * FROM charolais " _
        & "WHERE HSIZE =[hs] AND HCLA=[hc] AND DCLA=[dc] AND SCLA=[sc] ;" 'AND MPPD3=[mp3] AND MPPD5=[mp5];"
    
    strSql2 = strParm & "SELECT HSIZE,HCLA,DCLA,SCLA,MPPD3,MPPD5 FROM charolais " _
        & "WHERE HSIZE =[hs] AND HCLA=[hc] AND DCLA=[dc] AND SCLA=[sc] ;" 'AND MPPD3=[mp3] AND MPPD5=[mp5];"
    
    ' Create a QueryDef object based on the SQL statement.
    Set qdf = dbs.CreateQueryDef _
        ("qF1", strSql)
    
     Set qdf2 = dbs.CreateQueryDef _
        ("qF2", strSql2)
        
    If Len(Text1.Text) Then
        ' MsgBox ("HSIZE not empty")
        qdf("hs") = Val(Text1.Text)
        qdf2("hs") = Val(Text1.Text)
    Else
       'MsgBox ("HSIZE empty")
       qdf("hs") = 0
       qdf2("hs") = 0
    End If
    
    If Len(Text2.Text) Then
        'MsgBox ("HCLA not empty")
        qdf("hc") = Val(Text2.Text)
        qdf2("hc") = Val(Text2.Text)
    Else
        'MsgBox ("HCLA empty")
         qdf("hc") = 0
         qdf2("hc") = 0
    End If
    
    If Len(Text3.Text) Then
        'MsgBox ("DCLA not empty")
        qdf("dc") = Val(Text3.Text)
        qdf2("dc") = Val(Text3.Text)
    Else
        'MsgBox ("DCLA empty")
         qdf("dc") = 0
        qdf2("dc") = 0
    End If
    
    If Len(Text4.Text) Then
        'MsgBox ("SCLA not empty")
        qdf("sc") = Val(Text4.Text)
        qdf2("sc") = Val(Text4.Text)
    Else
        'MsgBox ("SCLA empty")
         qdf("sc") = 0
        qdf2("sc") = 0
    End If
    
       'Set qdf1 = dbs.CreateQueryDef("")
       'qdf1.SQL = "SELECT HSIZE FROM charolais where HSIZE=any(select HSIZE from charolais)"
        
      ' qdf("hs") =
       'Set Data1.Recordset = dbs.OpenRecordset("SELECT  * FROM charolais where HSIZE=any(select HSIZE from charolais);")
                    
       'qdf("mp3") = Val(Text5.Text)
       'qdf("mp5") = Val(Text6.Text)
       Set Data1.Recordset = qdf.OpenRecordset(dbOpenDynaset)
       Set Data3.Recordset = qdf2.OpenRecordset(dbOpenDynaset)
     dbs.QueryDefs.Delete "qF1"
     dbs.QueryDefs.Delete "qF2"
          
End Sub


Private Sub Form_Load()
   gsconnect = OpenFile.globalConnect
   gsdatabase = OpenFile.globalDataBase
   gsrecordsource = OpenFile.globalRecordsource
        
    Data1.Connect = gsconnect
    Data1.DatabaseName = gsdatabase
    Data1.RecordSource = gsrecordsource
    Data1.RecordsetType = 1     'dynaset
    Data1.Options = 0
    Data1.Refresh

Set dbs = OpenDatabase(gsdatabase)
End Sub
