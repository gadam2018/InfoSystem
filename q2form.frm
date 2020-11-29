VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form q2form 
   Caption         =   "Form1"
   ClientHeight    =   2448
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2448
   ScaleWidth      =   3744
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Return"
      Height          =   372
      Left            =   3120
      TabIndex        =   2
      Top             =   5640
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   372
      Left            =   3120
      TabIndex        =   1
      Top             =   5040
      Width           =   1452
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "q2form.frx":0000
      Height          =   3132
      Left            =   240
      OleObjectBlob   =   "q2form.frx":0014
      TabIndex        =   0
      Top             =   600
      Width           =   11652
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   492
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Width           =   3132
   End
End
Attribute VB_Name = "q2form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gsdatabase
Dim dbs As Database


Private Sub Command8_Click()
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

Sub selectCalf()

End Sub

Private Sub Command1_Click()  'SEARCH
'SelectNX1
selectCalf
MsgBox ("Query is executed")
End Sub

Private Sub Command2_Click()  'RETURN
 Unload Me
'Load Master
'Master.Show
' dbs.Close
End Sub

Private Sub data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'This will display the current record position for dynasets and snapshots
  Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
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
