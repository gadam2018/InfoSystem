VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form SearchForm 
   Caption         =   "SearchForm"
   ClientHeight    =   8496
   ClientLeft      =   696
   ClientTop       =   1464
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8496
   ScaleWidth      =   11040
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   5760
      TabIndex        =   4
      Top             =   120
      Width           =   5292
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   1932
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "SearchForm.frx":0000
      Height          =   7212
      Left            =   0
      OleObjectBlob   =   "SearchForm.frx":0014
      TabIndex        =   1
      Top             =   720
      Width           =   12012
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Finish "
      Height          =   372
      Left            =   8640
      TabIndex        =   0
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Data Data1 
      Caption         =   "Database"
      Connect         =   "Access"
      DatabaseName    =   "C:\adamg\Datach10.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   276
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Datach10"
      Top             =   120
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label Label4 
      Caption         =   "SQL WHERE like syntax clause, i.e ='KAP5' AND  HERDN0=602"
      Height          =   252
      Left            =   5760
      TabIndex        =   6
      Top             =   480
      Width           =   5292
   End
   Begin VB.Label Label3 
      Caption         =   "Search for:"
      Height          =   252
      Left            =   4800
      TabIndex        =   5
      Top             =   120
      Width           =   852
   End
   Begin VB.Label Label2 
      Caption         =   "Search in:"
      Height          =   252
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   852
   End
   Begin VB.Menu FindFirst 
      Caption         =   "FindFirst"
   End
   Begin VB.Menu FindNext 
      Caption         =   "FindNext"
   End
   Begin VB.Menu Findprevious 
      Caption         =   "FindPrevious"
   End
   Begin VB.Menu FindLast 
      Caption         =   "FindLast"
   End
   Begin VB.Menu End 
      Caption         =   "End"
   End
End
Attribute VB_Name = "SearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim db As Database
Dim gsconnect, gsdatabase, gsrecordsource As String

Private Sub Combo1_Click()
    Data1.Refresh
End Sub
Private Function GenerateSQL() 'As String
GenerateSQL() = Combo1.Text & " " & Text2.Text
End Function
Private Sub End_Click()
Unload Me
End Sub
Private Sub FindFirst_Click()
On Error GoTo sqlerror
'Data1.Recordset.FindFirst GenerateSQL() ' *** STACK OVERFLOW ***
 Data1.Recordset.FindFirst Combo1.Text & " " & Text2.Text

If Data1.Recordset.NoMatch Then
    MsgBox "No such record found"
Else
     'TO CREATE A FOCUS TO CURRENT FOUND RESULT
    'MsgBox Data1.Recordset.Fields(Combo1.Text)
End If
Exit Sub

sqlerror:
   MsgBox Err.Description
End Sub

Private Sub FindLast_Click()
On Error GoTo sqlerror
 Data1.Recordset.FindLast Combo1.Text & " " & Text2.Text

If Data1.Recordset.NoMatch Then
    MsgBox "No such record found"
Else
    'MsgBox Data1.Recordset.Fields(Combo1.Text)
    
End If
Exit Sub

sqlerror:
   MsgBox Err.Description
End Sub

Private Sub FindNext_Click()
On Error GoTo sqlerror
 Data1.Recordset.FindNext Combo1.Text & " " & Text2.Text

If Data1.Recordset.NoMatch Then
    MsgBox "No such record found"
Else
    'MsgBox Data1.Recordset.Fields(Combo1.Text)
End If
Exit Sub

sqlerror:
   MsgBox Err.Description
End Sub

Private Sub Findprevious_Click()
On Error GoTo sqlerror
 Data1.Recordset.Findprevious Combo1.Text & " " & Text2.Text

If Data1.Recordset.NoMatch Then
    MsgBox "No such record found"
Else
    'MsgBox Data1.Recordset.Fields(Combo1.Text)
End If
Exit Sub

sqlerror:
   MsgBox Err.Description
End Sub

Private Sub Form_Load()
Dim tbl As TableDef
Dim fld As Field

gsconnect = OpenFile.globalConnect
gsdatabase = OpenFile.globalDataBase
gsrecordsource = OpenFile.globalRecordsource
 
On Error GoTo LoadErr
      
Data1.Connect = gsconnect
Data1.DatabaseName = gsdatabase
Data1.RecordSource = gsrecordsource
Data1.RecordsetType = 1 'dynaset   '0 table
Data1.Options = 0
Data1.Refresh

'set db = OpenDatabase(Data1.DatabaseName)
Set db = Data1.Database
Set tbl = db.TableDefs(Data1.RecordSource)

'MsgBox (db.TableDefs(0).Name)
'For Each tbl In db.TableDefs
 ' Combo2.AddItem tbl.Name
'Next
Combo1.Clear
For Each fld In tbl.Fields
  Combo1.AddItem fld.Name
Next
 Combo1.ListIndex = 0 '?
Exit Sub

LoadErr:
    MsgBox "Error:" & Err & " " & Err.Description
    Unload Me
End Sub

