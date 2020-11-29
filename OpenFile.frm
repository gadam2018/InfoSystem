VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form OpenFile 
   Caption         =   "OpenFile"
   ClientHeight    =   4644
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   6456
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4644
   ScaleWidth      =   6456
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   372
      Left            =   3360
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   656
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   3240
      TabIndex        =   7
      Text            =   "charolaisx$"
      Top             =   480
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Data DataFile 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   3240
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   2652
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3960
      TabIndex        =   4
      Top             =   2640
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3960
      TabIndex        =   3
      Top             =   2040
      Width           =   1212
   End
   Begin VB.DirListBox Dir1 
      Height          =   2232
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   2172
   End
   Begin VB.FileListBox File1 
      Height          =   1800
      Left            =   600
      Pattern         =   "*.mdb;*.xls;*.dbf"
      TabIndex        =   1
      Top             =   2760
      Width           =   2172
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   2172
   End
   Begin VB.Label Label2 
      Caption         =   "Please wait, loading in progress....."
      Height          =   252
      Left            =   3360
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   2892
   End
   Begin VB.Label Label1 
      Caption         =   "RecordSource (i.e Table or WorkSheet)"
      Height          =   252
      Left            =   3240
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   3012
   End
End
Attribute VB_Name = "OpenFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public globalConnect, globalDataBase, globalRecordsource
Dim appWorld As Excel.Application
Dim wbWorld As Excel.Workbook
Dim DB As Database
Dim xb As Workbook
Private Sub Combo1_Click() '*** TABLES SELECTION ***
  If globalConnect = "Access" Then
              globalRecordsource = Combo1.Text
  Else
      If globalConnect = "Excel 8.0;" Then
              globalRecordsource = Combo1.Text & "$"
      End If
  End If
End Sub

Private Sub Command1_Click() ' *** OK ***

Load MSFG
 'DataFile.RecordSource = globalRecordsource
 'DataFile.RecordsetType = 1     'dynaset
 'DataFile.Options = 0
 'DataFile.Refresh
'Dim intCounter As Integer ' Counter to set Progressbar1.Value
'With ProgressBar1
 '       .Max = DataFile.Recordset.RecordCount
 '       .Visible = True
  '  End With
  '  Do While Not DataFile.Recordset.EOF
   '     intCounter = intCounter + 1
  '      ProgressBar1.Value = intCounter ' Update ProgressBar1.
  '     DataFile.Recordset.MoveNext    ' Move to next rsdata record.
  '  Loop
  '   ProgressBar1.Visible = False
MSFG.Show
       
'Do
'ProgressBar1.Visible = True
'label2.visible=true
'ProgressBar1.Value = ProgressBar1.Value + 5
'If MSFG.ActiveControl = True Then
'   check = True
 '  ProgressBar1.Visible = False
   'label2.visible=false
'End If
'Loop Until check = True
Unload Me
End Sub

Private Sub Command2_Click() ' *** CANCEL ***
Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub


Private Sub File1_Click() '*** FILE SELECTION ***
Dim tbl As TableDef
Dim xbl As Worksheet
Dim fxls, fmdb, fdbf
    Label1.Visible = False
    Text1.Visible = False
    Combo1.Visible = False
         
    SelectedFile = File1.Path & "\" & File1.FileName
    globalDataBase = SelectedFile
    
    fmdb = InStr(globalDataBase, ".mdb")   ' Find 'mdb' string in text.
    If fmdb Then   ' If ACCESS File
        globalConnect = "Access"
        DataFile.Connect = "Access"
        DataFile.DatabaseName = globalDataBase
        Master.ExcelAnalysis.Enabled = False
        Master.Open2.Enabled = False
        Master.Close1.Enabled = True
        Master.SortFilter.Enabled = True
        Master.Search.Enabled = True
        Master.ADRU.Enabled = True
        Master.Advanced.Enabled = True
        Label1.Visible = True
        Combo1.Visible = True
        Set DB = OpenDatabase(SelectedFile)
        'Set DB = OpenDatabase(DataFile.DatabaseName)
        Combo1.Clear
        For Each tbl In DB.TableDefs
          If Left(tbl.Name, 4) <> "MSys" Then
               Combo1.AddItem tbl.Name
          End If
        Next
        Combo1.ListIndex = 0
     End If
     
     fxls = InStr(globalDataBase, ".xls")   ' Find 'xls' string in text.
     If fxls Then   ' If EXCEL File
         globalConnect = "Excel 8.0;"
         DataFile.Connect = "Excel 8.0;"
         DataFile.DatabaseName = globalDataBase
         Master.ExcelAnalysis.Enabled = True
         Master.Open2.Enabled = True 'Open Work File In Excel
         Master.Close1.Enabled = True
         Master.SortFilter.Enabled = True
         Master.Search.Enabled = True
         Master.ADRU.Enabled = True
         Master.Advanced.Enabled = False
         Label1.Visible = True
         Combo1.Visible = True
         Set appWorld = GetObject(, "Excel.Application") 'look for a running copy of Excel
         If Err.Number <> 0 Then 'If Excel is not running then
            Set appWorld = CreateObject("Excel.Application") 'run it
         End If
         Err.Clear   ' Clear Err object in case error occurred.
         On Error GoTo 0 'Resume normal error processing
         Set wbWorld = appWorld.Workbooks.Open(SelectedFile)
         ' Fill the Continents combo box with the names
         ' of the sheets in the workbook.
         Dim sht As Excel.Worksheet
       ' Iterate through the collection of sheets and add
       ' the name of each sheet to the combo box.
         Combo1.Clear
         For Each sht In wbWorld.Sheets
           OpenFile.Combo1.AddItem sht.Name
         Next
         Combo1.ListIndex = 0
       ' Select the first item and display it in the combo box.
       ' OpenFile.Combo1.Text = OpenFile.Combo1.List(0)
         Set sht = Nothing
         Set appWorld = Nothing
         Set wbWorld = Nothing
        End If
        
        
     
     fdbf = InStr(globalDataBase, ".dbf")   ' Find 'dbf' string in text.
     If fdbf Then   ' If DBASE File
        Text1.Text = File1.FileName  '? e.g "Charolai"  'globalRecordsource
        globalRecordsource = Text1.Text
        globalConnect = "dBASE 5.0;"
        DataFile.Connect = "dBASE 5.0;"
        globalDataBase = File1.Path 'e.g. "c:\adamg"
        DataFile.DatabaseName = globalDataBase
        Master.ExcelAnalysis.Enabled = False
        Master.Open2.Enabled = False
        Master.Close1.Enabled = True
        Master.SortFilter.Enabled = True
        Master.Search.Enabled = True
        Master.ADRU.Enabled = True
        Label1.Visible = True
        Text1.Visible = True
        'Set DB = OpenDatabase(SelectedFile)
    End If
                       
 On Error GoTo LoadErr
 Exit Sub
  
LoadErr:
    MsgBox "Error:" & Err & " " & Err.Description
    Unload Me
End Sub



Private Sub Form_Load()
globalConnect = Null
globalDataBase = Null
globalRecordsource = Null
Drive1.Drive = "c"
Dir1.Path = CurDir   'e.g. "c:\AgroModel"
End Sub


