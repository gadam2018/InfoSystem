VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ExcelAnalysis 
   Caption         =   "Excel Analysis"
   ClientHeight    =   5676
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8784
   ScaleWidth      =   12192
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "ExcelAnalysis.frx":0000
      Height          =   8772
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12132
      _ExtentX        =   21400
      _ExtentY        =   15473
      _Version        =   393216
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   276
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8520
      Visible         =   0   'False
      Width           =   1332
   End
End
Attribute VB_Name = "ExcelAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  gsconnect = OpenFile.globalConnect
  gsdatabase = OpenFile.globalDataBase
  gsrecordsource = OpenFile.globalRecordsource
 
    Dim bParmQry As Integer
    Dim qdfTmp As QueryDef
    
    On Error GoTo LoadErr
    
    Data1.Connect = gsconnect
    Data1.DatabaseName = gsdatabase
    Data1.RecordSource = gsrecordsource
    Data1.RecordsetType = 1     'dynaset
    Data1.Options = 0
    Data1.Refresh
    
    Exit Sub
   
LoadErr:
    MsgBox "Error:" & Err & " " & Err.Description
    Unload Me
End Sub

Private Sub MSFlexGrid1_Click()
MsgBox (MSFlexGrid1.MouseCol & MSFlexGrid1.MouseRow)
End Sub
