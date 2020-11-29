VERSION 5.00
Begin VB.Form Stat_Analysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StatAnalysis"
   ClientHeight    =   5736
   ClientLeft      =   36
   ClientTop       =   372
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   5460
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4128
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   1452
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   372
      Left            =   3360
      TabIndex        =   4
      Top             =   4800
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "End"
      Height          =   372
      Left            =   3360
      TabIndex        =   3
      Top             =   5280
      Width           =   1452
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   2640
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   2412
   End
   Begin VB.Label reslabel 
      Alignment       =   2  'Center
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2640
      TabIndex        =   8
      Top             =   1800
      Width           =   2412
   End
   Begin VB.Label resultlabel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   2400
      TabIndex        =   7
      Top             =   2400
      Width           =   2892
   End
   Begin VB.Label Label3 
      Caption         =   "Click to select an Excel Stat function from the list below:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   2532
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Height          =   612
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label Label1 
      Caption         =   "STAT ANALYSIS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   16.2
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   2892
   End
End
Attribute VB_Name = "Stat_Analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub List1_Click() 'Excel Stat Function Selection
Text1.Text = ""
Text1.Visible = False
Label2.Caption = ""
resultlabel.Caption = ""
Label2.Visible = False
If List1.ListIndex = -1 Then
   MsgBox ("Please click to select an Excel Stat Function from the list below")
Else
   '***** IF AVERAGE Function *****
   If List1.Text = "Average()" Then
      resultlabel.Caption = ""
      Label2.Visible = True
      Text1.Visible = True
      Label2.Caption = "Give the numbers or select numeric data cells"
   End If
   '***** IF  Function *****
   'If List1.Text = "()" Then
   'End If
End If
End Sub

Private Sub Command2_Click() 'OK
'********** IF AVERAGE Function *********
If List1.Text = "Average()" Then
   totalsum = 0
   For j = MSFG.MSFlexGrid1.Col To MSFG.MSFlexGrid1.ColSel
     For i = MSFG.MSFlexGrid1.Row To MSFG.MSFlexGrid1.RowSel
      If IsNumeric(MSFG.MSFlexGrid1.TextMatrix(i, j)) Then
        totalsum = totalsum + MSFG.MSFlexGrid1.TextMatrix(i, j)
      Else
       MsgBox ("Please select numeric data")
       Exit Sub
      End If
     Next
    Next
   ' Total number of rows, total number of columns, total number of cells
     totalrows = (MSFG.MSFlexGrid1.RowSel - MSFG.MSFlexGrid1.Row) + 1
     totalcols = (MSFG.MSFlexGrid1.ColSel - MSFG.MSFlexGrid1.Col) + 1
     totalcells = totalrows * totalcols
     arithmean = totalsum / totalcells
     'MsgBox (totalsum)
     resultlabel.Caption = arithmean
    Text1.Text = ""
    On Error GoTo CalcError
    Exit Sub
 End If
    
CalcError:
    MsgBox ("Excel Stat Function returned the following error:" & vbCrLf & Err.Description)
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
  List1.AddItem "Avedev()"
  List1.AddItem "Average()"
  List1.AddItem "Averagea()"
  List1.AddItem "Betadist()"
  List1.AddItem "Correl()"
  List1.Refresh
  Label2.Visible = False
  Text1.Visible = False
  Text1.Text = ""
  Label2.Caption = ""
End Sub
