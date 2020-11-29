VERSION 5.00
Begin VB.Form ExcelFunctions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ExcelFunctions"
   ClientHeight    =   5436
   ClientLeft      =   36
   ClientTop       =   372
   ClientWidth     =   6456
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5436
   ScaleWidth      =   6456
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   372
      Left            =   3480
      TabIndex        =   5
      Top             =   4440
      Width           =   1452
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   3000
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.CommandButton Command1 
      Caption         =   "End"
      Height          =   372
      Left            =   3480
      TabIndex        =   2
      Top             =   4920
      Width           =   1452
   End
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
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1452
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
      Left            =   2880
      TabIndex        =   7
      Top             =   1920
      Width           =   2892
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
      Left            =   3120
      TabIndex        =   6
      Top             =   1320
      Width           =   2412
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2640
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   3132
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Click to select an EXCEL function from the list below:"
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
      TabIndex        =   1
      Top             =   0
      Width           =   2532
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ExcelFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() ' END
   Unload Me
End Sub


Private Sub Form_Load() 'Load Excel Functions
  List1.AddItem "Abs()"
  List1.AddItem "Acos()"
  List1.AddItem "Average()"
  List1.AddItem "Correl()"
  List1.AddItem "Min()"
  List1.Refresh
  Label2.Visible = False
  Text1.Visible = False
  Text1.Text = ""
  Label2.Caption = ""
End Sub

Private Sub List1_Click() 'Excel Function Selection
Text1.Text = ""
resultlabel.Caption = ""
Text1.Visible = False
Label2.Caption = ""
Label2.Visible = False
If List1.ListIndex = -1 Then
   MsgBox ("Please click to select an Excel Function from the list below")
Else
   '***** IF ABS Function *****
   If List1.Text = "Abs()" Then
      resultlabel.Caption = ""
      Label1.Visible = True
      Text1.Visible = True
      Label1.Caption = "Give the number or Click on a cell"
   End If
   '***** IF ACOS Function *****
   If List1.Text = "Acos()" Then
   End If
   '***** IF AVERAGE Function *****
   If List1.Text = "Average()" Then
      resultlabel.Caption = ""
      Label1.Visible = True
      Text1.Visible = True
      Label1.Caption = "Give the numbers or select numeric data cells"
   End If
End If
End Sub

Private Sub Command2_Click() 'OK
'********** IF ABS Function *********
If List1.Text = "Abs()" Then
    arg1 = Text1.Text
    arg2 = MSFG.MSFlexGrid1.Text
    If arg1 <> "" And IsNumeric(arg1) Then  ' Show the result
        resultlabel.Caption = Excel.Evaluate(Abs(arg1))
    Else
        If arg2 <> "" And IsNumeric(arg2) Then
           resultlabel.Caption = Excel.Evaluate(Abs(arg2))
        End If
    End If
    Text1.Text = ""
    arg1 = Null
    arg2 = Null
    On Error GoTo CalcError
    Exit Sub
 End If
 
 '********** IF ACOS Function *********
 
 
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
    arg1 = Null
    arg2 = Null
    On Error GoTo CalcError
    Exit Sub
End If

CalcError:
    MsgBox ("Excel Function returned the following error:" & vbCrLf & Err.Description)
End Sub
