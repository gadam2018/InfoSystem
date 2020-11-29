VERSION 5.00
Begin VB.Form BInStat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Build-in Statistical Functions"
   ClientHeight    =   4932
   ClientLeft      =   36
   ClientTop       =   372
   ClientWidth     =   5256
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4932
   ScaleWidth      =   5256
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   372
      Left            =   3000
      TabIndex        =   3
      Top             =   3720
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   2880
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "End"
      Height          =   372
      Left            =   3000
      TabIndex        =   1
      Top             =   4200
      Width           =   1332
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2292
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
      Left            =   2520
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
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
      TabIndex        =   5
      Top             =   1680
      Width           =   2892
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   2400
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   2772
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "BInStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() ' END
   Unload Me
End Sub


Private Sub Form_Load() 'Load my Stat Functions ()
  List1.AddItem "ArithmeticMean()"
  List1.AddItem "GeometricMean()"
  List1.Refresh
  Label1.Visible = False
  'Text1.Visible = False
  'Text1.Text = ""
  Label1.Caption = ""
  resultlabel.Caption = ""
End Sub



Private Sub List1_Click()
'Text1.Text = ""
'Text1.Visible = False
Label1.Caption = ""
resultlabel.Caption = ""
Label1.Visible = False
reslabel.Visible = False
If List1.ListIndex = -1 Then
   MsgBox ("Please click to select a Statistical Function from the list below")
Else
   '***** IF ArithmeticMean Function *****
   If List1.Text = "ArithmeticMean()" Then
      Label1.Visible = True
      reslabel.Visible = True
      'Text1.Visible = True
      Label1.Caption = "Please select the numeric data cells."
   End If
   '***** IF GeometricMean Function *****
   If List1.Text = "GeometricMean()" Then
   End If
End If
End Sub

Private Sub Command2_Click() 'OK
'********** IF ArithmeticMean Function *********
If List1.Text = "ArithmeticMean()" Then
    'arg1 = Text1.Text
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
    'arg2 = MSFG.MSFlexGrid1.Clip
    'MsgBox (arg2)
     'arg2 = MSFG.MSFlexGrid1.Text
    'If arg1 <> "" And IsNumeric(arg1) Then  ' Show the result
        'MsgBox (ArithmeticMean(arg1))
    'Else
        'If arg2 <> "" And IsNumeric(arg2) Then
           ' MsgBox (ArithmeticMean(arg2))
       ' End If
    'End If
   ' Text1.Text = ""
    'arg1 = Null
    'arg2 = Null
    On Error GoTo CalcError
    Exit Sub
 End If
'********** IF GeometricMean Function *********
If List1.Text = "Atn()" Then
End If

CalcError:
    MsgBox ("Statistical Function returned the following error:" & vbCrLf & Err.Description)
End Sub
