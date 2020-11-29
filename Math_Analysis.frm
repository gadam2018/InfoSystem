VERSION 5.00
Begin VB.Form Math_Analysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MathAnalysis"
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
      Left            =   600
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
      Top             =   960
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
      TabIndex        =   8
      Top             =   2160
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
      Left            =   2520
      TabIndex        =   7
      Top             =   1560
      Width           =   2412
   End
   Begin VB.Label Label3 
      Caption         =   "Click to select an Excel Math function from the list below:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   2532
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Height          =   252
      Left            =   2640
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label Label1 
      Caption         =   "MATH ANALYSIS"
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
Attribute VB_Name = "Math_Analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub List1_Click() 'Excel Math Function Selection
Text1.Text = ""
Text1.Visible = False
Label2.Caption = ""
resultlabel.Caption = ""
Label2.Visible = False
If List1.ListIndex = -1 Then
   MsgBox ("Please click to select an Excel Math Function from the list below")
Else
   '***** IF ABS Function *****
   If List1.Text = "Abs()" Then
      resultlabel.Caption = ""
      Label2.Visible = True
      Text1.Visible = True
      Label2.Caption = "Give the number or Click on a cell"
   End If
   '***** IF ACOS Function *****
   If List1.Text = "Acos()" Then
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
If List1.Text = "Acos()" Then
End If

CalcError:
    MsgBox ("Excel Math Function returned the following error:" & vbCrLf & Err.Description)
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
  List1.AddItem "Abs()"
  List1.AddItem "Acos()"
  List1.AddItem "Asin()"
  List1.AddItem "Asinh()"
  List1.AddItem "Atan()"
  List1.Refresh
  Label2.Visible = False
  Text1.Visible = False
  Text1.Text = ""
  Label2.Caption = ""
End Sub
