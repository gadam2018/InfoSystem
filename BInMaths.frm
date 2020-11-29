VERSION 5.00
Begin VB.Form BInMaths 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Build-in Mathematical Functions"
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
      Top             =   3840
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "End"
      Height          =   372
      Left            =   3000
      TabIndex        =   1
      Top             =   4320
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
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   1332
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
      Left            =   2040
      TabIndex        =   6
      Top             =   1680
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
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   2412
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
      Height          =   492
      Left            =   1920
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   3132
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "BInMaths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() ' END
   Unload Me
End Sub


Private Sub Form_Load() 'Load VBASIC Math Functions (16)
  List1.AddItem "Abs()"
  List1.AddItem "Atn()"
  List1.AddItem "Cos()"
  List1.AddItem "Exp()"
  List1.AddItem "Int()"
  List1.AddItem "Fix()"
  List1.AddItem "Round()"
  List1.AddItem "Log()"
  List1.AddItem "Oct()"
  List1.AddItem "Hex()"
  List1.AddItem "Rnd()"
  List1.AddItem "Sgn()"
  List1.AddItem "Sin()"
  List1.AddItem "Sqr()"
  List1.AddItem "Tan()"
  List1.AddItem "Val()"
  List1.Refresh
  Label1.Visible = False
  Text1.Visible = False
  Text1.Text = ""
  Label1.Caption = ""
  resultlabel.Caption = ""
End Sub

Private Sub List1_Click()
Text1.Text = ""
Text1.Visible = False
resultlabel.Caption = ""
Label1.Caption = ""
Label1.Visible = False
If List1.ListIndex = -1 Then
   MsgBox ("Please click to select a Math Function from the list below")
Else
   '***** IF ABS Function *****
   If List1.Text = "Abs()" Then
      Label1.Visible = True
      Text1.Visible = True
      Label1.Caption = "Give the number or Click on a cell"
   End If
   '***** IF ATN Function *****
   If List1.Text = "Atn()" Then
   End If
End If
End Sub

Private Sub Command2_Click() 'OK
'********** IF ABS Function *********
If List1.Text = "Abs()" Then
    arg1 = Text1.Text
    arg2 = MSFG.MSFlexGrid1.Text
    If arg1 <> "" And IsNumeric(arg1) Then  ' Show the result
        resultlabel.Caption = Abs(arg1)
    Else
        If arg2 <> "" And IsNumeric(arg2) Then
            resultlabel.Caption = Abs(arg2)
        End If
    End If
    Text1.Text = ""
    arg1 = Null
    arg2 = Null
    On Error GoTo CalcError
    Exit Sub
 End If
'********** IF ATN Function *********
If List1.Text = "Atn()" Then
End If

CalcError:
    MsgBox ("Math Function returned the following error:" & vbCrLf & Err.Description)
End Sub
