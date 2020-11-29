VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Begin VB.Form GraphForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Function Plot"
   ClientHeight    =   4992
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6744
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4992
   ScaleWidth      =   6744
   StartUpPosition =   3  'Windows Default
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   0
      Top             =   720
      _ExtentX        =   995
      _ExtentY        =   995
      AllowUI         =   -1  'True
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1575
      TabIndex        =   5
      Text            =   "Cos(3*X)*Sin(5*X)"
      Top             =   450
      Width           =   5040
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1575
      TabIndex        =   4
      Text            =   "Exp(2/X)*Cos(2*X)"
      Top             =   105
      Width           =   5040
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Draw both functions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4530
      TabIndex        =   3
      Top             =   4470
      Width           =   2085
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Draw second function"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2310
      TabIndex        =   2
      Top             =   4470
      Width           =   2085
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw first function"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   105
      TabIndex        =   1
      Top             =   4470
      Width           =   2085
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3480
      Left            =   105
      ScaleHeight     =   286
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   539
      TabIndex        =   0
      Top             =   855
      Width           =   6510
   End
   Begin VB.Label Label2 
      Caption         =   "Function #2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   7
      Top             =   495
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Function #1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   6
      Top             =   150
      Width           =   1365
   End
End
Attribute VB_Name = "GraphForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function FunctionEval1(ByVal X As Long) As Double
    'MsgBox ("in function X=" & X)
    'X = CDec(X)
    ScriptControl1.ExecuteStatement "X=" & X
    FunctionEval1 = ScriptControl1.Eval(Trim(Text1.Text))
    'MsgBox ("in function FunctionEval1=" & FunctionEval1)
    'FunctionEval1 = Exp(2 / X) * Cos(2 * X)
End Function

Function FunctionEval2(ByVal X As Long) As Double
    ScriptControl1.AddCode "X=" & X
    FunctionEval2 = ScriptControl1.Eval(Trim(Text2.Text))
End Function

Private Sub Command1_Click() 'DRAW FIRST FUNCTION
Dim t  As Double
Dim XMin As Double, XMax As Double, YMin As Double, YMax As Double
Dim XPixels As Integer

On Error GoTo FncError
    YMin = 1E+101: YMax = -1E+101
    XMin = 2: XMax = 10
    Picture1.Cls
    Picture1.ScaleMode = 3
    XPixels = Picture1.ScaleWidth - 1
    Me.Caption = "Calculating range..."
    Screen.MousePointer = vbHourglass
    ' Calculate Min and Max for Y axis
    For i = 1 To XPixels
        t = XMin + (XMax - XMin) * i / XPixels
        'MsgBox ("t=" & t)
        functionval = FunctionEval1(t)
       ' MsgBox ("functionval=" & functionval)
        If functionval > YMax Then YMax = functionval
        If functionval < YMin Then YMin = functionval
    Next
    Me.Caption = "Plotting function..."
    ' Set up a user defined scale mode
    Picture1.Scale (XMin, YMin)-(XMax, YMax)
    Picture1.ForeColor = RGB(0, 0, 255)
    ' Move to the first point
    Picture1.PSet (XMin, FunctionEval1(XMin))
    ' Plot the function
    For i = 0 To XPixels
        t = XMin + (XMax - XMin) * i / XPixels
        'Picture1.PSet (t, FunctionEval1(t))
        Picture1.Line -(t, FunctionEval1(t))
    Next
    Me.Caption = "Function Plot"
    Screen.MousePointer = vbDefault
    Exit Sub
FncError:
    MsgBox "There was an error in evaluating the function"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click() ' DRAW SECOND FUNCTION
Dim t As Double
Dim XMin As Double, XMax As Double, YMin As Double, YMax As Double
Dim XPixels As Integer

    YMin = 1E+101: YMax = -1E+101
    XMin = 2: XMax = 10
    Picture1.Cls
    Picture1.ScaleMode = 3
    XPixels = Picture1.ScaleWidth - 1
    Me.Caption = "Calculating range..."
    Screen.MousePointer = vbHourglass
    ' Calculate Min and Max for Y axis
    For i = 0 To XPixels
        t = XMin + (XMax - XMin) * i / XPixels
        functionval = FunctionEval2(t)
        If functionval > YMax Then YMax = functionval
        If functionval < YMin Then YMin = functionval
    Next
    Me.Caption = "Plotting function..."
    ' Set up a user defined scale mode
        Picture1.Scale (XMin, YMin)-(XMax, YMax)
    Picture1.ForeColor = RGB(255, 0, 0)
    ' Move to the first point
    Picture1.PSet (XMin, FunctionEval1(XMin))
    ' Plot the function
    For i = 0 To XPixels - 1
        t = XMin + (XMax - XMin) * i / XPixels
        functionval = FunctionEval2(t)
        'Picture1.PSet (t, functionVal)
        Picture1.Line -(t, functionval)
    Next
    Me.Caption = "Function Plot"
    Screen.MousePointer = vbDefault
    Exit Sub
FncError:
    MsgBox "There was an error in evaluating the function"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command3_Click() 'DRAW BOTH FUNCTIONS
Dim t As Double
Dim XMin As Double, XMax As Double, YMin As Double, YMax As Double
Dim XPixels As Integer

    YMin = 1E+101: YMax = -1E+101
    XMin = 2: XMax = 10
    Picture1.Cls
    Picture1.ScaleMode = 3
    XPixels = Picture1.ScaleWidth - 1
    Me.Caption = "Calculating range..."
    Screen.MousePointer = vbHourglass
    ' Calculate Min and Max for Y axis
    For i = 1 To XPixels
        t = XMin + (XMax - XMin) * i / XPixels
        functionval = FunctionEval1(t)
        If functionval > YMax Then YMax = functionval
        If functionval < YMin Then YMin = functionval
    Next

    Me.Caption = "Plotting functions..."
    ' Set up a user defined scale mode
    Picture1.Scale (XMin, YMin)-(XMax, YMax)

    Picture1.ForeColor = RGB(0, 0, 255)
    ' Move to the first point
    Picture1.PSet (XMin, FunctionEval1(XMin))
    ' Plot the function
    For i = 0 To XPixels
        t = XMin + (XMax - XMin) * i / XPixels
        'Picture1.PSet (t, FunctionEval1(t))
        Picture1.Line -(t, FunctionEval1(t))
    Next

    Picture1.ForeColor = RGB(255, 0, 0)
    Picture1.PSet (XMin, FunctionEval2(XMin))
    ' Plot the function
    For i = 0 To XPixels
        t = XMin + (XMax - XMin) * i / XPixels
        'Picture1.PSet (t, FunctionEval2(t))
        Picture1.Line -(t, FunctionEval2(t))
    Next
    Me.Caption = "Function Plot"
    Screen.MousePointer = vbDefault
    Exit Sub
FncError:
    MsgBox "There was an error in evaluating the function"
    Screen.MousePointer = vbHourglass
End Sub

Private Sub ScriptControl1_Error()
    Debug.Print ScriptControl1.Error.Number
    Debug.Print ScriptControl1.Error.Text
End Sub
