VERSION 5.00
Begin VB.Form StartForm 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7212
   ClientLeft      =   252
   ClientTop       =   1416
   ClientWidth     =   12216
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "StartForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7212
   ScaleWidth      =   12216
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2400
      Top             =   720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   4272
      Left            =   2400
      TabIndex        =   0
      Top             =   1800
      Width           =   7296
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "CROP and ANIMAL PRODUCTION"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   2520
         TabIndex        =   10
         Top             =   2280
         Width           =   4548
      End
      Begin VB.Image imgLogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   1428
         Left            =   360
         Picture         =   "StartForm.frx":000C
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   2052
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "MANAGEMENT ENVIRONMENT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   9
         Top             =   3480
         Width           =   4320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "AGRICULTURAL DATA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   3108
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "LAB OF BIOMETRY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3360
         TabIndex        =   7
         Top             =   2640
         Width           =   2640
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "OF"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   420
         Left            =   4440
         TabIndex        =   5
         Top             =   1440
         Width           =   504
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "FACULTY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   420
         Left            =   3840
         TabIndex        =   6
         Top             =   1080
         Width           =   1752
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "AGRICULTURE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   22.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   504
         Left            =   3120
         TabIndex        =   4
         Top             =   1800
         Width           =   3348
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00FF8080&
         Caption         =   "Copyright August 1999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   5280
         TabIndex        =   1
         Top             =   3960
         Width           =   1932
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Version 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   276
         Left            =   960
         TabIndex        =   2
         Top             =   3840
         Width           =   960
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UNIVERSITY OF THESSALIA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   456
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   4800
      End
   End
End
Attribute VB_Name = "StartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  
End Sub

Private Sub Frame1_Click()
    Unload Me
    'Load frmSplash
    'frmSplash.Show
    'Load Master
    'Master.Show
    Load Frontpage
    Frontpage.Show
End Sub

Private Sub Timer1_Timer()
Unload Me
'Load frmSplash
'frmSplash.Show
'Load Master
'Master.Show
 Load Frontpage
 Frontpage.Show
End Sub
