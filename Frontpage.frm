VERSION 5.00
Begin VB.Form Frontpage 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "AgroModel"
   ClientHeight    =   5292
   ClientLeft      =   192
   ClientTop       =   1428
   ClientWidth     =   7008
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5292
   ScaleWidth      =   7008
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "stzortz@uth.gr"
      Height          =   252
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   11
      ToolTipText     =   "Email to Biometrical Lab of Thessalia University"
      Top             =   7896
      Width           =   1332
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1812
      Left            =   9000
      TabIndex        =   3
      Top             =   1800
      Width           =   2652
      Begin VB.Image Image4 
         Height          =   1392
         Left            =   312
         Picture         =   "Frontpage.frx":0000
         Stretch         =   -1  'True
         ToolTipText     =   "Other Agricultural Topics"
         Top             =   216
         Width           =   2040
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         Height          =   1812
         Left            =   0
         Top             =   0
         Width           =   2652
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00C00000&
         Height          =   1704
         Left            =   60
         Top             =   60
         Width           =   2532
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1812
      Left            =   6240
      TabIndex        =   2
      Top             =   1800
      Width           =   2652
      Begin VB.Image Image3 
         Height          =   1392
         Left            =   312
         Picture         =   "Frontpage.frx":7433
         Stretch         =   -1  'True
         ToolTipText     =   "View Agricultural Aids"
         Top             =   216
         Width           =   2040
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         Height          =   1812
         Left            =   0
         Top             =   0
         Width           =   2652
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00C00000&
         Height          =   1704
         Left            =   60
         Top             =   48
         Width           =   2532
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1812
      Left            =   3480
      TabIndex        =   1
      Top             =   1800
      Width           =   2652
      Begin VB.Image Image2 
         Height          =   1392
         Left            =   312
         Picture         =   "Frontpage.frx":E36B
         Stretch         =   -1  'True
         ToolTipText     =   "Enter Plant Databases"
         Top             =   216
         Width           =   2040
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         Height          =   1812
         Left            =   0
         Top             =   0
         Width           =   2652
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C00000&
         Height          =   1704
         Left            =   60
         Top             =   48
         Width           =   2532
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1812
      Left            =   720
      TabIndex        =   0
      Top             =   1800
      Width           =   2652
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C00000&
         Height          =   1704
         Left            =   60
         Top             =   60
         Width           =   2532
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         Height          =   1812
         Left            =   0
         Top             =   0
         Width           =   2652
      End
      Begin VB.Image Image1 
         Height          =   1392
         Left            =   312
         Picture         =   "Frontpage.frx":164AE
         Stretch         =   -1  'True
         ToolTipText     =   "Enter Animal Databases"
         Top             =   216
         Width           =   2040
      End
   End
   Begin VB.Shape Shape21 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   5040
      Top             =   4320
      Width           =   2652
   End
   Begin VB.Shape Shape20 
      BorderColor     =   &H00C00000&
      Height          =   420
      Left            =   5076
      Top             =   4344
      Width           =   2568
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   792
      Left            =   5880
      Picture         =   "Frontpage.frx":1CA49
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   912
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "For any comments contact:"
      Height          =   252
      Left            =   4680
      TabIndex        =   13
      Top             =   7920
      Width           =   2052
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 1999 Agricultural Dept. University of Thessalia"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   7.8
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4560
      TabIndex        =   12
      Top             =   8160
      Width           =   3972
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frontpage.frx":1D511
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   2880
      TabIndex        =   10
      Top             =   7080
      Width           =   7212
   End
   Begin VB.Shape Shape17 
      BorderWidth     =   2
      Height          =   852
      Left            =   4680
      Top             =   5040
      Width           =   3288
   End
   Begin VB.Shape Shape16 
      BorderColor     =   &H00C00000&
      Height          =   420
      Left            =   9036
      Top             =   3744
      Width           =   2568
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   9000
      Top             =   3720
      Width           =   2652
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Other Topics"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   9000
      TabIndex        =   9
      ToolTipText     =   "Other Agricultural Topics"
      Top             =   3744
      Width           =   2652
   End
   Begin VB.Shape Shape14 
      BorderColor     =   &H00C00000&
      Height          =   420
      Left            =   6276
      Top             =   3744
      Width           =   2568
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   6240
      Top             =   3720
      Width           =   2652
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Agricultural Aids"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   6240
      TabIndex        =   8
      ToolTipText     =   "View Agricultural Aids"
      Top             =   3744
      Width           =   2652
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H00C00000&
      Height          =   420
      Left            =   3516
      Top             =   3744
      Width           =   2568
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   3480
      Top             =   3720
      Width           =   2652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Plant Science"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   3480
      TabIndex        =   7
      ToolTipText     =   "Enter Plant Databases"
      Top             =   3744
      Width           =   2652
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00C00000&
      Height          =   420
      Left            =   756
      Top             =   3744
      Width           =   2568
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   720
      Top             =   3720
      Width           =   2652
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AgroModel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1092
      Left            =   3720
      TabIndex        =   6
      Top             =   600
      Width           =   4812
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LAB OF BIOMETRY    UNIVERSITY OF THESSALIA  GREECE 1999"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   4680
      TabIndex        =   5
      Top             =   5040
      Width           =   3300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Animal Science"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   720
      TabIndex        =   4
      ToolTipText     =   "Enter Animal Databases"
      Top             =   3744
      Width           =   2652
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Links"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   5040
      TabIndex        =   14
      ToolTipText     =   "Links to Agriculture"
      Top             =   4344
      Width           =   2652
   End
End
Attribute VB_Name = "Frontpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'URLstring = "http://www.uth.gr"
'Window.Navigate URLstring
End Sub


Private Sub Image1_Click()
Load Master
Master.Show
Unload Me
End Sub


Private Sub Image2_Click()
Load Master
Master.Show
Unload Me
End Sub

Private Sub Label1_Click()
Load Master
Master.Show
Unload Me
End Sub


Private Sub Label11_Click() ' Links
Load Master
Master.Show
Load links
links.Show
Unload Me
End Sub

Private Sub Label2_Click()
Load Master
Master.Show
Unload Me
End Sub
