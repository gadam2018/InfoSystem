VERSION 5.00
Begin VB.Form helpform 
   Caption         =   "HELP"
   ClientHeight    =   4980
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   7092
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   7092
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   372
      Left            =   3000
      TabIndex        =   1
      Top             =   4560
      Width           =   852
   End
   Begin VB.TextBox helptext 
      Height          =   4212
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   6132
   End
End
Attribute VB_Name = "helpform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSys As New FileSystemObject

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
HelpFile = App.Path & "\help.txt"
Set InStream = FSys.OpentextFile(HelpFile, ForReading)
While InStream.AtEndOfStream = False
tline = InStream.readline
txt = txt & tline & vbCrLf
Wend
helptext = txt
Set InStream = Nothing
'hlptext.Text = Input$(LOF(fnum), fnum)
End Sub
