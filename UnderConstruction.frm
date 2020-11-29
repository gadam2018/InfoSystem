VERSION 5.00
Begin VB.Form UnderConstruction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Under Construction"
   ClientHeight    =   1320
   ClientLeft      =   36
   ClientTop       =   372
   ClientWidth     =   3588
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "UnderConstruction.frx":0000
   ScaleHeight     =   1320
   ScaleWidth      =   3588
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   372
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   732
   End
End
Attribute VB_Name = "UnderConstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
