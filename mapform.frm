VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0460CA20-346F-11CF-8682-00805F7CED21}#1.1#0"; "MO10.OCX"
Begin VB.Form mapform 
   Caption         =   "mapform"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MapObjects.Map Map1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _Version        =   65537
      _ExtentX        =   12515
      _ExtentY        =   11880
      _StockProps     =   225
      BackColor       =   16777215
      BorderStyle     =   1
      Contents        =   "mapform.frx":0000
   End
End
Attribute VB_Name = "mapform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim iLayer As New ImageLayer

 ' CommonDialog1.Filter = "Windows Bitmap (*.bmp)|*.bmp|TIFF Image(*.tif)|*.tif"
  
  'CommonDialog1.Filter = "Images|*.BMP;*.GIF;*.JPG;*.tif;*.DIB|All Files|*.*"
  CommonDialog1.FilterIndex = 1
  CommonDialog1.ShowOpen
  
  If CommonDialog1.FileName <> "" Then
    iLayer.File = CommonDialog1.FileName
    
    ' move the existing layer to the top
    If Map1.Layers.Add(iLayer) Then
      Map1.Layers.MoveToTop 1
    End If
    
  End If
End Sub

