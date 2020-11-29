VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form links 
   Caption         =   "Agricultural Links"
   ClientHeight    =   6768
   ClientLeft      =   3060
   ClientTop       =   3348
   ClientWidth     =   7452
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6768
   ScaleWidth      =   7452
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tToolBar 
      Align           =   1  'Align Top
      Height          =   432
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7452
      _ExtentX        =   13145
      _ExtentY        =   762
      ButtonWidth     =   656
      ButtonHeight    =   635
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3732
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   5388
      ExtentX         =   9525
      ExtentY         =   6588
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6120
      Top             =   2160
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   0
      ScaleHeight     =   1140
      ScaleWidth      =   7452
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   432
      Width           =   7452
      Begin VB.ComboBox cbo2 
         Height          =   288
         Left            =   0
         TabIndex        =   6
         Text            =   "http://www.minagric.gr"
         Top             =   840
         Width           =   3732
      End
      Begin VB.ComboBox cboAddress 
         Height          =   288
         Left            =   0
         TabIndex        =   2
         Text            =   "http://www.uth.gr"
         Top             =   240
         Width           =   3732
      End
      Begin VB.Label Label1 
         Caption         =   "Plant Links"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Tag             =   "&Address:"
         Top             =   600
         Width           =   3072
      End
      Begin VB.Label lblAddress 
         Caption         =   "Animal Links"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   1
         Tag             =   "&Address:"
         Top             =   0
         Width           =   3072
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2670
      Top             =   2325
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "links.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "links.frx":02E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "links.frx":05C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "links.frx":08A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "links.frx":0B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "links.frx":0E6A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "links"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean


Private Sub Form_Load()
    On Error Resume Next
    Me.Show
    tToolBar.Refresh
    Form_Resize

    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15

    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If
cboAddress.AddItem "http://www.uth.gr"
cboAddress.AddItem "http://www.agworks.com"
cboAddress.AddItem "http://www.cattlepages.com"
cboAddress.AddItem "http://www.farmjournal.com"
cboAddress.AddItem "http://www.ag.iastate.edu/international/conf.html"
cboAddress.AddItem "http://www.ucd.ie/cigr/"
cboAddress.ListIndex = 0
cbo2.AddItem "http://www.minagric.gr"
cbo2.AddItem "http://www.agworks.com"
cbo2.ListIndex = 0
End Sub



Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub

Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub cbo2_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cbo2.Text
End Sub

Private Sub cbo2_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cbo2_Click
    End If
End Sub

Private Sub Form_Resize()
    cboAddress.Width = Me.ScaleWidth - 100
    cbo2.Width = Me.ScaleWidth - 100
    brwWebBrowser.Width = Me.ScaleWidth - 100
    brwWebBrowser.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height) - 100
End Sub


Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Working..."
    End If
End Sub
'This should not be workable
Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
     
    timTimer.Enabled = True
     
    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.Refresh
        Case "Home"
            brwWebBrowser.GoHome
        Case "Search"
            brwWebBrowser.GoSearch
        Case "Stop"
            timTimer.Enabled = False
            brwWebBrowser.Stop
            Me.Caption = brwWebBrowser.LocationName
    End Select

End Sub


Private Sub tToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
     
    timTimer.Enabled = True
     
    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.Refresh
        Case "Home"
            brwWebBrowser.GoHome
        Case "Search"
            brwWebBrowser.GoSearch
        Case "Stop"
            timTimer.Enabled = False
            brwWebBrowser.Stop
            Me.Caption = brwWebBrowser.LocationName
    End Select
End Sub
