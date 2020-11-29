VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Datach10Form 
   Caption         =   "Datach10"
   ClientHeight    =   5316
   ClientLeft      =   1116
   ClientTop       =   396
   ClientWidth     =   6744
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8784
   ScaleWidth      =   12192
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   12192
      TabIndex        =   1
      Top             =   8136
      Width           =   12192
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4505
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   3409
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   2313
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   1217
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   121
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Data datPrimaryRS 
      Align           =   2  'Align Bottom
      Caption         =   " "
      Connect         =   "Access"
      DatabaseName    =   "C:\adamg\Datach10.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"Datach10Form.frx":0000
      Top             =   8436
      Width           =   12192
   End
   Begin MSDBGrid.DBGrid grdDataGrid 
      Align           =   1  'Align Top
      Bindings        =   "Datach10Form.frx":0420
      Height          =   8052
      Left            =   0
      OleObjectBlob   =   "Datach10Form.frx":057E
      TabIndex        =   0
      Top             =   0
      Width           =   12192
   End
End
Attribute VB_Name = "Datach10Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
  datPrimaryRS.Recordset.MoveLast
  grdDataGrid.SetFocus
  'MSFlexGrid1.SetFocus
  SendKeys "{down}"
End Sub

Private Sub cmdDelete_Click()
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  datPrimaryRS.Refresh
End Sub

Private Sub cmdUpdate_Click()
  datPrimaryRS.UpdateRecord
  datPrimaryRS.Recordset.Bookmark = datPrimaryRS.Recordset.LastModified
End Sub

Private Sub cmdClose_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub datPrimaryRS_Error(DataErr As Integer, Response As Integer)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Error$(DataErr)
  Response = 0  'Throw away the error
End Sub

Private Sub datPrimaryRS_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'This will display the current record position for dynasets and snapshots
  datPrimaryRS.Caption = "Record: " & (datPrimaryRS.Recordset.AbsolutePosition + 1)
End Sub

Private Sub datPrimaryRS_Validate(Action As Integer, Save As Integer)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
      Screen.MousePointer = vbDefault
  End Select
  Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_Load()
  gsconnect = OpenFile.globalConnect
  gsdatabase = OpenFile.globalDataBase
  gsrecordsource = OpenFile.globalRecordsource
 
    Dim bParmQry As Integer
    Dim qdfTmp As QueryDef
    
    On Error GoTo LoadErr
    
    datPrimaryRS.Connect = gsconnect
    datPrimaryRS.DatabaseName = gsdatabase
    datPrimaryRS.RecordSource = gsrecordsource
    datPrimaryRS.RecordsetType = 1     'dynaset
    datPrimaryRS.Options = 0
    datPrimaryRS.Refresh

    'If Len(datPrimaryRS.RecordSource) > 50 Then
    '    Me.Caption = "SQL Statement"
   ' Else
      '  Me.Caption = datPrimaryRS.RecordSource
   ' End If

    Exit Sub
   
LoadErr:
    MsgBox "Error:" & Err & " " & Err.Description
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  'MSFlexGrid1.Height = Me.ScaleHeight - datPrimaryRS.Height - picButtons.Height - 30
  grdDataGrid.Height = Me.ScaleHeight - datPrimaryRS.Height - picButtons.Height - 30
End Sub

'Private Sub grdDataGrid_Click()
'MsgBox (grdDataGrid.Col & grdDataGrid.Row)
'End Sub

'Private Sub MSFlexGrid1_Click()
'MsgBox (MSFlexGrid1.MouseCol & MSFlexGrid1.MouseRow)
'End Sub
