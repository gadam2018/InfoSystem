VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDataGrid 
   ClientHeight    =   4584
   ClientLeft      =   1656
   ClientTop       =   1548
   ClientWidth     =   6144
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4590
   ScaleMode       =   0  'User
   ScaleWidth      =   6150
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   336
      ScaleWidth      =   6144
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6144
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   330
         Left            =   4398
         TabIndex        =   5
         Tag             =   "&Close"
         Top             =   0
         Width           =   1437
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Filter"
         Height          =   330
         Left            =   2924
         TabIndex        =   4
         Tag             =   "&Filter"
         Top             =   0
         Width           =   1462
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "&Sort"
         Height          =   330
         Left            =   1462
         TabIndex        =   3
         Tag             =   "&Sort"
         Top             =   0
         Width           =   1462
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   330
         Left            =   0
         TabIndex        =   2
         Tag             =   "&Refresh"
         Top             =   0
         Width           =   1462
      End
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
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
      RecordSource    =   "Datach10"
      Top             =   4236
      Width           =   6144
   End
   Begin MSDBGrid.DBGrid grdDataGrid 
      Align           =   1  'Align Top
      Bindings        =   "frmDataGrid.frx":0000
      Height          =   3648
      Left            =   0
      OleObjectBlob   =   "frmDataGrid.frx":00D2
      TabIndex        =   0
      Top             =   336
      Width           =   6144
   End
End
Attribute VB_Name = "frmDataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit


Dim msSortCol As String
Dim mbCtrlKey As Integer


Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFilter_Click()
    On Error GoTo FilterErr

    Dim recRecordset1 As Recordset, recRecordset2 As Recordset
    Dim sFilterStr As String

    If Data1.RecordsetType = vbRSTypeTable Then
        Beep
        MsgBox "You Cannot Filter a Table Recordset!", 48
        Exit Sub
    End If
    
    Set recRecordset1 = Data1.Recordset                        'copy the recordset
    
    sFilterStr = InputBox("Enter Filter Expression:")
    If Len(sFilterStr) = 0 Then Exit Sub

    Screen.MousePointer = vbHourglass
    recRecordset1.Filter = sFilterStr
    Set recRecordset2 = recRecordset1.OpenRecordset(recRecordset1.Type) 'establish the filter
    Set Data1.Recordset = recRecordset2                        'assign back to original recordset object

    Screen.MousePointer = vbDefault
    Exit Sub

FilterErr:
    Screen.MousePointer = vbDefault
    MsgBox "Error:" & Err & " " & Err.Description
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo RefErr
    
    
    Data1.RecordSource = gsrecordsource
    Data1.Refresh
    grdDataGrid.Refresh
    'Data1.Recordset.Requery
    
    Exit Sub
    
RefErr:
    MsgBox "Error:" & Err & " " & Err.Description
End Sub


Private Sub cmdSort_Click()
    On Error GoTo SortErr

    Dim recRecordset1 As Recordset, recRecordset2 As Recordset
    Dim SortStr As String

    If Data1.RecordsetType = vbRSTypeTable Then
        Beep
        MsgBox "You Cannot Sort a Table Recordset!", 48
        Exit Sub
    End If

    Set recRecordset1 = Data1.Recordset                        'copy the recordset
    
    If Len(msSortCol) = 0 Then
        SortStr = InputBox("Enter Sort Column:")
        If Len(SortStr) = 0 Then Exit Sub
    Else
        SortStr = msSortCol
    End If

    Screen.MousePointer = vbHourglass
    recRecordset1.Sort = SortStr
    
    'establish the Sort
    Set recRecordset2 = recRecordset1.OpenRecordset(recRecordset1.Type)
    Set Data1.Recordset = recRecordset2
    
    Screen.MousePointer = vbDefault
    Exit Sub

SortErr:
    Screen.MousePointer = vbDefault
    MsgBox "Error:" & Err & " " & Err.Description
End Sub


Private Sub Form_Load()

   gsconnect = OpenFile.globalConnect
   gsdatabase = OpenFile.globalDataBase
   gsrecordsource = OpenFile.globalRecordsource
    
    Dim bParmQry As Integer
    Dim qdfTmp As QueryDef
    
    On Error GoTo LoadErr
      
    Data1.Connect = gsconnect
    Data1.DatabaseName = gsdatabase
    Data1.RecordSource = gsrecordsource
    Data1.RecordsetType = 1     'dynaset
    Data1.Options = 0
    Data1.Refresh

    'If Len(Data1.RecordSource) > 50 Then
     '   Me.Caption = "SQL Statement"
    'Else
       ' Me.Caption = Data1.RecordSource
  '  End If
       
    Exit Sub

LoadErr:
    MsgBox "Error:" & Err & " " & Err.Description
    Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> 1 Then
        grdDataGrid.Height = Me.Height - (425 + picButtons.Height)
    End If
End Sub

Private Sub grdDataGrid_BeforeDelete(Cancel As Integer)
    If MsgBox("Delete Current Row?", vbYesNo + vbQuestion) <> vbYes Then
        Cancel = True
    End If
End Sub

Private Sub grdDataGrid_BeforeUpdate(Cancel As Integer)
    If MsgBox("Commit changes?", vbYesNo + vbQuestion) <> vbYes Then
        Cancel = True
    End If
End Sub

Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
    'let's sort on this column
    If Data1.RecordsetType = vbRSTypeTable Then Exit Sub
    
    'check for the use of the ctrl key for descending sort
    If mbCtrlKey Then
        msSortCol = "[" & Data1.Recordset(ColIndex).Name & "] desc"
        mbCtrlKey = 0 'reset it
    Else
        msSortCol = "[" & Data1.Recordset(ColIndex).Name & "]"
    End If
    cmdSort_Click
    msSortCol = vbNullString 'reset it
    
End Sub

Private Sub grdDataGrid_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    mbCtrlKey = Shift
End Sub

Private Sub data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'This will display the current record position for dynasets and snapshots
  Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
End Sub


