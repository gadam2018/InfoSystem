VERSION 5.00
Begin VB.MDIForm Master 
   BackColor       =   &H00808000&
   Caption         =   "GisTool"
   ClientHeight    =   6705
   ClientLeft      =   135
   ClientTop       =   465
   ClientWidth     =   8430
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu Basic 
      Caption         =   "Data Management"
      Begin VB.Menu ImageLoadForm 
         Caption         =   "Load ImageLoadForm"
      End
      Begin VB.Menu Open1 
         Caption         =   "Load mapfile"
      End
      Begin VB.Menu Open1new 
         Caption         =   "Load WorkFile (.mdb, .xls, .dbf)"
      End
      Begin VB.Menu Close1 
         Caption         =   "Close WorkFile"
      End
      Begin VB.Menu ADRU 
         Caption         =   "Add, Delete, Refresh, Update"
      End
      Begin VB.Menu SortFilter 
         Caption         =   "Sort & Filter"
      End
      Begin VB.Menu Search 
         Caption         =   "Search"
      End
      Begin VB.Menu Reports 
         Caption         =   "Reports and Printouts"
         Begin VB.Menu AnimalReports 
            Caption         =   "Animal Science Reports"
            Begin VB.Menu CSDH_rep 
               Caption         =   "CSDH Ids Table Report"
               Checked         =   -1  'True
            End
            Begin VB.Menu Herds_rep 
               Caption         =   "Herds Table Report"
            End
            Begin VB.Menu SIRES_rep 
               Caption         =   "Sires Table Report"
            End
         End
         Begin VB.Menu PlantReports 
            Caption         =   "Plant Science Reports"
            Begin VB.Menu CercoReport 
               Caption         =   "Cercospora Report"
            End
         End
      End
      Begin VB.Menu QuitModel 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu DAnalysis 
      Caption         =   "Data Analysis"
      Begin VB.Menu BuiltInFunc 
         Caption         =   "Built in Functions"
         Begin VB.Menu Math 
            Caption         =   "Mathematical"
         End
         Begin VB.Menu Stat 
            Caption         =   "Statistical"
         End
         Begin VB.Menu Finance 
            Caption         =   "Financial"
            Begin VB.Menu vbfv 
               Caption         =   "fv"
            End
         End
         Begin VB.Menu Graph 
            Caption         =   "Graph"
         End
      End
      Begin VB.Menu Calculate 
         Caption         =   "Calculate"
         Begin VB.Menu ExcelFormula 
            Caption         =   "Formula in Excel"
         End
         Begin VB.Menu ExcelFunction 
            Caption         =   "Excel Function"
         End
      End
      Begin VB.Menu ExcelAnalysis 
         Caption         =   "Excel Analysis"
         Begin VB.Menu MathAnalysis 
            Caption         =   "Mathematical Analysis"
         End
         Begin VB.Menu StatAnalysis 
            Caption         =   "Statistical Analysis"
         End
      End
      Begin VB.Menu Open2 
         Caption         =   "Show WorkFile in Excel"
      End
   End
   Begin VB.Menu Advanced 
      Caption         =   "Data processing"
      Begin VB.Menu ExecSQL 
         Caption         =   "Execute SQL"
      End
      Begin VB.Menu Open3 
         Caption         =   "Open mdb Query"
         Enabled         =   0   'False
      End
      Begin VB.Menu q2 
         Caption         =   "Open mdb query2"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Packages 
      Caption         =   "Tools"
      Begin VB.Menu Mathematica 
         Caption         =   "MATHEMATICA Analysis"
      End
      Begin VB.Menu Mstatc 
         Caption         =   "Mstatc Analysis"
      End
      Begin VB.Menu Spss 
         Caption         =   "SPSS Analysis"
      End
      Begin VB.Menu Statistica 
         Caption         =   "STATISTICA Analysis"
      End
      Begin VB.Menu Systat 
         Caption         =   "Systat Analysis"
      End
   End
   Begin VB.Menu InternetLinks 
      Caption         =   "Links"
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
   End
   Begin VB.Menu End 
      Caption         =   "End"
   End
End
Attribute VB_Name = "Master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ExcelApp As Excel.Application
Public WordApp As Word.Application
Dim wsheet As Worksheet
Dim wbook As Workbook
Dim wdoc As Document

Private Sub ADRU_Click()
Load Datach10Form
Datach10Form.Show
End Sub


Private Sub CercoReport_Click()
Cerco_report.Show
End Sub

Private Sub Close1_Click()
globalConnect = Null
globalDataBase = Null
globalRecordsource = Null
Master.SortFilter.Enabled = False 'Reset these options
Master.Search.Enabled = False
Master.ADRU.Enabled = False
Unload MSFG
Master.Close1.Enabled = False
End Sub

Private Sub CSDH_rep_Click()
CSDH_report.Show
End Sub

Private Sub End_Click() '*** EXIT THE APPLICATION ***
msg = "Do you want to exit application ?"   ' Define message.
style = vbYesNo ' Define buttons.
title = "Quit AgroModel?"  ' Define title.

Response = MsgBox(msg, style, title)
If Response = vbYes Then    ' User chose Yes.
Unload Me
Load frmAbout
frmAbout.Show
Else
End If
End Sub

Private Sub ExcelFormula_Click() '*** Calculate(evaluate) an EXCEL formula/function
Dim expression
Set ExcelApp = CreateObject("Excel.Application")
If ExcelApp Is Nothing Then
   MsgBox ("Could not start Excel")
   End
End If
expression = InputBox("Enter expression (i.e. formula: 10/5, function: min(10,4), etc) to evaluate")
If vbOK Then
    If Trim(expression) <> "" Then
        MsgBox ExcelApp.Evaluate(expression)
    End If
Else
End If
On Error GoTo CalcError

GoTo Terminate
Exit Sub
CalcError:
    MsgBox ("Excel returned the following error:" & vbCrLf & Err.Description)
Terminate:
   ExcelApp.Quit
   'Set ExcelApp = Nothing
End Sub

Private Sub ExcelFunction_Click() '****
    Load ExcelFunctions
    ExcelFunctions.Show
End Sub

Private Sub ExecSQL_Click()
Load SQLForm
SQLForm.Show
End Sub

Private Sub Graph_Click()
Load GraphForm
GraphForm.Show
End Sub
Private Sub Help_Click() ' *** NEW HELP ***
Load helpform
helpform.Show
End Sub
Private Sub Help_OLD_Click() ' *** OLD HELP ***
Set WordApp = CreateObject("Word.Application")
If WordApp Is Nothing Then
   MsgBox ("Could not start Word")
   End
End If
On Error GoTo OpenWordAppError
'WordApp.Documents.Open ("C:\AgroModel\help.doc")
WordApp.Documents.Open (CurDir & "\" & "help.doc")
WordApp.Visible = True
'WordApp.Documents.Close ("C:\AgroModel\help.doc")
GoTo wordterminate

OpenWordAppError:
    MsgBox ("Word returned the following error:" & vbCrLf & Err.Description)

wordterminate:
  MsgBox ("Help File is closed")
 ' WordApp.Quit
End Sub

Private Sub Herds_rep_Click()
Herds_report.Show
End Sub

Private Sub ImageLoadForm_Click()
Load Form1
Form1.Show
End Sub

Private Sub InternetLinks_Click()
Load links
links.Show
End Sub

Private Sub Math_Click()
Load BInMaths
BInMaths.Show
End Sub

Private Sub MathAnalysis_Click()
Load Math_Analysis
Math_Analysis.Show
End Sub

Private Sub MDIForm_Load()
Master.Close1.Enabled = False 'Reset these options
Master.SortFilter.Enabled = False
Master.Search.Enabled = False
Master.ADRU.Enabled = False
'Master.Advanced.Enabled = False

End Sub

Private Sub Modelling_Click()
'MsgBox ("Sorry, it's Under Construction")
Load UnderConstruction
UnderConstruction.Show
End Sub

Private Sub Open1_Click() ' Open a FILE (.mdb, or .xls)
'Load OpenFile
'OpenFile.Show
Load mapform
mapform.Show
End Sub

Private Sub Open1new_Click()
Load OpenFile
OpenFile.Show
End Sub

Private Sub Open2_Click() ' **** Open in EXCEL ****
Dim wkbookstring As String

Unload MSFG '*** Unload MSFG Form
            '*** REQUIRED in order for the file
            '*** to be loaded in EXCEL

Set ExcelApp = CreateObject("Excel.Application")
If ExcelApp Is Nothing Then
   MsgBox ("Could not start Excel")
   End
End If
'wkbookstring = InputBox("Enter Excel workbook fullpath and name" & vbCr & "(" & "by default=c:\adamg\animalxls.xls")
wkbookstring = OpenFile.globalDataBase

'MsgBox (wkbookstring)
If Len(wkbookstring) = 0 Then
   'ExcelApp.Workbooks.Open ("c:\AgroModel\animalxls.xls") 'BY DEFAULT
    GoTo TerminateExcelApp
Else
   ExcelApp.Workbooks.Open (wkbookstring)
End If
On Error GoTo OpenExcelAppError
ExcelApp.Visible = True
'Set wbook = ExcelApp.ActiveWorkbook
'Set wsheet = ExcelApp.ActiveSheet

GoTo TerminateExcelApp

OpenExcelAppError:
    MsgBox ("Excel returned the following error:" & vbCrLf & Err.Description)
TerminateExcelApp:
MsgBox ("Excel File Work is Finished") ' REQUIRED in order to stop!!!
    ExcelApp.Quit
    ' Set ExcelApp = Nothing
Load MSFG
MSFG.Show
End Sub

Private Sub Open3_Click()
Load QueriesForm
QueriesForm.Show
End Sub


Private Sub q2_Click()
Load q2form
q2form.Show
End Sub



Private Sub QuitModel_Click()
msg = "Do you want to quit application ?"   ' Define message.
style = vbYesNo ' Define buttons.
title = "Quit AgroModel?"  ' Define title.

Response = MsgBox(msg, style, title)
If Response = vbYes Then    ' User chose Yes.
Unload Me
Load frmAbout
frmAbout.Show
Else
End If
End Sub

Private Sub Search_Click()
'If OpenFile.globalConnect Then
  ' MsgBox ("You have to Load a Work File first")
  ' Unload Me
'Else
  Load SearchForm
  SearchForm.Show
'End If
End Sub

Private Sub SIRES_rep_Click()
SIRES_report.Show
End Sub

Private Sub SortFilter_Click()
Load frmDataGrid
frmDataGrid.Show
End Sub


Private Sub Stat_Click()
Load BInStat
BInStat.Show
End Sub

Private Sub StatAnalysis_Click()
Load Stat_Analysis
Stat_Analysis.Show
End Sub



Private Sub vbfv_Click() 'FV(rate, nper, pmt[, pv[, type]])
Dim Fmt, Payment, APR, TotPmts, PayType, pvalue, fvalue
Const ENDPERIOD = 0, BEGINPERIOD = 1   ' When payments are made.
Fmt = "###,###,##0.00"   ' Define money format.
Payment = InputBox("How much do you plan to save each month?")
APR = InputBox("Enter the expected interest annual percentage rate.")
If APR > 1 Then APR = APR / 100   ' Ensure proper form.
TotPmts = InputBox("For how many months do you expect to save?")
PayType = MsgBox("Do you make payments at the end of month?", vbYesNo)
If PayType = vbNo Then PayType = BEGINPERIOD Else PayType = ENDPERIOD
pvalue = InputBox("How much is in this savings account now?")
fvalue = FV(APR / 12, TotPmts, -Payment, -pvalue, PayType)
MsgBox "Your savings will be worth " & Format(fvalue, Fmt) & "."
End Sub

Private Sub Mathematica_Click()
Dim retval
retval = Shell(App.Path & "\WNMATH22\FE.exe")
End Sub
Private Sub Mstatc_Click()
Dim retval
retval = Shell(App.Path & "\mstatc\mstatc.exe")
End Sub
Private Sub Spss_Click()
Dim retval
retval = Shell(App.Path & "\SPSS\spsswin.exe")
End Sub
Private Sub Statistica_Click()
Dim retval
retval = Shell(App.Path & "\stat\Sta_win.exe")
End Sub
Private Sub Systat_Click()
Dim retval
retval = Shell(App.Path & "\systat\systat.exe")
End Sub

