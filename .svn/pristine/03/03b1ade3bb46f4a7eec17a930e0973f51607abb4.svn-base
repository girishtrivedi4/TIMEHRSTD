VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDailyTry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Process"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   240
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFin 
      Cancel          =   -1  'True
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2520
      TabIndex        =   11
      Top             =   7440
      Width           =   1515
   End
   Begin VB.CommandButton cmdPro 
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   840
      TabIndex        =   10
      Top             =   7440
      Width           =   1515
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   585
      Left            =   3360
      TabIndex        =   19
      Top             =   6720
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   1032
      _Version        =   393216
      FixedCols       =   0
      ForeColorFixed  =   16711680
      Enabled         =   0   'False
      ScrollBars      =   0
   End
   Begin VB.Frame frEmp 
      Caption         =   "Employee Selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   5445
      Begin VB.CommandButton cmdSR 
         Caption         =   "Select Range"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3855
         TabIndex        =   6
         Top             =   1560
         Width           =   1515
      End
      Begin VB.CommandButton cmdUR 
         Caption         =   "Unselect Range"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3855
         TabIndex        =   7
         Top             =   2112
         Width           =   1515
      End
      Begin VB.CommandButton cmdSA 
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3855
         TabIndex        =   8
         Top             =   2664
         Width           =   1515
      End
      Begin VB.CommandButton cmdUA 
         Caption         =   "Unselect All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3855
         TabIndex        =   9
         Top             =   3216
         Width           =   1515
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   4365
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Click to Toggle Selection"
         Top             =   1140
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   7699
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSForms.ComboBox CboLocation 
         Height          =   315
         Left            =   3840
         TabIndex        =   22
         Top             =   720
         Width           =   1515
         VariousPropertyBits=   612390939
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2672;556"
         TextColumn      =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3960
         TabIndex        =   21
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lblPro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Processing Please Wait..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   5520
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.Label lblDeptCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin MSForms.ComboBox cboDept 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   2595
         VariousPropertyBits=   612390939
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4577;556"
         TextColumn      =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboTo 
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Top             =   720
         Width           =   1125
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1984;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboFrom 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         Width           =   1125
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1984;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2400
         TabIndex        =   18
         Top             =   720
         Width           =   210
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   600
         TabIndex        =   17
         Top             =   720
         Width           =   435
      End
   End
   Begin VB.Frame frDates 
      Caption         =   "Processing Dates"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   4605
      Begin VB.TextBox txtTo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         TabIndex        =   1
         Tag             =   "D"
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox txtFrom 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Tag             =   "D"
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label lblToD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   14
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblFromD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmDailyTry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim strSelEmp As String
Dim STRECODE As String
Dim adrsC As New ADODB.Recordset
Dim Con As New ADODB.Connection


Private Sub cboFrom_Change()
On Error Resume Next
cboTo.ListIndex = cboFrom.ListIndex
End Sub

Private Sub cboFrom_Click()
If cboFrom.ListIndex < 0 Then Exit Sub
''For Mauritius 19-08-2003
If cboTo.ListIndex = 0 Then Exit Sub
''
On Error Resume Next
 cboTo.ListIndex = cboFrom.ListIndex              ' for Removing standard error
'Resume Next
End Sub

Private Sub CboLocation_Change()
On Error GoTo ERR_P
'cbodept.Text = "ALL"
If cboDept.ListIndex < 0 Then Exit Sub               '' If No Department
Call FillLocaGrid
Call SelUnselAll(vbWhite, MSF1)
Exit Sub
ERR_P:
    ShowError ("Department::" & Me.Caption)
End Sub

Private Sub FillLocaGrid()     '' Fills Employee Combo and Grid
On Error GoTo ERR_P
Dim intEmpCnt As Integer, intTmpCnt As Integer
Dim strArrEmp() As String
Dim strDeptTmp As String, strTempforCF As String

intEmpCnt = 0
If CboLocation.Text = "ALL" Then
    strDeptTmp = "ALL"
Else
    strDeptTmp = CboLocation.List(CboLocation.ListIndex, 1)
End If


Select Case UCase(Trim(strDeptTmp))
    Case "", "ALL"
        strTempforCF = "Select Empcode,Name from empmst where (joindate is not null and joindate<=" & strDTEnc & _
        DateCompStr(Date) & strDTEnc & ") Order by EmpCode"
       'strTempforCF = "select Empcode,name from empmst where location=" & (cboLoc.Text) & " and company=" & (cboComp.Text) & " order by Empcode"
    Case Else
        If strCurrentUserType = HOD Then

            strTempforCF = "Select Empcode,Name from empmst " & strCurrData & " and (joindate is not null and joindate<=" & strDTEnc & DateCompStr(Date) & _
              strDTEnc & ") and Empmst." & SELCRIT1 & " = " & strDeptTmp & " Order by EmpCode"
         
        Else
            'this If condition add by  for datatype
'            If blnFlagForDept = True Then
                strTempforCF = "select Empcode,name from empmst where Empmst." & SELCRIT1 & "=" & _
                strDeptTmp & " order by Empcode"    'Empcode,name
'            End If
        End If
End Select

If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open strTempforCF, ConMain, adOpenStatic, adLockReadOnly
If (adrsEmp.EOF And adrsEmp.BOF) Then
    cboFrom.clear
    cboTo.clear
    MSF1.Rows = 1
    Exit Sub
End If
intEmpCnt = adrsEmp.RecordCount
intTmpCnt = intEmpCnt
MSF1.Rows = intEmpCnt + 1
ReDim strArrEmp(intTmpCnt - 1, 1)
For intEmpCnt = 0 To intTmpCnt - 1
    strArrEmp(intEmpCnt, 0) = adrsEmp(0)
    strArrEmp(intEmpCnt, 1) = adrsEmp(1)
    MSF1.TextMatrix(intEmpCnt + 1, 0) = adrsEmp(0)
    MSF1.TextMatrix(intEmpCnt + 1, 1) = adrsEmp(1)
    adrsEmp.MoveNext
Next
cboFrom.List = strArrEmp
cboTo.List = strArrEmp
cboFrom.ListIndex = 0
cboTo.ListIndex = cboTo.ListCount - 1
Erase strArrEmp
Exit Sub
ERR_P:
    ShowError ("Fill Combo Grid :: " & Me.Caption)
'    Resume Next
End Sub

Private Sub cmdFin_Click()
    Unload Me
End Sub


Private Sub cmdPro_Click()
'On Error GoTo ERR_P
'' Phase 1
If Not CheckDates Then      '' If not valid Dates then Exit
    Exit Sub
Else                        '' Set Dates to the Types
    typDT.dtFrom = DateCompDate(txtFrom.Text)
    typDT.dtTo = DateCompDate(txtTo.Text)
End If
    If Not CheckEmployee Then Exit Sub  '' If not valid Employees then Exit
    If Not CheckShifts Then Exit Sub    '' If not valid Monthly Shift Files then Exit
'    If Not CheckPrevious Then Exit Sub  '' If Previous Day processing not done then Exit
' Phase 2
Call AddActivityLog(lg_NoModeAction, 2, 25)     '' Process Log
Call AuditInfo("DAILY PROCESS", Me.Caption, "Daily Process For The Period " & txtFrom.Text & "To " & txtTo.Text)
'' Select Records from .Dat File into the RAW Dat Table
lblPro.Visible = True       '' Make Processing Label Visible
Call ChangeLabelCaption(1)  '' Set appropriate Caption
'Call AppendDatFileNew(lstDat, Me)
If Not AppendDataFile(Me) Then
    lblPro.Visible = False  '' Make processing Label Invisible
    MsgBox NewCaptionTxt("17011", adrsC), vbCritical
    Exit Sub
End If
  
DoEvents: Me.Refresh
''
Call FilterEmpty            '' Clears all the Records from tbldata which are blank

'' Phase 3
Call ChangeLabelCaption(2)  '' Set appropriate Caption
Call FillInstalltypes       '' Fills Details from Parameters to their respective Type

DoEvents: Me.Refresh
''
Call OpenMasters(strSelEmp)            '' Opens Necessary Master Tables

DoEvents: Me.Refresh
''
Call FilterOnDates          '' Filter on the basis of Dates
Me.Refresh
'' Phase 4
Call TruncateTable("DailyPro")

Call GetDataPunches(strSelEmp)         '' Puts Data in processing Table

''
Call FilterOnCard           '' Filter on the basis of Cards
Call PutFlag                '' Puts Flag to all Records
''Phase 5
'Call StartProcessing(MSF2, strSelEmp, Me) '' Starts the actual Data Processing
Call StartProcessing(MSF2, STRECODE, Me)  '' Starts the actual Data Processing
lblPro.Visible = False      '' Make processing Label Invisible
Call TruncateTable("DailyPro")  '' Clear Table DailyPro Afetr Daily Process
'''For Mauritius 14-08-2003
MsgBox NewCaptionTxt("17012", adrsC), vbInformation
StrGroup1 = ""
''

blnIrregular = False
''
txtFrom.SetFocus            '' Set Focus to the From Date Text Box
Exit Sub
ERR_P:
    ShowError ("Process :: " & Me.Caption)
End Sub

Private Sub cmdSA_Click()
    Call SelUnselAll(&HC0FFFF, MSF1)
End Sub

Private Sub cmdSR_Click()
    Call SelUnsel(&HC0FFFF, MSF1, cboFrom, cboTo)
End Sub

Private Sub cmdUA_Click()
    Call SelUnselAll(vbWhite, MSF1)
End Sub

Private Sub cmdUR_Click()
    Call SelUnsel(vbWhite, MSF1, cboFrom, cboTo)
End Sub

Private Sub Form_Activate()
On Error Resume Next
txtFrom.Enabled = True
txtTo.Enabled = True
txtFrom.SetFocus
Call GetRights
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)        '' Set the Form Icon
txtFrom.Enabled = False
txtTo.Enabled = False
Call SetToolTipText(Me)     '' Set the ToolTipText
Call RetCaptions            '' Set the Control Captions
'' Empty Grid
Call FillCombos
'' Empty .Dat List

'' Set Current Dates.
txtFrom.Text = DateDisp(CStr(Date))
txtTo.Text = DateDisp(CStr(Date))
lblLocation.Visible = GetFlagStatus("pratham")
CboLocation.Visible = GetFlagStatus("pratham")

End Sub

Private Sub cboDept_Click()
On Error GoTo ERR_P
'CboLocation.Text = "ALL"
If cboDept.ListIndex < 0 Then Exit Sub               '' If No Department
Call FillComboGrid
Call SelUnselAll(vbWhite, MSF1)
Exit Sub
ERR_P:
    ShowError ("Department::" & Me.Caption)
End Sub

Private Sub FillCombos()
On Error GoTo ERR_P

Call SetCritCombos(cboDept)
Call SetCritCombos1(CboLocation)
If strCurrentUserType <> HOD Then cboDept.Text = "ALL"

Exit Sub
ERR_P:
    ShowError ("FillCombos::" & Me.Caption)
End Sub

Private Sub FillComboGrid()     '' Fills Employee Combo and Grid
On Error GoTo ERR_P
Dim intEmpCnt As Integer, intTmpCnt As Integer
Dim strArrEmp() As String
Dim strDeptTmp As String, strTempforCF As String

intEmpCnt = 0

If cboDept.Text = "ALL" Then
    strDeptTmp = "ALL"
Else
    strDeptTmp = cboDept.List(cboDept.ListIndex, 1)
    strDeptTmp = EncloseQuotes(strDeptTmp)
End If


Select Case UCase(Trim(strDeptTmp))
    Case "", "ALL"
        strTempforCF = "Select Empcode,Name from empmst where (joindate is not null and joindate<=" & strDTEnc & _
        Format(Now, "dd/MMM/yyyy") & strDTEnc & ") Order by EmpCode"
       'strTempforCF = "select Empcode,name from empmst where location=" & (cboLoc.Text) & " and company=" & (cboComp.Text) & " order by Empcode"
    Case Else
        If strCurrentUserType = HOD Then

            strTempforCF = "Select Empcode,Name from empmst " & strCurrData & " and (joindate is not null and joindate<=" & strDTEnc & DateCompStr(Date) & _
              strDTEnc & ") and Empmst." & SELCRIT & " = " & strDeptTmp & " Order by EmpCode"
         
        Else
            'this If condition add by  for datatype
'            If blnFlagForDept = True Then
                strTempforCF = "select Empcode,name from empmst where Empmst." & SELCRIT & "=" & _
                strDeptTmp & " order by Empcode"    'Empcode,name
'            End If
        End If
End Select

If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open strTempforCF, ConMain, adOpenStatic, adLockReadOnly
If (adrsEmp.EOF And adrsEmp.BOF) Then
    cboFrom.clear
    cboTo.clear
    MSF1.Rows = 1
    Exit Sub
End If
intEmpCnt = adrsEmp.RecordCount
intTmpCnt = intEmpCnt
MSF1.Rows = intEmpCnt + 1
ReDim strArrEmp(intTmpCnt - 1, 1)
For intEmpCnt = 0 To intTmpCnt - 1
    strArrEmp(intEmpCnt, 0) = adrsEmp(0)
    strArrEmp(intEmpCnt, 1) = adrsEmp(1)
    MSF1.TextMatrix(intEmpCnt + 1, 0) = adrsEmp(0)
    MSF1.TextMatrix(intEmpCnt + 1, 1) = adrsEmp(1)
    adrsEmp.MoveNext
Next
cboFrom.List = strArrEmp
cboTo.List = strArrEmp
cboFrom.ListIndex = 0
cboTo.ListIndex = cboTo.ListCount - 1
Erase strArrEmp
Exit Sub
ERR_P:
    ShowError ("Fill Combo Grid :: " & Me.Caption)
'    Resume Next
End Sub

Private Sub MSF1_Click()
If MSF1.Rows = 1 Then Exit Sub
If MSF1.CellBackColor = &HC0FFFF Then
    With MSF1
        .Col = 0
        .CellBackColor = vbWhite
        .Col = 1
        .CellBackColor = vbWhite
    End With
Else
    With MSF1
        .Col = 0
        .CellBackColor = &HC0FFFF
        .Col = 1
        .CellBackColor = &HC0FFFF
    End With
End If
End Sub

Private Function CheckDates() As Boolean    '' Function to Check if Dates are in Valid Range
On Error GoTo ERR_P
Dim strDateM
CheckDates = True                           '' or Not
If Trim(txtFrom.Text) = "" Then
    MsgBox NewCaptionTxt("00016", adrsMod), vbExclamation
    CheckDates = False
    txtFrom.SetFocus
    Exit Function
End If
If Trim(txtTo.Text) = "" Then
    MsgBox NewCaptionTxt("00017", adrsMod), vbExclamation
    CheckDates = False
    txtTo.SetFocus
    Exit Function
End If
'' Check for Software Lockdate
strDateM = DateCompDate(txtTo.Text)
strDateM = Right(Year(strDateM), 2) & Format(Month(strDateM), "00") & Format(Day(strDateM), "00")
If CLng(InVar.strLok) <= CLng(strDateM) Then
    MsgBox " Error : TimeHR-0007221 : Contact IV SOFTTECH. ", vbInformation, "PELK-0001221 "

        CheckDates = False
    Exit Function
End If
' End
If DateCompDate(txtTo.Text) < DateCompDate(txtFrom.Text) Then
    MsgBox NewCaptionTxt("00018", adrsMod), vbExclamation, App.EXEName
    CheckDates = False
    txtTo.SetFocus
    Exit Function
End If
'If DateDiff("m", Year_Start, DateCompDate(txtFrom.Text)) > 11 Or DateDiff("m", Year_Start, DateCompDate(txtFrom.Text)) < 0 Then
'        MsgBox NewCaptionTxt("00019", adrsMod) & (txtFrom.Text) & NewCaptionTxt("00021", adrsMod), vbExclamation
'        CheckDates = False
'        txtFrom.SetFocus
'        Exit Function
'End If
'If DateDiff("m", Year_Start, DateCompDate(txtTo.Text)) > 11 Or DateDiff("m", Year_Start, DateCompDate(txtTo.Text)) < 0 Then
'        MsgBox NewCaptionTxt("00020", adrsMod) & (txtTo.Text) & NewCaptionTxt("00021", adrsMod), vbExclamation
'        CheckDates = False
'        txtTo.SetFocus
'        Exit Function
'End If
Exit Function
ERR_P:
    ShowError ("Check Dates :: " & Me.Caption)
End Function

Private Function CheckEmployee() As Boolean     '' Function to Check if Employees are

strSelEmp = ""                                  '' Selected or not

CheckEmployee = True
    If MSF1.Rows = 1 Then
        MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation, App.EXEName
        CheckEmployee = False
        cmdSR.SetFocus
        Exit Function
    End If
    Call TruncateTable("ECode") ' 09-03
    MSF1.Col = 0
    STRECODE = ""
    For i = 1 To MSF1.Rows - 1
        MSF1.Row = i
        If MSF1.CellBackColor = SELECTED_COLOR Then
            'strSelEmp = strSelEmp & "'" & MSF1.Text & "',"
            STRECODE = STRECODE & "'" & MSF1.Text & "',"
            ConMain.Execute "insert into ECode values('" & MSF1.Text & "')"   ' 09-03
        End If
    Next
    strSelEmp = "select empcode from ECode "
    If STRECODE = "" Then
        MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation
        CheckEmployee = False
        cmdSR.SetFocus
    Else
        STRECODE = Left(STRECODE, Len(STRECODE) - 1)
    End If
    If strSelEmp = "" Then
        MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation
        CheckEmployee = False
        cmdSR.SetFocus
    Else
        strSelEmp = Left(strSelEmp, Len(strSelEmp) - 1)
        strSelEmp = "(" & strSelEmp & ")"
    End If
End Function

Private Function CheckShifts() As Boolean   '' Checks the Necessary Shift Files needed
On Error GoTo ERR_P                         '' For Processing
Dim bytM As Byte, bytTmpCnt As Byte
Dim strArrShf() As String
Dim strTabName As String
If Year(DateCompDate(txtTo.Text)) <> Year(DateCompDate(txtFrom.Text)) Then
    bytM = (Month(DateCompDate(txtTo.Text)) + 12) - Month(DateCompDate(txtFrom.Text))
Else
    bytM = Month(DateCompDate(txtTo.Text)) - Month(DateCompDate(txtFrom.Text))
End If
ReDim strArrShf(bytM)
strArrShf(0) = Left(MonthName(Month(DateCompDate(txtFrom.Text))), 3) & _
Right(Year(DateCompDate(txtFrom.Text)), 2) & "Shf"
CheckShifts = True
For bytM = 1 To UBound(strArrShf)
    strArrShf(bytM) = CStr(MonthNumber(strArrShf(bytM - 1)))
    If CInt(strArrShf(bytM)) = 12 Then strArrShf(bytM) = "0"
    strArrShf(bytM) = Left(MonthName(CInt(strArrShf(bytM)) + 1), 3)
    If UCase(strArrShf(bytM)) = "JAN" Then
        strArrShf(bytM) = strArrShf(bytM) & _
        Format((CInt(Mid(strArrShf(bytM - 1), 4, 2)) + 1), "00") & "shf"
    Else
        strArrShf(bytM) = strArrShf(bytM) & _
        Format((CInt(Mid(strArrShf(bytM - 1), 4, 2))), "00") & "shf"
    End If
Next
bytTmpCnt = 0
For bytM = 0 To UBound(strArrShf)
    If Not FindTable(strArrShf(bytM)) Then
        If Not CreateS(strArrShf(bytM)) Then bytTmpCnt = bytTmpCnt + 1
    End If
Next
If bytTmpCnt > 0 Then
    MsgBox NewCaptionTxt("17016", adrsC), vbExclamation
    CheckShifts = False
End If
Exit Function
ERR_P:
    ShowError ("Check Shifts ::")
    CheckShifts = False
End Function

Private Function CreateS(ByVal strShiftTab As String) As Boolean    '' Creates Shift File
On Error GoTo ERR_P
If MsgBox(NewCaptionTxt("17017", adrsC) & ForMatFull(Left(strShiftTab, 3)) & " " & _
IIf(Month(DateCompDate(txtFrom.Text)) < Val(pVStar.Yearstart), Val(pVStar.YearSel + 1), _
Val(pVStar.YearSel)) & NewCaptionTxt("17018", adrsC) & vbCrLf & _
NewCaptionTxt("17019", adrsC), vbYesNo + vbQuestion) = vbYes Then
ShiftFileCreate:
    Call SaveSetting(App.EXEName, "PrjSettings", "ShiftCreated", 0)
    bytShfMode = 4
    '' Assign Month Name
    '' Assign Dept Type
    bytLstInd = cboDept.ListIndex
    strRotPass = ForMatFull(Left(strShiftTab, 3)) & IIf(Month(DateCompDate(txtFrom.Text)) < _
    Val(pVStar.Yearstart), Val(pVStar.YearSel + 1), pVStar.YearSel)
    frmShiftCr.Show vbModal
    If GetSetting(App.EXEName, "PrjSettings", "ShiftCreated", 0) = 1 Then
        CreateS = True
    Else
        CreateS = False
    End If
Else
    CreateS = False
End If
Exit Function
ERR_P:
    ShowError ("CreateS :: " & Me.Caption)
End Function

Private Function ForMatFull(ByVal strMMMM As String) As String  '' Converts MMM Month Name to
Select Case UCase(strMMMM)                                      '' Full Name Format
    Case "JAN"
        ForMatFull = "January"
    Case "FEB"
        ForMatFull = "February"
    Case "MAR"
        ForMatFull = "March"
    Case "APR"
        ForMatFull = "April"
    Case "MAY"
        ForMatFull = "May"
    Case "JUN"
        ForMatFull = "June"
    Case "JUL"
        ForMatFull = "July"
    Case "AUG"
        ForMatFull = "August"
    Case "SEP"
        ForMatFull = "September"
    Case "OCT"
        ForMatFull = "October"
    Case "NOV"
        ForMatFull = "November"
    Case "DEC"
        ForMatFull = "December"
End Select
End Function

Private Sub MSF1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then Call MSF1_Click
End Sub


Private Sub txtFrom_Click()
varCalDt = ""
varCalDt = Trim(txtFrom.Text)
txtFrom.Text = ""
Call ShowCalendar
End Sub

Private Sub txtFrom_GotFocus()
    Call GF(txtFrom)
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    Call CDK(txtFrom, KeyAscii)
End Sub

Private Sub txtFrom_Validate(Cancel As Boolean)
If Not ValidDate(txtFrom) Then txtFrom.SetFocus: Cancel = True
End Sub

Private Sub txtTo_Click()
varCalDt = ""
varCalDt = Trim(txtTo.Text)
txtTo.Text = ""
Call ShowCalendar
End Sub

Private Sub txtTo_GotFocus()
    Call GF(txtTo)
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    Call CDK(txtTo, KeyAscii)
End Sub

Private Sub ChangeLabelCaption(Optional bytCapFlag As Byte = 1)
Select Case bytCapFlag
    Case 1
        '' Retreiving Records from the Dat File ...
        lblPro.Caption = NewCaptionTxt("17007", adrsC, 0)
    Case 2
        '' Processing Records :: Please Wait...
        lblPro.Caption = NewCaptionTxt("17008", adrsC, 0)
End Select
DoEvents
Me.Refresh
End Sub

Private Sub txtTo_Validate(Cancel As Boolean)
    If Not ValidDate(txtTo) Then txtTo.SetFocus: Cancel = True
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '17%'", ConMain, adOpenStatic
frDates.Caption = NewCaptionTxt("17002", adrsC)         '' Processing Dates
Call SetCritLabel(lblDeptCap)
Call CapGrid
Call SetGridDetails(Me, frEmp, MSF1, lblFrom, lblTo)
End Sub

Private Sub CapGrid()
'' Sizing
MSF1.ColWidth(1) = MSF1.ColWidth(1) * 2.65
MSF2.ColWidth(1) = MSF2.ColWidth(1) * 1.5
'' Aligning
MSF1.ColAlignment(0) = flexAlignLeftTop
MSF2.ColAlignment(0) = flexAlignCenterTop
MSF2.ColAlignment(1) = flexAlignCenterTop
'' Naming
MSF2.TextMatrix(0, 0) = NewCaptionTxt("00030", adrsMod)   '' Date
MSF2.TextMatrix(0, 1) = NewCaptionTxt("00047", adrsMod)   '' Code
End Sub

Private Sub GetRights()
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(4, 6, 3, 1)
If strTmp = "1" Then
    cmdPro.Enabled = True
Else
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    cmdPro.Enabled = False
End If
Exit Sub
ERR_P:
    ShowError ("GetRights :: " & Me.Caption)
    cmdPro.Enabled = False
End Sub


Public Sub SetCritCombos1(ByRef cboDept As Object)         '
'' On Error Resume Next
Select Case UCase(SELCRIT1)
    Case "DEPT"
        Call ComboFill(cboDept, 2, 2)
    Case "CAT"
        Call ComboFill(cboDept, 3, 2)
    Case "COMPANY"
        Call ComboFill(cboDept, 5, 2)
    Case "" & strKGroup & ""
        Call ComboFill(cboDept, 8, 2)
    Case "LOCATION"
        Call ComboFill(cboDept, 11, 2)
End Select
If strCurrentUserType <> HOD Then cboDept.AddItem "ALL"
End Sub
