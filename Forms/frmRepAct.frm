VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRepAct 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activity Log Reports"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   60
      Top             =   390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frGroup 
      Caption         =   "Grouping"
      Height          =   525
      Left            =   0
      TabIndex        =   18
      Top             =   1620
      Width           =   7635
      Begin VB.CheckBox chkNew 
         Caption         =   "Start New Page &When Group Changes"
         Height          =   225
         Left            =   3330
         TabIndex        =   22
         Top             =   210
         Width           =   3705
      End
      Begin VB.OptionButton optUser 
         Caption         =   "User"
         Height          =   195
         Left            =   2310
         TabIndex        =   21
         Top             =   210
         Width           =   1515
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Date "
         Height          =   195
         Left            =   1290
         TabIndex        =   20
         Top             =   210
         Width           =   1635
      End
      Begin VB.OptionButton optNoGroup 
         Caption         =   "Non&e"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   210
         Width           =   1455
      End
   End
   Begin VB.Frame frAct 
      Height          =   615
      Left            =   0
      TabIndex        =   23
      Top             =   2160
      Width           =   7665
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6450
         TabIndex        =   28
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "&File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   27
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4110
         TabIndex        =   26
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Pre&view"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2940
         TabIndex        =   25
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton cmdSelPri 
         Caption         =   "Selec&t Printer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   24
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.Frame frTime 
      Caption         =   "Time"
      Height          =   1155
      Left            =   4650
      TabIndex        =   11
      Top             =   450
      Width           =   3015
      Begin VB.TextBox txtToT 
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   810
         Width           =   630
      End
      Begin VB.TextBox txtFromT 
         Height          =   255
         Left            =   1140
         TabIndex        =   15
         Top             =   810
         Width           =   630
      End
      Begin VB.OptionButton optFromToT 
         Caption         =   "Between"
         Height          =   225
         Left            =   90
         TabIndex        =   13
         Top             =   540
         Width           =   1935
      End
      Begin VB.OptionButton optAnyT 
         Caption         =   "A&ny Time"
         Height          =   255
         Left            =   90
         TabIndex        =   12
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label lblToT 
         AutoSize        =   -1  'True
         Caption         =   "T&o"
         Height          =   195
         Left            =   1890
         TabIndex        =   16
         Top             =   810
         Width           =   195
      End
      Begin VB.Label lblFromT 
         AutoSize        =   -1  'True
         Caption         =   "Fro&m"
         Height          =   195
         Left            =   750
         TabIndex        =   14
         Top             =   840
         Width           =   345
      End
   End
   Begin VB.Frame frdate 
      Caption         =   "Date"
      Height          =   1155
      Left            =   0
      TabIndex        =   4
      Top             =   450
      Width           =   4635
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   3150
         TabIndex        =   10
         Tag             =   "D"
         Top             =   780
         Width           =   1305
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   1500
         TabIndex        =   8
         Tag             =   "D"
         Top             =   780
         Width           =   1305
      End
      Begin VB.OptionButton optFromTo 
         Caption         =   "Between"
         Height          =   255
         Left            =   90
         TabIndex        =   6
         Top             =   540
         Width           =   1815
      End
      Begin VB.OptionButton optAny 
         Caption         =   "Any &Date"
         Height          =   225
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         Caption         =   "T&o"
         Height          =   195
         Left            =   2910
         TabIndex        =   9
         Top             =   810
         Width           =   195
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         Caption         =   "Fro&m"
         Height          =   195
         Left            =   900
         TabIndex        =   7
         Top             =   810
         Width           =   345
      End
   End
   Begin VB.ComboBox cboForms 
      Height          =   315
      Left            =   1028
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   30
      Width           =   3195
   End
   Begin MSForms.ComboBox cboUser 
      Height          =   315
      Left            =   5145
      TabIndex        =   3
      Top             =   30
      Width           =   2295
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4048;556"
      ListWidth       =   6000
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   0
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1500;4500"
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "&User"
      Height          =   195
      Left            =   4365
      TabIndex        =   2
      Top             =   60
      Width           =   330
   End
   Begin VB.Label lblAct 
      AutoSize        =   -1  'True
      Caption         =   "&Activity on"
      Height          =   195
      Left            =   248
      TabIndex        =   0
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "frmRepAct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strLogQuery As String
Dim objRec As Object

Private Sub chkNew_Click()
    SaveSetting "Vstar", "PrjSettings", "Print on Next Page for LOG", chkNew.Value
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFile_Click()
    Call ReportsModLog(3)
End Sub

Private Sub cmdPreview_Click()
    Call ReportsModLog(1)
End Sub

Private Sub cmdPrint_Click()
    Call ReportsModLog(2)
End Sub

Private Sub cmdSelPri_Click()
On Error GoTo ERR_P
CommonDialog1.PrinterDefault = True
CommonDialog1.Flags = cdlSetNotSupported
CommonDialog1.ShowPrinter
Printer.TrackDefault = True
Exit Sub
ERR_P:
    ShowError ("Select Printer :: " & Me.Caption)
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Call SetFormIcon(Me)    '' Sets the Forms Icon
Call SetToolTipText(Me) '' Sets the ToolTipText for Date Boxes
'' Call GetRights       '' Gets and Sets the Rights
Call LoadSpecifics      '' Load Specific Procedure
Exit Sub
ERR_P:
    ShowError ("Load :: " & Me.Caption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If DELOG.cnLog.State = 0 Then DELOG.cnLog.Open
End Sub

Private Sub optAny_Click()
    Call AdjustControls(0)      '' Adjust Controls of Date
End Sub

Private Sub optAnyT_Click()
    Call AdjustControlsT(0)      '' Adjust Controls of Date
End Sub

Private Sub optFromTo_Click()
    Call AdjustControls(1)      '' Adjust Controls of Date
    If bytMode <> 1 Then txtFrom.SetFocus
End Sub

Private Sub optFromToT_Click()
    Call AdjustControlsT(1)      '' Adjust Controls of Date
    If bytMode <> 1 Then txtFromT.SetFocus
End Sub

Private Sub txtFrom_Click()
varCalDt = ""
varCalDt = Trim(txtFrom.Text)
txtFrom.Text = ""
Load CalendarFrm
CalendarFrm.Show 1
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

Private Sub txtFromT_GotFocus()
    Call GF(txtFromT)
End Sub

Private Sub txtFromT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 4)
End If
End Sub

Private Sub txtToT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 4)
End If
End Sub

Private Sub txtToT_GotFocus()
    Call GF(txtToT)
End Sub

Private Sub txtTo_Click()
varCalDt = ""
varCalDt = Trim(txtTo.Text)
txtTo.Text = ""
Load CalendarFrm
CalendarFrm.Show 1
End Sub

Private Sub txtTo_GotFocus()
    Call GF(txtTo)
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
     Call CDK(txtTo, KeyAscii)
End Sub

Private Sub txtTo_Validate(Cancel As Boolean)
    If Not ValidDate(txtTo) Then txtTo.SetFocus: Cancel = True
End Sub

Private Sub LoadSpecifics()             '' Procedure to perform Load Specific Actions
Call FillFormCombo                      '' Fill Forms Combo
Call FillUserCombo                      '' Fill User Combo
bytMode = 1                             '' Load Mode
txtFrom.Text = DateDisp(CStr(Date))     '' Set Current Date
txtTo.Text = DateDisp(CStr(Date))       '' Set Current Date
optAny.Value = True                     '' Set to Any Date
optAnyT.Value = True                    '' Set to Any Time
optNoGroup.Value = True                 '' Set to No Grouping
txtFromT.Text = "0.00"                  '' Set Value to 0.00
txtToT.Text = "0.00"                    '' Set Value to 0.00
chkNew.Value = GetSetting("Vstar", "PrjSettings", "Print on Next Page for LOG", 1)
bytMode = 2                             '' No Mode
End Sub

Private Sub AdjustControls(Optional bytFlg As Byte = 1)
Select Case bytFlg
    Case 0
        lblFrom.Enabled = False
        lblTo.Enabled = False
        txtFrom.Enabled = False
        txtTo.Enabled = False
    Case 1
        lblFrom.Enabled = True
        lblTo.Enabled = True
        txtFrom.Enabled = True
        txtTo.Enabled = True
End Select
End Sub

Private Sub AdjustControlsT(Optional bytFlg As Byte = 1)
Select Case bytFlg
    Case 0
        lblFromT.Enabled = False
        lblToT.Enabled = False
        txtFromT.Enabled = False
        txtToT.Enabled = False
    Case 1
        lblFromT.Enabled = True
        lblToT.Enabled = True
        txtFromT.Enabled = True
        txtToT.Enabled = True
End Select
End Sub

Private Sub FillFormCombo()     '' Fills the Activity On Combo
On Error GoTo ERR_P
cboForms.AddItem "All Forms"
If adrsLog.State = 1 Then adrsLog.Close
adrsLog.Open "Select TransDesc From Transdet"
Do While Not adrsLog.EOF
    cboForms.AddItem adrsLog("TransDesc")
    adrsLog.MoveNext
Loop
If cboForms.ListCount > 0 Then cboForms.ListIndex = 0
Exit Sub
ERR_P:
    ShowError ("FillFormCombo :: " & Me.Caption)
End Sub

Private Sub FillUserCombo()     '' Fills the User Combo
On Error GoTo ERR_P
Dim bytTmp As Byte, strTmp() As String
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "Select UserCode,UserName from UserInfo Order by UserCode,UserName" _
, VstarDataEnv.cnDJConn, adOpenStatic
If (adrsTemp.EOF And adrsTemp.BOF) Then
    ReDim strTmp(0, 1)
    strTmp(0, 0) = "All Users"
    strTmp(0, 1) = ""
Else
    ReDim strTmp(adrsTemp.RecordCount, 1)
    strTmp(0, 0) = "All Users"
    strTmp(0, 1) = ""
    For bytTmp = 1 To adrsTemp.RecordCount
        strTmp(bytTmp, 0) = adrsTemp("UserCode")    '' User Code
        strTmp(bytTmp, 1) = adrsTemp("userName")    '' user Name
        adrsTemp.MoveNext
    Next
End If
cboUser.List = strTmp
If cboUser.ListCount > 0 Then cboUser.ListIndex = 0
Exit Sub
ERR_P:
    ShowError ("FillUserCombo :: " & Me.Caption)
End Sub

Private Sub MakeQueryStringLog()        '' Make Query String
On Error GoTo ERR_P
Dim strForms As String, strUserTmp As String, strDatesTmp As String, strTimeTmp As String
'' Forms
If cboForms.ListIndex > 0 Then
    strForms = " and activity.TransSource=" & cboForms.ListIndex
Else
    strForms = ""
End If
'' User
If cboUser.ListIndex > 0 Then
    strUserTmp = " and UserName='" & cboUser.List(cboUser.ListIndex, 1) & "'"
Else
    strUserTmp = ""
End If
'' Dates
If optAny.Value = True Then
    strDatesTmp = ""
Else
    strDatesTmp = " and (Logdate between #" & DateCompStr(txtFrom.Text) & "# and #" & _
    DateCompStr(txtTo.Text) & "#)"
End If
'' Time
If optAnyT.Value = True Then
    strTimeTmp = ""
Else
    strTimeTmp = " and (LogTime between " & txtFromT.Text & " and " & _
    txtToT.Text & ")"
End If
'' Group
If optNoGroup.Value = True Then     '' No Grouping
    sqlStr = "RecordNumber"
End If
If optDate.Value = True Then        '' On Date
    sqlStr = "logDate"
End If
If optUser.Value = True Then        '' On User
    sqlStr = "UserName"
End If
strLogQuery = "SHAPE { SELECT distinct *,transdesc from Activity,transdet where " & _
"activity.transSource=TransDet.TransSource " & _
strForms & strUserTmp & strDatesTmp & strTimeTmp & " order by username,logDate} AS cmdAct " & _
"COMPUTE cmdAct by '" & sqlStr & "'"
Exit Sub
ERR_P:
End Sub

Private Sub ReportsModLog(ByVal bytLogAction As Byte)
On Error GoTo ERR_P
If Not ValidEntries Then Exit Sub
Call MakeQueryStringLog
If Not RecordsFoundLog Then
    Exit Sub
End If
rptLog.Refresh
Unload DELOG
Select Case bytLogAction
    Case 1 '' Preview
        rptLog.Show vbModal
    Case 2 '' Print
        rptLog.PrintReport
    Case 3 '' File
        CommonDialog1.Filter = "*.txt|*.txt"
        CommonDialog1.Flags = cdlOFNOverwritePrompt
        CommonDialog1.ShowSave
        rptLog.ExportReport rptKeyText, CommonDialog1.FileName
        MsgBox "File Saved Successfully", vbInformation, App.EXEName
End Select
LogIn:
If DELOG.rscmdAct_Grouping.State = 1 Then DELOG.rscmdAct_Grouping.Close
Exit Sub
ERR_P:
    Select Case Err.Number
        Case 8507
            GoTo LogIn
        Case Else
            CommonDialog1.FileName = ""
            ShowError ("LogReportsMod :: " & Me.Caption)
    End Select
End Sub

Private Function RecordsFoundLog() As Boolean
On Error GoTo ERR_P
RecordsFoundLog = True
Set objRec = Nothing
Set objRec = DELOG.rscmdAct_Grouping
If objRec.State = 1 Then objRec.Close
objRec.Open strLogQuery
If objRec.EOF And objRec.BOF Then
    Unload DELOG
    RecordsFoundLog = False
    MsgBox " No Records Found ", vbInformation, App.EXEName
End If
Exit Function
ERR_P:
    ShowError ("LogRecords :: " & Me.Caption)
    RecordsFoundLog = False
End Function

Private Function ValidEntries() As Boolean
On Error GoTo ERR_P
ValidEntries = False
'' Objects
If cboForms.ListIndex < 0 Then Exit Function
'' User
If cboUser.ListIndex < 0 Then Exit Function
'' Date
If optFromTo.Value = True Then
    If Trim(txtFrom.Text) = "" Then
        MsgBox "Please Enter From Date", vbExclamation, App.EXEName
        txtFrom.SetFocus
        Exit Function
    End If
    If Trim(txtTo.Text) = "" Then
        MsgBox "Please Enter To Date", vbExclamation, App.EXEName
        txtTo.SetFocus
        Exit Function
    End If
End If
'' Time
If optFromToT.Value = True Then
    txtFromT.Text = IIf(Trim(txtFromT.Text) = "", "0.00", Format(txtFromT.Text, "0.00"))
    txtToT.Text = IIf(Trim(txtToT.Text) = "", "0.00", Format(txtToT.Text, "0.00"))
End If
ValidEntries = True
Exit Function
ERR_P:
End Function
