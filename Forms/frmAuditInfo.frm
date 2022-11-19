VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAuditInfo 
   Caption         =   "Audit Information"
   ClientHeight    =   1305
   ClientLeft      =   3450
   ClientTop       =   2535
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1305
   ScaleWidth      =   4875
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraAudit 
      Height          =   1410
      Left            =   0
      TabIndex        =   6
      Top             =   -90
      Width           =   4920
      Begin MSWinsockLib.Winsock winip 
         Left            =   3240
         Top             =   1440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3780
         TabIndex        =   5
         Top             =   855
         Width           =   1050
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "Select All User"
         Height          =   240
         Left            =   90
         TabIndex        =   2
         Top             =   675
         Width           =   1410
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3780
         TabIndex        =   4
         Top             =   225
         Width           =   1050
      End
      Begin VB.TextBox txtToDt 
         Height          =   330
         Left            =   2250
         TabIndex        =   1
         Top             =   180
         Width           =   1185
      End
      Begin VB.TextBox txtFromDt 
         Height          =   330
         Left            =   585
         TabIndex        =   0
         Top             =   180
         Width           =   1185
      End
      Begin VB.ComboBox cboLoginId 
         Height          =   315
         Left            =   1170
         TabIndex        =   3
         Top             =   945
         Width           =   1590
      End
      Begin VB.Label lblToDt 
         AutoSize        =   -1  'True
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1935
         TabIndex        =   9
         Top             =   270
         Width           =   300
      End
      Begin VB.Label lblFromDt 
         AutoSize        =   -1  'True
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblId 
         AutoSize        =   -1  'True
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   990
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmAuditInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAudit As New ADODB.Recordset

Private Sub chkAll_Click()
If chkAll.Value = 1 Then
    cboLoginId.Visible = False
    lblId.Visible = False
Else
    cboLoginId.Visible = True
    lblId.Visible = True
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdShow_Click()
On Error GoTo ERR_P
If Not CheckDate Then Exit Sub
Me.MousePointer = vbHourglass
typRep.strPeriFr = txtFromDt.Text
typRep.strPeriTo = txtToDt.Text
If chkAll.Value = 1 Then
    empstr3 = "Select * from AuditLog where AccessDate>=" & strDTEnc & Format(txtFromDt.Text, "DD/MMM/YYYY") & strDTEnc & " and AccessDate<=" & strDTEnc & Format(txtToDt.Text, "DD/MMM/YYYY") & strDTEnc & " order by LoginId,AccessDate,AccessTime"
Else
    empstr3 = "Select * from AuditLog where AccessDate>=" & strDTEnc & Format(txtFromDt.Text, "DD/MMM/YYYY") & strDTEnc & " and AccessDate<=" & strDTEnc & Format(txtToDt.Text, "DD/MMM/YYYY") & strDTEnc & " and LoginId='" & cboLoginId.Text & "' order by LoginId,AccessDate,AccessTime"
End If
strCName = InVar.strCOM

Set Report = crxApp.OpenReport(App.path & "\Reports\AuditInfo.rpt", 1)

Report.FormulaFields.GetItemByName("Header").Text = "'" & "Audit Report For The Period Of  " & DateDisp(typRep.strPeriFr) & " To " & DateDisp(typRep.strPeriTo) & "'"
Report.FormulaFields.GetItemByName("Cname").Text = "'" & strCName & "'"
frmCRV.Caption = "Audit Report For The Period Of  " & DateDisp(typRep.strPeriFr) & " To " & DateDisp(typRep.strPeriTo)
Report.DiscardSavedData
If adrsCrep.State = 1 Then adrsCrep.Close
adrsCrep.Open empstr3, ConMain, adOpenStatic, adLockOptimistic
If Not (adrsCrep.BOF And adrsCrep.EOF) = False Then
    If adrsCrep.State = 1 Then adrsCrep.Close
    blnIntz = False
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
    Me.MousePointer = vbNormal
    Exit Sub
Else
    blnIntz = True
    Report.Database.SetDataSource adrsCrep
End If
If blnIntz Then Set CRV = frmCRV.CRV: 'Call PrnAtul:
CRV.ReportSource = Report: bytPrint = 2: Call SetFormIcon(frmCRV)
If Not RecordsFound Then Exit Sub
Exit Sub
ERR_P:
    ShowError ("Show AuditInfo :: " & Me.Caption)
    'Resume Next
End Sub
Private Function RecordsFound() As Boolean
On Error GoTo ERR_P
RecordsFound = True
If bytPrint = 2 Then
    If blnIntz = True Then
        CRV.ViewReport
        Do While CRV.IsBusy
            DoEvents
        Loop
        frmCRV.Show vbModal
    Else
        RecordsFound = False
        Exit Function
    End If
End If
Me.MousePointer = vbNormal
Exit Function
ERR_P:
    ShowError ("Records Found :: " & Me.Caption)
    Set Report = Nothing
    RecordsFound = False
End Function
Private Sub Form_Load()
txtFromDt.Text = DateDisp(Date)
txtToDt.Text = DateDisp(Date)
Call SetFormIcon(Me)            '' Set the Forms Icon
'Call ReportCon
Call FillCombo
End Sub

Private Sub FillCombo()
On Error GoTo ERR_P
If rsAudit.State = 1 Then rsAudit.Close
rsAudit.Open "Select UserName from UserAccs", ConMain
If Not (rsAudit.EOF And rsAudit.BOF) Then
    cboLoginId.Text = rsAudit(0)
    Do While Not rsAudit.EOF
        cboLoginId.AddItem rsAudit(0)
        rsAudit.MoveNext
    Loop
End If
Exit Sub
ERR_P:
    ShowError ("AuditInfo :: " & Me.Caption)
End Sub

Private Sub txtFromDt_Click()
varCalDt = ""
varCalDt = Trim(txtFromDt.Text)
txtFromDt.Text = ""
Call ShowCalendar
End Sub

Private Sub txtFromDt_GotFocus()
    Call GF(txtFromDt)
End Sub

Private Sub txtFromDt_KeyPress(KeyAscii As Integer)
    Call CDK(txtFromDt, KeyAscii)
End Sub
Private Sub txtFromDt_Validate(Cancel As Boolean)
If Not ValidDate(txtFromDt) Then
    txtFromDt.SetFocus
    Cancel = True
End If
End Sub

Private Sub txtToDt_Click()
varCalDt = ""
varCalDt = Trim(txtToDt.Text)
txtToDt.Text = ""
Call ShowCalendar
End Sub

Private Sub txtToDt_GotFocus()
    Call GF(txtToDt)
End Sub

Private Sub txtToDt_KeyPress(KeyAscii As Integer)
    Call CDK(txtToDt, KeyAscii)
End Sub

Private Sub txtToDt_Validate(Cancel As Boolean)
If Not ValidDate(txtToDt) Then
txtToDt.SetFocus
Cancel = True
End If
End Sub

Private Function CheckDate() As Boolean
On Error GoTo ERR_P
CheckDate = True                              '' FUNCTION FOR PERIODIC REPORT VALIDATIONS
If txtFromDt.Text = "" Then
    MsgBox NewCaptionTxt("00016", adrsMod), vbInformation
    txtFromDt.SetFocus
    CheckDate = False
    Exit Function
End If
If txtToDt.Text = "" Then
    MsgBox NewCaptionTxt("00017", adrsMod), vbInformation
    txtToDt.SetFocus
    CheckDate = False
    Exit Function
End If
If CDate(txtFromDt.Text) > CDate(txtToDt.Text) Then
    MsgBox NewCaptionTxt("00018", adrsMod), vbInformation
    txtFromDt.SetFocus
    CheckDate = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("CheckDate ::" & Me.Caption)
    CheckDate = False
    Resume Next
End Function
