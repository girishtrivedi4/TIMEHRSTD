VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLateCorr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Late Employees List"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frDates 
      Caption         =   "Date"
      Height          =   675
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   5805
      Begin VB.TextBox txtFrom 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2250
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "D"
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label lblFromD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&From Date"
         Height          =   195
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblToD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&To Date"
         Height          =   195
         Left            =   3180
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   585
      End
   End
   Begin VB.Frame frEmp 
      Caption         =   "Employee Selection"
      Height          =   4935
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   5805
      Begin VB.CommandButton cmdSR 
         Caption         =   "Select &Range"
         Height          =   435
         Left            =   3840
         TabIndex        =   4
         Top             =   1080
         Width           =   1755
      End
      Begin VB.CommandButton cmdUR 
         Caption         =   "&Unselect Range"
         Height          =   465
         Left            =   3840
         TabIndex        =   7
         Top             =   1500
         Width           =   1755
      End
      Begin VB.CommandButton cmdSA 
         Caption         =   "&Select All"
         Height          =   435
         Left            =   3870
         TabIndex        =   8
         Top             =   2100
         Width           =   1755
      End
      Begin VB.CommandButton cmdUA 
         Caption         =   "U&nselect All"
         Height          =   435
         Left            =   3870
         TabIndex        =   9
         Top             =   2520
         Width           =   1755
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   4005
         Left            =   30
         TabIndex        =   11
         ToolTipText     =   "Click to Toggle Selection"
         Top             =   900
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   7064
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fro&m"
         Height          =   195
         Left            =   540
         TabIndex        =   14
         Top             =   630
         Width           =   345
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T&o"
         Height          =   195
         Left            =   3000
         TabIndex        =   13
         Top             =   630
         Width           =   195
      End
      Begin MSForms.ComboBox cboFrom 
         Height          =   315
         Left            =   1410
         TabIndex        =   2
         Top             =   570
         Width           =   1365
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2408;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboTo 
         Height          =   315
         Left            =   3600
         TabIndex        =   3
         Top             =   570
         Width           =   1365
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2408;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboDept 
         Height          =   315
         Left            =   1410
         TabIndex        =   1
         Top             =   210
         Width           =   1395
         VariousPropertyBits=   612390939
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2461;556"
         TextColumn      =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblDeptCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdPro 
      Caption         =   "&Process"
      Enabled         =   0   'False
      Height          =   465
      Left            =   1320
      TabIndex        =   5
      Top             =   5640
      Width           =   1545
   End
   Begin VB.CommandButton cmdFin 
      Caption         =   "F&inish"
      Height          =   465
      Left            =   2880
      TabIndex        =   6
      Top             =   5640
      Width           =   1545
   End
End
Attribute VB_Name = "frmLateCorr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFName As String, strSelEmp As String
Dim adrsC As New ADODB.Recordset
Dim i As Integer

Private Sub cboFrom_Change()
    cboTo.ListIndex = cboTo.ListCount - 1
End Sub

Private Sub cboFrom_Click()
    If cboFrom.ListIndex < 0 Then Exit Sub
    If cboTo.ListIndex = 0 Then Exit Sub
    cboTo.ListIndex = cboFrom.ListIndex
End Sub

Private Sub cmdFin_Click()
    Unload Me
End Sub

Private Sub cmdPro_Click()
If Not CheckEmployee Then Exit Sub  '' If not valid Employees then Exit
Call UpdateTrn
MsgBox "Late Hrs Correction done successfully for the selected employees."
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
txtFrom.SetFocus
Call GetRights
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)        '' Set the Form Icon
txtFrom.Enabled = False
txtFrom.Text = DateDisp(CStr(Date))
Call SetToolTipText(Me)     '' Set the ToolTipText
Call RetCaptions            '' Set the Control Captions
'' Empty Grid
Call FillCombos
'' Set Current Dates.
txtFrom.Text = DateDisp(CStr(Date))
End Sub

Private Sub cboDept_Click()
On Error GoTo ERR_P
If cbodept.ListIndex < 0 Then Exit Sub               '' If No Department
Call FillComboGrid
Call SelUnselAll(vbWhite, MSF1)
Exit Sub
ERR_P:
    ShowError ("Department::" & Me.Caption)
End Sub

Private Sub FillCombos()
On Error GoTo ERR_P
Call SetCritCombos(cbodept)
cbodept.ListIndex = cbodept.ListCount - 1
Exit Sub
ERR_P:
    ShowError ("FillCombos::" & Me.Caption)
End Sub

Private Sub FillComboGrid()     '' Fills Employee Combo and Grid
10    On Error GoTo ERR_P
      Dim intEmpCnt As Integer, intTmpCnt As Integer
      Dim strArrEmp() As String
      Dim strDeptTmp As String, strTempforCF As String
20    intEmpCnt = 0
30    strDeptTmp = cbodept.List(cbodept.ListIndex, 0)
40    strDeptTmp = EncloseQuotes(strDeptTmp)
50    strFName = Left(MonthName(Month(txtFrom.Text)), 3) & Right(Year(txtFrom.Text), 2) & "trn"
60    If Not FindTable(strFName) Then Exit Sub
70    Select Case UCase(Trim(strDeptTmp))
          Case "", "ALL"
80            strTempforCF = "Select empmst.Empcode,Name from empmst," & strFName & " where latehrs > 0 and " & _
              strFName & "." & strKDate & " = " & strDTEnc & Format(DateCompStr(txtFrom.Text), "DD/MMM/YYYY") & _
              strDTEnc & " and empmst.empcode=" & strFName & ".empcode and (joindate is not null and joindate<=" & _
              strDTEnc & Format(DateCompStr(Date), "DD/MMM/YYYY") & strDTEnc & ") Order by empmst.EmpCode"
90        Case Else
100           If strCurrentUserType = HOD Then
110               strTempforCF = "Select empmst.Empcode,Name from empmst," & strFName & "" & strCurrData & _
                  " and " & strFName & "." & strKDate & " = " & strDTEnc & Format(DateCompStr(txtFrom.Text), "DD/MMM/YYYY") & _
                  strDTEnc & " and latehrs > 0 and empmst.empcode=" & strFName & _
                  ".empcode and (joindate is not null and joindate<=" & strDTEnc & Format(DateCompStr(Date), "DD/MMM/YYYY") & _
                    strDTEnc & ") and Empmst." & SELCRIT & " = " & strDeptTmp & " Order by empmst.EmpCode"
120           Else
130               strTempforCF = "Select empmst.Empcode,Name from empmst," & strFName & " where latehrs > 0 and empmst.empcode=" & strFName & ".empcode and " & strFName & "." & _
              strKDate & " = " & strDTEnc & Format(DateCompStr(txtFrom.Text), "DD/MMM/YYYY") & strDTEnc & " and (joindate is not null and joindate<=" & strDTEnc & _
                  Format(DateCompStr(Date), "DD/MMM/YYYY") & strDTEnc & ") and " & SELCRIT & " = " & strDeptTmp & "  Order by empmst.EmpCode"
140           End If
150   End Select
160   If adrsEmp.State = 1 Then adrsEmp.Close
170   adrsEmp.Open strTempforCF, VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
180   If (adrsEmp.EOF And adrsEmp.BOF) Then
190       cboFrom.Clear
200       cboTo.Clear
210       MSF1.Rows = 1
220       Exit Sub
230   End If
240   intEmpCnt = adrsEmp.RecordCount
250   intTmpCnt = intEmpCnt
260   MSF1.Rows = intEmpCnt + 1
270   ReDim strArrEmp(intTmpCnt - 1, 1)
280   For intEmpCnt = 0 To intTmpCnt - 1
290       strArrEmp(intEmpCnt, 0) = adrsEmp(0)
300       strArrEmp(intEmpCnt, 1) = adrsEmp(1)
310       MSF1.TextMatrix(intEmpCnt + 1, 0) = adrsEmp(0)
320       MSF1.TextMatrix(intEmpCnt + 1, 1) = adrsEmp(1)
330       adrsEmp.MoveNext
340   Next
360   cboTo.List = strArrEmp
350   cboFrom.List = strArrEmp
370   cboFrom.ListIndex = 0
380   cboTo.ListIndex = cboTo.ListCount - 1
390   Erase strArrEmp
400   Exit Sub
ERR_P:
410       ShowError ("FillComboGrid :: " & Me.Caption & vbCrLf & _
                "Erl:" & Erl)
End Sub

Private Sub MSF1_Click()
If MSF1.Rows = 1 Then Exit Sub
If MSF1.CellBackColor = &HC0FFFF Then
    With MSF1
        .col = 0
        .CellBackColor = vbWhite
        .col = 1
        .CellBackColor = vbWhite
    End With
Else
    With MSF1
        .col = 0
        .CellBackColor = &HC0FFFF
        .col = 1
        .CellBackColor = &HC0FFFF
    End With
End If
End Sub

Private Function CheckEmployee() As Boolean     '' Function to Check if Employees are
strSelEmp = ""                                  '' Selected or not
CheckEmployee = True
If Not FindTable("lvinfo" & Right(pVStar.YearSel, 2)) Then
    CheckEmployee = False
    Exit Function
End If
If MSF1.Rows = 1 Then
    MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation, App.EXEName
    CheckEmployee = False
    cmdSR.SetFocus
    Exit Function
End If
MSF1.col = 0
For i = 1 To MSF1.Rows - 1
    MSF1.row = i
    If MSF1.CellBackColor = SELECTED_COLOR Then
        strSelEmp = strSelEmp & "'" & MSF1.Text & "',"
    End If
Next
If strSelEmp = "" Then
    MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation
    CheckEmployee = False
    cmdSR.SetFocus
Else
    strSelEmp = Left(strSelEmp, Len(strSelEmp) - 1)
    strSelEmp = "(" & strSelEmp & ")"
End If
End Function

Private Sub MSF1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then Call MSF1_Click
End Sub

Private Sub txtFrom_Click()
    varCalDt = ""
    varCalDt = Trim(txtFrom.Text)
    txtFrom.Text = ""
    Call ShowCalendar
    If DateDiff("m", Year_Start, DateCompDate(txtFrom.Text)) > 11 Or _
    DateDiff("m", Year_Start, DateCompDate(txtFrom.Text)) < 0 Then
        MsgBox NewCaptionTxt("00019", adrsMod) & txtFrom.Text & NewCaptionTxt("00021", adrsMod), _
        vbExclamation
        txtFrom.SetFocus
    Else
           Call FillCombos
    End If
End Sub

Private Sub txtFrom_GotFocus()
    Call GF(txtFrom)
    varCalDt = ""
    varCalDt = Trim(txtFrom.Text)
    If DateDiff("m", Year_Start, DateCompDate(txtFrom.Text)) > 11 Or _
    DateDiff("m", Year_Start, DateCompDate(txtFrom.Text)) < 0 Then
        MsgBox NewCaptionTxt("00019", adrsMod) & txtFrom.Text & NewCaptionTxt("00021", adrsMod), _
        vbExclamation
        txtFrom.SetFocus
    Else
           Call FillCombos
    End If
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    Call CDK(txtFrom, KeyAscii)
End Sub

Private Sub txtFrom_Validate(Cancel As Boolean)
    If Not ValidDate(txtFrom) Then
        txtFrom.SetFocus: Cancel = True
    Else
        varCalDt = ""
        varCalDt = Trim(txtFrom.Text)
        If DateDiff("m", Year_Start, DateCompDate(txtFrom.Text)) > 11 Or _
        DateDiff("m", Year_Start, DateCompDate(txtFrom.Text)) < 0 Then
            MsgBox NewCaptionTxt("00019", adrsMod) & txtFrom.Text & NewCaptionTxt("00021", adrsMod), _
            vbExclamation
            txtFrom.SetFocus
        Else
               Call FillCombos
        End If
    End If
End Sub

Private Sub RetCaptions()
    If adrsC.State = 1 Then adrsC.Close
    adrsC.Open "Select * From NewCaptions Where ID Like '17%'", VstarDataEnv.cnDJConn, adOpenStatic
    frDates.Caption = NewCaptionTxt("17002", adrsC)
    Call SetCritLabel(lblDeptCap)
    Call SetGridDetails(Me, frEmp, MSF1, lblFrom, lblTo)
    lblFromD.Caption = NewCaptionTxt("17003", adrsC)        '' From Date
    Call CapGrid
End Sub

Private Sub CapGrid()
    MSF1.ColWidth(1) = MSF1.ColWidth(1) * 2.65
    MSF1.ColAlignment(0) = flexAlignLeftTop
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

Public Sub UpdateTrn()
    VstarDataEnv.cnDJConn.Execute "Update " & strFName & " set latehrs=0 where empcode in " & strSelEmp & " and " & strFName & "." & _
            strKDate & " = " & strDTEnc & DateCompStr(txtFrom.Text) & strDTEnc & " "
    Call FillCombos
End Sub
