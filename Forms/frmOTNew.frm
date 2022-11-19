VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmOTNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OT Authorization"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frEmp 
      Height          =   1725
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton cmdSR 
         Caption         =   "Select &Range"
         Height          =   435
         Left            =   3930
         TabIndex        =   18
         Top             =   630
         Width           =   1755
      End
      Begin VB.CommandButton cmdUR 
         Caption         =   "&Unselect Range"
         Height          =   465
         Left            =   3930
         TabIndex        =   17
         Top             =   1080
         Width           =   1755
      End
      Begin VB.CommandButton cmdSA 
         Caption         =   "&Select All"
         Height          =   435
         Left            =   6210
         TabIndex        =   16
         Top             =   630
         Width           =   1755
      End
      Begin VB.CommandButton cmdUA 
         Caption         =   "U&nselect All"
         Height          =   435
         Left            =   6210
         TabIndex        =   15
         Top             =   1080
         Width           =   1755
      End
      Begin VB.CommandButton cmdGet 
         Caption         =   "&Get"
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
         Left            =   3120
         TabIndex        =   14
         Top             =   160
         Width           =   495
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
         Height          =   330
         Left            =   1410
         TabIndex        =   12
         Tag             =   "D"
         Top             =   160
         Width           =   1400
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fro&m"
         Height          =   195
         Left            =   600
         TabIndex        =   24
         Top             =   1020
         Width           =   345
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T&o"
         Height          =   195
         Left            =   780
         TabIndex        =   23
         Top             =   1380
         Width           =   195
      End
      Begin MSForms.ComboBox cboFrom 
         Height          =   315
         Left            =   1410
         TabIndex        =   22
         Top             =   960
         Width           =   1395
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2461;556"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboTo 
         Height          =   315
         Left            =   1410
         TabIndex        =   21
         Top             =   1320
         Width           =   1395
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2461;556"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboDept 
         Height          =   315
         Left            =   1410
         TabIndex        =   20
         Top             =   600
         Width           =   1395
         VariousPropertyBits=   612390939
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2461;556"
         TextColumn      =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblDeptCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   660
         Width           =   825
      End
      Begin VB.Label lblFromD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&For Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   13
         Top             =   160
         Width           =   705
      End
   End
   Begin VB.Frame frOT 
      Caption         =   "OT Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   0
      TabIndex        =   2
      Top             =   6240
      Visible         =   0   'False
      Width           =   6945
      Begin VB.CheckBox chkOT 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   4
         Top             =   270
         Width           =   1875
      End
      Begin VB.TextBox txtOT 
         Height          =   315
         Left            =   5580
         TabIndex        =   6
         Top             =   240
         Width           =   765
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Default         =   -1  'True
         Height          =   405
         Left            =   4440
         TabIndex        =   9
         Top             =   660
         Width           =   1155
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Command1"
         Height          =   405
         Left            =   5670
         TabIndex        =   10
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox txtRem 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1530
         MaxLength       =   15
         TabIndex        =   8
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label lblOT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4470
         TabIndex        =   5
         Top             =   300
         Width           =   615
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblOTRem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OT Remark"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   690
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   435
      Left            =   7440
      TabIndex        =   11
      Top             =   6750
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   1740
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   12
      Cols            =   8
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   4194368
      ForeColorFixed  =   8454143
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOTNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset
''
Private strFileName As String * 9       '' String File Name
Private blnFileFound As Boolean         '' Boolean for Invalid Transaction
Private sngCalcOT As Single             '' For Calculated OT
Private bytOtConf As Byte               '' For OT Confirmation
Private strSelEmp As String                '' For IN in Empcode query

Private Sub Form_Load()
Call SetFormIcon(Me)        '' Sets the Form Icon
Call RetCaptions            '' Sets the Captions
Call FillCombos             '' Fills All the Combos on the Form
Call GetRights              '' Gets and Sets the Rights
Call LoadSpecifics
End Sub

Private Sub RetCaptions()
On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '62%'", VstarDataEnv.cnDJConn, adOpenStatic
Me.Caption = NewCaptionTxt("62001", adrsC)
'' Employee Details

frOT.Caption = NewCaptionTxt("62002", adrsC)                        '' Remarks
chkOT.Caption = NewCaptionTxt("62001", adrsC)                       '' OT Authorized
lblOT.Caption = NewCaptionTxt("62007", adrsC)                       '' Overtime Hrs.
lblOTRem.Caption = NewCaptionTxt("00125", adrsMod)
cmdSave.Caption = NewCaptionTxt("00007", adrsMod)                   '' Save
cmdCancel.Caption = NewCaptionTxt("00003", adrsMod)                 '' Cancel
cmdExit.Caption = NewCaptionTxt("00008", adrsMod)                   '' Save
Call CapGrid
End Sub

Private Sub CapGrid()
On Error Resume Next
With MSF1                       '' Captions
    .TextMatrix(0, 0) = NewCaptionTxt("00061", adrsMod)         '' Employee Code
    .TextMatrix(0, 1) = NewCaptionTxt("00031", adrsMod)         '' Shift
    .TextMatrix(0, 2) = NewCaptionTxt("00034", adrsMod)         '' Arrival
    .TextMatrix(0, 3) = NewCaptionTxt("00036", adrsMod)         '' Departure
    .TextMatrix(0, 4) = NewCaptionTxt("62006", adrsC)           '' Work Hours
    .TextMatrix(0, 5) = NewCaptionTxt("62005", adrsC)           '' Overtime Authorized
    .TextMatrix(0, 6) = NewCaptionTxt("00038", adrsMod)         '' Overtime
    .TextMatrix(0, 7) = NewCaptionTxt("00124", adrsMod)         '' Remarks
    '' Widths
    .ColWidth(0) = .ColWidth(8) * 1.3
    .ColWidth(1) = .ColWidth(1) * 0.6
    .ColWidth(2) = .ColWidth(2) * 0.85
    .ColWidth(3) = .ColWidth(3) * 0.9
    .ColWidth(5) = .ColWidth(5) * 1.2
    .ColWidth(6) = .ColWidth(6) * 0.75
    .ColWidth(7) = .ColWidth(7) * 1
    '' Alignments
    .ColAlignment(1) = flexAlignCenterCenter
    .ColAlignment(5) = flexAlignCenterCenter
End With
End Sub

Private Sub GetRights()
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(4, 13, 3, 1)
If strTmp = "1" Then
    cmdSave.Enabled = True
Else
    cmdSave.Enabled = False
End If
Exit Sub
ERR_P:
    ShowError ("GetRights::" & Me.Caption)
    cmdSave.Enabled = False
End Sub

Private Sub LoadSpecifics()
On Error GoTo ERR_P
txtFrom.Text = DateDisp(Date)

Exit Sub
ERR_P:
    ShowError ("LoadSpecifics : " & Me.Caption)
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

Private Sub cboDept_Click()
On Error GoTo ERR_P
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
cboDept.ListIndex = cboDept.ListCount - 1
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
strDeptTmp = cboDept.List(cboDept.ListIndex, 0)
strDeptTmp = EncloseQuotes(strDeptTmp)
Select Case UCase(Trim(strDeptTmp))
    Case "", "ALL"
        strTempforCF = "Select Empcode,Name from empmst ORDER BY Empcode"
    Case Else
        If strCurrentUserType = HOD Then
            strTempforCF = "Select Empcode,Name from empmst " & strCurrData & " and Empmst." & SELCRIT & "=" & _
                strDeptTmp & " Order by EmpCode"
        Else
            strTempforCF = "Select Empcode,Name from empmst WHERE Empmst." & SELCRIT & " = " & strDeptTmp & _
            " Order by EmpCode"
        End If
End Select
If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open strTempforCF, VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
If (adrsEmp.EOF And adrsEmp.BOF) Then
    cboFrom.Clear
    cboTo.Clear
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

Private Sub cmdCancel_Click()
cmdExit.Cancel = True
frOT.Visible = False
End Sub

Private Sub cmdsave_Click()
On Error GoTo ERR_P
Dim strTmp As String, strTmp1 As String
If Val(txtOT.Text) = 0 Then
    strTmp = ""
    strTmp1 = ""
Else
    strTmp = IIf(chkOT.Value = 1, "Y", "N")
    strTmp1 = Trim(txtRem.Text)
End If
If Val(txtOT.Text) > sngCalcOT Then
    MsgBox NewCaptionTxt("62004", adrsC), vbExclamation
    txtOT.SetFocus
    Exit Sub
End If
If Not CheckEmployee Then Exit Sub
VstarDataEnv.cnDJConn.Execute "Update " & strFileName & " Set OTConf='" & strTmp _
& "',Ovtim=" & Val(txtOT.Text) & ",OTRem='" & strTmp1 & "' Where " & strKDate & " =" & _
strDTEnc & DateCompStr(txtFrom.Text) & strDTEnc & " and Empcode in (" & strSelEmp & ")"
cmdExit.Cancel = True
frOT.Visible = False
Call FillGrid
Exit Sub
ERR_P:
    ShowError ("Save::" & Me.Caption)
End Sub

Private Function CheckEmployee() As Boolean     '' Function to Check if Employees are
strSelEmp = ""                                  '' Selected or not

CheckEmployee = True
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


Private Sub FillGrid()          '' Fills the Grid
On Error GoTo ERR_P
If adrsLeave.State = 1 Then adrsLeave.Close
adrsLeave.Open "Select Empcode," & strKDate & " ,Shift,ArrTim,Deptim,WrkHrs,Ovtim,OTConf,OTRem from " & _
strFileName & " Where OvTim>0" & _
" order by " & strKDate & " ", VstarDataEnv.cnDJConn, adOpenStatic
MSF1.Rows = 1
Do While Not adrsLeave.EOF
    If Not adrsLeave.EOF Then
        MSF1.Rows = MSF1.Rows + 1
    Else
        Exit Do
    End If
    With MSF1
        '' Code for Grid Display
        .TextMatrix(MSF1.Rows - 1, 0) = DateDisp(adrsLeave("Date"))                 '' Date
        .TextMatrix(MSF1.Rows - 1, 1) = IIf(IsNull(adrsLeave("Shift")), "", _
                                        adrsLeave("Shift"))                         '' Shift
        .TextMatrix(MSF1.Rows - 1, 2) = IIf(IsNull(adrsLeave("ArrTim")), "0.00", _
                                        Format(adrsLeave("ArrTim"), "0.00"))        '' Arrival
        .TextMatrix(MSF1.Rows - 1, 3) = IIf(IsNull(adrsLeave("DepTim")), "0.00", _
                                        Format(adrsLeave("DepTim"), "0.00"))        '' Departure
        .TextMatrix(MSF1.Rows - 1, 4) = IIf(IsNull(adrsLeave("WrkHrs")), "0.00", _
                                        Format(adrsLeave("WrkHrs"), "0.00"))        '' Work Hours
        .TextMatrix(MSF1.Rows - 1, 5) = IIf(adrsLeave("OTConf") = "Y", _
                                        NewCaptionTxt("00100", adrsMod), _
                                        NewCaptionTxt("00101", adrsMod))            '' Yes or no
        .TextMatrix(MSF1.Rows - 1, 6) = IIf(IsNull(adrsLeave("OvTim")), "0.00", _
                                        Format(adrsLeave("OvTim"), "0.00"))         '' Overtime
        .TextMatrix(MSF1.Rows - 1, 7) = IIf(IsNull(adrsLeave("OTRem")), "", _
                                        adrsLeave("OTRem"))                         '' Overtime
    End With
    adrsLeave.MoveNext
Loop
If MSF1.Rows = 1 Then
    MsgBox NewCaptionTxt("62008", adrsC), vbExclamation
End If
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub MSF1_dblClick()
If MSF1.Rows = 1 Then Exit Sub
MSF1.col = 0
If MSF1.Text = NewCaptionTxt("00030", adrsMod) Then Exit Sub
sngCalcOT = Val(MSF1.TextMatrix(MSF1.row, 6))
If sngCalcOT <= 0 Then
    MsgBox NewCaptionTxt("62009", adrsC), vbExclamation
    frOT.Visible = False
    Exit Sub
End If
bytOtConf = IIf(MSF1.TextMatrix(MSF1.row, 5) = NewCaptionTxt("00101", adrsMod), 0, 1)
Call ShowOTDetails
End Sub

Public Sub ShowOTDetails()
frOT.Visible = True
lblDate.Caption = MSF1.TextMatrix(MSF1.row, 0)
chkOT.Value = bytOtConf
txtOT.Text = sngCalcOT
txtRem.Text = MSF1.TextMatrix(MSF1.row, 7)
cmdCancel.Cancel = True
End Sub

Private Sub MSF1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 32, 13
        Call MSF1_dblClick
End Select
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

Private Sub txtOT_GotFocus()
    Call GF(txtOT)
End Sub

Private Sub txtOT_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtOT)
End If
End Sub

Private Sub txtRem_GotFocus()
    Call GF(txtRem)
End Sub

Private Sub txtRem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 7)
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

