VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmExportTata 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   3675
   ClientTop       =   2475
   ClientWidth     =   5010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frAssume 
      Height          =   855
      Left            =   5520
      TabIndex        =   19
      Top             =   3960
      Width           =   4995
      Begin VB.TextBox txtFDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1830
         MaxLength       =   2
         TabIndex        =   21
         Text            =   "23"
         Top             =   500
         Width           =   280
      End
      Begin VB.CheckBox chkPrev 
         Caption         =   "Consider data from last month's file"
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
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3465
      End
      Begin VB.Label lblonwards 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Onwards"
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
         Left            =   2190
         TabIndex        =   23
         Top             =   495
         Width           =   795
      End
      Begin VB.Label lblFromDay 
         BackStyle       =   0  'Transparent
         Caption         =   "from the day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   400
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame frMonth 
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   -60
      Width           =   4995
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1500
      End
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   3330
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
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
         Left            =   180
         TabIndex        =   0
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   2640
         TabIndex        =   3
         Top             =   270
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   495
      Left            =   3600
      TabIndex        =   18
      Top             =   5040
      Width           =   1395
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1080
      TabIndex        =   17
      Top             =   5160
      Width           =   1395
   End
   Begin VB.Frame frEmp 
      Caption         =   "Employee Selection"
      Height          =   4365
      Left            =   0
      TabIndex        =   5
      Top             =   660
      Width           =   4995
      Begin VB.CommandButton cmdSR 
         Caption         =   "Select &Range"
         Height          =   435
         Left            =   3630
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdUR 
         Caption         =   "&Unselect Range"
         Height          =   465
         Left            =   3630
         TabIndex        =   14
         Top             =   1500
         Width           =   1335
      End
      Begin VB.CommandButton cmdSA 
         Caption         =   "&Select All"
         Height          =   435
         Left            =   3630
         TabIndex        =   15
         Top             =   2100
         Width           =   1335
      End
      Begin VB.CommandButton cmdUA 
         Caption         =   "U&nselect All"
         Height          =   435
         Left            =   3630
         TabIndex        =   16
         Top             =   2520
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   3435
         Left            =   30
         TabIndex        =   12
         ToolTipText     =   "Click to Toggle Selection"
         Top             =   900
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   6059
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
         TabIndex        =   8
         Top             =   630
         Width           =   345
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T&o"
         Height          =   195
         Left            =   2790
         TabIndex        =   10
         Top             =   630
         Width           =   195
      End
      Begin MSForms.ComboBox cboFrom 
         Height          =   315
         Left            =   1350
         TabIndex        =   9
         Top             =   570
         Width           =   1365
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
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
         Left            =   3300
         TabIndex        =   11
         Top             =   570
         Width           =   1365
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
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
         Left            =   1350
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   5520
      Width           =   5055
   End
End
Attribute VB_Name = "frmExportTata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''00049 please select employee
'' Monthly Process Form Module
'' ----------------------

Option Explicit
Dim strTableName As String
Dim adrsC As New ADODB.Recordset
Dim LvFiles As LeaveFile
Private Sub Form_Load()
ReDim strArrStatus(0 To 62)
ReDim sngArrTotl(5)
Call SetFormIcon(Me)
Call RetCaptions
Call SetFrame
Call SetValues
Call FillCombos
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '36%'", VstarDataEnv.cnDJConn, adOpenStatic
Me.Caption = "Export" 'NewCaptionTxt("36001", adrsC)              '' Monthly Process
Call SetCritLabel(lblDeptCap)
Call SetGridDetails(Me, frEmp, MSF1, lblFrom, lblTo)
'ChkLateEarl.Caption = NewCaptionTxt("36006", adrsC)     '' Execute Late/Early Rules
cmdProcess.Caption = "&Export" 'NewCaptionTxt("36010", adrsC)      '' Process
cmdExit.Caption = NewCaptionTxt("00039", adrsMod)         '' Finish
If InVar.blnAssum = "1" Then
    lblMonth = NewCaptionTxt("00026", adrsMod)            '' Month
    lblYear = NewCaptionTxt("00027", adrsMod)             '' Year
    chkPrev.Caption = NewCaptionTxt("36007", adrsC)     '' Consider data from last month's file
    lblFromDay.Caption = NewCaptionTxt("36008", adrsC)  '' from the day
    lblonwards.Caption = NewCaptionTxt("36009", adrsC)  '' Onwards
Else
    'lblFrDate = NewCaptionTxt("00010", adrsMod)           '' From Date
    'lblToDate = NewCaptionTxt("00011", adrsMod)           '' To Date
End If
Call CapGrid
End Sub

Private Sub SetFrame()
On Error GoTo ERR_P
If InVar.blnAssum = "1" Then
    txtFDate.Enabled = False
    'frMonth(1).Enabled = False
Else
    frMonth(1).Left = frMonth(0).Left
    frMonth(1).Top = frMonth(0).Top
    frMonth(0).Enabled = False
    frAssume.Visible = False
    Me.Height = Me.Height - (frAssume.Height)
    cmdProcess.Top = frAssume.Top + 50
    cmdExit.Top = cmdProcess.Top
End If
Exit Sub
ERR_P:
    ShowError ("SetFrame :: " & Me.Caption)
End Sub

Private Sub SetValues()
On Error GoTo ERR_P
Dim bytCnt As Byte

If InVar.blnAssum = "1" Then
    For bytCnt = 0 To 99
        cmbYear.AddItem (1997 + bytCnt)
    Next
    For bytCnt = 1 To 12
        cmbMonth.AddItem MonthName(bytCnt)
    Next
    cmbYear.Text = Year(Date)
    cmbMonth.Text = MonthName(Month(Date))
End If
Exit Sub
ERR_P:
    ShowError ("SetValues :: " & Me.Caption)
End Sub

Private Sub FillCombos()
On Error GoTo ERR_P
Call SetCritCombos(cboDept)
cboDept.ListIndex = cboDept.ListCount - 1
Exit Sub
ERR_P:
    ShowError ("FillCombos::" & Me.Caption)
End Sub

Private Sub cboDept_Click()
On Error GoTo ERR_P
If cboDept.ListIndex < 0 Then Exit Sub               '' If No Department
Call FillComboGrid
Call SelUnselAll(UNSELECTED_COLOR, MSF1)
Exit Sub
ERR_P:
    ShowError ("Department::" & Me.Caption)
End Sub

Private Sub FillComboGrid()
On Error GoTo ERR_P
Dim adrsTmp As New ADODB.Recordset, intEmpCnt As Integer, intTmpCnt As Integer
Dim strDeptTmp As String, strTempforCF As String
Dim strArrEmp() As String
Call ComboFill(cboFrom, 12, 2, cboDept.List(cboDept.ListIndex, 0))
Call ComboFill(cboTo, 12, 2, cboDept.List(cboDept.ListIndex, 0))
If cboFrom.ListCount > 0 Then cboFrom.ListIndex = 0
If cboTo.ListCount > 0 Then cboTo.ListIndex = cboTo.ListCount - 1
strDeptTmp = cboDept.List(cboDept.ListIndex, 0)
strDeptTmp = EncloseQuotes(strDeptTmp)
Select Case UCase(Trim(strDeptTmp))
    Case "", "ALL"
        strTempforCF = "select Empcode,name from empmst order by Empcode"               'Empcode,name
    Case Else
        ''For Mauritius 09-08-2003
        ''Original ->strTempforCF = "select Empcode,name from empmst Where " & SELCRIT & "=" & _
        strDeptTmp & " order by Empcode"                               'Empcode,name
        If strCurrentUserType = HOD Then
            strTempforCF = "select Empcode,name from empmst " & strCurrData & " and Empmst." & SELCRIT & "='" & _
                strDeptTmp & "' order by Empcode"    'Empcode,name
        Else
            strTempforCF = "select Empcode,name from empmst where Empmst." & SELCRIT & "='" & _
                strDeptTmp & "' order by Empcode"    'Empcode,name
        End If
End Select
If adrsTmp.State = 1 Then adrsTmp.Close
adrsTmp.Open strTempforCF, VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
If (adrsTmp.EOF And adrsTmp.BOF) Then
    cboFrom.Clear
    cboTo.Clear
    MSF1.Rows = 1
    Exit Sub
End If
intEmpCnt = adrsTmp.RecordCount
intTmpCnt = intEmpCnt
MSF1.Rows = intEmpCnt + 1
ReDim strArrEmp(intTmpCnt - 1, 1)
For intEmpCnt = 0 To intTmpCnt - 1
    strArrEmp(intEmpCnt, 0) = adrsTmp(0)
    strArrEmp(intEmpCnt, 1) = adrsTmp(1)
    MSF1.TextMatrix(intEmpCnt + 1, 0) = adrsTmp(0)
    MSF1.TextMatrix(intEmpCnt + 1, 1) = adrsTmp(1)
    adrsTmp.MoveNext
Next
cboFrom.List = strArrEmp
cboTo.List = strArrEmp
cboFrom.ListIndex = 0
cboTo.ListIndex = cboTo.ListCount - 1
Erase strArrEmp
Exit Sub
ERR_P:
    ShowError ("Fill Employee Combos :: " & Me.Caption)
    Resume Next
End Sub

Private Sub Form_Activate()
    Call GetRights
End Sub

Private Sub GetRights()
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(4, 7, 4, 1)
If strTmp = "1" Then
    cmdProcess.Enabled = True
Else
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    cmdProcess.Enabled = False
End If
Exit Sub
ERR_P:
    ShowError ("GetRights :: " & Me.Caption)
    cmdProcess.Enabled = False
End Sub

Private Sub cmdProcess_Click()
On Error GoTo ERR_P
If Not monValid Then Exit Sub
If Not ValidAssume Then Exit Sub
'this activity log add log same as monthly process
Call AddActivityLog(lg_NoModeAction, 2, 26)     '' Process Log
Call AuditInfo("MONTHLY PROCESS", Me.Caption, "Export Data Processing: Done Monthly Process")
If monProcess Then
    MsgBox "Export Completed", vbInformation
Else
    lblStatus.Caption = "Some Problem"
End If
Exit Sub
ERR_P:
    ShowError ("Process :: " & Me.Caption)
End Sub

Private Function monValid() As Boolean
On Error GoTo ERR_P
monValid = True
Dim strYear As String
If Not CheckEmployee Then           ''check if any employee is selected.
    monValid = False
    Exit Function
End If
If Val(pVStar.Yearstart) > MonthNumber(cmbMonth.Text) Then
    strYear = CStr(Val(cmbYear.Text) - 1)
Else
    strYear = cmbYear.Text
End If
LvFiles.strLvTrn = "Lvtrn" & Right(strYear, 2)
If Not FindTable(LvFiles.strLvTrn) Then
    MsgBox LvFiles.strLvTrn & " This File Not Present", vbInformation
    monValid = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("MonValid :: " & Me.Caption)
End Function

Private Function CheckEmployee() As Boolean     '' Function to Check if Employees are
Dim intEmpTmp As Integer                        '' Selected or not
intEmpTmp = 0
CheckEmployee = True
If MSF1.Rows = 1 Then
    MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation, App.EXEName
    CheckEmployee = False
    cmdSR.SetFocus
    Exit Function
End If
MSF1.Col = 0
typMnlVar.strEmpList = ""
For i = 1 To MSF1.Rows - 1
    MSF1.row = i
    If MSF1.CellBackColor = SELECTED_COLOR Then
        intEmpTmp = intEmpTmp + 1
        typMnlVar.strEmpList = typMnlVar.strEmpList & "'" & Trim(MSF1.Text) & "',"
    End If
Next
If Trim(typMnlVar.strEmpList) <> "" Then
    typMnlVar.strEmpList = Left(typMnlVar.strEmpList, Len(typMnlVar.strEmpList) - 1)
End If
If intEmpTmp = 0 Then
    MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation
    CheckEmployee = False
    cmdSR.SetFocus
End If
End Function

Private Function ValidAssume() As Boolean
On Error GoTo ERR_P
Dim strTempDate As String
'' Current Month Transaction file
strTableName = MakeName(cmbMonth.Text, cmbYear.Text, "trn") & "E"
If Not FindTable(strTableName) Then
    MsgBox NewCaptionTxt("36019", adrsC) & cmbMonth.Text & " " & _
        cmbYear.Text & NewCaptionTxt("00055", adrsMod), vbExclamation
    ValidAssume = False
    Exit Function
End If
ValidAssume = True
Exit Function
ERR_P:
ValidAssume = False
ShowError ("ValidAssume :: " & Me.Caption)
End Function

Private Function GetDateOfDay(ByVal bytDay As Byte, ByVal strMonth As String, _
strYear As String) As String        '' Function to make Date
On Error GoTo ERR_P
Select Case bytDateF
    Case 1      '' American (MM/DD/YY)
        GetDateOfDay = Format(MonthNumber(strMonth), "00") & "/" & Format(bytDay, "00") & _
        "/" & strYear
    Case 2      '' British  (DD/MM/YY)
        GetDateOfDay = Format(bytDay, "00") & "/" & Format(MonthNumber(strMonth), "00") & _
        "/" & strYear
End Select
Exit Function
ERR_P:
    ShowError ("Get Date of day :: Rotation Module")
End Function

Private Function monProcess() As Boolean
On Error GoTo ERR_P
    Dim strquery As String
    
    strquery = "SELECT e.name,l.*, " & _
    " DAY(l.lst_date) AS TD,encash as PLE,[SD-01],[SD-02],[SD-03] " & _
    " FROM " & strTableName & " m," & LvFiles.strLvTrn & " l,empmst e " & _
    " WHERE m.empcode=l.empcode AND l.lst_date=" & strDTEnc & _
    "" & FdtLdt(MonthNumber(cmbMonth.Text), cmbYear.Text, "L") & "" & _
    strDTEnc & " AND e.empcode=m.empcode AND m.empcode IN (" & typMnlVar.strEmpList & ")"
    
    If Export(strquery) Then
        lblStatus.Caption = "First Export Completed"
    End If
    
    strquery = "SELECT e.empcode,e.name,CAST(YEAR(l.lst_date) " & _
    " AS VARCHAR(20)) + CAST(DAY(l.lst_date) AS VARCHAR(20)) AS Months, " & _
    " m.days,l.otpd_hrs AS OT,m.[SD-01] AS Wo_Day,m.[SD-02] AS Att_Bon_Per," & _
    " m.[SD-02] AS Att_Bon_Trainee" & _
    " FROM empmst e, " & strTableName & " m," & LvFiles.strLvTrn & _
    " l WHERE e.empcode=m.empcode AND l.lst_date=" & strDTEnc & _
    "" & FdtLdt(MonthNumber(cmbMonth.Text), cmbYear.Text, "L") & "" & _
    strDTEnc & " AND l.empcode=e.empcode AND m.empcode IN (" & typMnlVar.strEmpList & ")"
    
    If Export(strquery) Then
        lblStatus.Caption = "Second and First Both Export Completed"
    End If
    
    monProcess = True
Exit Function
ERR_P:
    ShowError ("MonProcess :: " & Me.Caption)
    monProcess = False
    ''Resume Next
End Function

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub MSF1_Click()
If MSF1.Rows = 1 Then Exit Sub
If MSF1.CellBackColor = SELECTED_COLOR Then
    With MSF1
        .Col = 0
        .CellBackColor = UNSELECTED_COLOR
        .Col = 1
        .CellBackColor = UNSELECTED_COLOR
    End With
Else
    With MSF1
        .Col = 0
        .CellBackColor = SELECTED_COLOR
        .Col = 1
        .CellBackColor = SELECTED_COLOR
    End With
End If
End Sub

Private Sub MSF1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then Call MSF1_Click
End Sub

Private Sub CapGrid()
'' Sizing
MSF1.ColWidth(1) = MSF1.ColWidth(1) * 2.65
'' Aligning
MSF1.ColAlignment(0) = flexAlignLeftTop
End Sub

Private Sub cmdSA_Click()
Call SelUnselAll(SELECTED_COLOR, MSF1)
End Sub

Private Sub cmdSR_Click()
Call SelUnsel(SELECTED_COLOR, MSF1, cboFrom, cboTo)
End Sub

Private Sub cmdUA_Click()
Call SelUnselAll(UNSELECTED_COLOR, MSF1)
End Sub

Private Sub cmdUR_Click()
Call SelUnsel(UNSELECTED_COLOR, MSF1, cboFrom, cboTo)
End Sub
