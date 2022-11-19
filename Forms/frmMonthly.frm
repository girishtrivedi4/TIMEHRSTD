VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMonthly 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   3720
   ClientTop       =   2910
   ClientWidth     =   5730
   FillStyle       =   7  'Diagonal Cross
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkLunchLt 
      Caption         =   "Lunch Late"
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
      Left            =   4320
      TabIndex        =   9
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Frame frMonth 
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
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   4275
      Begin VB.TextBox txtToDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Text            =   "00/00/2001"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtFrDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   0
         Text            =   "01/00/2001"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         Left            =   2280
         TabIndex        =   22
         Top             =   427
         Width           =   270
      End
      Begin VB.Label lblFrDate 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         Left            =   240
         TabIndex        =   21
         Top             =   427
         Width           =   450
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msfProcess 
      Height          =   495
      Left            =   3360
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   873
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.CheckBox ChkLateEarl 
      Caption         =   "Execute Late Early Rules"
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
      Left            =   360
      TabIndex        =   10
      Top             =   6000
      Width           =   3675
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   7320
      Width           =   1755
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   7320
      Width           =   1755
   End
   Begin VB.Frame frAssume 
      Height          =   735
      Left            =   240
      TabIndex        =   29
      Top             =   6480
      Visible         =   0   'False
      Width           =   4995
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
         TabIndex        =   11
         Top             =   150
         Width           =   3465
      End
      Begin VB.TextBox txtFDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "23"
         Top             =   400
         Width           =   280
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
         TabIndex        =   30
         Top             =   400
         Width           =   1095
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
         TabIndex        =   31
         Top             =   400
         Width           =   795
      End
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
      Height          =   5385
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   5475
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
         Height          =   435
         Left            =   3960
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
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
         Height          =   435
         Left            =   3960
         TabIndex        =   6
         Top             =   2040
         Width           =   1455
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
         Height          =   435
         Left            =   3960
         TabIndex        =   7
         Top             =   2580
         Width           =   1455
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
         Height          =   435
         Left            =   3960
         TabIndex        =   8
         Top             =   3120
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   3675
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "Click to Toggle Selection"
         Top             =   1080
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   6482
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
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
         Left            =   4080
         TabIndex        =   33
         Top             =   360
         Width           =   705
      End
      Begin MSForms.ComboBox CboLocation 
         Height          =   315
         Left            =   3840
         TabIndex        =   32
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
      Begin MSForms.ComboBox cboDept 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   2595
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4577;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
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
         Left            =   720
         TabIndex        =   25
         Top             =   720
         Width           =   435
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
         TabIndex        =   26
         Top             =   720
         Width           =   210
      End
      Begin MSForms.ComboBox cboFrom 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         Width           =   1125
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1984;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboTo 
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Top             =   720
         Width           =   1005
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1773;556"
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
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame frMonth 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   4680
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   5835
      Begin VB.ComboBox cmbMonth 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   1500
      End
      Begin VB.ComboBox cmbYear 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3450
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label lblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
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
         Left            =   300
         TabIndex        =   15
         Top             =   405
         Width           =   495
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   2760
         TabIndex        =   18
         Top             =   390
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmMonthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''00049 please select employee
'' Monthly Process Form Module
'' ----------------------
Option Explicit
Private strLvBal As String, strLvInfo As String, strLvTrn As String
Private strMonTrnCurr As String, strMonTrnPrev As String
Private sngTotLates As Single, sngTotEarls As Single
Private STRECODE As String, strCatCode As String, strquery As String
Private sngPd_Dys As Single '', sngCmpOff As Single
Private sngIns_WOOt As Single, sngIns_HLOt As Single, sngIns_OTOt As Single
Private bytCase As Byte, bytDate As Byte
Private bytNight As Byte, bytLtNo As Byte, bytErNo As Byte
Private sngLtHrs As Single, sngErHrs As Single
Private bytLunchLtNo As Byte
Private sngLunchLtHrs As Single, sngTotLunchLates As Single
Private sngWrkHrs As Single, sngOtHrs As Single, sngOtH_pd As Single
Private sngArrTotl
Private strArrStatus
Public adrslv As New ADODB.Recordset
Private strFields
''
Dim adrsC As New ADODB.Recordset


Private blnCalculateFirstTime As Boolean
Private RsWeekOffPaid As New ADODB.Recordset        '' 27-12
Private weekoffnot As New ADODB.Recordset            '

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


Private Sub ChkLunchLt_Click()
If ChkLunchLt.Value = 1 Then
    If (MsgBox(NewCaptionTxt("36011", adrsC) & _
    vbCrLf & NewCaptionTxt("36012", adrsC) & _
    vbCrLf & NewCaptionTxt("00009", adrsMod), _
    vbYesNo + vbQuestion)) = vbNo Then ChkLunchLt.Value = 0
End If
End Sub


Private Sub Form_Load()

ReDim strArrStatus(0 To 62)
ReDim sngArrTotl(5)

'frAssume.Visible = True

ChkLunchLt.Visible = False

Call SetFormIcon(Me)
Call GetInstPara    '' GET PARAMETERS FROM INSTALL TABLE FOR OVERTIME CALC
Call RetCaptions
'Call SetFrame
Call SetValues
Call FillCombos
If bytRepMode = 6 And typOptIdx.bytPer = 7 Then    ' 14-01
    If Not AutoSel Then PerAtt = False
End If
lblLocation.Visible = GetFlagStatus("pratham")
CboLocation.Visible = GetFlagStatus("pratham")

End Sub

Private Function EnableDisablComm(blnI As Boolean)  'Added by  14-01
    Me.lblDeptCap.Visible = blnI: Me.cboDept.Visible = blnI
    Me.lblFrom.Visible = blnI: Me.cboFrom.Visible = blnI
    Me.lblTo.Visible = blnI: Me.cboTo.Visible = blnI
    Me.cmdSR.Visible = blnI: Me.cmdUR.Visible = blnI
    Me.cmdSA.Visible = blnI: Me.cmdUA.Visible = blnI
End Function
Private Function AutoSel()  ' 14-01
On Error GoTo ERR_P
    AutoSel = True
    Dim adrsTmp As New ADODB.Recordset
    Dim intEmpCnt As Integer, intTmpCnt As Integer
    
    EnableDisablComm (False)
    If adrsTmp.State = 1 Then adrsTmp.Close
    'adrsTmp.Open "select Empcode,name from " & rpTables & " where " & "''" & strSql & " order by Empcode", conmain, adOpenStatic, adLockReadOnly
    adrsTmp.Open "select Empcode,name from " & rpTables & " where empmst.empcode<>'NULL' " & strSql & " order by Empcode", ConMain, adOpenStatic, adLockReadOnly
    If (adrsTmp.EOF And adrsTmp.BOF) Then
        MSF1.Rows = 1
        Exit Function
    End If
    intEmpCnt = adrsTmp.RecordCount
    intTmpCnt = intEmpCnt
    MSF1.Rows = intEmpCnt + 1
    ReDim strArrEmp(intTmpCnt - 1, 1)
    For intEmpCnt = 0 To intTmpCnt - 1
        MSF1.TextMatrix(intEmpCnt + 1, 0) = adrsTmp(0)
        MSF1.TextMatrix(intEmpCnt + 1, 1) = adrsTmp(1)
        adrsTmp.MoveNext
    Next
    Call SelUnselAll(SELECTED_COLOR, MSF1)
    Exit Function
ERR_P:
    ShowError ("error in AutoSel")
    AutoSel = False
    'Resume Next
    Exit Function
End Function
Private Sub GetInstPara()
On Error GoTo ERR_P
Dim adrsInstall1 As New ADODB.Recordset
If adrsInstall1.State = 1 Then adrsInstall1.Close
adrsInstall1.Open "select definCut,cutdt from install", ConMain, adOpenStatic
If Not (adrsInstall1.EOF And adrsInstall1.BOF) Then
    If IIf(IsNull(adrsInstall1("defincut")) Or adrsInstall1("defincut") = "N", "N", adrsInstall1("defincut")) = "Y" Then
        bytDate = IIf(IsNull(adrsInstall1("cutdt")) = True, 0, adrsInstall1("cutdt"))
    Else
        bytDate = 0
    End If
    If bytDate = 31 Then
        txtFDate.Text = 1
    ElseIf bytDate = 0 Then
        chkPrev.Value = 0
        txtFDate.Text = 0
    Else
        txtFDate.Text = bytDate + 1
    End If
End If
Exit Sub
ERR_P:
    ShowError ("Get Install Parameters :: " & Me.Caption)
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '36%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("36001", adrsC)              '' Monthly Process
Call SetCritLabel(lblDeptCap)
ChkLateEarl.Caption = NewCaptionTxt("36006", adrsC)     '' Execute Late/Early Rules
cmdProcess.Caption = "Process"
cmdExit.Caption = "Finish"
If InVar.blnAssum = "1" Then
    lblMonth = NewCaptionTxt("00026", adrsMod)            '' Month
    lblYear = NewCaptionTxt("00027", adrsMod)             '' Year
    chkPrev.Caption = NewCaptionTxt("36007", adrsC)     '' Consider data from last month's file
End If
Call CapGrid
Call SetGridDetails(Me, frEmp, MSF1, lblFrom, lblTo)
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
Else
        txtFrDate.Text = DateDisp(Date)
        txtToDate.Text = DateDisp(Date)
End If
msfProcess.TextMatrix(0, 0) = NewCaptionTxt("00061", adrsMod)
msfProcess.ColAlignment(0) = flexAlignLeftCenter
msfProcess.ForeColor = vbBlue
msfProcess.ColWidth(0) = msfProcess.ColWidth(0) * 1.3
msfProcess.ColWidth(1) = msfProcess.ColWidth(1) * 0.9
msfProcess.Width = msfProcess.ColWidth(0) + msfProcess.ColWidth(1) + 100
msfProcess.Height = msfProcess.CellHeight + 100
msfProcess.Visible = False
Exit Sub
ERR_P:
    ShowError ("SetValues :: " & Me.Caption)
End Sub

Private Sub FillCombos()
On Error GoTo ERR_P
    Call SetCritCombos(cboDept)
    Call ComboFill(CboLocation, 11, 2)
If strCurrentUserType <> HOD Then cboDept.Text = "ALL"
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
Dim adrsTmp As New ADODB.Recordset
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
            strTempforCF = "select Empcode,name from empmst order by Empcode"               'Empcode,name
        Case Else
            ''For Mauritius 09-08-2003
            ''Original ->strTempforCF = "select Empcode,name from empmst Where " & SELCRIT & "=" & _
            strDeptTmp & " order by Empcode"                               'Empcode,name
            If strCurrentUserType = HOD Then

                strTempforCF = "select Empcode,name from empmst " & strCurrData & " and Empmst." & SELCRIT & "=" & _
                    strDeptTmp & " order by Empcode"    'Empcode,name
      
            Else
                'this If condition add by  for datatype
                If blnFlagForDept = True Then
                    strTempforCF = "select Empcode,name from empmst where Empmst." & SELCRIT & "='" & _
                    strDeptTmp & "' order by Empcode"    'Empcode,name
                Else
                    strTempforCF = "select Empcode,name from empmst where Empmst." & SELCRIT & "=" & _
                    strDeptTmp & " order by Empcode"
                End If
            End If
    End Select
    If GetFlagStatus("LocationRights") And strCurrentUserType <> ADMIN Then ' 28-01-09
        If UCase(Trim(strDeptTmp)) = "ALL" Then
            strTempforCF = "select Empcode,name from empmst where Empmst.Location in (Select Dept From UserAccs Where UserName = '" & UserName & "') order by Empcode"
        Else
            strTempforCF = "select Empcode,name from empmst where Empmst.Dept = " & cboDept.Text & " And Empmst.Location in (Select Dept From UserAccs Where UserName = '" & UserName & "') order by Empcode"
        End If
    End If

If adrsTmp.State = 1 Then adrsTmp.Close
adrsTmp.Open strTempforCF, ConMain, adOpenStatic, adLockReadOnly
If (adrsTmp.EOF And adrsTmp.BOF) Then
    cboFrom.clear
    cboTo.clear
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
    'Resume Next
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

Private Sub txtFrDate_Click()
varCalDt = ""
varCalDt = Trim(txtFrDate.Text)
txtFrDate.Text = ""
Call ShowCalendar
End Sub

Private Sub txtFrDate_GotFocus()
    Call GF(txtFrDate)
End Sub

Private Sub txtFrDate_KeyPress(KeyAscii As Integer)
    Call CDK(txtFrDate, KeyAscii)
End Sub

Private Sub txtFrDate_Validate(Cancel As Boolean)
If Not ValidDate(txtFrDate) Then
    txtFrDate.SetFocus
    Cancel = True
End If
End Sub

Private Sub txtToDate_Click()
varCalDt = ""
varCalDt = Trim(txtToDate.Text)
txtToDate.Text = ""
Call ShowCalendar
End Sub

Private Sub txtToDate_GotFocus()
    Call GF(txtToDate)
End Sub

Private Sub txtToDate_KeyPress(KeyAscii As Integer)
    Call CDK(txtToDate, KeyAscii)
End Sub

Private Sub txtToDate_Validate(Cancel As Boolean)
If Not ValidDate(txtToDate) Then
    txtToDate.SetFocus
    Cancel = True
End If
End Sub

Private Sub ChkLateEarl_Click()
If ChkLateEarl.Value = 1 Then
    If (MsgBox(NewCaptionTxt("36011", adrsC) & _
    vbCrLf & NewCaptionTxt("36012", adrsC) & _
    vbCrLf & NewCaptionTxt("00009", adrsMod), _
    vbYesNo + vbQuestion)) = vbNo Then ChkLateEarl.Value = 0
End If
End Sub

Private Sub chkPrev_Click()
If chkPrev.Value = 1 Then
    txtFDate.Enabled = True
    txtFDate.SetFocus
Else
    txtFDate.Enabled = False
End If
End Sub

Private Sub txtFDate_GotFocus()
With txtFDate
    .SelStart = 0
    .SelLength = Len(txtFDate.Text)
End With
End Sub

Private Sub txtFDate_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13: SendKeys (Chr(9))
    Case Else
        If InStr("1234567890", Chr(KeyAscii)) <= 0 Then KeyAscii = 0
End Select
End Sub

Public Sub cmdProcess_Click()
On Error GoTo ERR_P
If typOptIdx.bytPer = 7 And bytRepMode = 6 Then
    txtFrDate.Text = DateDisp(CDate(frmReports.txtFrPeri.Text))
    txtToDate.Text = DateDisp(CDate(frmReports.txtToPeri.Text))
End If
If Not monValid Then Exit Sub
If InVar.blnAssum = "1" Then
    If Not ValidAssume Then Exit Sub
Else
    If Not ValidActual Then Exit Sub
End If
Call AddActivityLog(lg_NoModeAction, 2, 26)     '' Process Log
If InVar.blnAssum = "0" Then
    Call AuditInfo("MONTHLY PROCESS", Me.Caption, "Done Monthly Process For The Period: " & txtFrDate.Text & " To " & txtToDate.Text)
Else
    Call AuditInfo("MONTHLY PROCESS", Me.Caption, "Done Monthly Process For: " & cmbMonth.Text & " " & cmbYear.Text)
End If
Call GetVals
 If Not monProcess Then
    msfProcess.Visible = False
    Exit Sub
 Else
 If typOptIdx.bytPer = 7 And bytRepMode = 6 Then Exit Sub
    msfProcess.Visible = False
    MsgBox NewCaptionTxt("36013", adrsC), vbInformation
    If InVar.blnAssum = "1" Then
        cmbMonth.SetFocus
    Else
        txtFrDate.SetFocus
    End If
End If

Exit Sub
ERR_P:
    msfProcess.Visible = False
    Resume Next
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

If InVar.blnAssum = "1" Then
    If chkPrev.Value = 1 And Val(txtFDate.Text) <= 0 Then
        MsgBox NewCaptionTxt("36014", adrsC) & vbCrLf & _
        NewCaptionTxt("36015", adrsC), vbInformation
        txtFDate.SetFocus
        monValid = False
        Exit Function
    End If
    If chkPrev.Value = 1 And Val(txtFDate.Text) > 31 Then
        MsgBox NewCaptionTxt("36016", adrsC), vbExclamation, App.EXEName
        txtFDate.SetFocus
        monValid = False
        Exit Function
    End If
End If
If InVar.blnAssum = "1" Then
    If Val(pVStar.Yearstart) > MonthNumber(cmbMonth.Text) Then
        strYear = CStr(Val(cmbYear.Text) - 1)
    Else
        strYear = cmbYear.Text
    End If
Else
    If Val(pVStar.Yearstart) > Month(txtToDate.Text) Then
        strYear = CStr(Year(txtToDate.Text) - 1)
    Else
        strYear = CStr(Year(txtToDate.Text))
    End If
End If
'' Leave transaction file
strLvTrn = "lvtrn" & Right(strYear, 2)
If Not FindTable(strLvTrn) Then
    MsgBox NewCaptionTxt("00054", adrsMod) & strYear & _
        NewCaptionTxt("00055", adrsMod), vbExclamation
    monValid = False
    Exit Function
End If
''Leave bal file
strLvBal = "lvbal" & Right(strYear, 2)
If Not FindTable(strLvBal) Then
    MsgBox NewCaptionTxt("36017", adrsC) & strYear & _
        NewCaptionTxt("00055", adrsMod), vbExclamation
    monValid = False
    Exit Function
End If
'' Leave info file
strLvInfo = "lvinfo" & Right(strYear, 2)
If Not FindTable(strLvInfo) Then
    MsgBox NewCaptionTxt("36017", adrsC) & strYear & _
        NewCaptionTxt("00055", adrsMod), vbExclamation
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
Call TruncateTable("Ecode") ' 09-03
For i = 1 To MSF1.Rows - 1
    MSF1.Row = i
    If MSF1.CellBackColor = SELECTED_COLOR Then
        intEmpTmp = intEmpTmp + 1
        'typMnlVar.strEmpList = typMnlVar.strEmpList & "'" & Trim(MSF1.Text) & "',"
        ConMain.Execute "insert into Ecode values('" & Trim(MSF1.Text) & "')" ' 09-03
    End If
    typMnlVar.strEmpList = "select empcode from Ecode "
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
strMonTrnCurr = MakeName(cmbMonth.Text, cmbYear.Text, "trn")
If Not FindTable(strMonTrnCurr) Then
    MsgBox NewCaptionTxt("36019", adrsC) & cmbMonth.Text & " " & _
        cmbYear.Text & NewCaptionTxt("00055", adrsMod), vbExclamation
    ValidAssume = False
    Exit Function
End If
If chkPrev.Value = 1 And Val(txtFDate.Text) > 0 Then
    strTempDate = DateAdd("m", -1, (GetDateOfDay(CByte(txtFDate.Text), cmbMonth.Text, _
        cmbYear.Text)))
    '' Previous Month Transaction file
    strMonTrnPrev = MakeName(MonthName(Month(strTempDate)), Year(DateCompDate(strTempDate)), "trn")
    If Not FindTable(strMonTrnPrev) Then
        strMonTrnPrev = strMonTrnCurr
    End If
Else
    strMonTrnPrev = strMonTrnCurr
End If
ValidAssume = True
Exit Function
ERR_P:
ValidAssume = False
ShowError ("ValidAssume :: " & Me.Caption)
End Function

Private Function ValidActual() As Boolean
On Error GoTo ERR_P
ValidActual = True
If Trim(txtFrDate.Text) = "" Then
    MsgBox NewCaptionTxt("00016", adrsMod), vbInformation
    txtFrDate.SetFocus
    ValidActual = False
    Exit Function
End If
If Trim(txtToDate.Text) = "" Then
    MsgBox NewCaptionTxt("00017", adrsMod), vbInformation
    txtToDate.SetFocus
    ValidActual = False
    Exit Function
End If
If CDate(txtToDate.Text) < CDate(txtFrDate.Text) Then
    MsgBox NewCaptionTxt("00018", adrsMod), vbInformation
    txtToDate.SetFocus
    ValidActual = False
    Exit Function
End If
If DateDiff("d", DateCompDate(txtFrDate.Text), DateCompDate(txtToDate.Text)) > 31 Then
    MsgBox NewCaptionTxt("36020", adrsC), vbInformation
    txtFrDate.SetFocus
    ValidActual = False
    Exit Function
End If
If DateDiff("m", DateCompDate(txtFrDate.Text), DateCompDate(txtToDate.Text)) > 1 Then
    MsgBox NewCaptionTxt("36021", adrsC), vbInformation
    txtFrDate.SetFocus
    ValidActual = False
    Exit Function
End If
'' Previous Month Transaction file
strMonTrnPrev = MakeName(MonthName(Month(DateCompDate(txtFrDate.Text))), _
    Year(DateCompDate(txtFrDate.Text)), "trn")
If Not FindTable(strMonTrnPrev) Then
    MsgBox NewCaptionTxt("36019", adrsC) & _
    MonthName(Month(DateCompDate(txtFrDate.Text))) & " " & _
    Year(DateCompDate(txtFrDate.Text)) & NewCaptionTxt("00055", adrsMod), vbExclamation
    ValidActual = False
    Exit Function
End If

'' Current Month Transaction file
strMonTrnCurr = MakeName(MonthName(Month(DateCompDate(txtToDate.Text))), _
    Year(DateCompDate(txtToDate.Text)), "trn")
If Not FindTable(strMonTrnCurr) Then
    MsgBox NewCaptionTxt("36019", adrsC) & _
    MonthName(Month(DateCompDate(txtToDate.Text))) & " " & _
    Year(DateCompDate(txtToDate.Text)) & NewCaptionTxt("00055", adrsMod), vbExclamation
    ValidActual = False
    Exit Function
End If

Exit Function
ERR_P:
ValidActual = False
ShowError ("ValidActual :: " & Me.Caption)
End Function

Private Sub GetVals()
On Error GoTo ERR_P
If InVar.blnAssum = "1" Then    '' Dates Preparation for Assume
    If chkPrev.Value = 1 And CInt(txtFDate.Text) > 0 Then
        If strMonTrnPrev = strMonTrnCurr Then
            typMnlVar.strFrtDate = FdtLdt(MonthNumber(cmbMonth.Text), cmbYear.Text, "f")
        Else
            typMnlVar.strFrtDate = DateAdd("m", -1, GetDateOfDay(CByte(txtFDate.Text), _
            cmbMonth.Text, cmbYear.Text))
        End If
    Else
        typMnlVar.strFrtDate = FdtLdt(MonthNumber(cmbMonth.Text), cmbYear.Text, "f")
    End If
    If bytDate = 0 Then
        typMnlVar.strLstDate = FdtLdt(MonthNumber(cmbMonth.Text), cmbYear.Text, "l")
    Else
        typMnlVar.strLstDate = GetDateOfDay(bytDate, cmbMonth.Text, cmbYear.Text)
    End If
    typMnlVar.strLvtDate = FdtLdt(MonthNumber(cmbMonth.Text), cmbYear.Text, "l")
Else                            '' Dates Preparation for Actual
    typMnlVar.strFrtDate = txtFrDate.Text
    typMnlVar.strLstDate = txtToDate.Text
    typMnlVar.strLvtDate = FdtLdt(Month(DateCompDate(txtToDate.Text)), Year(DateCompDate(txtToDate.Text)), "l")
End If

If ChkLateEarl.Value = 1 Then
    typMnlVar.bytExeLE = 1          '' Execute Late/Early
Else
    typMnlVar.bytExeLE = 2          '' Do not Execute
End If

Exit Sub
ERR_P:
    ShowError ("Get Values :: " & Me.Caption)
End Sub

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

Public Function monProcess() As Boolean
On Error GoTo ERR_P
monProcess = True

Dim blnCheckSD_2 As Boolean
Dim sngEncash As Single
''
If Not GetLvtTable Then monProcess = False: Exit Function     '' CREATE TEMP FILE "LVT" FROM StrLvTrn(1=2)
If Not OpenRSets Then monProcess = False: Exit Function          '' CREATE MASTER RECORDSETS REQUIRED WITHIN THE PROCESS
If Not GetTrnRs Then monProcess = False: Exit Function          '' CREATE RECORDSET FOR MONTHLY TRANSACTION TABLES TO PROCESS
Call FillField(strLvTrn)
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    msfProcess.Visible = True
    If GetFlagStatus("PERATTENDANCE") And typOptIdx.bytPer = 7 And bytRepMode = 6 Then
        If FindTable("prAbsentLvt") Then ConMain.Execute "drop table prAbsentLvt"
        Call CreateTableIntoAs("*", "leavtrn", "prAbsentLvt", " Where 1=2 ")
    End If
    If RsWeekOffPaid.State = 1 Then RsWeekOffPaid.Close          ''  26-12
        RsWeekOffPaid.Open "Select EmpMst.EmpCode, CatDesc.Cat, CatDesc.WeekOffPaid From EmpMst, CatDesc Where EmpMst.Cat = CatDesc.Cat AND Empmst.EmpCode  IN (" & typMnlVar.strEmpList & ")", ConMain, adOpenStatic

    Do While Not adrsTemp.EOF
        STRECODE = adrsTemp!Empcode
        strCatCode = GetCode(1, STRECODE)

        '' SET ALL VARIABLES TO NOTHING(nulls) FOR NEW EMPLOYEE
        Call RetAllVars
        msfProcess.TextMatrix(0, 1) = STRECODE
        msfProcess.Refresh
        '' MAKING LVT TABLE EMPTY AND REFILLING WITH CURRENT EMPCODE AND DATE
        Call TruncateTable("LVT")
        ConMain.Execute "insert into lvt (Empcode,lst_date) values (" & _
        "'" & STRECODE & "'" & "," & strDTEnc & Format(typMnlVar.strLvtDate, "dd/mmm/yyyy") & _
        strDTEnc & ")"

        Call GetStatusArr(STRECODE)     '' FILLING THE STATUS ARRAY FOR CURRENT EMPLOYEE
        '' Code to Get the Array Back
        '' Code to Get OT Rates for specific employee from OT Rule Master
        Call GetOTRates
        '' CALCULATION OF ABS,PRS,WOS,HLS,UNPDLV,PDLV
        If Not TotalCalc(STRECODE, strCatCode) Then monProcess = False: Exit Function
        
        ''If sngArrTotl(0) = 100 Then monProcess = False: Exit Function
        ''
        Do While (STRECODE = adrsTemp!Empcode And Not adrsTemp.EOF)
            '' NIGHT SHIFT CALCULATION
            If adrsTemp!Shift <> "" Then
                If GetCode(2, adrsTemp("Shift")) <> "0" And adrsTemp("Entry") > 0 Then bytNight = bytNight + 1
            End If
            '' COMP OFF CALCULATION
            '' NO OF LATE CALCULATION AND LATEHRS CALCULATION
            If (Trim(adrsTemp!presabs) = Trim(pVStar.PrsCode & pVStar.PrsCode) Or _
            Trim(adrsTemp!presabs) = Trim(pVStar.PrsCode & pVStar.AbsCode)) And _
            adrsTemp!latehrs > 0 And (IsNull(adrsTemp!aflg) Or adrsTemp!aflg = "") Then

                    bytLtNo = bytLtNo + 1

                sngLtHrs = TimAdd(sngLtHrs, adrsTemp!latehrs)
            End If

            '' NO OF EARLY CALCULATION AND EARLHRS CALCULATION
            If (Trim(adrsTemp!presabs) = Trim(pVStar.PrsCode & pVStar.PrsCode) Or _
            Trim(adrsTemp!presabs) = Trim(pVStar.AbsCode & pVStar.PrsCode)) And _
            adrsTemp!earlhrs > 0 And (IsNull(adrsTemp!Dflg) Or adrsTemp!Dflg = "") Then

                    bytErNo = bytErNo + 1

                sngErHrs = TimAdd(sngErHrs, adrsTemp!earlhrs)
            End If
            '' WORKHRS CALCULATION
            If Not IsNull(adrsTemp!wrkHrs) Then sngWrkHrs = TimAdd(sngWrkHrs, adrsTemp!wrkHrs)
            '' OVERTIME CALCULATION
            If adrsTemp!ovtim > 0 And adrsTemp("OTConf") = "Y" Then
                sngOtHrs = TimAdd(sngOtHrs, adrsTemp!ovtim)
                Select Case adrsTemp!presabs
                    Case pVStar.WosCode & pVStar.WosCode
                        sngOtH_pd = TimAdd(sngOtH_pd, dec2Hrs(hrs2Dec(adrsTemp!ovtim) _
                        * sngIns_WOOt))
                    Case pVStar.HlsCode & pVStar.HlsCode
                        sngOtH_pd = TimAdd(sngOtH_pd, dec2Hrs(hrs2Dec(adrsTemp!ovtim) _
                        * sngIns_HLOt))
                    Case Else
                        sngOtH_pd = TimAdd(sngOtH_pd, dec2Hrs(hrs2Dec(adrsTemp!ovtim) _
                        * sngIns_OTOt))
                End Select
            End If
     
            
            adrsTemp.MoveNext
            If adrsTemp.EOF Then Exit Do
        Loop
        
        sngTotLates = bytLtNo
        sngTotEarls = bytErNo
        
        If typMnlVar.bytExeLE = 1 Then '' IF REQUIRED THEN EXECUTE LATE/EARLY RULES
            ''sngPd_Dys = sngPd_Dys + MainForm.SC1.Run("LateEarlchk", strEcode, strCatCode, _
            DateSaveIns(typMnlVar.strLvtDate), strLvBal, strLvInfo, strDTEnc, sngTotLates, sngTotEarls)
            sngPd_Dys = sngPd_Dys + LateEarlChk1(STRECODE, strCatCode, _
            DateSaveIns(typMnlVar.strLvtDate), strLvBal, strLvInfo, strDTEnc, sngTotLates, sngTotEarls)
        End If

        '' PAID DAYS= PAID DAYS + NO OF PRESENT + NO OF WEEKOFF + NO OF HOLIDAY + NO OF PAID LEAVE
        RsWeekOffPaid.MoveFirst                            '  27-12
        RsWeekOffPaid.Find "EmpCode = '" & STRECODE & "'"
        If RsWeekOffPaid("WeekOffPaid") = "N" Then
            sngPd_Dys = sngPd_Dys + sngArrTotl(1) + sngArrTotl(3) + sngArrTotl(5)
            If sngArrTotl(1) = 0 And GetFlagStatus("WOHLABS") Then
                sngPd_Dys = 0
            End If
        Else

    
         sngPd_Dys = sngPd_Dys + sngArrTotl(1) + sngArrTotl(2) + sngArrTotl(3) + sngArrTotl(5)
     
        End If
               
        '' Add appropriate paid days and presents in final total
        If InVar.blnAssum = "1" Then
             Call AddDelAsume
        End If
        '' UPDATE LVT WITH NEW VALUES
        
        Dim strTmpTrnFld As String
        strTmpTrnFld = ""

        ConMain.Execute " update lvt set paiddays = " & sngPd_Dys & _
        ",ot_hrs= " & sngOtHrs & ", otpd_hrs =" & sngOtH_pd & ",lt_no=" & bytLtNo & _
        ",lt_hrs=" & sngLtHrs & ",erl_no=" & bytErNo & ",erl_hrs=" & sngErHrs & _
        ",wrk_hrs=" & sngWrkHrs & ",night=" & bytNight & "," & pVStar.AbsCode & _
        "=" & sngArrTotl(0) & "," & pVStar.PrsCode & "=" & sngArrTotl(1) & _
        "," & pVStar.WosCode & "=" & sngArrTotl(2) & "," & pVStar.HlsCode & _
        "=" & sngArrTotl(3) & strTmpTrnFld & " where Empcode = '" & STRECODE & "'"
        
        If bytRepMode = 6 And typOptIdx.bytPer = 7 Then
            ConMain.Execute "insert into prAbsentLvt select * from lvt"
        Else
            ConMain.Execute "delete  from " & strLvTrn & " where lst_date=" & _
            strDTEnc & Format(typMnlVar.strLvtDate, "dd/mmmm/yyyy") & strDTEnc & " and Empcode='" & STRECODE & "'", i
            '' COPY LVT INTO LVTRN FOR THE CURRENT EMPLOYEE?
            ConMain.Execute "insert into " & strLvTrn & " select * from lvt"
        End If
    Loop
End If
Exit Function
ERR_P:
    ShowError ("MonProcess :: " & Me.Caption)
    monProcess = False
    Resume Next
End Function


Private Function GetLvtTable() As Boolean
On Error GoTo ERR_P
If FindTable("lvt") Then ConMain.Execute "drop table lvt"
''conmain.Execute "select * into lvt from " & strLvTrn & " where 1=2"
Call CreateTableIntoAs("*", strLvTrn, "lvt", " Where 1=2 ")
GetLvtTable = True
Exit Function
ERR_P:
    ShowError ("Get Leave Table :: " & Me.Caption)
End Function

Private Function OpenRSets() As Boolean
On Error GoTo ERR_P
If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "select Empcode,cat from empmst where Empcode in (" & typMnlVar.strEmpList & _
")", ConMain, adOpenStatic

If adRsInstall.State = 1 Then adRsInstall.Close
adRsInstall.Open "select shift,night from instshft where shift <> '100'", ConMain, adOpenStatic

If adRsintshft.State = 1 Then adRsintshft.Close
adRsintshft.Open "select * from lateerl where (" & strKDate & " between  " & strDTEnc & _
Format(typMnlVar.strFrtDate, "dd/mmm/yyyy") & strDTEnc & " and " & strDTEnc & _
Format(typMnlVar.strLstDate, "dd/mmm/yyyy") & strDTEnc & ") and (latehrs>0 or earlhrs>0) AND EMPCODE" & _
" IN (" & typMnlVar.strEmpList & ")", ConMain, adOpenStatic
OpenRSets = True
Exit Function
ERR_P:
    ShowError ("Open Record Sets :: " & Me.Caption)
    'Resume
End Function

Private Function GetTrnRs() As Boolean
On Error GoTo ERR_P
If InVar.blnAssum = "1" Then
    Call GetTrnRsAssume
Else
    Call GetTrnRsActual
End If
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strquery, ConMain, adOpenStatic
GetTrnRs = True
Exit Function
ERR_P:
    ShowError ("Get Trn Record Sets:: " & Me.Caption)
End Function

Private Sub GetTrnRsAssume()
On Error GoTo ERR_P
If strMonTrnPrev = strMonTrnCurr Then
    If bytDate = 0 Then         ''  For Entire Current month
        bytCase = 1
        Call SetQuery(bytCase)
    Else
        bytCase = 3             ''  For Current month but upto specified date
        Call SetQuery(bytCase)
    End If
Else                            ''  From Previous month's some specified date to-
    bytCase = 2                 ''  -Current month's specified date
    Call SetQuery(bytCase)
End If
Exit Sub
ERR_P:
    ShowError ("Get Trn Record Sets:: " & Me.Caption)
End Sub

Private Sub GetTrnRsActual()
On Error GoTo ERR_P
If strMonTrnPrev = strMonTrnCurr Then   ''From & To Dates fall in the same current Month
    bytCase = 4
    Call SetQuery(bytCase)
Else                                    ''  From Previous month's some specified date to-
    bytCase = 2                         ''  -Current month's specified date
    Call SetQuery(bytCase)
End If
Exit Sub
ERR_P:
    ShowError ("Get Trn Record Sets:: " & Me.Caption)
End Sub

Private Function GetCode(ByVal bytRSet As Byte, ByVal strIn As String, Optional strDate As String) As String
On Error GoTo ERR_P     ''  This function returns required value from any specified Recordset.
GetCode = ""
Select Case bytRSet
    Case 1          '' GET CATCODE FROM EMPMST FOR EMPCODE
        GetCode = "Nothing"
        adrsEmp.MoveFirst
        adrsEmp.Find "Empcode = '" & strIn & "'"
        GetCode = adrsEmp!cat
    Case 2          '' GET NIGHT FLAG FROM INSTSHFT FOR SHIFT
        GetCode = "0"
        adRsInstall.MoveFirst
        adRsInstall.Find "shift = '" & strIn & "'"
        If Not adRsInstall.EOF Then
            GetCode = IIf(adRsInstall!night = 0 Or adRsInstall!night = False, "0", "1")
        Else
            GetCode = "0"
        End If
    Case 3          '' GET LATE FROM LATEEARL TABLE
        GetCode = "nothing"
        If Not (adRsintshft.BOF And adRsintshft.EOF) Then
            adRsintshft.MoveFirst
            Do While Not adRsintshft.EOF
                If adRsintshft!Empcode = strIn And adRsintshft!Date = DateCompDate(strDate) _
                And adRsintshft!latehrs > 0 Then
                    GetCode = "1"
                    Exit Do
                End If
                adRsintshft.MoveNext
            Loop
        End If
    Case 4          '' GET EARLY FROM LATEEARL TABLE
        GetCode = "nothing"
        If Not (adRsintshft.BOF And adRsintshft.EOF) Then
            adRsintshft.MoveFirst
            Do While Not adRsintshft.EOF
                If adRsintshft!Empcode = strIn And adRsintshft!Date = DateCompDate(strDate) _
                And adRsintshft!earlhrs > 0 Then
                    GetCode = "1"
                    Exit Do
                End If
                adRsintshft.MoveNext
            Loop
        End If
End Select
Exit Function
ERR_P:
    ShowError ("GetCode :: " & Me.Caption)
End Function

'' The following function is Heart of Monthly Process
'' Do take care while attempting any change to this function.
'' This function calculates Total no. of Presents, Absents, Weekoffs, Holidays ,
'' -Paid and Unpaid Leaves.

Private Function TotalCalc(ByVal StrEmpCd As String, ByVal strEmpCat As String) As Boolean
On Error GoTo ERR_P
TotalCalc = True
Dim bytArrcnt As Byte, bytPoArCnt As Byte
Dim bytArrPos(31) As Byte
Dim lastday As Byte
''Call GetStatusArr(StrEmpCd)     '' FILLING THE STATUS ARRAY FOR CURRENT EMPLOYEE
bytPoArCnt = 0
For bytArrcnt = 1 To UBound(strArrStatus)
    If strArrStatus(bytArrcnt) = "" Then
        lastday = bytArrcnt - 1
        Exit For
    End If
    Select Case Left(strArrStatus(bytArrcnt), 2)
        Case pVStar.AbsCode                         '' SNGARRTOTL ARRAY POSITIONS
            sngArrTotl(0) = sngArrTotl(0) + 0.5     '' 0  Absent
        Case pVStar.PrsCode
            sngArrTotl(1) = sngArrTotl(1) + 0.5     '' 1  Present
        Case pVStar.WosCode
            sngArrTotl(2) = sngArrTotl(2) + 0.5     '' 2  Week Offs
        Case pVStar.HlsCode
            sngArrTotl(3) = sngArrTotl(3) + 0.5     '' 3  Holidays
        Case Else
            Select Case LvCalc(Left(strArrStatus(bytArrcnt), 2), strEmpCat)
                Case 0                                  ''If Leave is not found in LvTrn
                    TotalCalc = False
                    Exit Function
                Case 1
                    sngArrTotl(0) = sngArrTotl(0) + 0.5 '' 0  Absent
                Case 2
                    sngArrTotl(4) = sngArrTotl(4) + 0.5 '' 4  Unpaid Leaves
                Case 3
                    sngArrTotl(5) = sngArrTotl(5) + 0.5 '' 5  Paid Leaves
'                    If SubLeaveFlag = 1 And Left(strArrStatus(bytArrcnt), 2) = "CM" Then  ' 15-10
'                        sngArrTotl(5) = sngArrTotl(5) + 0.5
'                    End If
            End Select
    End Select
    Select Case Right(strArrStatus(bytArrcnt), 2)
        Case pVStar.AbsCode
            sngArrTotl(0) = sngArrTotl(0) + 0.5
        Case pVStar.PrsCode
            sngArrTotl(1) = sngArrTotl(1) + 0.5
        Case pVStar.WosCode
            sngArrTotl(2) = sngArrTotl(2) + 0.5
        Case pVStar.HlsCode
            sngArrTotl(3) = sngArrTotl(3) + 0.5
        Case Else
            Select Case LvCalc(Right(strArrStatus(bytArrcnt), 2), strEmpCat)
                Case 0
                    TotalCalc = False
                    Exit Function
                Case 1                              '' 0  Absent
                    sngArrTotl(0) = sngArrTotl(0) + 0.5
                Case 2                              '' 4  Unpaid Leaves
                    sngArrTotl(4) = sngArrTotl(4) + 0.5
                Case 3                              '' 5  Paid Leaves
                    sngArrTotl(5) = sngArrTotl(5) + 0.5
'                    If SubLeaveFlag = 1 And Left(strArrStatus(bytArrcnt), 2) = "CM" Then  ' 15-10
'                        sngArrTotl(5) = sngArrTotl(5) + 0.5
'                    End If
            End Select
    
    End Select
    If strArrStatus(bytArrcnt) = pVStar.WosCode & pVStar.WosCode Or _
    strArrStatus(bytArrcnt) = pVStar.HlsCode & pVStar.HlsCode Then
        bytArrPos(bytPoArCnt) = bytArrcnt
        bytPoArCnt = bytPoArCnt + 1
    End If
   
Next

If GetFlagStatus("WOABS") Then

    Dim adrsMonTrn As New ADODB.Recordset
    Dim ArrStatus(31) As String
    Dim ArrIn(31) As Integer
    Dim x As Integer
    If adrsMonTrn.State = 1 Then adrsMonTrn.Close
    Call SetQuery(10, StrEmpCd)
    adrsMonTrn.Open strquery, ConMain, adOpenKeyset, adLockReadOnly
    If Not (adrsMonTrn.EOF And adrsMonTrn.BOF) Then
        adrsMonTrn.MoveFirst
        x = 0
        Do While Not adrsMonTrn.EOF
            ArrStatus(x) = adrsMonTrn!presabs
            ArrIn(x) = adrsMonTrn!arrtim + adrsMonTrn!deptim
            x = x + 1
            adrsMonTrn.MoveNext
            If adrsMonTrn.EOF Then Exit Do
        Loop
    End If
    
        For x = 2 To UBound(ArrStatus)
            If ArrStatus(x) = pVStar.WosCode & pVStar.WosCode And ArrIn(x) = 0 Then
                If ArrStatus(x - 1) = pVStar.AbsCode & pVStar.AbsCode Or ArrStatus(x - 2) = pVStar.AbsCode & pVStar.AbsCode Then
                    sngArrTotl(2) = sngArrTotl(2) - 1
                    sngArrTotl(0) = sngArrTotl(0) + 1
                End If
             ElseIf ArrStatus(x) = pVStar.WosCode & pVStar.WosCode And ArrIn(x) <> 0 Then
                    sngArrTotl(1) = sngArrTotl(1) + 1
            End If
        Next
End If



    If InVar.blnAssum = "1" Then    'for month and year option
        Dim Fdate As String, CrDate As String
        CrDate = DateCompStr(FdtLdt(MonthNumber(cmbMonth.Text), cmbYear.Text, "L"))
        Call FullCrLeave("Monthly", StrEmpCd, strEmpCat, CrDate, CrDate, sngArrTotl(1))
        'Call FullCrLeave("Monthly", StrEmpCd, strEmpCat, CrDate, CrDate, sngArrTotl(1))
    Else        'For from and to date
        Fdate = DateCompStr(txtToDate.Text)
        CrDate = DateCompStr(FdtLdt(Month(txtToDate.Text), Year(txtToDate.Text), "L"))
        Call FullCrLeave("Monthly", StrEmpCd, strEmpCat, Fdate, CrDate, sngArrTotl(1))
    End If


Dim bytPcnt As Byte, bytInCnt As Byte
Dim bytStart As Byte, bytPCount As Byte
Dim bytNext As Byte
Dim pcnt As Integer

If bytArrPos(0) = 1 And InVar.blnAssum <> "1" Then
    Dim adrsTemp As New ADODB.Recordset
    Dim dte As Date
    Dim strtrn As String
    dte = DateAdd("D", -1, DateCompDate(txtFrDate.Text))
    If Month(dte) <> Month(DateCompDate(txtFrDate.Text)) Then
        strtrn = MakeName(MonthName(Month(dte)), Year(dte), "trn")
        If Not FindTable(strtrn) Then
            GoTo Ex
        End If
    Else
        strtrn = strMonTrnPrev
    End If

    strquery = "SELECT presabs,wrkhrs FROM " & strMonTrnPrev & " WHERE empcode='" & StrEmpCd & "' And " & strKDate & "=" & strDTEnc & "" & Format(dte, "dd/mmm/yyyy") & "" & strDTEnc & ""
    
    adrsTemp.Open strquery, ConMain, adOpenStatic, adLockOptimistic
        
    If Not (adrsTemp.EOF And adrsTemp.BOF) Then
        If FilterNull(adrsTemp.Fields("presabs")) = pVStar.AbsCode & _
            pVStar.AbsCode Then
            Dim bytTemp As Byte
            For bytTemp = 0 To UBound(bytArrPos)
                If Not bytArrPos(bytTemp) = bytTemp + 1 Then
                    sngArrTotl(0) = sngArrTotl(0) + bytTemp
                    Exit For
                Else
                    Select Case strArrStatus(bytTemp + 1)
                        Case pVStar.WosCode & pVStar.WosCode
                            sngArrTotl(2) = sngArrTotl(2) - 1
                        Case pVStar.HlsCode & pVStar.HlsCode
                            sngArrTotl(3) = sngArrTotl(3) - 1
                    End Select
                End If
            Next
        End If
    End If
End If
Ex:
''end by
bytNext = 0
bytStart = bytArrPos(0)
bytInCnt = 0: bytPCount = 0
For bytPcnt = 0 To UBound(bytArrPos)
    If bytArrPos(bytPcnt) = 0 Then Exit For
    If CInt(bytArrPos(bytPcnt + 1)) - CInt(bytStart) < CInt((bytPcnt + 1)) _
        - CInt(bytInCnt) Then
        If bytArrPos(bytPcnt) + 1 > UBound(strArrStatus) Then Exit For
            If bytStart < 1 Then bytNext = 1
            If bytNext = 0 Then
                If Not GetFlagStatus("PrimacyHLPaid") Then
                If strArrStatus(bytStart - 1) = strArrStatus(bytArrPos(bytPcnt) + 1) Then
                    Select Case strArrStatus(bytStart - 1)
                        Case pVStar.AbsCode & pVStar.AbsCode
                            If RsWeekOffPaid("WeekOffPaid") = "Y" Then
                                If Not GetFlagStatus("WOABS") Then
                                    sngArrTotl(0) = sngArrTotl(0) + (bytPcnt + 1) - bytInCnt
                                End If
                            End If
                    End Select
                    If strArrStatus(bytStart - 1) = pVStar.AbsCode & pVStar.AbsCode Then
                        For bytPCount = bytStart - 1 To bytArrPos(bytPcnt) + 1
                            Select Case strArrStatus(bytPCount)
                                Case pVStar.WosCode & pVStar.WosCode
                                    If RsWeekOffPaid("WeekOffPaid") = "Y" Then
                                        If Not GetFlagStatus("WOABS") Then
                                            sngArrTotl(2) = sngArrTotl(2) - 1
                                        End If
                                    End If
                                Case pVStar.HlsCode & pVStar.HlsCode
                                    sngArrTotl(3) = sngArrTotl(3) - 1
                             End Select
                        Next
                    End If
                End If
                bytStart = bytArrPos(bytPcnt + 1)
                bytInCnt = bytPcnt + 1
                Else

                    If strArrStatus(bytStart - 1) <> "" And (strArrStatus(bytStart - 1) <> pVStar.PrsCode & pVStar.PrsCode Or strArrStatus(bytArrPos(bytPcnt) + 1) <> pVStar.PrsCode & pVStar.PrsCode) Then
                        If strArrStatus(bytStart - 1) <> pVStar.PrsCode & pVStar.PrsCode Then
                            For bytPCount = bytStart - 1 To bytArrPos(bytPcnt) + 1
                                If strArrStatus(bytPCount + 1) <> pVStar.PrsCode & pVStar.PrsCode Then
                                        sngArrTotl(0) = sngArrTotl(0) + 1
                                        sngArrTotl(2) = sngArrTotl(2) - 1
                                        Exit For
                                End If
                            Next
                        End If
                    End If

                    bytStart = bytArrPos(bytPcnt + 1)
                    bytInCnt = bytPcnt + 1
                End If
            End If '' bytnext
        bytNext = 0
    Else
        If bytArrPos(bytPcnt + 1) - bytStart <> (bytPcnt + 1) - bytInCnt Then
            If bytArrPos(bytPcnt) + 1 > UBound(strArrStatus) Then Exit For
            If bytStart < 1 Then bytNext = 1
            If bytNext = 0 Then
                If Not GetFlagStatus("PrimacyHLPaid") Then
                If strArrStatus(bytStart - 1) = strArrStatus(bytArrPos(bytPcnt) + 1) Then
                    Select Case strArrStatus(bytStart - 1)
                        Case pVStar.AbsCode & pVStar.AbsCode
                             If RsWeekOffPaid("WeekOffPaid") = "Y" Then
                                If Not GetFlagStatus("WOABS") Then
                                    sngArrTotl(0) = sngArrTotl(0) + (bytPcnt + 1) - bytInCnt
                                End If
                            End If
                    End Select
                    If strArrStatus(bytStart - 1) = pVStar.AbsCode & pVStar.AbsCode Then
                        For bytPCount = bytStart - 1 To bytArrPos(bytPcnt) + 1
                            Select Case strArrStatus(bytPCount)
                                Case pVStar.WosCode & pVStar.WosCode
                                    If RsWeekOffPaid("WeekOffPaid") = "Y" Then
                                        If Not GetFlagStatus("WOABS") Then
                                            sngArrTotl(2) = sngArrTotl(2) - 1
                                        End If
                                    End If
                                Case pVStar.HlsCode & pVStar.HlsCode
                                     sngArrTotl(3) = sngArrTotl(3) - 1
                            End Select
                        Next
                    End If
                End If
                bytStart = bytArrPos(bytPcnt + 1)
                bytInCnt = bytPcnt + 1
                Else
                
                    If strArrStatus(bytStart - 1) <> "" And (strArrStatus(bytStart - 1) <> pVStar.PrsCode & pVStar.PrsCode Or strArrStatus(bytArrPos(bytPcnt) + 1) <> pVStar.PrsCode & pVStar.PrsCode) Then
                        If strArrStatus(bytStart - 1) <> pVStar.PrsCode & pVStar.PrsCode Then
                            For bytPCount = bytStart - 1 To bytArrPos(bytPcnt) + 1
                                If strArrStatus(bytStart + 1) <> pVStar.PrsCode & pVStar.PrsCode And strArrStatus(bytStart) = pVStar.WosCode & pVStar.WosCode Then
                                        sngArrTotl(0) = sngArrTotl(0) + 1
                                        sngArrTotl(2) = sngArrTotl(2) - 1
                                        Exit For
                                End If
                            Next
                        End If
                    End If
                    bytStart = bytArrPos(bytPcnt + 1)
                    bytInCnt = bytPcnt + 1
                
                End If
            End If '' bytnext
        bytNext = 0
        End If
    End If
Next


Exit Function
ERR_P:
    ShowError ("Total Calc :: " & Me.Caption)
    TotalCalc = False
    Resume Next
End Function

Private Sub GetStatusArr(ByVal StrEmpCd As String)  ''Retrives presabs into array named
On Error GoTo ERR_P                                 ''strArrStatus
Dim bytArrcnt As Byte
'Dim sysdate As String
Dim adrsMonTrn As New ADODB.Recordset
If adrsMonTrn.State = 1 Then adrsMonTrn.Close
Call SetQuery(bytCase + 5, StrEmpCd)
adrsMonTrn.Open strquery, ConMain, adOpenKeyset, adLockReadOnly
If Not (adrsMonTrn.EOF And adrsMonTrn.BOF) Then
    bytArrcnt = 1
    Do While Not adrsMonTrn.EOF
        strArrStatus(bytArrcnt) = adrsMonTrn!presabs
        bytArrcnt = bytArrcnt + 1
        adrsMonTrn.MoveNext
        If adrsMonTrn.EOF Then Exit Do
    Loop
End If

Exit Sub
ERR_P:
    ShowError ("Get Status Array :: " & Me.Caption)
    'Resume Next
End Sub

Public Function LvCalc(ByVal strLeave As String, ByVal strEmpCat As String) As Byte
On Error GoTo ERR_P
Dim TempAdrs As New ADODB.Recordset
If Not IsFieldThere(Trim(strLeave), strFields) Then
    LvCalc = 0 '' IF LEAVE NOT FOUND IN LVTRN TERMINATE MONTHLY PROCESS
    MsgBox strLeave & NewCaptionTxt("36022", adrsC) & _
    vbCrLf & vbTab & NewCaptionTxt("36023", adrsC), vbInformation
    'Exit Function
End If
If TempAdrs.State = 1 Then TempAdrs.Close
TempAdrs.Open " select paid from leavdesc where lvcode = '" & strLeave & _
"' and cat = '" & strEmpCat & "'", ConMain, adOpenStatic
If TempAdrs.EOF And TempAdrs.BOF Then
    LvCalc = 1 '' IF LEAVE NOT FOUND IN LEAVEDESC ADD 0.5 EMPLOYEE'S ABSENT TOTAL
    Exit Function
Else
    If TempAdrs!paid = "Y" Then
        LvCalc = 3  '' ADD TO PAID LEAVE
    Else
        LvCalc = 2  '' ADD TO UNPAID LEAVE
    End If
    
    ConMain.Execute "update lvt set " & strLeave & "=" & _
    TrnLeaves(strLeave) & " + 0.5 "
    If SubLeaveFlag = 1 And strLeave = "CM" Then  ' 15-10
        ConMain.Execute "update lvt set " & strLeave & "=" & TrnLeaves(strLeave) & " + 0.5 "
    End If
End If
Exit Function
ERR_P:
    ShowError ("Leave Calc :: " & Me.Caption)
End Function

Public Function TrnLeaves(colName As String) As Single
On Error GoTo ERR_P
Dim adrsLeave1 As New ADODB.Recordset
If adrsLeave1.State = 1 Then adrsLeave1.Close
adrsLeave1.Open "select " & colName & " from lvt ", ConMain, adOpenStatic
TrnLeaves = IIf(IsNull(adrsLeave1(0)), 0, adrsLeave1(0))
Exit Function
ERR_P:
    ShowError ("Transaction Leaves :: " & Me.Caption)
End Function


Private Sub RetAllVars()
sngTotLates = 0: sngTotEarls = 0
sngPd_Dys = 0
bytNight = 0: bytLtNo = 0: bytErNo = 0
sngLtHrs = 0: sngErHrs = 0
sngWrkHrs = 0: sngOtHrs = 0: sngOtH_pd = 0

For sngTotLates = 0 To UBound(sngArrTotl)
    sngArrTotl(sngTotLates) = 0
Next
sngTotLates = 0
ReDim strArrStatus(0 To 62)

sngTotLunchLates = 0: sngLunchLtHrs = 0: bytLunchLtNo = 0
End Sub


Private Sub AddDelAsume()
On Error GoTo ERR_P
Dim bytVal1 As Byte, sngVal2 As Single, dtTmpDate As Date

If bytDate = 0 Then
Else    ''Adding specified No of Paid days and Presents if To date is not EOM
    bytVal1 = Day(DateCompDate(typMnlVar.strLvtDate))
    'bytDate this is cutoff days
    bytVal1 = IIf((bytVal1 - bytDate) < 0, 0, (bytVal1 - bytDate))
    sngPd_Dys = sngPd_Dys + bytVal1          ''Add to Paid Days
    sngArrTotl(1) = sngArrTotl(1) + bytVal1  ''Add to Presents
    ''Deducting specified No of Paid days and Presents if previous month's data
    ''- is also processed
    If chkPrev.Value = 1 Then
        If txtFDate.Text = 0 Then Exit Sub
            dtTmpDate = DateAdd("M", -1, DateCompDate(typMnlVar.strLvtDate))
            sngVal2 = Day(FdtLdt(Month(dtTmpDate), Year(dtTmpDate), "L"))
            sngVal2 = sngVal2 - Val(txtFDate.Text) + 1
            sngPd_Dys = sngPd_Dys - sngVal2          ''Deduct from Paid Days
            sngArrTotl(1) = sngArrTotl(1) - sngVal2  ''Deduct from Presents
    End If
End If
Exit Sub
ERR_P:
ShowError ("AddAsume  :: " & Me.Caption)
End Sub

Private Sub cmdExit_Click()
If bytRepMode = 6 And typOptIdx.bytPer = 7 Then PerAtt = True
Unload Me
End Sub

Private Function SetQuery(ByVal bytOpt As Byte, Optional ByVal STRECODE As String)
On Error GoTo ERR_P
Dim strdate1 As String, strDate2 As String, strdate3 As String
Dim strEmp As String, strOrder As String

strEmp = " Empcode in (" & typMnlVar.strEmpList & ") "
strOrder = " order by Empcode," & strKDate & ""
strdate1 = " " & strMonTrnPrev & "." & strKDate & " >= " & strDTEnc & _
            Format(typMnlVar.strFrtDate, "dd/mmm/yyyy") & strDTEnc & " "
strDate2 = " " & strMonTrnCurr & "." & strKDate & " <= " & strDTEnc & _
            Format(typMnlVar.strLstDate, "dd/mmm/yyyy") & strDTEnc & " "
strdate3 = " " & strMonTrnCurr & "." & strKDate & " >= " & strDTEnc & Format(typMnlVar.strFrtDate, "dd/mmm/yyyy") & _
           strDTEnc & " and " & strMonTrnCurr & "." & strKDate & "<= " & strDTEnc & _
           Format(typMnlVar.strLstDate, "dd/mmm/yyyy") & strDTEnc
Select Case bytOpt
    Case 1  ''Select all Employee for entire current month
        strquery = "Select * from " & strMonTrnCurr & " Where " & strEmp & strOrder & " "
    Case 2  ''Select for previous month and current month
        strquery = "Select * from " & strMonTrnPrev & " where " & strdate1 & " and " & strEmp & _
        " union select *  from " & strMonTrnCurr & " where " & strDate2 & " and " & strEmp & strOrder & " "
    Case 3  ''Select for current month but upto specified date
        strquery = "Select * from " & strMonTrnCurr & " where " & strDate2 & " and " & strEmp & strOrder & ""
    Case 4  ''Select for current month's specified dates(Actual)
        strquery = "Select * from " & strMonTrnCurr & " where " & strdate3 & " and " & strEmp & strOrder & ""
''To Fill Array(strArrStatus) of Presabs
    Case 6  ''Select PresAbs for entire current month
        strquery = "Select presabs," & strKDate & " from " & strMonTrnCurr & " where Empcode = '" & _
            STRECODE & "' order by " & strKDate & ""
    Case 7  ''Select PresAbs for for previous month and current month
        strquery = "Select presabs," & strKDate & " from " & strMonTrnPrev & " where " & strdate1 & _
            " and Empcode = '" & STRECODE & "' union select presabs," & strKDate & " from " & _
            strMonTrnCurr & " where " & strDate2 & " and Empcode = '" & STRECODE & "' order by " & strKDate & ""
    Case 8  ''Select PresAbs for current month but upto specified date
        strquery = "Select presabs," & strKDate & " from " & strMonTrnCurr & " where " & strDate2 & _
            " and Empcode = '" & STRECODE & "' order by " & strKDate & ""
    Case 9  ''Select for current month's specified dates(Actual)
        strquery = "Select presabs," & strKDate & " from " & strMonTrnCurr & " where " & strdate3 & _
            " and Empcode = '" & STRECODE & "' order by " & strKDate & ""
    Case 10
        strquery = "Select * from " & strMonTrnCurr & " where Empcode = '" & STRECODE & "' order by " & strKDate & ""
End Select
Exit Function
ERR_P:
ShowError ("SetQuery :: " & Me.Caption)
strquery = ""
End Function

Private Sub FillField(ByVal strFileName As String)
On Error GoTo ERR_P
Dim adrs As New ADODB.Recordset, intCnt As Integer
If adrs.State = 1 Then adrs.Close
adrs.Open "Select * from  " & strFileName & " where 1=2", ConMain
If adrs.Fields.Count > 0 Then
    ReDim strFields(adrs.Fields.Count)
    For intCnt = 0 To adrs.Fields.Count - 1
        strFields(intCnt) = adrs.Fields(intCnt).name
    Next
End If
Exit Sub
ERR_P:
    ShowError ("FillFiled :: " & Me.Caption)
End Sub

Private Function IsFieldThere(ByVal strFieldName As String, strArrFields) As Boolean
On Error GoTo ERR_P
Dim bytCnt As Byte
IsFieldThere = False
For bytCnt = 0 To UBound(strArrFields)
    If strArrFields(bytCnt) = strFieldName Then IsFieldThere = True: Exit Function
Next
Exit Function
ERR_P:
IsFieldThere = False
End Function

Private Sub GetOTRates()
On Error GoTo ERR_P
Dim adrsOTRule As New ADODB.Recordset
adrsOTRule.Open "Select WDRates ,WORates,HLRates from OtRul where otcode = " & _
"(select otcode from empmst where Empcode = '" & STRECODE & "')", ConMain, adOpenStatic, adLockOptimistic
If Not (adrsOTRule.EOF And adrsOTRule.BOF) Then
    sngIns_OTOt = IIf(IsNull(adrsOTRule(0)), 0, adrsOTRule(0))
    sngIns_WOOt = IIf(IsNull(adrsOTRule(1)), 0, adrsOTRule(1))
    sngIns_HLOt = IIf(IsNull(adrsOTRule(2)), 0, adrsOTRule(2))
Else
    sngIns_OTOt = 1
    sngIns_WOOt = 1
    sngIns_HLOt = 1
End If
Exit Sub
ERR_P:
MsgBox Err.Description
End Sub

Private Sub MSF1_Click()
If bytRepMode = 6 And typOptIdx.bytPer = 7 Then Exit Sub
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
Function LateEarlChk1(ECode, Cat_code, strLvtDate, strLvBal, strLvInfo, strDTEnc, _
  sngTotLates, sngTotEarls)
''On Error Resume Next
Dim lateRul, erlRul, erMth, LtMth
Dim evLate, evEarl, Late_cut, Late_Earl
Dim DedLv_pd_late, DedLv_pd_erl
Dim fstlv_erl, SecLv_erl, TrdLv_erl
Dim FstLv, SecLv, TrdLv
Dim DaysToCut_Lt, DaysToCut_Erl
Dim QueryStr, tempvar, sng_PDDAy
Dim adrslv, AdrsCat, cnDJConn
Set cnDJConn = ConMain

Set adrslv = CreateObject("ADODB.RECORDSET")
Set AdrsCat = CreateObject("ADODB.RECORDSET")

If AdrsCat.State = 1 Then AdrsCat.Close
AdrsCat.Open "select laterule,earlrule,ltinmnth,erinmnth,letcut,erlcut,everlet," & _
"evererl,dedlet,fstletpr,secletpr,trdletpr,fsterlpr,secerlpr,trderlpr,dederl " & _
"from catdesc where cat= " & "'" & Cat_code & "'", ConMain, adOpenKeyset
If Not (AdrsCat.EOF And AdrsCat.BOF) Then
    If IsNull(AdrsCat("laterule")) Then:    lateRul = 0: Else: lateRul = AdrsCat("laterule")
    If IsNull(AdrsCat("earlrule")) Then: erlRul = 0: Else: erlRul = AdrsCat("earlrule")
    If IsNull(AdrsCat("erinmnth")) Then: erMth = 0: Else: erMth = AdrsCat("erinmnth")
    If IsNull(AdrsCat("ltinmnth")) Then: LtMth = 0: Else: LtMth = AdrsCat("ltinmnth")
    If IsNull(AdrsCat("everlet")) Then: evLate = 0: Else: evLate = AdrsCat("everlet")
    If IsNull(AdrsCat("evererl")) Then: evEarl = 0: Else: evEarl = AdrsCat("evererl")
    If IsNull(AdrsCat("letcut")) Then: Late_cut = 0: Else: Late_cut = AdrsCat("letcut")
    If IsNull(AdrsCat("erlCut")) Then: Late_Earl = 0: Else: Late_Earl = AdrsCat("erlCut")
    If IsNull(AdrsCat("fstletpr")) Then: FstLv = 0: Else: FstLv = AdrsCat("fstletpr")
    If IsNull(AdrsCat("secletpr")) Then: SecLv = 0: Else: SecLv = AdrsCat("secletpr")
    If IsNull(AdrsCat("trdletpr")) Then: TrdLv = 0: Else: TrdLv = AdrsCat("trdletpr")
    fstlv_erl = AdrsCat("fsterlpr")
    SecLv_erl = AdrsCat("secerlpr")
    TrdLv_erl = AdrsCat("trderlpr")
    If IsNull(AdrsCat("dedlet")) Then: DedLv_pd_late = 0: Else: DedLv_pd_late = AdrsCat("dedlet")
    If IsNull(AdrsCat("dederl")) Then: DedLv_pd_erl = 0: Else: DedLv_pd_erl = AdrsCat("dederl")
    AdrsCat.Close
    If (lateRul = "Y" And sngTotLates >= LtMth) Or (erlRul = "Y" And sngTotEarls >= erMth) Then
    If Not UCase(GetFlagStatus("CombineLateEarly")) Then  'Fire 2019
        If Not IsNull(LtMth) Then sngTotLates = sngTotLates - LtMth
        If Not IsNull(erMth) Then sngTotEarls = sngTotEarls - erMth
    Else
        sngTotLates = sngTotLates + sngTotEarls - LtMth - erMth
        sngTotEarls = 0
    End If
            If evLate > 0 And sngTotLates >= evLate Then
                DaysToCut_Lt = CInt(Left(CStr(sngTotLates / evLate), 2)) * Late_cut
            ElseIf evLate > 0 And sngTotLates < evLate Then
                DaysToCut_Lt = 0
            End If
            If evEarl > 0 And sngTotEarls >= evEarl Then
                DaysToCut_Erl = CInt(Left(CStr(sngTotEarls / evEarl), 2)) * Late_Earl
            ElseIf evEarl > 0 And sngTotEarls < evEarl Then
                DaysToCut_Erl = 0
            End If

        '' late: trcd=6  ''53
        If DaysToCut_Lt > 0 And lateRul = "Y" Then
            If DedLv_pd_late = "PD" Then sng_PDDAy = sng_PDDAy - DaysToCut_Lt
            If DedLv_pd_late = "LV" Then
                If FstLv <> "" Then QueryStr = FstLv
                If SecLv <> "" And QueryStr <> "" Then QueryStr = QueryStr & "," & SecLv
                If TrdLv <> "" And QueryStr <> "" Then QueryStr = QueryStr & "," & TrdLv
                If QueryStr <> "" Then '********
                    If bytBackEnd = 2 And GetFlagStatus("OPTIONALDSN") Then   ' 19-02-10
                        If FstLv = SecLv And SecLv = TrdLv Then
                            QueryStr = FstLv
                        ElseIf (FstLv = SecLv And SecLv <> TrdLv) Or (FstLv <> SecLv And SecLv = TrdLv) Then
                            If FstLv = "" Or TrdLv = "" Then
                                QueryStr = FstLv & TrdLv
                            Else
                                QueryStr = FstLv & "," & TrdLv
                            End If
                        ElseIf FstLv <> SecLv And FstLv = TrdLv Then
                            QueryStr = FstLv & "," & SecLv
                        Else
                            If TrdLv <> "" Then
                                QueryStr = FstLv & "," & SecLv & "," & TrdLv
                            End If
                        End If
                    End If
                    
                    If adrslv.State = 1 Then adrslv.Close
                    adrslv.Open "select " & QueryStr & "  from " & strLvBal & " where Empcode=" & "'" & _
                    ECode & "'", ConMain, adOpenKeyset, adLockOptimistic
                    
                    '' Begin transaction
                    '' TRCD=6 FOR DEDUCTING LEAVE FROM LEAVBAL WHEN THE CATMST.DEDLET = "LV"
                    
                    If FstLv <> "" Then
                        Call Removal(ECode, strLvtDate, "6", FstLv, strDTEnc, strLvInfo, strLvBal)
                        tempvar = 0
                        adrslv.Requery
                        
                        If adrslv.RecordCount > 0 Then
                            If adrslv(0) > 0 Then
                                Do While Not (adrslv(0) = 0 Or DaysToCut_Lt = 0)
                                    If DaysToCut_Lt >= 1 And adrslv(0) >= 1 Then
                                        ConMain.Execute "Update " & strLvBal & " set " & _
                                        adrslv(0).name & "=" & adrslv(0) & "-1 where Empcode='" & ECode & "'"
                                        DaysToCut_Lt = DaysToCut_Lt - 1
                                        tempvar = tempvar + 1
                                    Else
                                        ConMain.Execute "Update " & strLvBal & " set " & _
                                        adrslv(0).name & "=" & adrslv(0) & "-" & 0.5 & _
                                        " where Empcode='" & _
                                        ECode & "'"
                                        DaysToCut_Lt = DaysToCut_Lt - 0.5
                                        tempvar = tempvar + 0.5
                                    End If
                                    adrslv.Requery
                                Loop
                                cnDJConn.Execute "update " & strLvBal & " set " & _
                                FstLv & "=" & adrslv(0) & " where Empcode=" & "'" & ECode & "'"
                              Call Insertion(ECode, strLvtDate, "6", FstLv, tempvar, strDTEnc, strLvInfo, strLvBal)
                            End If
                        End If
                    End If
                    If adrslv.Fields.Count > 1 Then
                        
                        If DaysToCut_Lt > 0 And SecLv <> "" Then
                            Call Removal(ECode, strLvtDate, "6", adrslv(1).name, strDTEnc, strLvInfo, strLvBal)
                            tempvar = 0
                            adrslv.Requery
                            If Not adrslv.EOF Then
                            If adrslv(1) > 0 Then
                                Do While Not (adrslv(1) = 0 Or DaysToCut_Lt = 0)
                                    If DaysToCut_Lt >= 1 And adrslv(1) >= 1 Then
                                        cnDJConn.Execute "Update " & strLvBal & " set " & _
                                        adrslv(1).name & "=" & adrslv(1) & "-1 where Empcode='" & ECode & "'"
                                        DaysToCut_Lt = DaysToCut_Lt - 1
                                        tempvar = tempvar + 1
                                     Else
                                        cnDJConn.Execute "Update " & strLvBal & " set " & _
                                        adrslv(1).name & "=" & adrslv(1) & "-" & 0.5 & _
                                        " where Empcode='" & ECode & "'"
                                        DaysToCut_Lt = DaysToCut_Lt - 0.5
                                        tempvar = tempvar + 0.5
                                    End If
                                    adrslv.Requery
                                Loop
                                cnDJConn.Execute "update " & strLvBal & " set " & _
                                adrslv(1).name & "=" & adrslv(1) & " where Empcode=" & "'" & ECode & "'"
                                
'
                                Call Insertion(ECode, strLvtDate, "6", adrslv(1).name, tempvar, strDTEnc, strLvInfo, strLvBal)
                            End If
                            End If
                        End If
                    End If
                    If adrslv.Fields.Count > 2 Then
                        
                        If DaysToCut_Lt > 0 And TrdLv <> "" Then
                            Call Removal(ECode, strLvtDate, "6", adrslv(2).name, strDTEnc, strLvInfo, strLvBal)
                            tempvar = 0
                            adrslv.Requery
                            If adrslv(2) > 0 Then
                                Do While Not (adrslv(2) = 0 Or DaysToCut_Lt = 0)
                                    If DaysToCut_Lt >= 1 And adrslv(2) >= 1 Then
                                        cnDJConn.Execute "Update " & strLvBal & " set " & _
                                        adrslv(2).name & "=" & adrslv(2) & "-1 where Empcode='" & _
                                        ECode & "'"
                                        DaysToCut_Lt = DaysToCut_Lt - 1
                                        tempvar = tempvar + 1
                                    Else
                                        cnDJConn.Execute "Update " & strLvBal & " set " & _
                                        adrslv(2).name & "=" & adrslv(2) & "-" & 0.5 & _
                                        " where Empcode='" & _
                                        ECode & "'"
                                        DaysToCut_Lt = DaysToCut_Lt - 0.5
                                        tempvar = tempvar + 0.5
                                    End If
                                    adrslv.Requery
                                Loop
                                cnDJConn.Execute "update " & strLvBal & " set " & _
                                adrslv(2).name & "=" & adrslv(2) & " where Empcode=" & "'" & ECode & "'"
                             Call Insertion(ECode, strLvtDate, "6", adrslv(2).name, tempvar, strDTEnc, strLvInfo, strLvBal)
                            End If
                            
                        End If
                    End If
                    If DaysToCut_Lt > 0 Then sng_PDDAy = sng_PDDAy - DaysToCut_Lt
                End If
            End If
        End If       '' late>0 & "Y"
        '' Early trcd=7
        QueryStr = ""
        If DaysToCut_Erl > 0 And erlRul = "Y" Then
            If DedLv_pd_erl = "PD" Then sng_PDDAy = sng_PDDAy - DaysToCut_Erl
            If DedLv_pd_erl = "LV" Then
                If fstlv_erl <> "" Then QueryStr = fstlv_erl
                If SecLv_erl <> "" And QueryStr <> "" Then QueryStr = QueryStr & "," & SecLv_erl
                If TrdLv_erl <> "" And QueryStr <> "" Then QueryStr = QueryStr & "," & TrdLv_erl
                If QueryStr <> "" Then
                   If bytBackEnd = 2 And GetFlagStatus("OPTIONALDSN") Then   ' 19-02-10
                        If fstlv_erl = SecLv_erl And SecLv_erl = TrdLv_erl Then
                            QueryStr = fstlv_erl
                        ElseIf (fstlv_erl = SecLv_erl And SecLv_erl <> TrdLv) Or (fstlv_erl <> SecLv_erl And SecLv_erl = TrdLv_erl) Then
                            If fstlv_erl = "" Or TrdLv_erl = "" Then
                                QueryStr = fstlv_erl & TrdLv_erl
                            Else
                                QueryStr = fstlv_erl & "," & TrdLv_erl
                            End If
                        ElseIf fstlv_erl <> SecLv_erl And FstLv = TrdLv_erl Then
                            QueryStr = fstlv_erl & "," & SecLv_erl
                        Else
                            QueryStr = fstlv_erl & "," & SecLv_erl & "," & TrdLv_erl
                        End If
                    End If
                    If adrslv.State = 1 Then adrslv.Close
                    adrslv.Open "select " & QueryStr & "  from " & strLvBal & " where Empcode=" & _
                    "'" & ECode & "'", cnDJConn, adOpenDynamic, adLockOptimistic
                    '' TRCD=6 FOR DEDUCTING LEAVE FROM LEAVBAL WHEN THE CATMST.DEDLET = "LV"
                    
                    If fstlv_erl <> "" Then
                        Call Removal(ECode, strLvtDate, "7", fstlv_erl, strDTEnc, strLvInfo, strLvBal)
                        tempvar = 0
                        adrslv.Requery
                        If adrslv(0) > 0 Then
                            Do While Not (adrslv(0) = 0 Or DaysToCut_Erl = 0)
                                If DaysToCut_Erl >= 1 And adrslv(0) >= 1 Then
                                    cnDJConn.Execute "Update " & strLvBal & " set " & _
                                    adrslv(0).name & "=" & adrslv(0) & "-1 where Empcode='" & ECode & "'"
                                    DaysToCut_Erl = DaysToCut_Erl - 1
                                    tempvar = tempvar + 1
                               Else
                                    cnDJConn.Execute "Update " & strLvBal & " set " & _
                                    adrslv(0).name & "=" & adrslv(0) & "-" & 0.5 & _
                                    " where Empcode='" & ECode & "'"
                                    DaysToCut_Erl = DaysToCut_Erl - 0.5
                                    tempvar = tempvar + 0.5
                                End If
                                adrslv.Requery
                            Loop
                            cnDJConn.Execute "update " & strLvBal & " set " & _
                            fstlv_erl & "=" & adrslv(0) & " where Empcode=" & "'" & ECode & "'"
                            
                           Call Insertion(ECode, strLvtDate, "7", fstlv_erl, tempvar, strDTEnc, strLvInfo, strLvBal)
                        End If
                    End If
                    
                    If adrslv.Fields.Count > 1 Then
                        If DaysToCut_Erl > 0 And SecLv_erl <> "" Then
                            Call Removal(ECode, strLvtDate, "7", adrslv(1).name, strDTEnc, strLvInfo, strLvBal)
                            tempvar = 0
                            adrslv.Requery
                            If adrslv(1) > 0 Then
                                Do While Not (adrslv(1) = 0 Or DaysToCut_Erl = 0)
                                    If DaysToCut_Erl >= 1 And adrslv(1) >= 1 Then
                                        cnDJConn.Execute "Update " & strLvBal & " set " & _
                                        adrslv(1).name & "=" & adrslv(1) & "-1 where Empcode='" & ECode & "'"
                                        DaysToCut_Erl = DaysToCut_Erl - 1
                                        tempvar = tempvar + 1
                                    Else
                                        cnDJConn.Execute "Update " & strLvBal & " set " & _
                                        adrslv(1).name & "=" & adrslv(1) & "-" & 0.5 & _
                                        " where Empcode='" & ECode & "'"
                                        DaysToCut_Erl = DaysToCut_Erl - 0.5
                                        tempvar = tempvar + 0.5
                                    End If
                                    adrslv.Requery
                                Loop
                                cnDJConn.Execute "update " & strLvBal & " set " & _
                                adrslv(1).name & "=" & adrslv(1) & " where Empcode=" & "'" & ECode & "'"
                                
'
                                Call Insertion(ECode, strLvtDate, "7", adrslv(1).name, tempvar, strDTEnc, strLvInfo, strLvBal)
                            End If
                        End If
                    End If
                    
                    If adrslv.Fields.Count > 2 Then
                        If DaysToCut_Erl > 0 And TrdLv_erl <> "" Then
                           Call Removal(ECode, strLvtDate, "7", adrslv(2).name, strDTEnc, strLvInfo, strLvBal)
                           tempvar = 0
                           adrslv.Requery
                           If adrslv(2) > 0 Then
                                Do While Not (adrslv(2) = 0 Or DaysToCut_Erl = 0)
                                     If DaysToCut_Erl >= 1 And adrslv(2) >= 1 Then
                                         cnDJConn.Execute "Update " & strLvBal & " set " & _
                                         adrslv(2).name & "=" & adrslv(2) & "-1 where Empcode='" & ECode & "'"
                                         DaysToCut_Erl = DaysToCut_Erl - 1
                                         tempvar = tempvar + 1
                                     Else
                                         cnDJConn.Execute "Update " & strLvBal & " set " & _
                                         adrslv(2).name & "=" & adrslv(2) & "-" & 0.5 & _
                                         " where Empcode='" & _
                                         ECode & "'"
                                         DaysToCut_Erl = DaysToCut_Erl - 0.5
                                         tempvar = tempvar + 0.5
                                     End If
                                     adrslv.Requery
                                 Loop
                                cnDJConn.Execute "update " & strLvBal & " set " & _
                                TrdLv_erl & "=" & adrslv(2).name & " where Empcode=" & "'" & ECode & "'"
                                
'
                                Call Insertion(ECode, strLvtDate, "7", adrslv(2).name, tempvar, strDTEnc, strLvInfo, strLvBal)
                            End If
                        End If
                    End If
                    If DaysToCut_Erl > 0 Then  '' no adequate leave balance cut from paid days
                        sng_PDDAy = sng_PDDAy - DaysToCut_Erl
                    End If
                End If
            End If           '' "LV"
        End If       '' earl>0 & "Y"
    End If '' late/earl
End If     '' EOF  & BOF
'' SETTING TO NULLS
lateRul = "": erlRul = "": erMth = "": LtMth = ""
evLate = "": evEarl = "": Late_cut = "": Late_Earl = ""
DedLv_pd_late = "":  DedLv_pd_erl = ""
FstLv = "": SecLv = "":  TrdLv = ""
DaysToCut_Lt = 0: DaysToCut_Erl = 0
QueryStr = "":  tempvar = 0
'' ASSIGNING VALUE TO THE FUNCTION VARIABLE
LateEarlChk1 = sng_PDDAy

End Function

Sub Insertion(ECode, strLvtDate, trcd, LCode, Days, strDTEnc, strLvInfo, strLvBal)
On Error GoTo ERR_P
ConMain.Execute "insert into " & strLvInfo & "(Empcode,fromdate,todate,trcd,lcode,days) values " & _
"('" & ECode & "'," & strDTEnc & strLvtDate & strDTEnc & "," & strDTEnc & strLvtDate & _
strDTEnc & "," & trcd & ",'" & LCode & "'," & Days & ")"

Exit Sub
ERR_P:
    ShowError ("Insertion :" & Me.Caption)
End Sub

Sub Removal(ECode, strLvtDate, trcd, LCode, strDTEnc, strLvInfo, strLvBal)
On Error GoTo ERR_P
Dim adrsDD As New ADODB.Recordset
adrsDD.Open "Select days from " & strLvInfo & " where Empcode = '" & ECode & _
"' and fromdate = " & strDTEnc & strLvtDate & strDTEnc & " and trcd = " & trcd & _
"  and lcode = '" & LCode & "'", ConMain, adOpenStatic, adLockReadOnly
If Not (adrsDD.EOF And adrsDD.BOF) Then
    ConMain.Execute "update " & strLvBal & " set " & LCode & "=" & LCode & _
    " + " & adrsDD(0) & " where Empcode='" & ECode & "'"
 
End If
ConMain.Execute "delete from " & strLvInfo & " where Empcode = '" & _
ECode & "' and fromdate = " & strDTEnc & strLvtDate & strDTEnc & " and trcd = " & _
trcd & "  and lcode = '" & LCode & "'"

Exit Sub
ERR_P:
    ShowError ("Removal :" & Me.Caption)
End Sub

