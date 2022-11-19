VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmOT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OT Authorization"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   435
      Left            =   7440
      TabIndex        =   8
      Top             =   4590
      Width           =   1155
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
      TabIndex        =   13
      Top             =   4320
      Visible         =   0   'False
      Width           =   7065
      Begin VB.ComboBox cmbreason 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtRem 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1530
         MaxLength       =   15
         TabIndex        =   18
         Top             =   630
         Width           =   1545
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Command1"
         Height          =   405
         Left            =   5760
         TabIndex        =   7
         Top             =   660
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Default         =   -1  'True
         Height          =   405
         Left            =   4440
         TabIndex        =   6
         Top             =   660
         Width           =   1155
      End
      Begin VB.TextBox txtOT 
         Height          =   315
         Left            =   5580
         TabIndex        =   16
         Top             =   240
         Width           =   765
      End
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
      Begin VB.Label lblOTRem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OT Remark"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   690
         Width           =   825
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
         TabIndex        =   14
         Top             =   240
         Width           =   720
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
         TabIndex        =   15
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame frEmp 
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8775
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   4890
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   150
         Width           =   1575
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   7200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   165
         Width           =   1050
      End
      Begin MSForms.ComboBox cboEmp 
         Height          =   375
         Left            =   960
         TabIndex        =   0
         Top             =   120
         Width           =   3255
         VariousPropertyBits=   612390939
         DisplayStyle    =   3
         Size            =   "5741;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
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
         Left            =   60
         TabIndex        =   10
         Top             =   180
         Width           =   870
      End
      Begin VB.Label lblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month "
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
         Left            =   4320
         TabIndex        =   11
         Top             =   180
         Width           =   600
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
         Left            =   6660
         TabIndex        =   12
         Top             =   180
         Width           =   405
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   3615
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6376
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
Attribute VB_Name = "frmOT"
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

Private Sub cmdCancel_Click()
cmdExit.Cancel = True
frOT.Visible = False
End Sub

Private Sub cmdExit_Click()
Unload Me
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

ConMain.Execute "Update " & strFileName & " Set OTConf='" & strTmp _
& "',Ovtim=" & Val(txtOT.Text) & ",OTRem='" & strTmp1 & "' Where Empcode='" & cboEmp.List(cboEmp.ListIndex, 1) & _
"' and " & strKDate & " =" & strDTEnc & DateCompStr(lblDate.Caption) & strDTEnc
cmdExit.Cancel = True
frOT.Visible = False
Call FillGrid
Exit Sub
ERR_P:
    ShowError ("Save::" & Me.Caption)
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)        '' Sets the Form Icon
Call RetCaptions            '' Sets the Captions
Call FillCombos             '' Fills All the Combos on the Form
Call GetRights              '' Gets and Sets the Rights

txtRem.Visible = True
cmbreason.Visible = False

End Sub

Private Sub RetCaptions()
On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '62%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("62001", adrsC)
'' Employee Details
lblCode.Caption = "Employee Code"
chkOT.Caption = NewCaptionTxt("62001", adrsC)                       '' OT Authorized
lblOT.Caption = NewCaptionTxt("62007", adrsC)                       '' Overtime Hrs.
cmdsave.Caption = "Save"
cmdCancel.Caption = "Cancel"
cmdExit.Caption = "Exit"
Call CapGrid
End Sub

Private Sub CapGrid()
On Error Resume Next
With MSF1                       '' Captions
    .TextMatrix(0, 0) = NewCaptionTxt("00030", adrsMod)         '' Date
    .TextMatrix(0, 1) = NewCaptionTxt("00031", adrsMod)         '' Shift
    .TextMatrix(0, 2) = NewCaptionTxt("00034", adrsMod)         '' Arrival
    .TextMatrix(0, 3) = NewCaptionTxt("00036", adrsMod)         '' Departure
    .TextMatrix(0, 4) = NewCaptionTxt("62006", adrsC)           '' Work Hours
    .TextMatrix(0, 5) = NewCaptionTxt("62005", adrsC)           '' Overtime Authorized
    .TextMatrix(0, 6) = NewCaptionTxt("00038", adrsMod)         '' Overtime
    .TextMatrix(0, 7) = NewCaptionTxt("00124", adrsMod)         '' Remarks
    '' Widths
    .ColWidth(0) = .ColWidth(8) * 1.3
    .ColWidth(1) = .ColWidth(1) * 0.75
    .ColWidth(2) = .ColWidth(2) * 0.85
    .ColWidth(3) = .ColWidth(3) * 0.9
    .ColWidth(5) = .ColWidth(5) * 1.5
    .ColWidth(6) = .ColWidth(6) * 0.9
    .ColWidth(7) = .ColWidth(7) * 1.25
    '' Alignments
    .ColAlignment(1) = flexAlignCenterCenter
    .ColAlignment(5) = flexAlignCenterCenter
End With
End Sub

Private Sub FillCombos()            '' Fills All the ComboBoxes on the Form
On Error GoTo ERR_P
Dim intTmp As Integer, k As Integer
'' Employee Combo
Call FillEmpCombo
'' Month Combo
For intTmp = 1 To 12
    cboMonth.AddItem Choose(intTmp, "January", "February", "March", "April", "May", "June" _
    , "July", "August", "September", "October", "November", "December")
Next
'' Year Combo
For intTmp = 1997 To 2096
    cboYear.AddItem CStr(intTmp)
Next
bytMode = 0                                 '' Set Mode to 0 for Month and Year Selection
cboMonth.Text = MonthName(Month(Date))
cboYear.Text = Year(Date)
bytMode = 1                                 '' Set bytMode to View / Normal
sngCalcOT = 0
For k = 1 To 6
cmbreason.AddItem Choose(k, "SHORTAGE OF STRENGTH", "ABSENTISM", "SPECIAL JOB/BREAK-DOWN", "RUSH TO WORK", "PROJECT WORK", "HOLIDAY")
Next
Exit Sub
ERR_P:
    ShowError ("FillCombos::" & Me.Caption)
End Sub

Private Sub FillEmpCombo()      '' Fills the Employee Combo
On Error GoTo ERR_P
Call ComboFill(cboEmp, 1, 2)
Exit Sub
ERR_P:
    ShowError ("FillEmpCombo :: ") & Me.Caption
End Sub

Private Sub cboEmp_Click()
If cboEmp.ListIndex < 0 Then Exit Sub               '' If No Employee
If blnFileFound = False Then Exit Sub               '' If File is Not Found
Call FillGrid                                       '' Fill the Grid with Employees Record
End Sub

Private Sub cboMonth_Click()
If cboMonth.Text = "" Then Exit Sub
Call ValidMonthYear
If blnFileFound = False Then        '' If no File is Found
    MSF1.Rows = 1
Else
    If bytMode <> 0 Then            '' If File Found and not Load Mode
        If cboEmp.Text = "" Then Exit Sub
        Call FillGrid
    End If
End If
End Sub

Private Sub cboYear_Click()
    If cboYear.Text = "" Then Exit Sub
    Call ValidMonthYear
    If blnFileFound = False Then
        MSF1.Rows = 1
    Else
        If bytMode <> 0 Then
            If cboEmp.Text = "" Then Exit Sub
            Call FillGrid
        End If
    End If
End Sub

Private Sub ValidMonthYear()    '' Procedure to Check if Valid Monthly transaction File
On Error GoTo ERR_P             '' for the Selected Month and Year is Available or not
strFileName = Left(cboMonth.Text, 3) & Right(cboYear.Text, 2) & "Trn" & ""
If Not FindTable(Trim(strFileName)) Then
    If bytMode <> 0 Then
        MsgBox NewCaptionTxt("62003", adrsC) & cboMonth.Text, vbExclamation
    End If
        blnFileFound = False    '' File not Found
        Exit Sub
End If
blnFileFound = True             '' File Found
Exit Sub
ERR_P:
    ShowError ("ValidEmpMonthYear :: " & Me.Caption)
    blnFileFound = False
End Sub

Private Sub FillGrid()          '' Fills the Grid
On Error GoTo ERR_P
If adrsLeave.State = 1 Then adrsLeave.Close
adrsLeave.Open "Select Empcode," & strKDate & " ,Shift,ArrTim,Deptim,WrkHrs,Ovtim,OTConf,OTRem from " & _
strFileName & " Where Empcode='" & cboEmp.List(cboEmp.ListIndex, 1) & "' and OvTim>0" & _
" order by " & strKDate, ConMain, adOpenStatic
MSF1.Rows = 1
Do While Not adrsLeave.EOF
'    adrsLeave.Find ("Empcode='" & cboEmp.Text & "'")    '' Searches for the Specified
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
    MsgBox NewCaptionTxt("62008", adrsC) & cboEmp.Text, vbExclamation
End If
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub MSF1_DblClick()
If MSF1.Rows = 1 Then Exit Sub
MSF1.Col = 0
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
        Call MSF1_DblClick
End Select
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

Private Sub GetRights()
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(4, 13, 3, 1)
If strTmp = "1" Then
    cmdsave.Enabled = True
Else
    cmdsave.Enabled = False
End If
Exit Sub
ERR_P:
    ShowError ("GetRights::" & Me.Caption)
    cmdsave.Enabled = False
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
