VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Period"
   ClientHeight    =   5985
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkWOHL 
      Caption         =   "Overwrite Holidays"
      Height          =   255
      Index           =   1
      Left            =   4530
      TabIndex        =   6
      Top             =   510
      Width           =   2175
   End
   Begin VB.CheckBox chkWOHL 
      Caption         =   "Overwtite Week Off's"
      Height          =   255
      Index           =   0
      Left            =   4530
      TabIndex        =   5
      Top             =   180
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   495
      Left            =   5250
      TabIndex        =   16
      Top             =   1950
      Width           =   1455
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      Default         =   -1  'True
      Height          =   495
      Left            =   5220
      TabIndex        =   15
      Top             =   1050
      Width           =   1485
   End
   Begin VB.Frame frP 
      Height          =   1035
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   4485
      Begin VB.ComboBox cboTD 
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   630
         Width           =   705
      End
      Begin VB.ComboBox cboFD 
         Height          =   315
         Left            =   930
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   630
         Width           =   705
      End
      Begin MSForms.ComboBox cboYear 
         Height          =   300
         Left            =   3120
         TabIndex        =   1
         Top             =   150
         Width           =   1215
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2143;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboMonth 
         Height          =   300
         Left            =   720
         TabIndex        =   0
         Top             =   120
         Width           =   1605
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2831;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboShift 
         Height          =   315
         Left            =   3660
         TabIndex        =   4
         Top             =   630
         Width           =   765
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "1349;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblShift 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Left            =   3210
         TabIndex        =   22
         Top             =   690
         Width           =   315
      End
      Begin VB.Label lblTD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Day"
         Height          =   195
         Left            =   1680
         TabIndex        =   21
         Top             =   690
         Width           =   525
      End
      Begin VB.Label lblFD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Day"
         Height          =   195
         Left            =   60
         TabIndex        =   20
         Top             =   690
         Width           =   675
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1500
         TabIndex        =   27
         Top             =   600
         Width           =   45
      End
      Begin VB.Label lblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month "
         Height          =   195
         Left            =   60
         TabIndex        =   18
         Top             =   180
         Width           =   495
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         Height          =   195
         Left            =   2580
         TabIndex        =   19
         Top             =   210
         Width           =   330
      End
   End
   Begin VB.Frame frEmp 
      Caption         =   "Employee Selection"
      Height          =   4935
      Left            =   0
      TabIndex        =   23
      Top             =   1050
      Width           =   5205
      Begin VB.CommandButton cmdSR 
         Caption         =   "Select Range"
         Height          =   405
         Left            =   3690
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdUR 
         Caption         =   "Unselect Range"
         Height          =   435
         Left            =   3690
         TabIndex        =   12
         Top             =   1500
         Width           =   1455
      End
      Begin VB.CommandButton cmdSA 
         Caption         =   "Select All"
         Height          =   405
         Left            =   3690
         TabIndex        =   13
         Top             =   2100
         Width           =   1455
      End
      Begin VB.CommandButton cmdUA 
         Caption         =   "Unselect All"
         Height          =   405
         Left            =   3690
         TabIndex        =   14
         Top             =   2520
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   3855
         Left            =   30
         TabIndex        =   10
         ToolTipText     =   "Click to Toggle Selection"
         Top             =   1050
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   6800
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSForms.ComboBox cboDept 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   210
         Width           =   3555
         VariousPropertyBits=   612390939
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "6271;556"
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
         Left            =   150
         TabIndex        =   24
         Top             =   270
         Width           =   825
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   660
         Width           =   345
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T&o"
         Height          =   195
         Left            =   2790
         TabIndex        =   26
         Top             =   660
         Width           =   195
      End
      Begin MSForms.ComboBox cboFrom 
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   600
         Width           =   1845
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3254;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboTo 
         Height          =   315
         Left            =   3240
         TabIndex        =   9
         Top             =   600
         Width           =   1845
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3254;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strShfFName As String, strPECode As String
''
Dim adrsC As New ADODB.Recordset

Private Sub cboFD_Click()
If cboFD.ListIndex = -1 Then Exit Sub
cboTD.ListIndex = cboFD.ListIndex
End Sub

Private Sub cboMonth_Click()
Call FillFTD
End Sub
Private Sub cboFrom_Change()
    cboTo.ListIndex = cboTo.ListCount - 1
End Sub
Private Sub cboFrom_Click()
    If cboFrom.ListIndex < 0 Then Exit Sub
    If cboTo.ListIndex = 0 Then Exit Sub
    cboTo.ListIndex = cboFrom.ListIndex
End Sub

Private Sub cboYear_Click()
Call FillFTD
End Sub

Private Sub cboDept_Click()
On Error GoTo ERR_P
If cboDept.ListIndex < 0 Then Exit Sub  '' If No Department
Call FillEmpCombos
Call SelUnselAll(UNSELECTED_COLOR, MSF1)
Exit Sub
ERR_P:
    ShowError ("Department::" & Me.Caption)
End Sub

Private Sub cmdChange_Click()
Dim intTmp As Integer
If Not ValidateDet Then Exit Sub
With MSF1
    .Col = 0
    For intTmp = 1 To .Rows - 1
        .row = intTmp
        If .CellBackColor = SELECTED_COLOR Then Call UpdateShifts(.Text)
    Next
End With
MsgBox NewCaptionTxt("14007", adrsC), vbInformation
End Sub

Private Sub UpdateShifts(strTmp As String)
On Error GoTo ERR_P
Dim adrsTmp1 As New ADODB.Recordset, bytTmp As Byte, strTmp1 As String
adrsTmp1.Open "Select * from " & strShfFName & " where Empcode='" & strTmp & _
"'", ConMain, adOpenKeyset, adLockOptimistic
If adrsTmp1.EOF Then
    ConMain.Execute "insert into " & strShfFName & "(Empcode)" & _
    " Values('" & strTmp & "')"
    adrsTmp1.Requery
End If
For bytTmp = Val(cboFD.Text) To Val(cboTD.Text)
    strTmp1 = "D" & bytTmp
    Select Case adrsTmp1(strTmp1)
        Case pVStar.WosCode
            If chkWOHL(0).Value = 1 Then adrsTmp1(strTmp1) = cboShift.Text
        Case pVStar.HlsCode
            If chkWOHL(1).Value = 1 Then adrsTmp1(strTmp1) = cboShift.Text
        Case Else
            adrsTmp1(strTmp1) = cboShift.Text
    End Select
    adrsTmp1.Update
Next
If UCase(cboMonth.Text & cboYear.Text & strTmp) = UCase(strPECode) Then bytShfMode = 8
adrsTmp1.Close
Set adrsTmp1 = Nothing
Exit Sub
ERR_P:
    ShowError ("UpdateShifts :: " & Me.Caption)
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Call SetFormIcon(Me, True)        '' Sets the Forms Icon
Call RetCaptions
Call FillCombos             '' FillCombos
Call FillShiftCombo         '' Shift Combo
Call PutParamaters
Exit Sub
ERR_P:
    ShowError ("Load :: " & Me.Caption)
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

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '14%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("14001", adrsC)              '' Period
Call SetCritLabel(lblDeptCap)
lblMonth.Caption = NewCaptionTxt("00026", adrsMod)        '' Month
Call SetGridDetails(Me, frEmp, MSF1, lblFrom, lblTo)
Call CapGrid
End Sub

Private Sub FillCombos()
On Error GoTo ERR_P
Dim intTmpCnt As Integer
'' Fill  Month Combo
For intTmpCnt = 1 To 12
    cboMonth.AddItem Choose(intTmpCnt, "January", "February", "March", "April", "May", "June" _
    , "July", "August", "September", "October", "November", "December")
Next
'' Year Combo
For intTmpCnt = 1997 To 2096
    cboYear.AddItem CStr(intTmpCnt)
Next
cboMonth.Text = MonthName(Month(Date))
cboYear.Text = pVStar.YearSel
Call SetCritCombos(cboDept)
cboDept.ListIndex = bytLstInd
Exit Sub
ERR_P:
    ShowError ("FillCombos :: " & Me.Caption)
End Sub

Private Sub FillEmpCombos()
On Error GoTo ERR_P
Dim intEmpCnt As Integer, intTmpCnt As Integer
Dim strArrEmp() As String
Dim strDeptTmp As String, strTempforCF As String
Dim adrsTmp As New ADODB.Recordset
intEmpCnt = 0

If cboDept.Text = "ALL" Then
    strDeptTmp = "ALL"
Else
    strDeptTmp = cboDept.List(cboDept.ListIndex, 1)
    strDeptTmp = EncloseQuotes(strDeptTmp)
End If

Select Case UCase(Trim(strDeptTmp))
    Case "", "ALL"
            strTempforCF = "select Empcode,name from empmst order by empcode"
    Case Else
        ''For Mauritius 09-08-2003
        ''Original ->strTempforCF = "select Empcode,name from empmst Where " & SELCRIT & "=" & _
        strDeptTmp & " order by Empcode"                               'Empcode,name
        If strCurrentUserType = HOD Then
            strTempforCF = "select Empcode,name from empmst " & strCurrData & " and Empmst." & SELCRIT & _
                " = " & strDeptTmp & " order by Empcode"
        Else
            strTempforCF = "select Empcode,name from empmst Where empmst." & SELCRIT & _
            " = " & strDeptTmp & " order by Empcode"             'Empcode,name
       End If
End Select
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
    ShowError ("FillEmpCombos :: " & Me.Caption)
End Sub

Private Sub FillShiftCombo()        '' Fills Shift ComboBox
On Error GoTo ERR_P
Dim strArrTmp() As String, bytTmp As Byte
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select Shift,Shf_In,Shf_Out from Instshft where shift <> '100' Order by Shift", _
ConMain, adOpenStatic
If Not (adrsDept1.BOF And adrsDept1.EOF) Then
    cboShift.ColumnCount = 3
    cboShift.ListWidth = "5.5 cm"
    cboShift.ColumnWidths = "1.5 cm;2  cm;2 cm"
    ReDim Preserve strArrTmp(adrsDept1.RecordCount - 1, 2)
    For bytTmp = 0 To adrsDept1.RecordCount - 1
        strArrTmp(bytTmp, 0) = adrsDept1("Shift")                       '' Shift Code
        strArrTmp(bytTmp, 1) = Format(adrsDept1("Shf_In"), "00.00")     '' Shift In Time
        strArrTmp(bytTmp, 2) = Format(adrsDept1("Shf_Out"), "00.00")    '' Shift Out Time
        adrsDept1.MoveNext
    Next
    cboShift.List = strArrTmp
    cboShift.AddItem pVStar.WosCode
    cboShift.AddItem pVStar.HlsCode
    Erase strArrTmp
End If
Exit Sub
ERR_P:
    ShowError ("FillShiftCombo :: " & Me.Caption)
End Sub

Private Sub CapGrid()
'' Sizing
MSF1.ColWidth(1) = MSF1.ColWidth(1) * 2.65
'' Aligning
MSF1.ColAlignment(0) = flexAlignLeftTop
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

Private Function LeapOrNotRot(Optional intYear As Integer) As Boolean   '' Checks if a
If intYear Mod 4 = 0 Then   '' If Divisible by 4
    ' Is it a Century?
    If intYear Mod 100 = 0 Then     '' if Divisible by 100
        ' If a Century, must be Evenly Divisible by 400.
        If intYear Mod 400 = 0 Then     '' If Divisible by 400
            LeapOrNotRot = True                 '' Leap Year
        Else
            LeapOrNotRot = False                '' Non-Leap Year
        End If
    Else
        LeapOrNotRot = True                     '' Leap Year
    End If
Else
    LeapOrNotRot = False                        '' Non-Leap Year
End If
End Function

Private Function ValidateDet() As Boolean
On Error GoTo ERR_P
Dim intTmp As Integer, bytTmp As Integer
bytTmp = 0
If cboMonth.Text = "" Then
    MsgBox NewCaptionTxt("14008", adrsC), vbExclamation
    cboMonth.SetFocus
    Exit Function
End If
If cboYear.Text = "" Then
    MsgBox NewCaptionTxt("14009", adrsC), vbExclamation
    cboYear.SetFocus
    Exit Function
End If
strShfFName = Left(cboMonth.Text, 3) & Right(cboYear.Text, 2) & "Shf"
If Not FindTable(strShfFName) Then
    MsgBox NewCaptionTxt("14010", adrsC) & cboMonth.Text & " " & cboYear.Text & _
    vbCrLf & NewCaptionTxt("14011", adrsC), vbExclamation
    cboMonth.SetFocus
    Exit Function
End If
If Trim(cboFD.Text) = "" Then
    MsgBox NewCaptionTxt("14012", adrsC), vbExclamation
    cboFD.SetFocus
    Exit Function
End If
If Trim(cboTD.Text) = "" Then
    MsgBox NewCaptionTxt("14013", adrsC), vbExclamation
    cboTD.SetFocus
    Exit Function
End If
If Trim(cboFrom.Text) = "" Then
    MsgBox NewCaptionTxt("14012", adrsC), vbExclamation
    cboFrom.SetFocus
    Exit Function
End If
If Trim(cboTo.Text) = "" Then
    MsgBox NewCaptionTxt("14013", adrsC), vbExclamation
    cboTo.SetFocus
    Exit Function
End If
If Val(cboFD.Text) > Val(cboTD.Text) Then
    MsgBox NewCaptionTxt("14014", adrsC), vbExclamation
    cboFD.SetFocus
    Exit Function
End If
If Val(cboFrom.Text) > Val(cboTo.Text) Then
    cboFrom.SetFocus
    Exit Function
End If
If Trim(cboShift.Text) = "" Then
    MsgBox NewCaptionTxt("14015", adrsC), vbExclamation
    cboShift.SetFocus
    Exit Function
End If
With MSF1
    .Col = 0
    For intTmp = 1 To .Rows - 1
        .row = intTmp
        If .CellBackColor = SELECTED_COLOR Then bytTmp = bytTmp + 1
    Next
End With
If bytTmp = 0 Then
    MsgBox NewCaptionTxt("14016", adrsC), vbExclamation
    cboFrom.SetFocus
    Exit Function
End If
ValidateDet = True
Exit Function
ERR_P:
    ShowError ("ValidateDet :: " & Me.Caption)
End Function

Private Sub FillFTD()
On Error GoTo ERR_P
Dim bytTmp As Byte, bytTmp1 As Byte
If cboMonth.ListIndex = -1 Then Exit Sub
If cboYear.ListIndex = -1 Then Exit Sub
cboFD.clear
cboTD.clear
Select Case cboMonth.ListIndex
    Case 0, 2, 4, 6, 7, 9, 11
        bytTmp1 = 31
    Case 1
        If LeapOrNotRot(Val(cboYear.Text)) Then
            bytTmp1 = 29
        Else
            bytTmp1 = 28
        End If
    Case 3, 5, 8, 10
        bytTmp1 = 30
End Select
For bytTmp = 1 To bytTmp1
    cboFD.AddItem bytTmp
    cboTD.AddItem bytTmp
Next
Exit Sub
ERR_P:
    ShowError ("Fill From TO Day :: " & Me.Caption)
End Sub

Private Sub PutParamaters()
On Error GoTo ERR_P
Dim strTmp() As String, intTmp As Integer
strPECode = ""
strTmp = Split(strDjFileN, ":")
If strTmp(2) <> "" Then
    cboMonth.Text = strTmp(2)
    strPECode = cboMonth.Text
End If
If strTmp(3) <> "" Then
    cboYear.Text = strTmp(3)
    strPECode = strPECode & cboYear.Text
End If
With MSF1
    .Col = 0
    For intTmp = 1 To .Rows - 1
        .row = intTmp
        If .Text = strTmp(0) Then
            .CellBackColor = SELECTED_COLOR
            .Col = 1
            .CellBackColor = SELECTED_COLOR
            strPECode = strPECode & strTmp(0)
            Exit For
        End If
    Next
End With
Exit Sub
ERR_P:
    ShowError ("PutParamaters " & Me.Caption)
End Sub
