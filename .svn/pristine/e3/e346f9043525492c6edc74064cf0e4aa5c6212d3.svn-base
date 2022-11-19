VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmShiftCr 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frEmp 
      Caption         =   "Employee Selection"
      Height          =   3675
      Left            =   30
      TabIndex        =   13
      Top             =   390
      Width           =   4725
      Begin VB.CommandButton cmdSR 
         Caption         =   "Select Range"
         Height          =   435
         Left            =   3390
         TabIndex        =   5
         Top             =   1560
         Width           =   1305
      End
      Begin VB.CommandButton cmdUR 
         Caption         =   "Unselect Range"
         Height          =   465
         Left            =   3390
         TabIndex        =   6
         Top             =   1980
         Width           =   1305
      End
      Begin VB.CommandButton cmdSA 
         Caption         =   "Select All"
         Height          =   435
         Left            =   3390
         TabIndex        =   7
         Top             =   2580
         Width           =   1305
      End
      Begin VB.CommandButton cmdUA 
         Caption         =   "Unselect All"
         Height          =   435
         Left            =   3390
         TabIndex        =   8
         Top             =   3000
         Width           =   1305
      End
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   2445
         Left            =   30
         TabIndex        =   17
         ToolTipText     =   "Click to Toggle Selection"
         Top             =   1200
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   4313
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSForms.ComboBox cboDept 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   3105
         VariousPropertyBits=   612390939
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5477;556"
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
         Left            =   90
         TabIndex        =   14
         Top             =   330
         Width           =   825
      End
      Begin VB.Label lblEmp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From "
         Height          =   195
         Left            =   510
         TabIndex        =   15
         Top             =   750
         Width           =   390
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   2520
         TabIndex        =   16
         Top             =   750
         Width           =   195
      End
      Begin MSForms.ComboBox cboFrom 
         Height          =   315
         Left            =   930
         TabIndex        =   3
         Top             =   690
         Width           =   1215
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2143;556"
         ListWidth       =   6000
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboTo 
         Height          =   315
         Left            =   2790
         TabIndex        =   4
         Top             =   690
         Width           =   1215
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2143;556"
         ListWidth       =   6000
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   405
      Left            =   2400
      TabIndex        =   10
      Top             =   5460
      Width           =   2355
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   345
      Left            =   30
      TabIndex        =   19
      Top             =   4740
      Visible         =   0   'False
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   609
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   585
      Left            =   0
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   1032
      _Version        =   393216
      FixedCols       =   0
      ScrollBars      =   0
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   405
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "Create monthly shift schedule"
      Top             =   5460
      Width           =   2415
   End
   Begin MSForms.ComboBox cboYear 
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Top             =   0
      Width           =   1605
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2831;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboMonth 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   1605
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2831;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblYear 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   195
      Left            =   2730
      TabIndex        =   12
      Top             =   60
      Width           =   330
   End
   Begin VB.Label lblMonth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   60
      Width           =   450
   End
End
Attribute VB_Name = "frmShiftCr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSelEmp As String
Dim adrsC As New ADODB.Recordset

Private Sub cboFrom_Click()
    If cboFrom.ListIndex >= 0 Then
        If cboTo.ListCount > 0 Then cboTo.ListIndex = cboFrom.ListIndex
    End If
End Sub

Private Sub cboMonth_Change()

   If cboDept.Text <> "" Then
      Call FillEmpCombos
   End If

End Sub

Private Sub cboYear_Change()

   If cboDept.Text <> "" Then
      Call FillEmpCombos
   End If

End Sub

Private Sub cmdCreate_Click()
On Error GoTo ERR_P
Dim strShfType As String
'' Check if Table Exists
If Not FoundTable Then Exit Sub     '' Checks for Table Existence and Acts Accoringly
Call AddActivityLog(lg_NoModeAction, 2, 16)     '' Non-Standard LOG
Call AuditInfo("SHIFT CREATION", Me.Caption, "Created Shift For : " & cboMonth.Text & " " & cboYear.Text)
If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "Select Empcode,F_Shf,SCode,Shf_Date,STyp,Cat,JoinDate,LeavDate," & _
"" & strKOff & ",Off2,Wo_1_3,Wo_2_4,location from EmpMst  where EmpCode in ( " & _
strSelEmp & ") order by EmpCode,STyp", ConMain, 1, adLockOptimistic
'' Make the Grids Visible
MSF1.Visible = True
MSF2.Visible = True
Do While Not adrsEmp.EOF
    '' Start Date Checks
    Dim dttmp As Date
    '' If the Employee has Already Left
    If Not IsNull(adrsEmp("Leavdate")) Then
        dttmp = FdtLdt(cboMonth.ListIndex + 1, cboYear.Text, "F")
        If DateCompDate(adrsEmp("LeavDate")) <= dttmp Then GoTo Loop_Employee
    End If
    '' Get Current Months Last Process Date
    dttmp = FdtLdt(cboMonth.ListIndex + 1, cboYear.Text, "L")
    '' Check on JoinDate
    If DateCompDate(adrsEmp("JoinDate")) > dttmp Then GoTo Loop_Employee
    '' Check on Shift Date
    If DateCompDate(adrsEmp("Shf_Date")) > dttmp Then GoTo Loop_Employee

    MSF1.Redraw = True
    MSF1.TextMatrix(1, 0) = adrsEmp("EmpCode")
    MSF1.TextMatrix(1, 1) = DateDisp(adrsEmp("Shf_Date"))
    MSF1.Refresh
    Me.Refresh
    '' Get the employees Details
    Call FillEmployeeDetails(adrsEmp("EmpCode"))
    '' Adjust the First Day Accordingly
    typSENum.bytStart = 1
    If cboMonth.ListIndex + 1 = Month(adrsEmp("Shf_Date")) And Year(adrsEmp("Shf_Date")) = CInt(cboYear.Text) Then Call AdjustSENums(adrsEmp("Shf_Date"))
    If typEmpRot.strShifttype = "F" Then
        '' If Fixed Shifts
        Call FixedShifts(adrsEmp("EmpCode"), cboMonth.Text, cboYear.Text)
    Else
        '' if Rotation Shifts
        '' Fill Other Skip Pattern and Shift Pattern Array
        Call FillArrays
        Select Case strCapSND
            Case "O"        '' After Specific Number of Days
                Call SpecificDaysShifts(adrsEmp("EmpCode"), cboMonth.Text, cboYear.Text)
            Case "D"        '' Only on Fixed Days
                Call FixedDaysShifts(adrsEmp("EmpCode"), cboMonth.Text, cboYear.Text)
            Case "W"        '' Only On Fixed Week days
                Call WeekDaysShifts(adrsEmp("EmpCode"), cboMonth.Text, cboYear.Text)
        End Select
    End If
    '' Add that Record to the Shift File
    Call AddRecordsToShift(cboMonth.Text, cboYear.Text, adrsEmp("EmpCode"))
Loop_Employee:
    adrsEmp.MoveNext
Loop

'' Make the Grids Invisible
MSF1.Visible = False
MSF2.Visible = False
'' On Successfull Completion of Shifts Set the Registry status to 1
'' This is Needed if this Form is Called from Daily Processing
Call SaveSetting(App.EXEName, "PrjSettings", "ShiftCreated", 1)
MsgBox NewCaptionTxt("49006", adrsC) & " '" & UCase(cboMonth.Text) & "' ", vbInformation
Exit Sub
ERR_P:
    ShowError ("Create Shift :: " & Me.Caption)
    Call SaveSetting(App.EXEName, "PrjSettings", "ShiftCreated", 0)
    'Resume Next
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSA_Click()
Call SelUnselAll(SELECTED_COLOR, MSF3)
End Sub

Private Sub cmdSR_Click()
Call SelUnsel(SELECTED_COLOR, MSF3, cboFrom, cboTo)
End Sub

Private Sub cmdUA_Click()
Call SelUnselAll(UNSELECTED_COLOR, MSF3)
End Sub

Private Sub cmdUR_Click()
Call SelUnsel(UNSELECTED_COLOR, MSF3, cboFrom, cboTo)
End Sub

Private Sub Form_Activate()
If bytShfMode = 4 Then
    If cmdCreate.Enabled = True Then
        cmdCreate.Enabled = False
        cmdExit.Enabled = False
        cboDept.ListIndex = bytLstInd
        Call SelUnselAll(SELECTED_COLOR, MSF3)
        Call cmdCreate_Click
    End If
    bytShfMode = 0
    Unload Me
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdCreate.Enabled = True Then Call cmdCreate_Click
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)            '' Sets the Forms Icon
Call RetCaption                 '' Sets the Captios on the Form Controls
Call FillCombos                 '' Fill Month and Year ComboBoxes
Call GetRights                  '' Gets and Sets the Appropriate Rights
If bytShfMode = 4 Then
    '' if Coming Throgh the Daily Processing Form
    cboMonth.Text = Left(strRotPass, Len(strRotPass) - 4)
    cboYear.Text = Right(strRotPass, 4)
    cboMonth.Enabled = False
    cboYear.Enabled = False
Else
    cboMonth.Text = MonthName(Month(Date))
    cboYear.Text = pVStar.YearSel
End If
If strCurrentUserType <> HOD Then cboDept.Text = "ALL"
End Sub

Private Sub RetCaption()
On Error GoTo ERR_P
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "select * from NewCaptions where ID like '49%'", ConMain, adOpenStatic, adLockReadOnly
Me.Caption = "Monthly Shift Creation"
Call SetCritLabel(lblDeptCap)
Call SetGridDetails(Me, frEmp, MSF3, lblEmp, lblTo)
'cmdCreate.Caption = NewCaptionTxt("00053", adrsMod)
'cmdExit.Caption = NewCaptionTxt("00008", adrsMod)
cmdCreate.ToolTipText = NewCaptionTxt("49011", adrsC) ''Create monthly shift schedule
Call CapGrid                        '' Sets the Captions and Sizes the Grid
Exit Sub
ERR_P:
    ShowError ("RetCaption :: " & Me.Caption)
End Sub

Private Sub CapGrid()
'' Sizing
MSF1.ColWidth(0) = MSF1.ColWidth(0) * 2.37
MSF1.ColWidth(1) = MSF1.ColWidth(1) * 2.37
MSF2.ColWidth(0) = MSF2.ColWidth(0) * 4.73
'' Setting Captions
MSF1.TextMatrix(0, 0) = NewCaptionTxt("00061", adrsMod)
MSF1.TextMatrix(0, 1) = NewCaptionTxt("49004", adrsC)
'' "Please Wait...Processing Shifts"
MSF2.TextMatrix(0, 0) = NewCaptionTxt("49005", adrsC)
'' MSF3
MSF3.ColWidth(1) = MSF3.ColWidth(1) * 2.2
End Sub

Private Sub FillCombos()
Dim intTmp As Integer
With cboMonth           '' Month
    .AddItem "January"
    .AddItem "February"
    .AddItem "March"
    .AddItem "April"
    .AddItem "May"
    .AddItem "June"
    .AddItem "July"
    .AddItem "August"
    .AddItem "September"
    .AddItem "October"
    .AddItem "November"
    .AddItem "December"
End With
With cboYear            '' Year
    For intTmp = 1997 To 2096
        .AddItem CStr(intTmp)
    Next
End With
cboMonth.Text = MonthName(Month(Date))
cboYear.Text = pVStar.YearSel
Call SetCritCombos(cboDept)
'If strCurrentUserType <> HOD Then cbodept.AddItem "ALL"
End Sub

Private Sub cboDept_Click()
On Error GoTo ERR_P
If cboDept.ListIndex < 0 Then Exit Sub               '' If No Department
Call FillEmpCombos
Call SelUnselAll(UNSELECTED_COLOR, MSF3)
Exit Sub
ERR_P:
    ShowError ("Department::" & Me.Caption)
End Sub

Public Sub GetRights()      '' Gets and Sets the Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(4, 4, 2, 1)
If strTmp = "1" Then
    cmdCreate.Enabled = True
Else
    cmdCreate.Enabled = False
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
End If
Exit Sub
ERR_P:
    ShowError ("GetRights :: " & Me.Caption)
    cmdCreate.Enabled = False
End Sub

Private Function FoundTable() As Boolean
On Error GoTo ERR_P
Dim intTmp As Integer
intTmp = 0
If MSF3.Rows = 1 Then
    MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation
    If bytShfMode <> 4 Then cmdSR.SetFocus
    Exit Function
End If
Call TruncateTable("ECode") ' 09-03
MSF3.Col = 0: strSelEmp = ""
    For i = 1 To MSF3.Rows - 1
        MSF3.row = i
        If MSF3.CellBackColor = SELECTED_COLOR Then
            intTmp = intTmp + 1
            'strSelEmp = strSelEmp & "'" & MSF3.Text & "'" & ","
            ConMain.Execute "insert into ECode values('" & MSF3.Text & "')"   ' 09-03
        End If
        strSelEmp = "select empcode from ECode "
    Next
    If intTmp = 0 Then
        MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation
        If bytShfMode <> 4 Then cmdSR.SetFocus
        Exit Function
    End If
    strSelEmp = Left(strSelEmp, Len(strSelEmp) - 1)

If FindTable(Left(cboMonth.Text, 3) & Right(cboYear.Text, 2) & "Shf") Then
    If MsgBox(NewCaptionTxt("49007", adrsC) & " " & UCase(cboMonth.Text) & _
              vbCrLf & NewCaptionTxt("49008", adrsC), vbQuestion + vbYesNo) = vbYes Then
              '' if the User Chooses to Overwrite the Shift File
        '' Delete from Shf File
OverWriteShift:
        ConMain.Execute "Delete from " & Left(cboMonth.Text, 3) & _
        Right(cboYear.Text, 2) & "Shf where EmpCode in( " & strSelEmp & ")"
        Call GetSENums(cboMonth.Text, cboYear.Text) '' Get the Start and the End Numbers
        FoundTable = True
    Else
            '' if the User Chooses not to Overwrite the Shift File
        FoundTable = False
    End If
Else
    'conmain.Execute "Select * into " & Left(cboMonth.Text, 3) & _
    Right(cboYear.Text, 2) & "Shf" & " from shfinfo where 1=2"
    Call CreateTableIntoAs("*", "shfinfo", Left(cboMonth.Text, 3) & _
        Right(cboYear.Text, 2) & "Shf", " where 1=2")
    Call CreateTableIndexAs("MONYYSHF", Left(cboMonth.Text, 3), Right(cboYear.Text, 2))
    Call GetSENums(cboMonth.Text, cboYear.Text)     '' Get the Start and the End Numbers
    FoundTable = True
End If
Exit Function
ERR_P:
    ShowError ("FoundTable :: " & Me.Caption)
    FoundTable = False
End Function

Private Sub FillEmpCombos()
On Error GoTo ERR_P
Dim adrsTmp As New ADODB.Recordset, intEmpCnt As Integer, intTmpCnt As Integer
Dim strDeptTmp As String, strTempforCF As String

Call ComboFill(cboFrom, 19, 2, cboDept.List(cboDept.ListIndex, 0))
Call ComboFill(cboTo, 19, 2, cboDept.List(cboDept.ListIndex, 0))
'End If
If cboFrom.ListCount > 0 Then cboFrom.ListIndex = 0
If cboTo.ListCount > 0 Then cboTo.ListIndex = cboTo.ListCount - 1
If cboDept.Text = "ALL" Then
    strDeptTmp = cboDept.List(cboDept.ListIndex, 0)
Else
    strDeptTmp = cboDept.List(cboDept.ListIndex, 1)
End If
strDeptTmp = EncloseQuotes(strDeptTmp)
Select Case UCase(Trim(strDeptTmp))
    Case "", "ALL"
    
       strTempforCF = "select Empcode,name from empmst where leavdate is null or year(leavdate) > " & (frmShiftCr.cboYear.Text) & " or (year(leavdate)= " & (frmShiftCr.cboYear.Text) & " and month(leavdate) >= " & Format(MonthNumber(frmShiftCr.cboMonth.Text), "00") & ") order by Empcode"

    Case Else
        
        If strCurrentUserType = HOD Then
                strTempforCF = "select Empcode,name from empmst " & strCurrData & " and Empmst." & SELCRIT & _
                " = " & strDeptTmp & " order by Empcode"
       Else
                strTempforCF = "select Empcode,name from empmst Where empmst." & SELCRIT & _
               " = " & strDeptTmp & " order by Empcode"             'Empcode,name
       End If
                        '
    
 End Select
If GetFlagStatus("LocationRights") And strCurrentUserType <> ADMIN Then ' 28-01-09
    If UCase(Trim(strDeptTmp)) = "ALL" Then
        strTempforCF = "select Empcode,name from empmst " & strCurrData & " order by Empcode"
    Else
        strTempforCF = "select Empcode,name from empmst where Empmst.Dept = " & cboDept.Text & " And Empmst.Location in (" & UserLocations & ") order by Empcode"
    End If
End If

If adrsTmp.State = 1 Then adrsTmp.Close
adrsTmp.Open strTempforCF, ConMain, adOpenStatic, adLockReadOnly
If (adrsTmp.EOF And adrsTmp.BOF) Then
    MSF3.Rows = 1
    Exit Sub
End If
intEmpCnt = adrsTmp.RecordCount
intTmpCnt = intEmpCnt
MSF3.Rows = intEmpCnt + 1
For intEmpCnt = 0 To intTmpCnt - 1
    MSF3.TextMatrix(intEmpCnt + 1, 0) = adrsTmp(0)
    MSF3.TextMatrix(intEmpCnt + 1, 1) = adrsTmp(1)
    adrsTmp.MoveNext
Next
Exit Sub
ERR_P:
    ShowError ("Fill Employee Combos :: " & Me.Caption)
End Sub

Private Sub MSF3_Click()
If MSF3.Rows = 1 Then Exit Sub
If MSF3.CellBackColor = SELECTED_COLOR Then
    With MSF3
        .Col = 0
        .CellBackColor = UNSELECTED_COLOR
        .Col = 1
        .CellBackColor = UNSELECTED_COLOR
    End With
Else
    With MSF3
        .Col = 0
        .CellBackColor = SELECTED_COLOR
        .Col = 1
        .CellBackColor = SELECTED_COLOR
    End With
End If
End Sub

Private Sub MSF3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then Call MSF3_Click
End Sub
