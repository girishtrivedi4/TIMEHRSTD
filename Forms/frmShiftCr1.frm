VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmShiftCr 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2340
      TabIndex        =   11
      Top             =   1980
      Width           =   2325
   End
   Begin VB.ComboBox cboYear 
      Height          =   315
      Left            =   3030
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   30
      Width           =   1605
   End
   Begin VB.ComboBox cboMonth 
      Height          =   315
      Left            =   810
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   30
      Width           =   1605
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   345
      Left            =   30
      TabIndex        =   9
      Top             =   1530
      Visible         =   0   'False
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   609
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   8
      Top             =   870
      Visible         =   0   'False
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   1032
      _Version        =   393216
      FixedCols       =   0
      ScrollBars      =   0
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   30
      TabIndex        =   10
      ToolTipText     =   "Create monthly shift schedule"
      Top             =   1980
      Width           =   2325
   End
   Begin MSForms.ComboBox cboTo 
      Height          =   315
      Left            =   3420
      TabIndex        =   7
      Top             =   420
      Width           =   1185
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "2090;556"
      ListWidth       =   6000
      ColumnCount     =   2
      cColumnInfo     =   2
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1500;4500"
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      Caption         =   "To Employee"
      Height          =   195
      Left            =   2460
      TabIndex        =   6
      Top             =   480
      Width           =   930
   End
   Begin MSForms.ComboBox cboFrom 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   420
      Width           =   1185
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "2090;556"
      ListWidth       =   6000
      ColumnCount     =   2
      cColumnInfo     =   2
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1500;4500"
   End
   Begin VB.Label lblEmp 
      AutoSize        =   -1  'True
      Caption         =   "From Employee"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label lblYear 
      Caption         =   "&Year"
      Height          =   225
      Left            =   2550
      TabIndex        =   2
      Top             =   60
      Width           =   375
   End
   Begin VB.Label lblMonth 
      Caption         =   "&Month"
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   615
   End
End
Attribute VB_Name = "frmShiftCr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFrom_Click()
    If cboFrom.ListIndex >= 0 Then
        If cboTo.ListCount > 0 Then cboTo.ListIndex = cboFrom.ListIndex
    End If
End Sub

Private Sub cmdCreate_Click()
On Error GoTo ERR_P
Dim strShfType As String
'' Check if Table Exists
If Not FoundTable Then Exit Sub     '' Checks for Table Existence and Acts Accoringly
Call AddActivityLog(lg_NoModeAction, 2, 16)     '' Non-Standard LOG
If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "Select Empcode,F_Shf,SCode,Shf_Date,STyp,Cat,JoinDate,LeavDate," & _
"[Off],Off2,Wo_1_3,Wo_2_4 from EmpMst  where EmpCode between " & _
"'" & cboFrom.Text & "' and '" & cboTo.Text & "' order by EmpCode,STyp", VstarDataEnv.cnDJConn
'' Make the Grids Visible
MSF1.Visible = True
MSF2.Visible = True
Do While Not adrsEmp.EOF
    If IsNull(adrsEmp("Leavdate")) And IsEmpty(adrsEmp("LeavDate")) Then
        If Month(adrsEmp("Leavdate")) < cboMonth.ListIndex - 1 Then GoTo Loop_Employee
    End If      '' If the Employee has Already Left
    Select Case DateDiff("m", Year_Start, adrsEmp("JoinDate"))
        Case Is <= 0        '' Joined Before Year Start Date
            If DateDiff("m", Year_Start, adrsEmp("Shf_Date")) >= 1 And _
            DateDiff("m", Year_Start, adrsEmp("Shf_Date")) <= 11 Then
                If Val(cboYear.Text) < Year(adrsEmp("shf_date")) Then GoTo Loop_Employee
                If Month(adrsEmp("Shf_Date")) > cboMonth.ListIndex + 1 And _
                Val(cboYear.Text) = Year(adrsEmp("shf_date")) Then GoTo Loop_Employee
            ElseIf DateDiff("m", Year_Start, adrsEmp("Shf_Date")) > 11 Then
                GoTo Loop_Employee
            End If
        Case 1 To 11        '' Joined in Current Year
            If Month(adrsEmp("JoinDate")) > cboMonth.ListIndex + 1 Then GoTo Loop_Employee
            If Month(adrsEmp("Shf_Date")) > cboMonth.ListIndex + 1 Then GoTo Loop_Employee
        Case Else           '' Joined in a Futuristic Date
            GoTo Loop_Employee
    End Select
    ''If Month(adrsEmp("JoinDate")) > cboMonth.ListIndex + 1 Then GoTo Loop_Employee
    ''If Month(adrsEmp("Shf_Date")) > cboMonth.ListIndex + 1 Then GoTo Loop_Employee
    '' Reflect the Status
    MSF1.Redraw = True
    MSF1.TextMatrix(1, 0) = adrsEmp("EmpCode")
    MSF1.TextMatrix(1, 1) = DateDisp(adrsEmp("Shf_Date"))
    MSF1.Refresh
    Me.Refresh
    '' Get the employees Details
    Call FillEmployeeDetails(adrsEmp("EmpCode"))
    '' Adjust the First Day Accordingly
    typSENum.bytStart = 1
    If cboMonth.ListIndex + 1 = Month(adrsEmp("Shf_Date")) Then Call AdjustSENums(adrsEmp("Shf_Date"))
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
If Not blnAutoProcess Then _
MsgBox "Shift File for the Month of '" & cboMonth.Text & "' Processed Successfully" _
, vbInformation, App.EXEName
Exit Sub
ERR_P:
    ShowError ("Create Shift :: " & Me.Caption)
    Call SaveSetting(App.EXEName, "PrjSettings", "ShiftCreated", 0)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
If bytShfMode = 4 Then
    If cmdCreate.Enabled = True Then
        cmdCreate.Enabled = False
        cmdExit.Enabled = False
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
Call SetFormIcon(Me)        '' Sets the Forms Icon
Call RetCaption             '' Sets the Captios on the Form Controls
Call FillCombos             '' Fill Month and Year ComboBoxes
Call GetRights              '' Gets and Sets the Appropriate Rights
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
End Sub

Private Sub RetCaption()
On Error GoTo ERR_P
Me.Caption = CaptionTxt(18101)      '' Form Caption
lblMonth.Caption = CaptionTxt(1148)     '' Month
lblYear.Caption = CaptionTxt(1149)      '' Year
lblEmp.Caption = CaptionTxt(1188)       '' From Employee
lblTo.Caption = CaptionTxt(1189)        '' To Employee
cmdCreate.Caption = CaptionTxt(1152)    '' Create
cmdExit.Caption = CaptionTxt(1153)      '' Exit
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
MSF1.TextMatrix(0, 0) = CaptionTxt(1150)    '' Employee Code
MSF1.TextMatrix(0, 1) = CaptionTxt(1151)    '' Shift date
'' "Please Wait...Processing Shifts"
MSF2.TextMatrix(0, 0) = CaptionTxt(1154)
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
Call FillEmpCombos
End Sub

Public Sub GetRights()      '' Gets and Sets the Rights
On Error GoTo ERR_P
If UCase(Trim(userName)) <> strPrintUser Then
    If adrsRits.State = 1 Then adrsRits.Close
    adrsRits.Open "Select Lv_Rights from user_leave_rights where username=" & "'" & userName & _
    "'", VstarDataEnv.cnDJConn
    If Not (adrsRits.BOF And adrsRits.EOF) Then
        If Mid(adrsRits(0), 15, 1) = "1" Then
            cmdCreate.Enabled = True
        Else
            cmdCreate.Enabled = False
            MsgBox "User " & userName & " does not have Rights to Create Shift.", vbInformation, App.EXEName
        End If
    End If
    adrsRits.Close
Else
    cmdCreate.Enabled = True
End If
Exit Sub
ERR_P:
    ShowError ("GetRights :: " & Me.Caption)
    cmdCreate.Enabled = False
End Sub

Private Function FoundTable() As Boolean
On Error GoTo ERR_P
FoundTable = True
If FindTable(Left(cboMonth.Text, 3) & Right(cboYear.Text, 2) & "Shf") Then
    '' If the Table is Found
    If MsgBox("Shift File for the Month " & cboMonth.Text & " Already Exists " & _
              vbCrLf & " Do You Wish to Overwrite it", vbQuestion + vbYesNo) = vbYes Then
              '' if the User Chooses to Overwrite the Shift File
        ''Call TruncateTable(Left(cboMonth.Text, 3) & Right(cboYear.Text, 2) & "Shf")
        '' Delete from Shf File
        VstarDataEnv.cnDJConn.Execute "Delete from " & Left(cboMonth.Text, 3) & _
        Right(cboYear.Text, 2) & "Shf where EmpCode between " & _
        "'" & cboFrom.Text & "' and '" & cboTo.Text & "'"
        Call GetSENums(cboMonth.Text, cboYear.Text) '' Get the Start and the End Numbers
        FoundTable = True
    Else
            '' if the User Chooses not to Overwrite the Shift File
        FoundTable = False
    End If
Else
    VstarDataEnv.cnDJConn.Execute "Select * into " & Left(cboMonth.Text, 3) & _
    Right(cboYear.Text, 2) & "Shf" & " from shfinfo where " & "1=2"
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
Call ComboFill(cboFrom, 1, 2)   '' Fill Employee Code Combo
Call ComboFill(cboTo, 1, 2)     '' Fill Employee Code Combo
If cboFrom.ListCount > 0 Then cboFrom.ListIndex = 0
If cboTo.ListCount > 0 Then cboTo.ListIndex = cboTo.ListCount - 1
Exit Sub
ERR_P:
    ShowError ("Fill Employee Combos :: " & Me.Caption)
End Sub

'' Modes of bytShfMode
'' 0. No Mode
'' 1. Assigning Shift from Employee Master Form
'' 2. Assigning Shift from Shedule Master Form
'' 3. Assigning Shift from Change Schedule Form
'' 4. For Creating Shifts from Daily Process
