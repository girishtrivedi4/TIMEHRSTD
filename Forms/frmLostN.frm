VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLostN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lost Entry"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRemark 
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   0
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtRemark 
         Enabled         =   0   'False
         Height          =   915
         Left            =   855
         MaxLength       =   256
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   28
         Top             =   0
         Width           =   3795
      End
      Begin VB.Label lblRemark 
         AutoSize        =   -1  'True
         Caption         =   "Remark::"
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
         Left            =   45
         TabIndex        =   29
         Top             =   90
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command4"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   4620
      Width           =   1185
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command3"
      Height          =   375
      Left            =   1170
      TabIndex        =   12
      Top             =   4620
      Width           =   1185
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2340
      TabIndex        =   13
      Top             =   4620
      Width           =   1185
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   3510
      TabIndex        =   14
      Top             =   4620
      Width           =   1185
   End
   Begin TabDlg.SSTab TB1 
      Height          =   4605
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8123
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmLostN.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cboEmpCode"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEmpCode(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "MSF1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmLostN.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frLost"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame frLost 
         Height          =   4200
         Left            =   60
         TabIndex        =   26
         Top             =   330
         Width           =   4590
         Begin VB.ComboBox cboIO 
            Height          =   315
            Index           =   3
            Left            =   3600
            TabIndex        =   10
            Text            =   "I"
            Top             =   3600
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cboIO 
            Height          =   315
            Index           =   2
            Left            =   3600
            TabIndex        =   8
            Text            =   "I"
            Top             =   3120
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cboIO 
            Height          =   315
            Index           =   1
            Left            =   3600
            TabIndex        =   6
            Text            =   "l"
            Top             =   2640
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cboIO 
            Height          =   315
            Index           =   0
            Left            =   3600
            TabIndex        =   4
            Text            =   "I"
            Top             =   2040
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtPunchDate 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2000
            TabIndex        =   2
            Tag             =   "D"
            Text            =   " "
            Top             =   1380
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtPunchTime 
            Height          =   375
            Index           =   0
            Left            =   1995
            TabIndex        =   3
            Top             =   1950
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtPunchTime 
            Height          =   375
            Index           =   1
            Left            =   1995
            TabIndex        =   5
            Top             =   2505
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtPunchTime 
            Height          =   375
            Index           =   2
            Left            =   1995
            TabIndex        =   7
            Top             =   3045
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtPunchTime 
            Height          =   375
            Index           =   3
            Left            =   1995
            TabIndex        =   9
            Top             =   3600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin VB.Label lblLateCnt 
            Caption         =   "Late Count = "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2175
            TabIndex        =   30
            Top             =   180
            Width           =   2205
         End
         Begin VB.Label lblCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "a"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2000
            TabIndex        =   18
            Top             =   450
            Width           =   90
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time  "
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
            Index           =   3
            Left            =   240
            TabIndex        =   25
            Top             =   3630
            Width           =   540
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time  "
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
            Index           =   2
            Left            =   240
            TabIndex        =   24
            Top             =   3075
            Width           =   540
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time  "
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
            Index           =   1
            Left            =   240
            TabIndex        =   23
            Top             =   2535
            Width           =   540
         End
         Begin VB.Label lblEmpName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "a"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2000
            TabIndex        =   20
            Top             =   900
            Width           =   90
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time  "
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
            Index           =   0
            Left            =   255
            TabIndex        =   22
            Top             =   1980
            Width           =   540
         End
         Begin VB.Label lblLostDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date  "
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
            Left            =   270
            TabIndex        =   21
            Top             =   1500
            Width           =   525
         End
         Begin VB.Label lblEmpNameCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee  Name"
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
            Left            =   270
            TabIndex        =   19
            Top             =   900
            Width           =   1500
         End
         Begin VB.Label lblEmpCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Code"
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
            Index           =   0
            Left            =   270
            TabIndex        =   17
            Top             =   450
            Width           =   1380
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   3695
         Left            =   -74940
         TabIndex        =   1
         Top             =   840
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   6509
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         AllowBigSelection=   0   'False
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
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
      Begin VB.Label lblEmpCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empoyee"
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
         Index           =   1
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   825
      End
      Begin MSForms.ComboBox cboEmpCode 
         Height          =   375
         Left            =   -73920
         TabIndex        =   0
         Top             =   405
         Width           =   3495
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "6165;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmLostN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strLostT_Punch As String, strLostDate As String
''
Dim adrsC As New ADODB.Recordset

Private Sub Form_Load()

    lblLateCnt.Visible = False

Call SetFormIcon(Me)        '' Set the Form Icon
Call SetToolTipText(Me)     '' Set the ToolTipText
Call RetCaptions            '' Retreive Captions
Call FillComboEmp           '' Fill EmployeeCombo
TB1.Tab = 0                 '' Set the Tab to List
Call GetRights              '' Gets Rights for the Operations
bytMode = 1
Call ChangeMode             '' Take Action on the Appropriate Mode
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '32%'", ConMain, adOpenStatic
Me.Caption = "Manual Entry"
TB1.TabCaption(0) = NewCaptionTxt("00013", adrsMod)       '' List
TB1.TabCaption(1) = NewCaptionTxt("00014", adrsMod)       '' Details
lblTime(0).Caption = NewCaptionTxt("32003", adrsC)      '' Time of Punch
lblTime(1).Caption = lblTime(0).Caption                 '' Time of punch
lblTime(2).Caption = lblTime(0).Caption                 '' Time of punch
lblTime(3).Caption = lblTime(0).Caption                 '' Time of punch
lblLostDate.Caption = NewCaptionTxt("32002", adrsC)     '' Date of Punch
lblEmpNameCap.Caption = NewCaptionTxt("32004", adrsC)   '' Employee Name
Call SetGButtonCap(Me)                           '' Sets Appropriate Caption for the Buttons
Call CapGrid                                '' Captions For Grid
End Sub

Private Sub CapGrid()           '' Gives the Captions to the Grid
With MSF1
    '' Sets the Column Widhts
    .ColWidth(0) = .ColWidth(0) * 1.4
    .ColWidth(1) = .ColWidth(1) * 1.32
    .ColWidth(2) = .ColWidth(2) * 1.2
    '' Sets the Column Alignment
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    '' Sets the Appropriate Captions
    .TextMatrix(0, 0) = NewCaptionTxt("00061", adrsMod)   '' Employee Code 29103
    .TextMatrix(0, 1) = NewCaptionTxt("32002", adrsC)   '' Date of Punch 29104
    .TextMatrix(0, 2) = NewCaptionTxt("32003", adrsC)   '' Time of Punch 29105
End With
End Sub

Private Sub OpenMasterTable(Optional strEmpCode As String)              '' Open the Recordset for the Display purposes
On Error GoTo ERR_P
Dim strTmp As String
''
Dim rsTmp As New ADODB.Recordset
''
If strCurrentUserType = HOD Then

    strTmp = "Select " & strKDate & ",t_punch,Lost.Empcode,Lost.shift from Lost,Empmst " & strCurrData & " And Empmst.Empcode = Lost.Empcode "
''
    If Trim(strEmpCode) = "" Then
        strTmp = strTmp & " order by Lost.Empcode"
    Else
        strTmp = strTmp & " And Lost.Empcode = '" & strEmpCode & "' order by Lost.Empcode"
    End If
Else

    strTmp = "Select " & strKDate & ",t_punch,Lost.Empcode,Lost.shift from Lost "
''
    If Trim(strEmpCode) = "" Then
        strTmp = strTmp & " order by " & strKDate & " ,t_punch "
    Else
        strTmp = strTmp & " where Lost.Empcode = '" & strEmpCode & "' order by " & strKDate & " ,t_punch "
    End If
End If
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open strTmp, ConMain, adOpenStatic

If rsTmp.State = 1 Then rsTmp.Close
rsTmp.Open "Select * from install", ConMain, adOpenStatic, adLockReadOnly
If Not (rsTmp.EOF And rsTmp.BOF) Then
    If rsTmp("IO") = "Y" Then
        typPerm.blnIO = True
    Else
        typPerm.blnIO = False
    End If
End If

Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Sub FillGrid()          '' Fills the Lost Grid
On Error GoTo ERR_P
Dim intCounter As Integer
Call OpenMasterTable(lblCode.Caption)
'' Put Appropriate Rows in the Grid
If typPerm.blnIO Then
    MSF1.Cols = 4
    MSF1.TextMatrix(0, 3) = "IO"
End If
If adrsDept1.EOF Then
    MSF1.Rows = 1
    TB1.TabEnabled(1) = False       '' Disables Tab 1 If no Records are Found
    Exit Sub
End If
MSF1.Rows = adrsDept1.RecordCount + 1   '' Sets Rows Appropriately
adrsDept1.MoveFirst
For intCounter = 1 To adrsDept1.RecordCount     '' Fills the Grid
    With MSF1
        .TextMatrix(intCounter, 0) = adrsDept1("Empcode")
        .TextMatrix(intCounter, 1) = DateDisp(adrsDept1("date"))
        .TextMatrix(intCounter, 2) = IIf(IsNull(adrsDept1("t_punch")), "0.00", _
                                     Format(adrsDept1("t_punch"), "0.00"))
        'null filteration add by  MIS2007DF021
        If typPerm.blnIO Then
            .TextMatrix(intCounter, 3) = FilterNull(adrsDept1("shift"))
        End If
    End With
    adrsDept1.MoveNext
Next
TB1.TabEnabled(1) = True        '' If Records Found Enables the TAB 1
TB1.Tab = 0
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub FillComboEmp()
On Error GoTo ERR_P
Call ComboFill(cboEmpCode, 1, 2)    '' Fill Employee Code Combo

If cboEmpCode.ListCount > 0 Then cboEmpCode.ListIndex = 0
If typPerm.blnIO Then
    cboIO(0).Visible = True
    cboIO(1).Visible = True
    cboIO(2).Visible = True
    cboIO(3).Visible = True

    cboIO(0).AddItem "I"
    cboIO(1).AddItem "I"
    cboIO(2).AddItem "I"
    cboIO(3).AddItem "I"
    
    cboIO(0).AddItem "O"
    cboIO(1).AddItem "O"
    cboIO(2).AddItem "O"
    cboIO(3).AddItem "O"
    
End If
''
Exit Sub
ERR_P:
    ShowError ("Fill Employee Combo :: " & Me.Caption)
End Sub

Private Sub GetRights()     '' Gets and Sets the Appropriate Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 15, 1)
If Mid(strTmp, 1, 1) = "1" Then AddRights = True
If Mid(strTmp, 2, 1) = "1" Then EditRights = True
If Mid(strTmp, 3, 1) = "1" Then DeleteRights = True
Exit Sub
ERR_P:
    ShowError ("Rights ::" & Me.Caption)
    AddRights = False
    EditRights = False
    DeleteRights = False
End Sub

Private Sub Display()       '' Displays the Given Master Records
On Error GoTo ERR_P
cboEmpCode.Text = MSF1.TextMatrix(MSF1.row, 0)  '' Employee Code
txtPunchDate = MSF1.TextMatrix(MSF1.row, 1)     '' Date
txtPunchTime(0).Text = MSF1.TextMatrix(MSF1.row, 2)      '' Time

If typPerm.blnIO Then
    cboIO(0).Text = MSF1.TextMatrix(MSF1.row, 3)
End If
''
''HULK
txtPunchTime(1).Text = ""
txtPunchTime(2).Text = ""
txtPunchTime(3).Text = ""
txtRemark.Text = ""
cboIO(1).Enabled = False
cboIO(2).Enabled = False
cboIO(3).Enabled = False
''
'' Get Values in the Temporary Variables
strLostT_Punch = Format(txtPunchTime(0).Text, "0.00")
strLostDate = txtPunchDate.Text
Exit Sub
ERR_P:
    ShowError ("Display  :: " & Me.Caption)
End Sub

Private Sub cboEmpCode_Click()
On Error GoTo ERR_P
'' Displays Employee Name

If cboEmpCode.Text = "" Then Exit Sub
lblCode.Caption = cboEmpCode.List(cboEmpCode.ListIndex, 1)
lblEmpName.Caption = cboEmpCode.Text '
Call FillGrid
Exit Sub
ERR_P:
    ShowError ("Employee :: " & Me.Caption)
End Sub

Private Sub cmdAddSave_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        '' Check for Rights
        If Not AddRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        bytMode = 2
        Call ChangeMode
    Case 2          '' Add Mode
        If Not ValidateAddmaster Then Exit Sub  '' Validate For Add
        If Not SaveAddMaster Then Exit Sub      '' Save for Add
    Call AuditInfo("ADD", Me.Caption, "Added Lost Entry for date: " & txtPunchDate.Text)
        Call FillGrid                           '' Reflect the Grid
        bytMode = 1
        Call ChangeMode
    Case 3          '' Edit Mode
        If Not ValidateModMaster Then Exit Sub  '' Validate for Edit
        If Not SaveModMaster Then Exit Sub      '' Save for Edit
        Call AuditInfo("UPDATE", Me.Caption, "Edit Lost Entry for date: " & txtPunchDate.Text)
        Call FillGrid                           '' Reflect the Grid
        bytMode = 1
        Call ChangeMode
End Select
Exit Sub
ERR_P:
    ShowError ("AddSave :: " & Me.Caption)
End Sub

Private Sub cmdDel_Click()
On Error GoTo ERR_P
'' Check for Rights
If Not DeleteRights Then
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    Exit Sub
Else
    If TB1.TabEnabled(1) = False Then Exit Sub
    If TB1.Tab = 0 Then                         '' Do not Display Record if
        If TB1.TabEnabled(1) Then TB1.Tab = 1   '' Already Displayed
    End If
    
    If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion) _
    = vbYes Then        '' Delete the Record
        ConMain.Execute "delete from lost where Empcode=" & "'" & _
        cboEmpCode.Text & "'" & " and " & strKDate & "=" & _
        strDTEnc & DateCompStr(strLostDate) & strDTEnc & " and t_punch=" & strLostT_Punch
    End If
    Call FillGrid       '' Reflect the Grid
    bytMode = 1
    Call ChangeMode
    Call AuditInfo("DELETE", Me.Caption, "Deleted Lost Entry for date: " & txtPunchDate.Text)
End If
Exit Sub
ERR_P:
    ShowError ("Delete Record :: " & Me.Caption)
End Sub

Private Sub cmdEditCan_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        If TB1.TabEnabled(1) = False Then Exit Sub
        '' Check for Rights
        If Not EditRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        bytMode = 3
        Call ChangeMode
    Case 2       '' Add Mode
        If MSF1.Rows = 1 Then
            TB1.TabEnabled(1) = False
            TB1.Tab = 0
        End If
        bytMode = 1
        Call ChangeMode
    Case 3      '' Edit Mode
        bytMode = 1
        Call ChangeMode
End Select
Exit Sub
ERR_P:
    ShowError ("EditCancel :: " & Me.Caption)
End Sub

Private Sub ChangeMode()
Select Case bytMode
    Case 1  '' View
        Call ViewAction
    Case 2  '' Add
        Call AddAction
    Case 3  '' Modify
        Call EditAction
End Select
End Sub

Private Sub MSF1_DblClick()
If TB1.TabEnabled(1) = False Then Exit Sub
TB1.Tab = 1
End Sub

Private Sub TB1_Click(PreviousTab As Integer)
If bytMode = 2 Then
    Exit Sub
End If
If PreviousTab = 1 Then
    Exit Sub
End If
MSF1.Col = 0
If MSF1.Text = NewCaptionTxt("00061", adrsMod) Then Exit Sub
Call Display

End Sub

Private Sub txtPunchDate_Click()
varCalDt = ""
varCalDt = Trim(txtPunchDate.Text)
txtPunchDate.Text = ""
Call ShowCalendar
End Sub

Private Sub txtPunchDate_GotFocus()
    Call GF(txtPunchDate)
End Sub

Private Sub txtPunchDate_KeyPress(KeyAscii As Integer)
     Call CDK(txtPunchDate, KeyAscii)
End Sub

Private Sub txtPunchDate_Validate(Cancel As Boolean)
    If Not ValidDate(txtPunchDate) Then txtPunchDate.SetFocus: Cancel = True
End Sub

Private Sub txtPunchTime_GotFocus(Index As Integer)
    Call GF(txtPunchTime(Index))
End Sub

Private Sub txtPunchTime_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtPunchTime(Index))
End If
End Sub

Private Sub ViewAction()    '' Procedure for Viewing Mode
'' Enable Needed Controls
cmdAddSave.Enabled = True       '' Enable ADD/SAVE Button
cmdEditCan.Enabled = True       '' Enable Edit/Cancel Button
cmdDel.Enabled = True           '' Enable Disable Button
'' Disable Needed Controls
txtPunchDate.Enabled = False    '' Disable Date TextBox
txtPunchTime(0).Enabled = False    '' Disable Time TextBox
txtPunchTime(1).Enabled = False    '' Disable Time TextBox
txtPunchTime(2).Enabled = False    '' Disable Time TextBox
txtPunchTime(3).Enabled = False    '' Disable Time TextBox

If typPerm.blnIO Then
    cboIO(0).Enabled = False
    cboIO(1).Enabled = False
    cboIO(2).Enabled = False
    cboIO(3).Enabled = False
End If
''
cboEmpCode.Enabled = True      '' Disable Employee Code Combo
'' Give Captions to the Needed Controls
Call SetGButtonCap(Me)
TB1.Tab = 0     '' Set the Tab to the First Tab
End Sub

Private Sub AddAction()     '' Procedure for Addition Mode
If TB1.TabEnabled(1) = False Then TB1.TabEnabled(1) = True
TB1.Tab = 1
'' Enable Necessary Controls
txtPunchDate.Enabled = True     '' Enable Date TextBox
txtPunchTime(0).Enabled = True     '' Enable Time TextBox
txtPunchTime(1).Enabled = True     '' Enable Time TextBox
txtPunchTime(2).Enabled = True     '' Enable Time TextBox
txtPunchTime(3).Enabled = True     '' Enable Time TextBox

If typPerm.blnIO Then
    cboIO(0).Enabled = True
    cboIO(1).Enabled = True
    cboIO(2).Enabled = True
    cboIO(3).Enabled = True
End If

''
cboEmpCode.Enabled = True       '' Enable Employee Code Combo
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
 '' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Clear Necessary Controls
txtPunchDate.Text = DateDisp(Date)  '' Clear Date Control
txtPunchTime(0).Text = ""              '' Clear Time Control
txtPunchTime(1).Text = ""              '' Clear Time Control
txtPunchTime(2).Text = ""              '' Clear Time Control
txtPunchTime(3).Text = ""              '' Clear Time Control
txtRemark.Text = ""
txtPunchDate.SetFocus                 '' Set Focus on the Employee ComboBox
End Sub

Private Sub EditAction()    '' Procedure for Edit Mode
'' Enable Necessary Controls
txtPunchDate.Enabled = True     '' Enable Date TextBox
txtPunchTime(0).Enabled = True     '' Enable Time TextBox
txtPunchTime(1).Enabled = True     '' Enable Time TextBox
txtPunchTime(2).Enabled = True     '' Enable Time TextBox
txtPunchTime(3).Enabled = True     '' Enable Time TextBox

If typPerm.blnIO Then
    cboIO(0).Enabled = True
    cboIO(1).Enabled = True
    cboIO(2).Enabled = True
    cboIO(3).Enabled = True
End If
''
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
txtPunchDate.SetFocus       '' Set Focus on the Date TextBox
If TB1.Tab = 1 Then

    Exit Sub
Else
    If TB1.TabEnabled(1) Then TB1.Tab = 1
End If

End Sub

Private Function ValidateAddmaster() As Boolean     '' Validate If in Addition Mode
On Error GoTo ERR_P
If cboEmpCode.Text = "" Then
    MsgBox NewCaptionTxt("32005", adrsC), vbExclamation
    cboEmpCode.SetFocus
    Exit Function
End If
If txtPunchDate.Text = "" Then
    MsgBox NewCaptionTxt("00072", adrsMod), vbExclamation
    txtPunchDate.SetFocus
    Exit Function
End If
If Val(txtPunchTime(0).Text) < 0 Then
    MsgBox NewCaptionTxt("32006", adrsC), vbExclamation
    txtPunchTime(0).SetFocus
    Exit Function
End If
If Trim(txtPunchTime(0).Text) = "" Then
    MsgBox "Time cannot be blank", vbExclamation
    txtPunchTime(0).SetFocus
    Exit Function
End If
''
''Check for More than 23.59
If Not LessThan2359(0) Then Exit Function
If Not LessThan2359(1) Then Exit Function
If Not LessThan2359(2) Then Exit Function
If Not LessThan2359(3) Then Exit Function
'' Check for Min. More than .59
If Not LessThan59(0) Then Exit Function
If Not LessThan59(1) Then Exit Function
If Not LessThan59(2) Then Exit Function
If Not LessThan59(3) Then Exit Function
ValidateAddmaster = True
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster :: " & Me.Caption)
End Function

Private Function LessThan2359(ByVal bytIndex As Byte) As Boolean
If Val(txtPunchTime(bytIndex).Text) > 23.59 Then
    MsgBox NewCaptionTxt("00025", adrsMod), vbExclamation
    txtPunchTime(bytIndex).SetFocus
    Exit Function
End If
LessThan2359 = True
End Function

Private Function LessThan59(ByVal bytIndex As Byte) As Boolean
If Val(txtPunchTime(bytIndex).Text) - Int(Val(txtPunchTime(bytIndex).Text)) > 0.59 Then
    MsgBox NewCaptionTxt("00024", adrsMod), vbExclamation
    txtPunchTime(bytIndex).SetFocus
    LessThan59 = False
    Exit Function
End If
LessThan59 = True
End Function

Private Function ValidateModMaster() As Boolean     '' Validate If in Edit Mode
On Error GoTo ERR_P
If txtPunchDate.Text = "" Then
    MsgBox NewCaptionTxt("00072", adrsMod), vbExclamation
    txtPunchDate.SetFocus
    Exit Function
End If

If Val(txtPunchTime(0).Text) < 0 Then
    MsgBox NewCaptionTxt("32006", adrsC), vbExclamation
    txtPunchTime(0).SetFocus
    Exit Function
End If
If Trim(txtPunchTime(0).Text) = "" Then
    MsgBox "Time cannot be blank", vbExclamation
    txtPunchTime(0).SetFocus
    Exit Function
End If
''
''Check for Time More than 23.59
If Not LessThan2359(0) Then Exit Function
'' Check for Min. More than .59
If Not LessThan59(0) Then Exit Function
ValidateModMaster = True
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
End Function

Private Function SaveAddMaster() As Boolean
On Error GoTo ERR_P
Dim bytCnt As Byte
Dim strqry As String
SaveAddMaster = True        '' Insert
' TO AVOID ENTRIES FOR SAME TIME ON SAME DATE
strLostT_Punch = txtPunchTime(0).Text
If adrsDept1.State = 1 Then adrsDept1.Close
strqry = "Select * from Lost where Empcode='" & Trim(cboEmpCode.Text) & "' and " & strKDate & " =" & strDTEnc & DateSaveIns(txtPunchDate.Text) & strDTEnc & " and t_punch=" & strLostT_Punch
adrsDept1.Open strqry, ConMain, adOpenStatic
'MsgBox adrsDept1.RecordCount
If adrsDept1.EOF = True Then
For bytCnt = 0 To 3
If Trim(txtPunchTime(bytCnt)) <> "" Then
    
        ConMain.Execute "insert into Lost (Empcode," & strKDate & ",t_punch,shift) values ('" & _
        Trim(lblCode.Caption) & "'," & strDTEnc & DateSaveIns(txtPunchDate.Text) & strDTEnc & "," & _
        txtPunchTime(bytCnt).Text & ",'" & cboIO(bytCnt).Text & "')"
    ''
    strLostDate = txtPunchDate.Text
    strLostT_Punch = txtPunchTime(0).Text
End If
Next
Else
    MsgBox "Punch Time Is Already Entered on This Date", vbCritical, Me.Caption
End If

Exit Function
ERR_P:
    SaveAddMaster = False
    ShowError ("SaveAddMaster :: " & Me.Caption)
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
SaveModMaster = True        '' Update


    ConMain.Execute "update lost set " & strKDate & "=" & strDTEnc & _
    DateSaveIns(txtPunchDate.Text) & strDTEnc & "," & "t_punch=" & _
    txtPunchTime(0).Text & ",shift='" & cboIO(0).Text & "' where Empcode=" & "'" & cboEmpCode.Text & "'" & " and " & strKDate & "=" & _
    strDTEnc & DateCompStr(strLostDate) & strDTEnc & " and " & "t_punch=" & strLostT_Punch

''
Exit Function
ERR_P:
    SaveModMaster = False
    ShowError ("SaveModMaster :: " & Me.Caption)
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub
