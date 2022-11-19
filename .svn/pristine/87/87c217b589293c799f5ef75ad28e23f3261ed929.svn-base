VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRotation 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command5"
      Height          =   435
      Left            =   5670
      TabIndex        =   16
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command5"
      Height          =   435
      Left            =   3780
      TabIndex        =   15
      Top             =   3480
      Width           =   1905
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command5"
      Height          =   435
      Left            =   1890
      TabIndex        =   14
      Top             =   3480
      Width           =   1905
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command5"
      Height          =   435
      Left            =   0
      TabIndex        =   13
      Top             =   3480
      Width           =   1905
   End
   Begin TabDlg.SSTab TB1 
      Height          =   3450
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   6085
      _Version        =   393216
      Tabs            =   2
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
      TabCaption(0)   =   "List"
      TabPicture(0)   =   "frmRotation.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSF1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Details"
      TabPicture(1)   =   "frmRotation.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frDetails"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame frDetails 
         Height          =   3045
         Left            =   -74940
         TabIndex        =   19
         Top             =   330
         Width           =   7500
         Begin VB.Frame frMisc 
            Height          =   630
            Left            =   45
            TabIndex        =   20
            Top             =   120
            Width           =   7425
            Begin MSMask.MaskEdBox txtName 
               Height          =   345
               Left            =   3450
               TabIndex        =   1
               Top             =   180
               Width           =   3540
               _ExtentX        =   6244
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               AutoTab         =   -1  'True
               MaxLength       =   29
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtCode 
               Height          =   345
               Left            =   1620
               TabIndex        =   0
               Top             =   180
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               AutoTab         =   -1  'True
               MaxLength       =   3
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
            Begin VB.Label lblCode 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rotation Code"
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
               Left            =   90
               TabIndex        =   21
               Top             =   210
               Width           =   1230
            End
            Begin VB.Label lblName 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
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
               Left            =   2670
               TabIndex        =   22
               Top             =   210
               Width           =   510
            End
         End
         Begin VB.Frame frRot 
            Caption         =   "Shift Rotates"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1650
            Left            =   45
            TabIndex        =   23
            Top             =   795
            Width           =   7425
            Begin VB.OptionButton optSND 
               Caption         =   "only after specified number of days"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   45
               TabIndex        =   2
               Top             =   240
               Width           =   3435
            End
            Begin VB.OptionButton optFD 
               Caption         =   "On the following dates of every month"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   60
               TabIndex        =   5
               Top             =   630
               Width           =   3735
            End
            Begin VB.OptionButton optWD 
               Caption         =   "On following week days (SUN..SAT)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   45
               TabIndex        =   8
               Top             =   1110
               Width           =   3540
            End
            Begin VB.TextBox txtFD 
               Appearance      =   0  'Flat
               Height          =   405
               Left            =   3900
               Locked          =   -1  'True
               TabIndex        =   6
               Text            =   " "
               Top             =   660
               Width           =   3075
            End
            Begin VB.TextBox txtWD 
               Appearance      =   0  'Flat
               Height          =   405
               Left            =   3900
               Locked          =   -1  'True
               TabIndex        =   9
               Text            =   " "
               Top             =   1140
               Width           =   3075
            End
            Begin VB.CommandButton cmdSND 
               Caption         =   ". . ."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   7020
               TabIndex        =   4
               ToolTipText     =   "Click to Select the Number of Days"
               Top             =   210
               Width           =   345
            End
            Begin VB.CommandButton cmdFD 
               Caption         =   ". . ."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   7020
               TabIndex        =   7
               ToolTipText     =   "Click to Select the Dates of Every Month"
               Top             =   690
               Width           =   345
            End
            Begin VB.CommandButton cmdWD 
               Caption         =   ". . ."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   7020
               TabIndex        =   10
               ToolTipText     =   "Click to Select the Week"
               Top             =   1155
               Width           =   345
            End
            Begin VB.TextBox txtSND 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   3900
               Locked          =   -1  'True
               TabIndex        =   3
               Text            =   " "
               Top             =   210
               Width           =   3075
            End
         End
         Begin VB.Frame frShift 
            Height          =   570
            Left            =   30
            TabIndex        =   24
            Top             =   2415
            Width           =   7440
            Begin VB.TextBox txtChange 
               Appearance      =   0  'Flat
               Height          =   420
               Left            =   3900
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   120
               Width           =   3090
            End
            Begin VB.CommandButton cmdChange 
               Caption         =   ". . ."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   7035
               TabIndex        =   12
               ToolTipText     =   "Click to Select the Shiffts"
               Top             =   150
               Width           =   345
            End
            Begin VB.Label lblChange 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "The Shift Changes from one to another"
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
               Left            =   75
               TabIndex        =   25
               Top             =   225
               Width           =   3330
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2835
         Left            =   30
         TabIndex        =   18
         Top             =   510
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   5001
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   2
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
   End
End
Attribute VB_Name = "frmRotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset

Private Sub cmdChange_Click()
strRotPass = txtChange.Text
frmSelectShift.Show vbModal
txtChange.Text = strRotPass
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFD_Click()
strCapSND = optFD.Caption
strRotPass = txtFD.Text
frmRotSND.Show vbModal
txtFD.Text = strRotPass
End Sub

Private Sub cmdSND_Click()
strCapSND = optSND.Caption
strRotPass = txtSND.Text
frmRotSND.Show vbModal
txtSND.Text = strRotPass
End Sub

Private Sub cmdWD_Click()
strCapSND = optWD.Caption
strRotPass = txtWD.Text
frmRotWD.Show vbModal
txtWD.Text = strRotPass
End Sub

Private Sub optFD_Click()
txtFD.Visible = True
cmdFD.Visible = True
txtSND.Visible = False
txtWD.Visible = False
cmdSND.Visible = False
cmdWD.Visible = False
End Sub

Private Sub optSND_Click()
txtSND.Visible = True
cmdSND.Visible = True
txtFD.Visible = False
txtWD.Visible = False
cmdFD.Visible = False
cmdWD.Visible = False
End Sub

Private Sub optWD_Click()
txtWD.Visible = True
cmdWD.Visible = True
txtFD.Visible = False
txtSND.Visible = False
cmdFD.Visible = False
cmdSND.Visible = False
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)        '' Set the Form Icon
Call RetCaption             '' Retreive Captions
Call OpenMasterTable        '' Open Master Table
Call FillGrid               '' Fill Grid
TB1.Tab = 0                 '' Set the Tab to List
Call GetRights              '' Gets Rights for the Operations
bytMode = 1
Call ChangeMode             '' Take Action on the Appropriate Mode
End Sub

Private Sub RetCaption()
On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '41%'", ConMain, adOpenStatic, adLockReadOnly
Me.Caption = NewCaptionTxt("41001", adrsC)
TB1.TabCaption(0) = NewCaptionTxt("00013", adrsMod)
TB1.TabCaption(1) = NewCaptionTxt("00014", adrsMod)
Call SetOtherCaps                           '' Other Captions of other Controls
Call SetGButtonCap(Me)                           '' Sets Appropriate Caption for the Buttons
Call CapGrid                                '' Captions For Grid
End Sub

Private Sub SetOtherCaps()
lblCode.Caption = NewCaptionTxt("41002", adrsC)
lblName.Caption = NewCaptionTxt("00048", adrsMod)
frRot.Caption = NewCaptionTxt("41003", adrsC)
optSND.Caption = NewCaptionTxt("41004", adrsC)
optFD.Caption = NewCaptionTxt("41005", adrsC)
optWD.Caption = NewCaptionTxt("41006", adrsC)
lblChange.Caption = NewCaptionTxt("41007", adrsC)
cmdSND.ToolTipText = NewCaptionTxt("41018", adrsC)
cmdFD.ToolTipText = NewCaptionTxt("41019", adrsC)
cmdWD.ToolTipText = NewCaptionTxt("41020", adrsC)
cmdChange.ToolTipText = NewCaptionTxt("41021", adrsC)
End Sub

Private Sub CapGrid()           '' Gives the Captions to the Grid
With MSF1
    '' Sets the Column Widhts
    .ColWidth(0) = .ColWidth(0) * 0.75
    .ColWidth(1) = .ColWidth(1) * 1.1
    .ColWidth(2) = .ColWidth(2) * 0.85
    .ColWidth(3) = .ColWidth(3) * 2
    .ColWidth(4) = .ColWidth(4) * 2
    '' Sets the Column Alignment
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    .ColAlignment(4) = flexAlignLeftCenter
    '' Sets the Appropriate Captions
    .TextMatrix(0, 0) = NewCaptionTxt("00047", adrsMod)
    .TextMatrix(0, 1) = NewCaptionTxt("00048", adrsMod)
    .TextMatrix(0, 2) = NewCaptionTxt("41008", adrsC)
    .TextMatrix(0, 3) = NewCaptionTxt("41009", adrsC)
    .TextMatrix(0, 4) = NewCaptionTxt("41010", adrsC)
End With
End Sub

Private Sub OpenMasterTable()             '' Open the recordset for the Display purposes
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select SCode,Name,Skp,Pattern,Mon_Oth,Tot_Shf,Tot_Skp,Day_Skp " & _
"From Ro_Shift where scode <> '100' Order by SCode", ConMain, adOpenStatic
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Sub FillGrid()          '' Fills the Grid
On Error GoTo ERR_P
Dim intCounter As Integer
adrsDept1.Requery               '' Requeries the Recordset for any Updated Values
'' Put Appropriate Rows in the Grid
If adrsDept1.EOF Then
    MSF1.Rows = 1
    TB1.TabEnabled(1) = False       '' Disables Tab 1 If no Records are Found
    Exit Sub
End If
MSF1.Rows = adrsDept1.RecordCount + 1   '' Sets Rows Appropriately
adrsDept1.MoveFirst
For intCounter = 1 To adrsDept1.RecordCount     '' Fills the Grid
    With MSF1           '' 0 1 4 2 3
        .TextMatrix(intCounter, 0) = adrsDept1("SCode")
        .TextMatrix(intCounter, 1) = IIf(IsNull(adrsDept1("Name")), "", adrsDept1("Name"))
        .TextMatrix(intCounter, 2) = IIf(IsNull(adrsDept1("Mon_Oth")), "", adrsDept1("Mon_Oth"))
        .TextMatrix(intCounter, 3) = IIf(IsNull(adrsDept1("Skp")), "", adrsDept1("Skp"))
        .TextMatrix(intCounter, 4) = IIf(IsNull(adrsDept1("Pattern")), "", adrsDept1("Pattern"))
    End With
    adrsDept1.MoveNext
Next
TB1.TabEnabled(1) = True        '' If Records Found Enables the TAB 1
TB1.Tab = 0
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub GetRights()     '' Gets and Sets the Appropriate Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 8)
If Mid(strTmp, 1, 1) = "1" Then AddRights = True
If Mid(strTmp, 2, 1) = "1" Then EditRights = True
If Mid(strTmp, 3, 1) = "1" Then DeleteRights = True
Exit Sub
ERR_P:
    ShowError ("GetRights ::" & Me.Caption)
    AddRights = False
    EditRights = False
    DeleteRights = False
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

Private Sub ViewAction()    '' Procedure for Viewing Mode
'' Enable Needed Controls
cmdAddSave.Enabled = True       '' Enable ADD/SAVE Button
cmdEditCan.Enabled = True       '' Enable Edit/Cancel Button
cmdDel.Enabled = True           '' Enable Disable Button
'' Disable Needed Controls
txtCode.Enabled = False         '' Disable Code TextBox
txtName.Enabled = False         '' Disable Name TextBox
frMisc.Enabled = False          '' Disable Miscellaneous Frame
frRot.Enabled = False           '' Disable Rotation Frame
frshift.Enabled = False         '' Disable Shift Frame
'' Give Captions to the Needed Controls
Call SetGButtonCap(Me)
TB1.Tab = 0     '' Set the Tab to the First Tab
End Sub

Private Sub AddAction()     '' Procedure for Addition Mode
If TB1.TabEnabled(1) = False Then TB1.TabEnabled(1) = True
TB1.Tab = 1
'' Enable Necessary Controls
txtCode.Enabled = True          '' Enable Code TextBox
txtName.Enabled = True          '' Enable Name TextBox
frMisc.Enabled = True           '' Enable Miscellaneous Frame
frRot.Enabled = True            '' Enable Rotation Frame
frshift.Enabled = True          '' Enable Shift Frame
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Clear Necessary Controls
txtCode.Text = ""               '' Clear Code TextBox
txtName.Text = ""               '' Clear Description TextBox
txtSND.Text = ""                '' Clear SND TextBox
txtFD.Text = ""                 '' CLear FD TextBox
txtWD.Text = ""                 '' Clear WD TextBox
txtChange.Text = ""             '' Clear Shifts TextBox
optSND.Value = True             '' Set Value to SND Option Button
txtCode.SetFocus                '' Set Focus to the Code
End Sub

Private Sub EditAction()    '' Procedure for Edit Mode
'' Enable Necessary Controls
frMisc.Enabled = True       '' Enable Miscellaneous Frame
frRot.Enabled = True        '' Enable Rotation Frame
frshift.Enabled = True      '' Enable Shifts Frame
txtName.Enabled = True      '' Enable Name TextBox
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
txtName.SetFocus                '' Set Focus on the Name TextBox
If TB1.Tab = 1 Then
    Exit Sub
Else
    If TB1.TabEnabled(1) Then TB1.Tab = 1
End If
End Sub

Private Sub TB1_Click(PreviousTab As Integer)
If bytMode = 2 Then Exit Sub
If PreviousTab = 1 Then Exit Sub
MSF1.Col = 0
If MSF1.Text = NewCaptionTxt("00047", adrsMod) Then Exit Sub
Call Display
End Sub

Private Sub Display()       '' Displays the Given Master Records
On Error GoTo ERR_P
txtCode.Text = MSF1.TextMatrix(MSF1.Row, 0)     '' Rotation Code
txtName.Text = MSF1.TextMatrix(MSF1.Row, 1)     '' Rotation Name
adrsDept1.MoveFirst
adrsDept1.Find "SCode='" & txtCode.Text & "'"
If adrsDept1.EOF Then
    bytMode = 1
    Call ChangeMode
Else
    Select Case adrsDept1("Mon_Oth")            '' Mon_Oth
        Case "O"
            optSND.Value = True
            txtSND.Text = adrsDept1("Skp")
        Case "D"
            optFD.Value = True
            txtFD.Text = adrsDept1("Skp")
        Case "W"
            optWD.Value = True
            txtWD.Text = adrsDept1("Skp")
    End Select
    txtChange.Text = adrsDept1("Pattern")       '' Pattern
End If
Exit Sub
ERR_P:
    ShowError ("Display  :: " & Me.Caption)
End Sub

Private Sub MSF1_DblClick()
If TB1.TabEnabled(1) = False Then Exit Sub
TB1.Tab = 1
End Sub

Private Function ValidateAddmaster() As Boolean     '' Validate If in Addition Mode
On Error GoTo ERR_P
ValidateAddmaster = True
If Trim(txtCode.Text) = "" Then
    MsgBox NewCaptionTxt("41022", adrsC), vbExclamation
    txtCode.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If MSF1.Rows > 1 Then
    adrsDept1.MoveFirst
    adrsDept1.Find "SCode='" & txtCode.Text & "'"
    If Not adrsDept1.EOF Then
        MsgBox NewCaptionTxt("41011", adrsC), vbExclamation
        txtCode.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If
If Trim(txtName.Text) = "" Then
    MsgBox NewCaptionTxt("41012", adrsC), vbExclamation
    txtName.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If optSND.Value = True And txtSND.Text = "" Then
    MsgBox NewCaptionTxt("41013", adrsC), vbExclamation
    cmdSND.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If optFD.Value = True And txtFD.Text = "" Then
    MsgBox NewCaptionTxt("41014", adrsC), vbExclamation
    cmdFD.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If optWD.Value = True And txtWD.Text = "" Then
    MsgBox NewCaptionTxt("41015", adrsC), vbExclamation
    cmdWD.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If txtChange.Text = "" Then
    MsgBox NewCaptionTxt("41016", adrsC), vbExclamation
    cmdChange.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster :: " & Me.Caption)
    ValidateAddmaster = False
End Function

Private Function ValidateModMaster() As Boolean     '' Validate If in Addition Mode
On Error GoTo ERR_P
ValidateModMaster = True
If Trim(txtName.Text) = "" Then
    MsgBox NewCaptionTxt("41012", adrsC), vbExclamation
    txtName.SetFocus
    ValidateModMaster = False
    Exit Function
End If
If optSND.Value = True And txtSND.Text = "" Then
    MsgBox NewCaptionTxt("41013", adrsC), vbExclamation
    cmdSND.SetFocus
    ValidateModMaster = False
    Exit Function
End If
If optFD.Value = True And txtFD.Text = "" Then
    MsgBox NewCaptionTxt("41014", adrsC), vbExclamation
    cmdFD.SetFocus
    ValidateModMaster = False
    Exit Function
End If
If optWD.Value = True And txtWD.Text = "" Then
    MsgBox NewCaptionTxt("41015", adrsC), vbExclamation
    cmdWD.SetFocus
    ValidateModMaster = False
    Exit Function
End If
If txtChange.Text = "" Then
    MsgBox NewCaptionTxt("41016", adrsC), vbExclamation
    cmdChange.SetFocus
    ValidateModMaster = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
    ValidateModMaster = False
End Function

Private Function SaveAddMaster() As Boolean
On Error GoTo ERR_P
Dim strTypeofRot As String
Dim strRotShift As String
Dim bytTotalShiftsTmp As Byte       '' Total Shifts
Dim bytTotalSkipsTmp As Byte        '' Total Skips
Dim bytTotalDaysTmp As Byte         '' Total Number of Days
Dim strTmp() As String              '' Temporary Array
Dim bytCnt As Byte
SaveAddMaster = True        '' Insert
'' Calculate Numbers
If optSND.Value = True Then
    strTypeofRot = "O"
    strRotShift = txtSND.Text
    strTmp = Split(Trim(txtChange.Text), ".")
    bytTotalShiftsTmp = UBound(strTmp)          '' Total Number of Shifts
    strTmp = Split(Trim(txtSND.Text), ",")
    bytTotalSkipsTmp = UBound(strTmp)          '' Total Number of Skips
    For bytCnt = 0 To UBound(strTmp) - 1
        bytTotalDaysTmp = bytTotalDaysTmp + Val(strTmp(bytCnt))
    Next
End If
If optFD.Value = True Then
    strTypeofRot = "D"
    strRotShift = txtFD.Text
    strTmp = Split(Trim(txtChange.Text), ".")
    bytTotalShiftsTmp = UBound(strTmp)          '' Total Number of Shifts
    strTmp = Split(Trim(txtFD.Text), ",")
    bytTotalSkipsTmp = UBound(strTmp)          '' Total Number of Skips
    For bytCnt = 0 To UBound(strTmp) - 1
        bytTotalDaysTmp = bytTotalDaysTmp + Val(strTmp(bytCnt))
    Next
End If
If optWD.Value = True Then
    strTypeofRot = "W"
    strRotShift = txtWD.Text
    bytTotalShiftsTmp = 1
    bytTotalSkipsTmp = 1
    bytTotalDaysTmp = 1
End If
ConMain.Execute "insert into Ro_Shift Values('" & txtCode.Text & "','" & _
Trim(txtName.Text) & "','" & strRotShift & "','" & txtChange.Text & "','" & strTypeofRot & _
"'," & bytTotalShiftsTmp & "," & bytTotalSkipsTmp & "," & bytTotalDaysTmp & ")"
Exit Function
ERR_P:
    Select Case Err.Number
        Case -2147217900
            MsgBox NewCaptionTxt("41017", adrsC), vbExclamation
        Case Else
            ShowError ("SaveAddMaster :: " & Me.Caption)
    End Select
    SaveAddMaster = False
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
Dim strTypeofRot As String
Dim strRotShift As String
Dim bytTotalShiftsTmp As Byte       '' Total Shifts
Dim bytTotalSkipsTmp As Byte        '' Total Skips
Dim bytTotalDaysTmp As Byte         '' Total Number of Days
Dim strTmp() As String              '' Temporary Array
Dim bytCnt As Byte
SaveModMaster = True        '' Update
'' Calculate Numbers
If optSND.Value = True Then
    strTypeofRot = "O"
    strRotShift = txtSND.Text
    strTmp = Split(Trim(txtChange.Text), ".")
    bytTotalShiftsTmp = UBound(strTmp)          '' Total Number of Shifts
    strTmp = Split(Trim(txtSND.Text), ",")
    bytTotalSkipsTmp = UBound(strTmp)          '' Total Number of Skips
    For bytCnt = 0 To UBound(strTmp) - 1
        bytTotalDaysTmp = bytTotalDaysTmp + Val(strTmp(bytCnt))
    Next
End If
If optFD.Value = True Then
    strTypeofRot = "D"
    strRotShift = txtFD.Text
    strTmp = Split(Trim(txtChange.Text), ".")
    bytTotalShiftsTmp = UBound(strTmp)          '' Total Number of Shifts
    strTmp = Split(Trim(txtFD.Text), ",")
    bytTotalSkipsTmp = UBound(strTmp)          '' Total Number of Skips
    For bytCnt = 0 To UBound(strTmp) - 1
        bytTotalDaysTmp = bytTotalDaysTmp + Val(strTmp(bytCnt))
    Next
End If
If optWD.Value = True Then
    strTypeofRot = "W"
    strRotShift = txtWD.Text
    bytTotalShiftsTmp = 1
    bytTotalSkipsTmp = 1
    bytTotalDaysTmp = 1
End If

ConMain.Execute "update Ro_Shift set Name='" & Trim(txtName.Text) & "'," & _
"Skp='" & strRotShift & "',Pattern='" & txtChange.Text & "',Mon_Oth='" & strTypeofRot & _
"',tot_shf = " & bytTotalShiftsTmp & ",tot_skp=" & bytTotalSkipsTmp & ", day_skp=" & _
bytTotalDaysTmp & " Where SCode='" & txtCode.Text & "'"

Exit Function
ERR_P:
    SaveModMaster = False
    ShowError ("SaveModMaster :: " & Me.Caption)
End Function

Private Sub txtCode_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 6))))
End Select
End Sub

Private Sub txtCode_LostFocus()
If txtCode.Text = "100" Then
 MsgBox " This Rotation Code is reserved for Application"
 txtCode.Text = ""
 txtCode.SetFocus
 End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 5))))
End Select
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
        Call SaveAddLog                         '' Save the Add Log
        Call FillGrid                           '' Reflect the Grid
        bytMode = 1
        Call ChangeMode
    Case 3          '' Edit Mode
        If Not ValidateModMaster Then Exit Sub  '' Validate for Edit
        If Not SaveModMaster Then Exit Sub      '' Save for Edit
        Call SaveModLog                         '' Save the Edit Log
        Call FillGrid                           '' Reflect the Grid
        bytMode = 1
        Call ChangeMode
End Select
Exit Sub
ERR_P:
    ShowError ("AddSave :: " & Me.Caption)
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
    If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion) = vbYes Then            '' Delete the Record
        ConMain.Execute "Delete from Ro_Shift where SCode='" & _
        txtCode.Text & "'"
        Call AddActivityLog(lgDelete_Action, 1, 11)     '' Delete Log
        Call AuditInfo("DELETE", Me.Caption, "Deleted Rotation Shift: " & txtCode.Text)
    End If
    Call FillGrid       '' Reflect the Grid
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
ERR_P:
    If Err.Number = -2147467259 Or Err.Number = -2147217900 Or Err.Number = -2147217873 Then
            MsgBox "Rotation Shift Cannot be deleted because employees belong to this Rotating Shift.", vbCritical, Me.Caption
            Exit Sub
    End If
    ShowError ("Delete Record :: " & Me.Caption)
End Sub

Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo ERR_P
Call AddActivityLog(lgADD_MODE, 1, 11)     '' Add Activity
Call AuditInfo("ADD", Me.Caption, "Added Rotation Shift: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 11)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edited Rotation Shift: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub
