VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmOTRul 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OT Rules"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOTRul.frx":0000
   ScaleHeight     =   4575
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frRound 
      Caption         =   "Round-Off OT"
      Height          =   1515
      Left            =   0
      TabIndex        =   72
      Top             =   2520
      Width           =   7455
      Begin VB.TextBox txtRF4 
         Height          =   285
         Left            =   3930
         TabIndex        =   16
         Top             =   540
         Width           =   555
      End
      Begin VB.TextBox txtRT4 
         Height          =   285
         Left            =   3930
         TabIndex        =   22
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox txtR4 
         Height          =   285
         Left            =   3930
         TabIndex        =   27
         Top             =   1140
         Width           =   555
      End
      Begin VB.TextBox txtRT2 
         Height          =   285
         Left            =   2370
         TabIndex        =   20
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox txtRF2 
         Height          =   285
         Left            =   2370
         TabIndex        =   14
         Top             =   540
         Width           =   555
      End
      Begin VB.TextBox txtR3 
         Height          =   285
         Left            =   3150
         TabIndex        =   26
         Top             =   1140
         Width           =   555
      End
      Begin VB.TextBox txtR2 
         Height          =   285
         Left            =   2370
         TabIndex        =   25
         Top             =   1140
         Width           =   555
      End
      Begin VB.TextBox txtR1 
         Height          =   285
         Left            =   1590
         TabIndex        =   24
         Top             =   1140
         Width           =   555
      End
      Begin VB.TextBox txtR5 
         Height          =   285
         Left            =   4800
         TabIndex        =   28
         Top             =   1140
         Width           =   555
      End
      Begin VB.TextBox txtRT3 
         Height          =   285
         Left            =   3150
         TabIndex        =   21
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox txtRT1 
         Height          =   285
         Left            =   1590
         TabIndex        =   19
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox txtRT5 
         Height          =   285
         Left            =   4800
         TabIndex        =   23
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox txtRF3 
         Height          =   285
         Left            =   3150
         TabIndex        =   15
         Top             =   540
         Width           =   555
      End
      Begin VB.TextBox txtRF1 
         Height          =   285
         Left            =   1590
         TabIndex        =   13
         Top             =   540
         Width           =   555
      End
      Begin VB.Label lblRSpec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "While Rounding OT only MINUTES part will be rounded , leaving the hours part as it is"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   150
         TabIndex        =   73
         Top             =   240
         Width           =   6090
      End
      Begin VB.Label lblRTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   900
         TabIndex        =   17
         Top             =   900
         Width           =   195
      End
      Begin VB.Label lblRFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   780
         TabIndex        =   74
         Top             =   630
         Width           =   345
      End
      Begin VB.Label lblRMore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "More Than"
         Height          =   195
         Left            =   4680
         TabIndex        =   76
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lblRound 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Round upto"
         Height          =   195
         Left            =   270
         TabIndex        =   75
         Top             =   1140
         Width           =   840
      End
   End
   Begin VB.TextBox txtRDesc 
      Height          =   285
      Left            =   3600
      MaxLength       =   49
      TabIndex        =   1
      Top             =   90
      Width           =   3825
   End
   Begin VB.Frame frDeductions 
      Caption         =   "Deduct the following hours from Basic OT"
      Height          =   2475
      Left            =   7200
      TabIndex        =   47
      Top             =   4080
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Frame fraDedBrHrs 
         Caption         =   "Deduct Lunch Break Hrs"
         Height          =   1575
         Left            =   5450
         TabIndex        =   77
         Top             =   480
         Width           =   1950
         Begin VB.TextBox txtHLLBHrs 
            Height          =   285
            Left            =   960
            TabIndex        =   83
            Top             =   1080
            Width           =   585
         End
         Begin VB.TextBox txtWOLBHrs 
            Height          =   285
            Left            =   960
            TabIndex        =   82
            Top             =   720
            Width           =   585
         End
         Begin VB.TextBox txtWDLBHrs 
            Height          =   285
            Left            =   960
            TabIndex        =   81
            Top             =   360
            Width           =   585
         End
         Begin VB.Label lblHLLBH 
            AutoSize        =   -1  'True
            Caption         =   "Holiday"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   1080
            Width           =   525
         End
         Begin VB.Label lblWOLBH 
            AutoSize        =   -1  'True
            Caption         =   "Weekoff"
            Height          =   195
            Left            =   120
            TabIndex        =   79
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblWDLBH 
            AutoSize        =   -1  'True
            Caption         =   "Weekday"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   360
            Width           =   690
         End
      End
      Begin VB.CheckBox chkA4 
         Height          =   195
         Left            =   4770
         TabIndex        =   69
         Top             =   1740
         Width           =   225
      End
      Begin VB.CheckBox chkA3 
         Height          =   195
         Left            =   3870
         TabIndex        =   65
         Top             =   1740
         Width           =   225
      End
      Begin VB.CheckBox chkA2 
         Height          =   195
         Left            =   3060
         TabIndex        =   61
         Top             =   1740
         Width           =   225
      End
      Begin VB.CheckBox chkA1 
         Height          =   195
         Left            =   2250
         TabIndex        =   57
         Top             =   1740
         Width           =   225
      End
      Begin VB.CheckBox chkDedHL 
         Caption         =   "Apply deduction on Holiday"
         Height          =   195
         Left            =   3630
         TabIndex        =   71
         Top             =   2130
         Width           =   2325
      End
      Begin VB.CheckBox chkDedWO 
         Caption         =   "Apply deduction on Weekoff"
         Height          =   195
         Left            =   180
         TabIndex        =   70
         Top             =   2130
         Width           =   2355
      End
      Begin VB.TextBox txtF1 
         Height          =   285
         Left            =   2220
         TabIndex        =   50
         Top             =   570
         Width           =   555
      End
      Begin VB.TextBox txtT1 
         Height          =   285
         Left            =   2220
         TabIndex        =   52
         Top             =   870
         Width           =   555
      End
      Begin VB.TextBox txtD1 
         Height          =   285
         Left            =   2220
         TabIndex        =   54
         Top             =   1170
         Width           =   555
      End
      Begin VB.TextBox txtF2 
         Height          =   285
         Left            =   3030
         TabIndex        =   58
         Top             =   570
         Width           =   555
      End
      Begin VB.TextBox txtT2 
         Height          =   285
         Left            =   3030
         TabIndex        =   59
         Top             =   870
         Width           =   555
      End
      Begin VB.TextBox txtD2 
         Height          =   285
         Left            =   3030
         TabIndex        =   60
         Top             =   1170
         Width           =   555
      End
      Begin VB.TextBox txtF3 
         Height          =   285
         Left            =   3840
         TabIndex        =   62
         Top             =   570
         Width           =   555
      End
      Begin VB.TextBox txtT3 
         Height          =   285
         Left            =   3840
         TabIndex        =   63
         Top             =   870
         Width           =   555
      End
      Begin VB.TextBox txtT4 
         Height          =   285
         Left            =   4740
         TabIndex        =   67
         Top             =   870
         Width           =   555
      End
      Begin VB.TextBox txtD3 
         Height          =   285
         Left            =   3840
         TabIndex        =   64
         Top             =   1170
         Width           =   555
      End
      Begin VB.TextBox txtD4 
         Height          =   285
         Left            =   4740
         TabIndex        =   68
         Top             =   1170
         Width           =   555
      End
      Begin VB.Label lblOR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--OR--"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   570
         TabIndex        =   55
         Top             =   1470
         Width           =   525
      End
      Begin VB.Label lblDSpec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basic OT will be calculated after the LATE-EARLY calculations specified in category Master"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   150
         TabIndex        =   48
         Top             =   270
         Width           =   6480
      End
      Begin VB.Label lblAll 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deduct all OT"
         Height          =   195
         Left            =   480
         TabIndex        =   56
         Top             =   1710
         Width           =   990
      End
      Begin VB.Label lblDeduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deduct specified"
         Height          =   195
         Left            =   270
         TabIndex        =   53
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblMore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "More Than"
         Height          =   195
         Left            =   4620
         TabIndex        =   66
         Top             =   630
         Width           =   780
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   1110
         TabIndex        =   49
         Top             =   600
         Width           =   345
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   1260
         TabIndex        =   51
         Top             =   900
         Width           =   195
      End
   End
   Begin VB.Frame frCheck 
      Caption         =   "Give OT on "
      Height          =   1395
      Left            =   0
      TabIndex        =   40
      Top             =   450
      Width           =   5205
      Begin VB.TextBox txtHLRate 
         Height          =   285
         Left            =   1710
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1050
         Width           =   555
      End
      Begin VB.TextBox txtWORate 
         Height          =   285
         Left            =   1710
         MaxLength       =   4
         TabIndex        =   5
         Top             =   660
         Width           =   555
      End
      Begin VB.TextBox txtWDRate 
         Height          =   285
         Left            =   1710
         MaxLength       =   4
         TabIndex        =   3
         Top             =   300
         Width           =   555
      End
      Begin VB.CheckBox chkHL 
         Caption         =   "Holiday @"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   1080
         Width           =   1305
      End
      Begin VB.CheckBox chkWO 
         Caption         =   "Weekoff @"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   720
         Width           =   1365
      End
      Begin VB.CheckBox chkWD 
         Caption         =   "Weekdays @"
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   330
         Width           =   1455
      End
      Begin VB.Label lblTimes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "times, total work hours of that day."
         Height          =   195
         Index           =   2
         Left            =   2370
         TabIndex        =   43
         Top             =   1050
         Width           =   2415
      End
      Begin VB.Label lblTimes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "times, total work hours of that day."
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   42
         Top             =   690
         Width           =   2415
      End
      Begin VB.Label lblTimes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "times, total work hours of that day."
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   41
         Top             =   330
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   435
      Left            =   5310
      TabIndex        =   33
      Top             =   4110
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   3990
      TabIndex        =   32
      Top             =   4110
      Width           =   1305
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   435
      Left            =   2670
      TabIndex        =   31
      Top             =   4110
      Width           =   1305
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   435
      Left            =   1350
      TabIndex        =   30
      Top             =   4110
      Width           =   1305
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   435
      Left            =   30
      TabIndex        =   29
      Top             =   4110
      Width           =   1305
   End
   Begin VB.Frame frMaxOT 
      Caption         =   "Maximum OT can be upto"
      Height          =   705
      Left            =   5220
      TabIndex        =   45
      Top             =   1140
      Width           =   2235
      Begin VB.TextBox txtMaxOT 
         Height          =   285
         Left            =   150
         TabIndex        =   10
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lblHours 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hours"
         Height          =   195
         Left            =   900
         TabIndex        =   18
         Top             =   330
         Width           =   420
      End
   End
   Begin VB.Frame frAuthorized 
      Caption         =   "Authorized by default"
      Height          =   675
      Left            =   5220
      TabIndex        =   44
      Top             =   450
      Width           =   2235
      Begin VB.OptionButton OptNo 
         Caption         =   "NO"
         Height          =   195
         Left            =   930
         TabIndex        =   9
         Top             =   300
         Width           =   615
      End
      Begin VB.OptionButton OptYes 
         Caption         =   "YES"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame frLateEarly 
      Caption         =   "Late-Early Deductions"
      Height          =   585
      Left            =   0
      TabIndex        =   46
      Top             =   1860
      Width           =   7455
      Begin VB.CheckBox chkEarly 
         Caption         =   "Deduct Early Hours from OT"
         Height          =   195
         Left            =   4260
         TabIndex        =   12
         Top             =   240
         Width           =   3015
      End
      Begin VB.CheckBox chkLate 
         Caption         =   "Deduct Late Hours from OT"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3435
      End
   End
   Begin VB.Frame frNewRule 
      Height          =   495
      Left            =   5100
      TabIndex        =   36
      Top             =   30
      Visible         =   0   'False
      Width           =   1605
      Begin VB.TextBox txtRNo 
         Height          =   285
         Left            =   1050
         MaxLength       =   2
         TabIndex        =   38
         Top             =   150
         Width           =   465
      End
      Begin VB.Label lblRNo 
         AutoSize        =   -1  'True
         Caption         =   "OT Rule No."
         Height          =   195
         Left            =   60
         TabIndex        =   37
         Top             =   180
         Width           =   900
      End
   End
   Begin VB.Frame frRule 
      Height          =   495
      Left            =   0
      TabIndex        =   34
      Top             =   -60
      Width           =   2205
      Begin MSForms.ComboBox cboRule 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   150
         Width           =   795
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1402;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblRules 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OT Rules"
         Height          =   195
         Left            =   90
         TabIndex        =   35
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.Label lblRDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   2340
      TabIndex        =   39
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "frmOTRul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset

Private Sub cboRule_Click()
Call Display
End Sub

Private Sub cboRule_LostFocus()
If cboRule.Text = "100" Then
 MsgBox " This OT Rule is reserved for Application"
 cboRule.Text = ""
 cboRule.SetFocus
 End If
End Sub


Private Sub chkWD_Click()
If chkWD.Value = 1 Then
    txtWDRate.Enabled = True
Else
    txtWDRate.Enabled = False
    
    txtWDRate.Text = "1"
End If
End Sub

Private Sub chkWO_Click()
If chkWO.Value = 1 Then
    txtWORate.Enabled = True
Else
    txtWORate.Enabled = False
    
    txtWORate.Text = "1"
End If
End Sub

Private Sub chkHL_Click()
If chkHL.Value = 1 Then
    txtHLRate.Enabled = True
Else
    txtHLRate.Enabled = False
    
    txtHLRate.Text = "1"
End If
End Sub

Private Sub cmdAdd_Click()
On Error GoTo ERR_P
If Not AddRights Then
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    Exit Sub
Else
    bytMode = 2
    Call ChangeMode
End If
Exit Sub
ERR_P:
    ShowError ("Add::" & Me.Caption)
End Sub

Private Sub cmdCancel_Click()
Select Case bytMode
    Case 1  '' Edit / View
         Call Display
    Case 2  '' Add
        If cboRule.ListCount > 0 Then
            bytMode = 1
            Call ChangeMode
            Call Display
        End If
End Select
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ERR_P
If Not DeleteRights Then
    MsgBox NewCaptionTxt("00001", adrsMod)
    Exit Sub
Else
    If cboRule.Visible Then
        If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion) = vbYes Then
            ConMain.Execute "Delete From OTRul Where OTCode=" & cboRule.Text
            Call AuditInfo("DELETE", Me.Caption, "Delete OT Rule no.: " & cboRule.Text)
            Call OpenMasterTable
            Call FillCombo
            If cboRule.ListCount = 0 Then
                If Not AddRights Then
                    cmdAdd.Enabled = False
                    cmdSave.Enabled = False
                    cmdCancel.Enabled = False
                    cmdDelete.Enabled = False
                Else
                    bytMode = 2     '' Add
                    Call ChangeMode
                End If
            Else
                bytMode = 1     '' Edit / View
                cboRule.ListIndex = 0
            End If
        End If
    End If
End If
Exit Sub
ERR_P:
    If Err.Number = -2147467259 Or Err.Number = -2147217900 Or Err.Number = -2147217873 Then
            MsgBox "OTRule Cannot be deleted because employees belong to this OTRule.", vbCritical, Me.Caption
            Exit Sub
    End If
    ShowError ("Delete Record :: " & Me.Caption)
End Sub

Private Sub cmdsave_Click()
Select Case bytMode
    Case 1  '' Edit / View
        If Not EditRights Then
            MsgBox NewCaptionTxt("00001", adrsMod)
        Else
            If Not ValidateModMaster Then Exit Sub
            If Not SaveModMaster Then Exit Sub
            Call AuditInfo("UPDATE", Me.Caption, "Edit OT Rule no.: " & cboRule.Text)
            Call OpenMasterTable
        End If
    Case 2  '' Add
        If Not ValidateAddmaster Then Exit Sub
        If Not SaveAddMaster Then Exit Sub
        Call AuditInfo("ADD", Me.Caption, "Added OT Rule no.: " & cboRule.Text)
        Call OpenMasterTable
        Call FillCombo
        bytMode = 1
        Call ChangeMode
        If cboRule.ListCount > 0 Then cboRule.ListIndex = 0
End Select
End Sub

Private Function ValidateModMaster() As Boolean
On Error GoTo ERR_P
If cboRule.Text = "" Then Exit Function
If adrsDept1.RecordCount <= 0 Then
    MsgBox NewCaptionTxt("59018", adrsC)
    Exit Function
End If
adrsDept1.MoveFirst
adrsDept1.Find "OTCode=" & cboRule.Text
If adrsDept1.EOF Then
    MsgBox NewCaptionTxt("59018", adrsC)
    Exit Function
End If
If Trim(txtRDesc.Text) = "" Then
    MsgBox NewCaptionTxt("59015", adrsC)
    txtRDesc.SetFocus
    Exit Function
End If

If chkWD.Value = 1 And (txtWDRate.Text = "" Or txtWDRate.Text = 0) Then
    MsgBox "Rate value cannot be 0 for selected OT Day", vbInformation
    txtWDRate.SetFocus
    Exit Function
End If
If chkWO.Value = 1 And (txtWORate.Text = "" Or txtWORate.Text = 0) Then
    MsgBox "Rate value cannot be 0 for selected OT Day", vbInformation
    txtWORate.SetFocus
    Exit Function
End If
If chkHL.Value = 1 And (txtHLRate.Text = "" Or txtHLRate.Text = 0) Then
    MsgBox "Rate value cannot be 0 for selected OT Day", vbInformation
    txtHLRate.SetFocus
    Exit Function
End If
''
If Not CorrectDecs Then Exit Function
ValidateModMaster = True
Exit Function
ERR_P:
    ShowError ("ValidateModMaster::" & Me.Caption)
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
If chkWD.Value = 0 Then txtWDRate.Text = "1"
If chkWO.Value = 0 Then txtWORate.Text = "1"
If chkHL.Value = 0 Then txtHLRate.Text = "1"

ConMain.Execute "Update OTRul Set OTDesc='" & Trim(txtRDesc.Text) & "'," & _
"OTWD=" & chkWD.Value & ",OTWO=" & chkWO.Value & ",OTHL=" & chkHL.Value & ",WDRates=" & _
txtWDRate.Text & ",WORates=" & txtWORate.Text & ",HLRates=" & txtHLRate.Text & _
",Authorized='" & IIf(OptYes.Value, "Y", "N") & "',MaxOT=" & _
txtMaxOT.Text & ",DedLate=" & chkLate.Value & ",DedEarl=" & chkEarly.Value & ",From1=" & _
txtF1.Text & ",To1=" & txtT1.Text & ",Deduct1=" & txtD1.Text & ",All1=" & chkA1.Value & _
",From2=" & txtF2.Text & ",To2=" & txtT2.Text & ",Deduct2=" & txtD2.Text & ",All2=" & _
chkA2.Value & ",From3=" & txtF3.Text & ",To3=" & txtT3.Text & ",Deduct3=" & txtD3.Text & _
",All3=" & chkA3.Value & ",MoreThan=" & txtT4.Text & ",Deduct4=" & txtD4.Text & ",All4=" & _
chkA4.Value & ",WODeduct=" & chkDedWO.Value & ",HLDeduct=" & chkDedHL.Value & _
",RFrom1=" & txtRF1.Text & ",RTo1=" & txtRT1.Text & ",Round1=" & txtR1.Text & _
",RFrom2=" & txtRF2.Text & ",RTo2=" & txtRT2.Text & ",Round2=" & txtR2.Text & _
",RFrom3=" & txtRF3.Text & ",RTo3=" & txtRT3.Text & ",Round3=" & txtR3.Text & _
",RFrom4=" & txtRF4.Text & ",RTo4=" & txtRT4.Text & ",Round4=" & txtR4.Text & _
",RTo5=" & txtRT5.Text & ",Round5=" & txtR5.Text & " Where OTCode=" & cboRule.Text

SaveModMaster = True
Exit Function
ERR_P:
    ShowError ("SaveModMaster::" & Me.Caption)
End Function

Private Function SaveAddMaster() As Boolean
On Error GoTo ERR_P
If chkWD.Value = 0 Then txtWDRate.Text = "1"
If chkWO.Value = 0 Then txtWORate.Text = "1"
If chkHL.Value = 0 Then txtHLRate.Text = "1"

ConMain.Execute "insert into OTRul Values(" & txtRNo.Text & _
",'" & txtRDesc.Text & "'," & chkWD.Value & "," & chkWO.Value & "," & _
chkHL.Value & "," & txtWDRate.Text & "," & txtWORate.Text & "," & txtHLRate.Text & _
",'" & IIf(OptYes.Value, "Y", "N") & "'," & txtMaxOT.Text & "," & chkLate.Value & "," & _
chkEarly.Value & "," & txtF1.Text & "," & txtT1.Text & "," & txtD1.Text & "," & _
chkA1.Value & "," & txtF2.Text & "," & txtT2.Text & "," & txtD2.Text & "," & _
chkA2.Value & "," & txtF3.Text & "," & txtT3.Text & "," & txtD3.Text & "," & _
chkA3.Value & "," & txtT4.Text & "," & txtD4.Text & "," & chkA4.Value & "," & _
chkDedWO.Value & "," & chkDedHL.Value & "," & _
txtRF1.Text & "," & txtRT1.Text & "," & txtR1.Text & "," & _
txtRF2.Text & "," & txtRT2.Text & "," & txtR2.Text & "," & _
txtRF3.Text & "," & txtRT3.Text & "," & txtR3.Text & "," & _
txtRF4.Text & "," & txtRT4.Text & "," & txtR4.Text & "," & _
txtRT5.Text & "," & txtR5.Text & IIf(GetFlagStatus("DeductLunch") Or GetFlagStatus("LESSBKHRSFROMOT"), "," & txtWDLBHrs.Text & "," & txtWOLBHrs.Text & "," & txtHLLBHrs.Text, "") & ")"

SaveAddMaster = True
Exit Function
ERR_P:
    ShowError ("SaveAddMaster::" & Me.Caption)
End Function

Private Function ValidateAddmaster() As Boolean
On Error GoTo ERR_P
If Trim(txtRNo.Text) = "" Then
    MsgBox NewCaptionTxt("59013", adrsC)
    txtRNo.SetFocus
    Exit Function
Else
    If adrsDept1.RecordCount > 0 Then
        adrsDept1.MoveFirst
        adrsDept1.Find "OTCode=" & Trim(txtRNo.Text)
        If Not adrsDept1.EOF Then
            MsgBox NewCaptionTxt("59014", adrsC)
            txtRNo.SetFocus
            Exit Function
        End If
    End If
End If
If Trim(txtRDesc.Text) = "" Then
    MsgBox NewCaptionTxt("59015", adrsC)
    txtRDesc.SetFocus
    Exit Function
End If

If chkWD.Value = 1 And (txtWDRate.Text = "" Or txtWDRate.Text = 0) Then
    MsgBox "Rate value cannot be 0 for selected OT Day", vbInformation
    txtWDRate.SetFocus
    Exit Function
End If
If chkWO.Value = 1 And (txtWORate.Text = "" Or txtWORate.Text = 0) Then
    MsgBox "Rate value cannot be 0 for selected OT Day", vbInformation
    txtWORate.SetFocus
    Exit Function
End If
If chkHL.Value = 1 And (txtHLRate.Text = "" Or txtHLRate.Text = 0) Then
    MsgBox "Rate value cannot be 0 for selected OT Day", vbInformation
    txtHLRate.SetFocus
    Exit Function
End If
''
If Not CorrectDecs Then Exit Function
ValidateAddmaster = True
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster::" & Me.Caption)
End Function

Private Function CorrectDecs() As Boolean
On Error GoTo ERR_P
'' Deductions
'' Decimal Validations
txtF1.Text = Format(Val(txtF1.Text), "0.00")
If Val(Right(txtF1.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtF1.SetFocus
    Exit Function
End If
txtF2.Text = Format(Val(txtF2.Text), "0.00")
If Val(Right(txtF2.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtF2.SetFocus
    Exit Function
End If
txtF3.Text = Format(Val(txtF3.Text), "0.00")
If Val(Right(txtF3.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtF3.SetFocus
    Exit Function
End If
txtT1.Text = Format(Val(txtT1.Text), "0.00")
If Val(Right(txtT1.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtT1.SetFocus
    Exit Function
End If
txtT2.Text = Format(Val(txtT2.Text), "0.00")
If Val(Right(txtT2.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtT2.SetFocus
    Exit Function
End If
txtT3.Text = Format(Val(txtT3.Text), "0.00")
If Val(Right(txtT3.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtT3.SetFocus
    Exit Function
End If
txtT4.Text = Format(Val(txtT4.Text), "0.00")
If Val(Right(txtT4.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtT4.SetFocus
    Exit Function
End If
If chkA1.Value = 0 Then
    txtD1.Text = Format(Val(txtD1.Text), "0.00")
    If Val(Right(txtD1.Text, 2)) > 59 Then
        MsgBox NewCaptionTxt("00024", adrsMod)
        If txtD1.Enabled Then txtD1.SetFocus
        Exit Function
    End If
End If
If chkA2.Value = 0 Then
    txtD2.Text = Format(Val(txtD2.Text), "0.00")
    If Val(Right(txtD2.Text, 2)) > 59 Then
        MsgBox NewCaptionTxt("00024", adrsMod)
        If txtD2.Enabled Then txtD2.SetFocus
        Exit Function
    End If
End If
If chkA3.Value = 0 Then
    txtD3.Text = Format(Val(txtD3.Text), "0.00")
    If Val(Right(txtD3.Text, 2)) > 59 Then
        MsgBox NewCaptionTxt("00024", adrsMod)
        If txtD3.Enabled Then txtD3.SetFocus
        Exit Function
    End If
End If
If chkA4.Value = 0 Then
    txtD4.Text = Format(Val(txtD4.Text), "0.00")
    If Val(Right(txtD4.Text, 2)) > 59 Then
        MsgBox NewCaptionTxt("00024", adrsMod)
        If txtD4.Enabled Then txtD4.SetFocus
        Exit Function
    End If
End If
'' To timings cannot be less than deduct timings
If chkA1.Value = 0 Then
    If Val(txtD1.Text) > Val(txtT1.Text) Then
        MsgBox NewCaptionTxt("59016", adrsC)
        txtT1.SetFocus
        Exit Function
    End If
End If
If chkA2.Value = 0 Then
    If Val(txtD2.Text) > Val(txtT2.Text) Then
        MsgBox NewCaptionTxt("59016", adrsC)
        txtT2.SetFocus
        Exit Function
    End If
End If
If chkA3.Value = 0 Then
    If Val(txtD3.Text) > Val(txtT3.Text) Then
        MsgBox NewCaptionTxt("59016", adrsC)
        txtT3.SetFocus
        Exit Function
    End If
End If
If chkA4.Value = 0 Then
    If Val(txtD4.Text) > Val(txtT4.Text) Then
        MsgBox NewCaptionTxt("59016", adrsC)
        txtT4.SetFocus
        Exit Function
    End If
End If
'' Please enter to timings
If Val(txtF1.Text) > 0 And Val(txtT1.Text) = 0 Then
    MsgBox NewCaptionTxt("59017", adrsC)
    txtT1.SetFocus
    Exit Function
End If
If Val(txtF2.Text) > 0 And Val(txtT2.Text) = 0 Then
    MsgBox NewCaptionTxt("59017", adrsC)
    txtT2.SetFocus
    Exit Function
End If
If Val(txtF3.Text) > 0 And Val(txtT3.Text) = 0 Then
    MsgBox NewCaptionTxt("59017", adrsC)
    txtT3.SetFocus
    Exit Function
End If
'' Maximum OT
txtMaxOT.Text = Format(Val(txtMaxOT.Text), "0.00")
If Val(Right(txtMaxOT.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtMaxOT.SetFocus
    Exit Function
End If
'' Round Off OT Validations
'' Decimal Validations
txtRF1.Text = Format(Val(txtRF1.Text), "0.00")
If Val(Right(txtRF1.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtRF1.SetFocus
    Exit Function
End If
txtRT1.Text = Format(Val(txtRT1.Text), "0.00")
If Val(Right(txtRT1.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtRT1.SetFocus
    Exit Function
End If
txtR1.Text = Format(Val(txtR1.Text), "0.00")
If Val(Right(txtR1.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtR1.SetFocus
    Exit Function
End If
txtRF2.Text = Format(Val(txtRF2.Text), "0.00")
If Val(Right(txtRF2.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtRF2.SetFocus
    Exit Function
End If
txtRT2.Text = Format(Val(txtRT2.Text), "0.00")
If Val(Right(txtRT2.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtRT2.SetFocus
    Exit Function
End If
txtR2.Text = Format(Val(txtR2.Text), "0.00")
If Val(Right(txtR2.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtR2.SetFocus
    Exit Function
End If
txtRF3.Text = Format(Val(txtRF3.Text), "0.00")
If Val(Right(txtRF3.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtRF3.SetFocus
    Exit Function
End If
txtRT3.Text = Format(Val(txtRT3.Text), "0.00")
If Val(Right(txtRT3.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtRT3.SetFocus
    Exit Function
End If
txtR3.Text = Format(Val(txtR3.Text), "0.00")
If Val(Right(txtR3.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtR3.SetFocus
    Exit Function
End If
txtRF4.Text = Format(Val(txtRF4.Text), "0.00")
If Val(Right(txtRF4.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtRF4.SetFocus
    Exit Function
End If
txtRT4.Text = Format(Val(txtRT4.Text), "0.00")
If Val(Right(txtRT4.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtRT4.SetFocus
    Exit Function
End If
txtR4.Text = Format(Val(txtR4.Text), "0.00")
If Val(Right(txtR4.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtR4.SetFocus
    Exit Function
End If
txtRT5.Text = Format(Val(txtRT5.Text), "0.00")
If Val(Right(txtRT5.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtRT5.SetFocus
    Exit Function
End If
txtR5.Text = Format(Val(txtR5.Text), "0.00")
If Val(Right(txtR5.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtR5.SetFocus
    Exit Function
End If
'' From To Validations
If Val(txtRF1.Text) > Val(txtRT1.Text) Then
    MsgBox NewCaptionTxt("00113", adrsMod)
    txtRT1.SetFocus
    Exit Function
End If
If Val(txtRF2.Text) > Val(txtRT2.Text) Then
    MsgBox NewCaptionTxt("00113", adrsMod)
    txtRT2.SetFocus
    Exit Function
End If
If Val(txtRF3.Text) > Val(txtRT3.Text) Then
    MsgBox NewCaptionTxt("00113", adrsMod)
    txtRT3.SetFocus
    Exit Function
End If
If Val(txtRF4.Text) > Val(txtRT4.Text) Then
    MsgBox NewCaptionTxt("00113", adrsMod)
    txtRT4.SetFocus
    Exit Function
End If

CorrectDecs = True
Exit Function
ERR_P:
    ShowError ("CorrectDecs::" & Me.Caption)
End Function

Private Sub Form_Load()
On Error GoTo ERR_P

    fraDedBrHrs.Visible = False

Call SetFormIcon(Me)
Call GetRights
Call OpenMasterTable
Call RetCaptions
Call FillCombo
If cboRule.ListCount = 0 Then
    If Not AddRights Then
        MsgBox NewCaptionTxt("00001", adrsMod)
        cmdAdd.Enabled = False
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        cmdDelete.Enabled = False
        Exit Sub
    Else
        bytMode = 2     '' Add
    End If
Else
    bytMode = 1     '' Edit / View
    cboRule.ListIndex = 0
End If
Call ChangeMode
Exit Sub
ERR_P:
    ShowError (" Load  :: " & Me.Caption)
    Resume Next
End Sub

Private Sub OpenMasterTable()             '' Open the Recordset for the Display purposes
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select * from OTRul order by OTCode,OTDesc", ConMain, adOpenStatic
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Sub ChangeMode()
Select Case bytMode
    Case 1  '' Edit / View
        Call EditAction
    Case 2  '' Add
        Call AddAction
End Select
End Sub

Private Sub EditAction()
'' Enable the Necessary controls
cmdAdd.Enabled = True
cmdDelete.Enabled = True
'' Visible necessary Controls
frRule.Visible = True
'' Invisible necessary controls
frNewRule.Visible = False
End Sub

Private Sub AddAction()
On Error Resume Next
'' Enable the Necessary controls
cmdAdd.Enabled = False
cmdDelete.Enabled = False
'' Visible necessary Controls
frRule.Visible = False
'' Invisible necessary controls
frNewRule.Left = frRule.Left
frNewRule.Top = frRule.Top
frNewRule.Visible = True
txtRNo.Text = ""
Call ClearControls
txtRNo.SetFocus
End Sub

Private Sub ClearControls()
txtRDesc.Text = ""
chkWD.Value = 1
chkWO.Value = 1
chkHL.Value = 1
txtWDRate.Text = "1.00"
txtWORate.Text = "1.00"
txtHLRate.Text = "1.00"
'' frAuthorized
OptNo.Value = True
'' frMaxOT
txtMaxOT.Text = "0.00"
'' frLateEarly
chkLate.Value = 0
chkEarly.Value = 0
'' frDeductions
txtF1.Text = "0.00"
txtT1.Text = "0.00"
txtD1.Text = "0.00"
chkA1.Value = 0
txtF2.Text = "0.00"
txtT2.Text = "0.00"
txtD2.Text = "0.00"
chkA2.Value = 0
txtF3.Text = "0.00"
txtT3.Text = "0.00"
txtD3.Text = "0.00"
chkA3.Value = 0
txtT4.Text = "0.00"
txtD4.Text = "0.00"
chkA4.Value = 0
chkDedWO.Value = 1
chkDedHL.Value = 1
'' Round Off
txtRF1.Text = "0.00"
txtRT1.Text = "0.00"
txtR1.Text = "0.00"
txtRF2.Text = "0.00"
txtRT2.Text = "0.00"
txtR2.Text = "0.00"
txtRF3.Text = "0.00"
txtRT3.Text = "0.00"
txtR3.Text = "0.00"
txtRF4.Text = "0.00"
txtRT4.Text = "0.00"
txtR4.Text = "0.00"
txtRT5.Text = "0.00"
txtR5.Text = "0.00"
End Sub

Private Sub GetRights()     '' Gets and Sets the Appropriate Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 9)
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

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '59%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("00088", adrsMod)
lblRules.Caption = NewCaptionTxt("00088", adrsMod)
lblRNo.Caption = NewCaptionTxt("59001", adrsC)
lblRDesc.Caption = NewCaptionTxt("00052", adrsMod)
'' frCheck
frCheck.Caption = NewCaptionTxt("59002", adrsC)
chkWD.Caption = NewCaptionTxt("59020", adrsC)
chkWO.Caption = NewCaptionTxt("59021", adrsC)
chkHL.Caption = NewCaptionTxt("59022", adrsC)
lblTimes(0).Caption = NewCaptionTxt("59012", adrsC)
lblTimes(1).Caption = NewCaptionTxt("59012", adrsC)
lblTimes(2).Caption = NewCaptionTxt("59012", adrsC)
'' frAuthorized
frAuthorized.Caption = NewCaptionTxt("59007", adrsC)
OptYes.Caption = NewCaptionTxt("00100", adrsMod)
OptNo.Caption = NewCaptionTxt("00101", adrsMod)
'' frMaxOT
frMaxOT.Caption = NewCaptionTxt("59008", adrsC)
lblHours.Caption = NewCaptionTxt("00023", adrsMod)
'' frLateEarly
frLateEarly.Caption = NewCaptionTxt("59009", adrsC)
chkLate.Caption = NewCaptionTxt("59010", adrsC)
chkEarly.Caption = NewCaptionTxt("59011", adrsC)
'' frDeductions
frDeductions.Caption = NewCaptionTxt("59023", adrsC)
lblDSpec.Caption = NewCaptionTxt("59024", adrsC)
lblFrom.Caption = NewCaptionTxt("00045", adrsMod)
lblTo.Caption = NewCaptionTxt("00046", adrsMod)
lblDeduct.Caption = NewCaptionTxt("59025", adrsC)
lblOR.Caption = NewCaptionTxt("59026", adrsC)
lblAll.Caption = NewCaptionTxt("59027", adrsC)
lblMore.Caption = NewCaptionTxt("59003", adrsC)
chkDedWO.Caption = NewCaptionTxt("59004", adrsC)
chkDedHL.Caption = NewCaptionTxt("59005", adrsC)
'' frRound
frRound.Caption = NewCaptionTxt("59028", adrsC)
lblRSpec.Caption = NewCaptionTxt("59029", adrsC)
lblRFrom.Caption = NewCaptionTxt("00010", adrsMod)
lblRTo.Caption = NewCaptionTxt("00011", adrsMod)
lblRMore.Caption = NewCaptionTxt("59003", adrsC)
lblRound.Caption = NewCaptionTxt("59030", adrsC)
'' Command Buttons
'cmdAdd.Caption = NewCaptionTxt("00004", adrsMod)
'cmdsave.Caption = NewCaptionTxt("00007", adrsMod)
'cmdCancel.Caption = NewCaptionTxt("00003", adrsMod)
'cmdDelete.Caption = NewCaptionTxt("00006", adrsMod)
'cmdExit.Caption = NewCaptionTxt("00008", adrsMod)
End Sub

Private Sub FillCombo()
On Error GoTo ERR_P
Dim bytTmp As Byte
'' Fill the OT Combo
Call ComboFill(cboRule, 27, 2)
'' fill the Round-off Combo
Exit Sub
ERR_P:
    ShowError ("FillCombo::" & Me.Caption)
End Sub

Private Sub Display()
On Error GoTo ERR_P
If cboRule.Text = "" Then Exit Sub
If adrsDept1.RecordCount <= 0 Then Exit Sub
adrsDept1.MoveFirst
adrsDept1.Find "OTCode=" & cboRule.Text
If Not adrsDept1.EOF Then
    txtRDesc.Text = IIf(IsNull(adrsDept1("OTDesc")), "", adrsDept1("OTDesc"))
    '' Give OT on
    chkWD.Value = IIf(adrsDept1("OTWD") = 1, 1, 0)
    Call chkWD_Click
    chkWO.Value = IIf(adrsDept1("OTWO") = 1, 1, 0)
    Call chkWO_Click
    chkHL.Value = IIf(adrsDept1("OTHL") = 1, 1, 0)
    Call chkHL_Click
    '' OT Rates
    txtWDRate.Text = IIf(IsNull(adrsDept1("WDRates")), "0.00", Format(adrsDept1("WDRates")))
    txtWORate.Text = IIf(IsNull(adrsDept1("WORates")), "0.00", Format(adrsDept1("WORates")))
    txtHLRate.Text = IIf(IsNull(adrsDept1("HLRates")), "0.00", Format(adrsDept1("HLRates")))
    '' Authorized by Default
    If UCase(adrsDept1("Authorized")) = "Y" Then
        OptYes.Value = True
    Else
        OptNo.Value = True
    End If
    '' Maximum OT
    txtMaxOT.Text = IIf(IsNull(adrsDept1("MaxOT")), "0.00", Format(adrsDept1("MaxOT")))
    '' Late-Early Deductions
    chkLate.Value = IIf(adrsDept1("DedLate") = 1, 1, 0)
    chkEarly.Value = IIf(adrsDept1("DedEarl") = 1, 1, 0)
    '' Deductions
    txtF1.Text = IIf(IsNull(adrsDept1("From1")), "0.00", Format(adrsDept1("From1")))
    txtT1.Text = IIf(IsNull(adrsDept1("To1")), "0.00", Format(adrsDept1("To1")))
    txtD1.Text = IIf(IsNull(adrsDept1("Deduct1")), "0.00", Format(adrsDept1("Deduct1")))
    chkA1.Value = IIf(adrsDept1("All1") = 1, 1, 0)
    txtF2.Text = IIf(IsNull(adrsDept1("From2")), "0.00", Format(adrsDept1("From2")))
    txtT2.Text = IIf(IsNull(adrsDept1("To2")), "0.00", Format(adrsDept1("To2")))
    txtD2.Text = IIf(IsNull(adrsDept1("Deduct2")), "0.00", Format(adrsDept1("Deduct2")))
    chkA2.Value = IIf(adrsDept1("All2") = 1, 1, 0)
    txtF3.Text = IIf(IsNull(adrsDept1("From3")), "0.00", Format(adrsDept1("From3")))
    txtT3.Text = IIf(IsNull(adrsDept1("To3")), "0.00", Format(adrsDept1("To3")))
    txtD3.Text = IIf(IsNull(adrsDept1("Deduct3")), "0.00", Format(adrsDept1("Deduct3")))
    chkA3.Value = IIf(adrsDept1("All3") = 1, 1, 0)
    txtT4.Text = IIf(IsNull(adrsDept1("MoreThan")), "0.00", Format(adrsDept1("Morethan")))
    txtD4.Text = IIf(IsNull(adrsDept1("Deduct4")), "0.00", Format(adrsDept1("Deduct4")))
    chkA4.Value = IIf(adrsDept1("All4") = 1, 1, 0)
    chkDedWO.Value = IIf(adrsDept1("WODeduct") = 1, 1, 0)
    chkDedHL.Value = IIf(adrsDept1("HLDeduct") = 1, 1, 0)
    '' Round Off
    txtRF1.Text = IIf(IsNull(adrsDept1("RFrom1")), "0.00", Format(adrsDept1("RFrom1")))
    txtRT1.Text = IIf(IsNull(adrsDept1("RTo1")), "0.00", Format(adrsDept1("RTo1")))
    txtR1.Text = IIf(IsNull(adrsDept1("Round1")), "0.00", Format(adrsDept1("Round1")))
    txtRF2.Text = IIf(IsNull(adrsDept1("RFrom2")), "0.00", Format(adrsDept1("RFrom2")))
    txtRT2.Text = IIf(IsNull(adrsDept1("RTo2")), "0.00", Format(adrsDept1("RTo2")))
    txtR2.Text = IIf(IsNull(adrsDept1("Round2")), "0.00", Format(adrsDept1("Round2")))
    txtRF3.Text = IIf(IsNull(adrsDept1("RFrom3")), "0.00", Format(adrsDept1("RFrom3")))
    txtRT3.Text = IIf(IsNull(adrsDept1("RTo3")), "0.00", Format(adrsDept1("RTo3")))
    txtR3.Text = IIf(IsNull(adrsDept1("Round3")), "0.00", Format(adrsDept1("Round3")))
    txtRF4.Text = IIf(IsNull(adrsDept1("RFrom4")), "0.00", Format(adrsDept1("RFrom4")))
    txtRT4.Text = IIf(IsNull(adrsDept1("RTo4")), "0.00", Format(adrsDept1("RTo4")))
    txtR4.Text = IIf(IsNull(adrsDept1("Round4")), "0.00", Format(adrsDept1("Round4")))
    txtRT5.Text = IIf(IsNull(adrsDept1("RTo5")), "0.00", Format(adrsDept1("RTo5")))
    txtR5.Text = IIf(IsNull(adrsDept1("Round5")), "0.00", Format(adrsDept1("Round5")))
Else
    MsgBox NewCaptionTxt("59018", adrsC)
End If
Exit Sub
ERR_P:
End Sub

Private Sub txtD1_GotFocus()
Call GF(txtD1)
End Sub

Private Sub txtD2_GotFocus()
Call GF(txtD2)
End Sub

Private Sub txtD3_GotFocus()
Call GF(txtD3)
End Sub

Private Sub txtD4_GotFocus()
Call GF(txtD4)
End Sub

Private Sub txtF1_GotFocus()
Call GF(txtF1)
End Sub

Private Sub txtF2_GotFocus()
Call GF(txtF2)
End Sub

Private Sub txtF3_GotFocus()
Call GF(txtF3)
End Sub

Private Sub txtHLLBHrs_GotFocus()
Call GF(txtHLLBHrs)
End Sub

Private Sub txtHLLBHrs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtHLLBHrs)
End If
End Sub

Private Sub txtHLRate_GotFocus()
Call GF(txtHLRate)
End Sub

Private Sub txtMaxOT_GotFocus()
Call GF(txtMaxOT)
End Sub

Private Sub txtRDesc_GotFocus()
Call GF(txtRDesc)
End Sub

Private Sub txtRNo_GotFocus()
Call GF(txtRNo)
End Sub

Private Sub txtRNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 2))))
End If
End Sub
Private Sub txtRNo_LostFocus()
If txtRNo.Text = "100" Then
 MsgBox " This CO Rule is reserved for Application"
 txtRNo.Text = ""
 txtRNo.SetFocus
 End If
End Sub
Private Sub txtRDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 5))))
End If
End Sub

Private Sub txtRF1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtRF1)
End If
End Sub

Private Sub txtRT1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtRT1)
End If
End Sub

Private Sub txtR1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtR1)
End If
End Sub

Private Sub txtRF2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtRF2)
End If
End Sub

Private Sub txtRT2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtRT2)
End If
End Sub

Private Sub txtR2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtR2)
End If
End Sub

Private Sub txtRF3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtRF3)
End If
End Sub

Private Sub txtRT3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtRT3)
End If
End Sub

Private Sub txtR3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtR3)
End If
End Sub

Private Sub txtRF4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtRF4)
End If
End Sub

Private Sub txtRT4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtRT4)
End If
End Sub

Private Sub txtR4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtR4)
End If
End Sub

Private Sub txtRT5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtRT5)
End If
End Sub

Private Sub txtR5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtR5)
End If
End Sub

Private Sub txtD1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtD1)
End If
End Sub

Private Sub txtD2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtD2)
End If
End Sub

Private Sub txtD3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtD3)
End If
End Sub

Private Sub txtD4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtD4)
End If
End Sub

Private Sub txtF1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtF1)
End If
End Sub

Private Sub txtF2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtF2)
End If
End Sub

Private Sub txtF3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtF3)
End If
End Sub

Private Sub txtHLRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtHLRate)
End If
End Sub

Private Sub txtMaxOT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtMaxOT)
End If
End Sub

Private Sub txtT1_GotFocus()
Call GF(txtT1)
End Sub

Private Sub txtT1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtT1)
End If
End Sub

Private Sub txtT2_GotFocus()
Call GF(txtT2)
End Sub

Private Sub txtT2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtT2)
End If
End Sub

Private Sub txtT3_GotFocus()
Call GF(txtT3)
End Sub

Private Sub txtT3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtT3)
End If
End Sub

Private Sub txtT4_GotFocus()
Call GF(txtT4)
End Sub

Private Sub txtT4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtT4)
End If
End Sub

Private Sub txtWDLBHrs_GotFocus()
Call GF(txtWDLBHrs)
End Sub

Private Sub txtWDLBHrs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtWDLBHrs)
End If
End Sub

Private Sub txtWDRate_GotFocus()
Call GF(txtWDRate)
End Sub

Private Sub txtWDRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtWDRate)
End If
End Sub


Private Sub txtWOLBHrs_GotFocus()
Call GF(txtWOLBHrs)
End Sub

Private Sub txtWOLBHrs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtWOLBHrs)
End If
End Sub

Private Sub txtWORate_GotFocus()
Call GF(txtWDRate)
End Sub

Private Sub txtWORate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtWORate)
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub txtRF1_GotFocus()
Call GF(txtRF1)
End Sub

Private Sub txtRT1_GotFocus()
Call GF(txtRT1)
End Sub

Private Sub txtR1_GotFocus()
Call GF(txtR1)
End Sub

Private Sub txtRF2_GotFocus()
Call GF(txtRF2)
End Sub

Private Sub txtRT2_GotFocus()
Call GF(txtRT2)
End Sub

Private Sub txtR2_GotFocus()
Call GF(txtR2)
End Sub

Private Sub txtRF3_GotFocus()
Call GF(txtRF3)
End Sub

Private Sub txtRT3_GotFocus()
Call GF(txtRT3)
End Sub

Private Sub txtR3_GotFocus()
Call GF(txtR3)
End Sub

Private Sub txtRF4_GotFocus()
Call GF(txtRF4)
End Sub

Private Sub txtRT4_GotFocus()
Call GF(txtRT4)
End Sub

Private Sub txtR4_GotFocus()
Call GF(txtR4)
End Sub

Private Sub txtRT5_GotFocus()
Call GF(txtRT5)
End Sub

Private Sub txtR5_GotFocus()
Call GF(txtR5)
End Sub
