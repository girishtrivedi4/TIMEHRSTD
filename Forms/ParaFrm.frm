VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ParaFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Pathcmd3 
      Caption         =   "..."
      Height          =   375
      Left            =   7200
      TabIndex        =   38
      ToolTipText     =   "Click to Browse the default Wall Paper File"
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame ParaFrame 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5325
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   7725
      Begin VB.TextBox Pathtxt3 
         Height          =   375
         Left            =   3000
         ScrollBars      =   1  'Horizontal
         TabIndex        =   37
         Text            =   " "
         Top             =   4080
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.ComboBox WeekBgCombo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1515
      End
      Begin VB.ComboBox YrStartCombo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   330
         Width           =   1515
      End
      Begin VB.ComboBox CardSizeCombo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   930
      End
      Begin VB.ComboBox CodeSZCombo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   930
      End
      Begin VB.ComboBox CurrYRCombo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   945
      End
      Begin VB.TextBox PathTxt 
         Height          =   375
         Left            =   3000
         ScrollBars      =   1  'Horizontal
         TabIndex        =   9
         Text            =   " "
         Top             =   3360
         Width           =   4035
      End
      Begin VB.TextBox Pathtxt2 
         Height          =   375
         Left            =   3000
         ScrollBars      =   1  'Horizontal
         TabIndex        =   11
         Text            =   " "
         Top             =   3720
         Visible         =   0   'False
         Width           =   4035
      End
      Begin VB.CommandButton PathCmd 
         Caption         =   "..."
         Height          =   375
         Left            =   7080
         TabIndex        =   10
         ToolTipText     =   "Click to Browse the .DAT File Path"
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton Pathcmd2 
         Caption         =   "..."
         Height          =   375
         Left            =   7080
         TabIndex        =   12
         ToolTipText     =   "Click to Browse the default Wall Paper File"
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox EmailChk 
         BackColor       =   &H80000004&
         Caption         =   "Send Reports using Email"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   5
         Top             =   1215
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CheckBox chkAssum 
         Caption         =   "Apply Salary Cut-Off Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   13
         Top             =   4200
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.TextBox txtAssum 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3870
         MaxLength       =   2
         TabIndex        =   14
         Top             =   4470
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chkIO 
         Caption         =   "I/O based processing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   15
         Top             =   4560
         Width           =   2325
      End
      Begin VB.CheckBox chkIgnore 
         Caption         =   "Consider only first and last punch while processing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   16
         Top             =   4920
         Width           =   5085
      End
      Begin MSComDlg.CommonDialog ParaDialog 
         Left            =   240
         Top             =   2280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox IgnPunchMask 
         Height          =   375
         Left            =   5400
         TabIndex        =   6
         Top             =   1920
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ShftChgLtMask 
         Height          =   375
         Left            =   5400
         TabIndex        =   7
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ShftChgErlMask 
         Height          =   345
         Left            =   5400
         TabIndex        =   8
         Top             =   2880
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0.00"
         PromptChar      =   " "
      End
      Begin VB.Label lblDateF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Date Format :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   21
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label ParaFramLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "digit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   3600
         TabIndex        =   35
         Top             =   1237
         Width           =   360
      End
      Begin VB.Label ParaFramLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "digit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   3600
         TabIndex        =   34
         Top             =   757
         Width           =   360
      End
      Begin VB.Label ParaFramLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hours"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   6480
         TabIndex        =   33
         Top             =   1987
         Width           =   510
      End
      Begin VB.Label ParaFramLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week begins on"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   4485
         TabIndex        =   32
         Top             =   757
         Width           =   1410
      End
      Begin VB.Label ParaFramLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year Starting from"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   4320
         TabIndex        =   31
         Top             =   367
         Width           =   1575
      End
      Begin VB.Label ParaFramLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ignore next punch from  the previous punch till"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   1200
         TabIndex        =   30
         Top             =   1987
         Width           =   3975
      End
      Begin VB.Label ParaFramLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Punching Card Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   600
         TabIndex        =   29
         Top             =   1237
         Width           =   1725
      End
      Begin VB.Label ParaFramLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Code Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   495
         TabIndex        =   28
         Top             =   757
         Width           =   1830
      End
      Begin VB.Label ParaFramLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Year is"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1020
         TabIndex        =   27
         Top             =   367
         Width           =   1305
      End
      Begin VB.Label ParaFramLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ignore next punch from  the previous punch till"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   1200
         TabIndex        =   26
         Top             =   2467
         Width           =   3975
      End
      Begin VB.Label ParaFramLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ignore next punch from  the previous punch till"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   1200
         TabIndex        =   25
         Top             =   2932
         Width           =   3975
      End
      Begin VB.Label ParaFramLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   6480
         TabIndex        =   24
         Top             =   2467
         Width           =   405
      End
      Begin VB.Label ParaFramLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   6480
         TabIndex        =   23
         Top             =   2932
         Width           =   405
      End
      Begin VB.Label ParaFramLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capturing Database Path"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   600
         TabIndex        =   22
         Top             =   3465
         Width           =   2175
      End
      Begin VB.Label lblAssum 
         AutoSize        =   -1  'True
         Caption         =   "Cut-Off Day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2730
         TabIndex        =   20
         Top             =   4530
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin VB.CommandButton ExitCancmd 
      Cancel          =   -1  'True
      Caption         =   " "
      Height          =   405
      Left            =   6000
      TabIndex        =   18
      Top             =   5790
      Width           =   1935
   End
   Begin VB.CommandButton EditSaveCmd 
      Caption         =   " "
      Height          =   405
      Left            =   4080
      TabIndex        =   17
      Top             =   5790
      Width           =   1935
   End
   Begin MSMask.MaskEdBox AfterShiftMask 
      Height          =   375
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0.00"
      PromptChar      =   " "
   End
End
Attribute VB_Name = "ParaFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Dim Install_Rights As Boolean
Dim adrsC As New ADODB.Recordset   '' L

Private Sub AfterShiftMask_Change()
        EditSaveCmd.Enabled = True
End Sub

Private Sub AfterShiftMask_GotFocus()
        SendKeys "{home}+{end}"
End Sub

Private Sub AfterShiftMask_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys Chr(9)
        Else
            KeyAscii = keycheck(KeyAscii, AfterShiftMask)
        End If
End Sub

Private Sub AfterShiftMask_Validate(Cancel As Boolean)
On Error GoTo ERR_P
AfterShiftMask = Format(AfterShiftMask, "00.00")
AfterShiftMask.SelLength = Len(AfterShiftMask)
Select Case Val(AfterShiftMask.Text)
        Case Is < 0
                MsgBox NewCaptionTxt("00060", adrsMod), vbExclamation
                Cancel = True
        Case Is > 23.59
                MsgBox NewCaptionTxt("00025", adrsMod), vbExclamation
                Cancel = True
        Case Else
                If Not ValidDecimal(Val(AfterShiftMask)) Then Cancel = True
End Select
Exit Sub
ERR_P:
    ShowError ("Validate :: " & Me.Caption)
End Sub


Private Sub CardSizeCombo_Change()
    EditSaveCmd.Enabled = True
End Sub

Private Sub CardSizeCombo_Click()
On Error GoTo ERR_P
If CByte(CardSizeCombo.Text) > pVStar.CardSize Then
    MsgBox NewCaptionTxt("54045", adrsC), vbInformation
End If
If CByte(CardSizeCombo.Text) < pVStar.CardSize Then
    MsgBox NewCaptionTxt("54046", adrsC), vbInformation
End If
EditSaveCmd.Enabled = True

Exit Sub
ERR_P:
    ShowError ("CardSizeCombo :: " & Me.Caption)
End Sub

Private Sub chkAssum_Click()
On Error Resume Next
If EditFlag Then
    EditSaveCmd.Enabled = True
    If chkAssum.Value = 1 Then
        txtAssum.Enabled = True
    Else
        txtAssum.Text = "0"
        txtAssum.Enabled = False
    End If
Else
    chkAssum.Enabled = False
    txtAssum.Enabled = False
End If
End Sub


Private Sub chkIgnore_Click()
If EditFlag Then
    EditSaveCmd.Enabled = True
End If
End Sub

Private Sub chkIO_Click()
If EditFlag Then
    EditSaveCmd.Enabled = True
End If
End Sub
''
Private Sub CodeSZCombo_Change()
        EditSaveCmd.Enabled = True
End Sub

Private Sub CodeSZCombo_Click()
If CByte(CodeSZCombo.Text) > pVStar.CodeSize Then
    MsgBox NewCaptionTxt("54045", adrsC), vbInformation
End If
If CByte(CodeSZCombo.Text) < pVStar.CodeSize Then
    MsgBox NewCaptionTxt("54046", adrsC), vbInformation
End If
EditSaveCmd.Enabled = True
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CurrYRCombo_Change()
        EditSaveCmd.Enabled = True
End Sub

Private Sub CurrYRCombo_Click()
        EditSaveCmd.Enabled = True
End Sub


Private Sub EditSaveCmd_Click()
On Error GoTo ErrParaES
If Not EditRights Then
        MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
        Exit Sub
End If
Dim ChkFlg As Boolean
If Not EditFlag Then
        ExitCancmd.Cancel = False
        EditFlag = True
        EditMode
        ExitCancmd.Caption = NewCaptionTxt("00003", adrsMod)
        EditSaveCmd.Caption = "Update"
        CurrYRCombo.SetFocus
        
ElseIf EditFlag Then     'true before save
    If txtAssum.Visible = True And txtAssum.Enabled = True Then
        If Val(txtAssum.Text) > 31 Then
            MsgBox NewCaptionTxt("54048", adrsC), vbExclamation
            txtAssum.SetFocus
            Exit Sub
        End If
    End If
    If Not ValidDecimal(Val(IgnPunchMask)) Then Exit Sub
    If Not ValidDecimal(Val(ShftChgLtMask)) Then Exit Sub
    If Not ValidDecimal(Val(ShftChgErlMask)) Then Exit Sub
    
    If chkAssum.Value = 1 And Val(txtAssum.Text) = 0 Then
        MsgBox "Please Enter The Cut Off Date"
        txtAssum.SetFocus
        Exit Sub
    End If
    
        If adRsInstall.State = 1 Then adRsInstall.Close
        adRsInstall.Open "Select * from install", ConMain, adOpenKeyset, adLockOptimistic
        adRsInstall.Fields("upto").Value = Val(AfterShiftMask)
        adRsInstall("e_codesize") = CodeSZCombo.Text
        adRsInstall("e_cardsize") = CardSizeCombo.Text
        adRsInstall("cur_year") = CurrYRCombo.Text
        adRsInstall("hl_ot") = 1
        adRsInstall("wo_ot") = 1
        adRsInstall("ot_ot") = 1
        adRsInstall("filt_time") = Val(IgnPunchMask)
        adRsInstall("yearfrom") = Month("01-" & Left(YrStartCombo.Text, 3) & _
        "-" & Year(Date))
        Select Case strConv(WeekBgCombo, vbProperCase)
                Case "Monday": adRsInstall("weekFrom") = 2
                Case "Tuesday": adRsInstall("weekFrom") = 3
                Case "Wednesday": adRsInstall("weekFrom") = 4
                Case "Thursday": adRsInstall("weekFrom") = 5
                Case "Friday": adRsInstall("weekFrom") = 6
                Case "Saturday": adRsInstall("weekFrom") = 7
                Case "Sunday": adRsInstall("weekFrom") = 1
        End Select
        adRsInstall("deductlter") = 0
        adRsInstall("otround") = 0
        adRsInstall("postlt") = Val(ShftChgLtMask)
        adRsInstall("posterl") = Val(ShftChgErlMask)
        If MultiDB = False Then
            adRsInstall("datpath") = Trim(PathTxt)
        Else
            Dim str As String
            str = Trim(PathTxt)
            If Trim(PathTxt.Text) <> "" Then str = Trim(PathTxt.Text)
            If Trim(Pathtxt2.Text) <> "" Then str = str + "|" + Trim(Pathtxt2.Text)
            If Trim(Pathtxt3.Text) <> "" Then str = str + "|" + Trim(Pathtxt3.Text)
            adRsInstall("datpath") = str ' Trim(PathTxt) + "|" + Trim(Pathtxt2.Text) + "|" + Trim(Pathtxt3.Text)
        End If
        adRsInstall("dec1") = 0
        adRsInstall("dec1a") = 0
        adRsInstall("dec2") = 0
        adRsInstall("dec2a") = 0
        adRsInstall("dec3") = 0
        adRsInstall("dec3a") = 0
        adRsInstall("dec4") = 0
        adRsInstall("dec4a") = 0
        adRsInstall("dec5") = 0
        adRsInstall("round1") = 0
        adRsInstall("round2") = 0
        adRsInstall("round3") = 0
        adRsInstall("round4") = 0
        adRsInstall("round5") = 0
        'adRsInstall("walpaper") = Trim(WallTxt)
        adRsInstall("email") = EmailChk.Value
        adRsInstall("defincut") = IIf(chkAssum.Value = 1, "Y", "N")
        adRsInstall("cutdt") = IIf(txtAssum.Text = "", 0, Val(txtAssum.Text))
        
        adRsInstall("IO") = IIf(chkIO.Value = 1, "Y", "N")
        typPerm.blnIO = IIf(chkIO.Value = 1, True, False)
        adRsInstall("IgnoreP") = IIf(chkIgnore.Value = 1, "Y", "N")
        
        ''
        adRsInstall.Update
        Call AddActivityLog(lgEdit_Mode, 2, 1)
        Call AuditInfo("Update", Me.Caption, "Edit Installation Parameter")
         '21 lvupdateyear,22 allowedit,23 deductlter,24 otround, 25 email, 26 defincod,27 datpath
         adRsInstall.Close
         Set adRsInstall = ConMain.Execute("Select * from install")
         With pVStar
                .CardSize = adRsInstall("e_cardsize")
                .CodeSize = adRsInstall("e_codesize")
                .WeekStart = adRsInstall("weekfrom")
                .Use_Mail = IIf(adRsInstall("email") = 0, False, True)
                .YearSel = adRsInstall("cur_year")
                .Yearstart = adRsInstall("yearfrom")
        End With
        adRsInstall.Close
        FillControl
        EditFlag = False
        ExitEditMode
        EditSaveCmd.Caption = "Update"
        ExitCancmd.Caption = NewCaptionTxt("00008", adrsMod)
        ExitCancmd.Cancel = True
        Call StandardChange      ''  Add By Girish 23-12
End If
ChkFlg = False
Exit Sub
ErrParaES:
    ShowError ("Edit/Save :: " & Me.Caption)
    Resume Next
End Sub

Private Sub EmailChk_Click()
        EditSaveCmd.Enabled = True
End Sub

Private Sub ExitCancmd_Click()
If EditFlag = False Then
        Unload ParaFrm
ElseIf EditFlag = True Then
        EditFlag = False
        ExitEditMode
        FillControl
        EditSaveCmd.Enabled = True
        EditSaveCmd.Caption = NewCaptionTxt("00005", adrsMod)
        ExitCancmd.Caption = NewCaptionTxt("00008", adrsMod)
        ExitCancmd.Cancel = True
End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrParaLoad
Call SetFormIcon(Me)
Dim YrLen%
'populate the year combo box
For YrLen% = 0 To 99
        CurrYRCombo.AddItem (2011 + YrLen%)
Next YrLen%
For i% = 1 To 9
        CodeSZCombo.AddItem Choose(i%, "2", "3", "4", "5", "6", "7", "8", "9", "10")
Next i%
For i% = 1 To 7
        CardSizeCombo.AddItem Choose(i%, "2", "3", "4", "5", "6", "7", "8")
Next i%
''
''
With YrStartCombo
        '.AddItem "", 0
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
'Populate the week days
With WeekBgCombo
        .AddItem "Sunday"
        .AddItem "Monday"
        .AddItem "Tuesday"
        .AddItem "Wednesday"
        .AddItem "Thursday"
        .AddItem "Friday"
        .AddItem "Saturday"
End With
'********
ExitEditMode
Call RetCaption

lblDateF.Caption = NewCaptionTxt("54044", adrsC) & strDateFO
FillControl
'********
'' For Rights
Dim strTmp As String
strTmp = RetRights(3, 1, , 1)
EditRights = False
If strTmp = "1" Then EditRights = True


YrLen = 0
    
Exit Sub
ErrParaLoad:
    ShowError ("Load :: " & Me.Caption)
    'Resume Next
End Sub

Private Sub RetCaption()
On Error Resume Next

If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '54%'", ConMain, adOpenStatic, adLockReadOnly
ParaFrm.Caption = NewCaptionTxt("54001", adrsC)

ParaFrame.Caption = NewCaptionTxt("54003", adrsC)

ParaFramLbl(0) = NewCaptionTxt("54004", adrsC)
ParaFramLbl(1) = NewCaptionTxt("54005", adrsC)
ParaFramLbl(2) = NewCaptionTxt("54006", adrsC)
ParaFramLbl(3) = NewCaptionTxt("54007", adrsC)
ParaFramLbl(4) = NewCaptionTxt("54008", adrsC)
ParaFramLbl(5) = NewCaptionTxt("54009", adrsC)
'ParaFramLbl(6) = NewCaptionTxt("54010", adrsC)
ParaFramLbl(7) = NewCaptionTxt("54011", adrsC)
ParaFramLbl(8) = NewCaptionTxt("00023", adrsMod)
ParaFramLbl(10) = NewCaptionTxt("00023", adrsMod)
ParaFramLbl(13) = NewCaptionTxt("54012", adrsC)
ParaFramLbl(14) = NewCaptionTxt("54013", adrsC)
ParaFramLbl(15) = NewCaptionTxt("00023", adrsMod)
ParaFramLbl(16) = NewCaptionTxt("00023", adrsMod)

'overtime
ParaFramLbl(19).Caption = NewCaptionTxt("00010", adrsMod)
ParaFramLbl(20).Caption = NewCaptionTxt("00011", adrsMod)
EditSaveCmd.Caption = "Update"
ExitCancmd.Caption = NewCaptionTxt("00008", adrsMod)

'ParaFramLbl(6).Caption = NewCaptionTxt("54034", adrsC)
ParaFramLbl(11).Caption = NewCaptionTxt("54035", adrsC)
ParaFramLbl(12).Caption = NewCaptionTxt("54035", adrsC)
EmailChk.Caption = NewCaptionTxt("54036", adrsC)
ParaFramLbl(17).Caption = NewCaptionTxt("54038", adrsC)
ParaFramLbl(18).Caption = NewCaptionTxt("54039", adrsC)
ParaFramLbl(21).Caption = NewCaptionTxt("54040", adrsC)
ParaFramLbl(22).Caption = NewCaptionTxt("54041", adrsC)
chkAssum.Caption = NewCaptionTxt("54042", adrsC)
lblAssum.Caption = NewCaptionTxt("54043", adrsC)
''' End

End Sub


Private Sub EditMode()
On Error GoTo ERR_P
CurrYRCombo.Enabled = True
CodeSZCombo.Enabled = True
CardSizeCombo.Enabled = True
YrStartCombo.Enabled = True
WeekBgCombo.Enabled = True
AfterShiftMask.Enabled = True
IgnPunchMask.Enabled = True
EmailChk.Enabled = True
PathTxt.Enabled = True
PathCmd.Enabled = True
ShftChgLtMask.Enabled = True
ShftChgErlMask.Enabled = True
chkAssum.Enabled = True
txtAssum.Enabled = True
If InVar.blnAssum = "1" Then
    chkAssum.Enabled = True
    txtAssum.Enabled = True
End If

chkIO.Enabled = True
'chkDI.Enabled = True
chkIgnore.Enabled = True
''

Pathtxt2.Enabled = True
Pathcmd3.Enabled = True
Pathcmd2.Enabled = True
Pathtxt3.Enabled = True
Exit Sub
ERR_P:
    ShowError ("EditMode :: " & Me.Caption)
End Sub

Private Sub ExitEditMode()
On Error GoTo ERR_P
CurrYRCombo.Enabled = False
CodeSZCombo.Enabled = False
CardSizeCombo.Enabled = False
YrStartCombo.Enabled = False
WeekBgCombo.Enabled = False
PathTxt.Enabled = False
PathCmd.Enabled = False
AfterShiftMask.Enabled = False
IgnPunchMask.Enabled = False
ShftChgLtMask.Enabled = False
ShftChgErlMask.Enabled = False
chkAssum.Enabled = False
EmailChk.Enabled = False
txtAssum.Enabled = False
If InVar.blnAssum = "1" Then
    chkAssum.Enabled = False
    txtAssum.Enabled = False
End If

chkIO.Enabled = False
chkIgnore.Enabled = False
Pathtxt2.Enabled = False
Pathcmd3.Enabled = False
Pathcmd2.Enabled = False
Pathtxt3.Enabled = False

Exit Sub
ERR_P:
    ShowError ("ExitEditMode :: " & Me.Caption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Install_Rights = False
EditFlag = False
End Sub

Private Sub IgnPunchMask_Change()
        EditSaveCmd.Enabled = True
End Sub

Private Sub IgnPunchMask_GotFocus()
        SendKeys "{home}+{end}"
End Sub

Private Sub IgnPunchMask_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys Chr(9)
        Else
            KeyAscii = keycheck(KeyAscii, IgnPunchMask)
        End If
End Sub

Private Sub IgnPunchMask_Validate(Cancel As Boolean)
On Error GoTo ERR_P
Select Case Val(IgnPunchMask)
        Case Is < 0
                MsgBox NewCaptionTxt("00060", adrsMod), vbExclamation
                Cancel = True
        Case Is > 23.59
                MsgBox NewCaptionTxt("00025", adrsMod), vbExclamation
                Cancel = True
        Case Else
                If Not ValidDecimal(Val(IgnPunchMask)) Then Cancel = True
End Select
Exit Sub
ERR_P:
    ShowError ("Validate :: " & Me.Caption)
End Sub

Private Sub PathCmd_Click()
    ParaDialog.Filter = "Database Files (*.mdb)|*.mdb"
    ParaDialog.ShowOpen
    PathTxt.Text = ParaDialog.FileName
End Sub

Private Sub PathCmd2_Click()
    If PathTxt.Text = "" Then Exit Sub
    ParaDialog.Filter = "Database Files (*.mdb)|*.mdb"
    ParaDialog.ShowOpen
    Pathtxt2.Text = ParaDialog.FileName
End Sub

Private Sub PathCmd3_Click()
    If Pathtxt2.Text = "" Then Exit Sub
    ParaDialog.Filter = "Database Files (*.mdb)|*.mdb"
    ParaDialog.ShowOpen
    Pathtxt3.Text = ParaDialog.FileName
End Sub

Private Sub PathTxt_Change()
        EditSaveCmd.Enabled = True
End Sub

Private Sub PathTxt_DblClick()
If Trim(Pathtxt2.Text) <> "" Then Exit Sub
PathTxt.Text = ""
End Sub

Private Sub PathTxt_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys Chr(9)
End Sub

Private Sub Pathtxt2_DblClick()
If Trim(Pathtxt3.Text) <> "" Then Exit Sub
Pathtxt2.Text = ""
End Sub

Private Sub Pathtxt3_DblClick()
    Pathtxt3.Text = ""
End Sub

Private Sub ShftChgErlMask_Change()
        EditSaveCmd.Enabled = True
End Sub

Private Sub ShftChgErlMask_GotFocus()
        SendKeys "{home}+{end}"
End Sub

Private Sub ShftChgErlMask_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys Chr(9)
        Else
            KeyAscii = keycheck(KeyAscii, ShftChgErlMask)
        End If
End Sub

Private Sub ShftChgErlMask_Validate(Cancel As Boolean)
        If Not ValidDecimal(Val(ShftChgErlMask)) Then Cancel = True
End Sub

Private Sub ShftChgLtMask_Change()
        EditSaveCmd.Enabled = True
End Sub

Private Sub ShftChgLtMask_GotFocus()
        SendKeys "{home}+{end}"
End Sub

Private Sub ShftChgLtMask_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys Chr(9)
        Else
            KeyAscii = keycheck(KeyAscii, ShftChgLtMask)
        End If
End Sub

Private Sub ShftChgLtMask_Validate(Cancel As Boolean)
        If Not ValidDecimal(Val(ShftChgLtMask)) Then Cancel = True
End Sub

Private Sub txtAssum_Change()
EditSaveCmd.Enabled = True
End Sub

Private Sub txtAssum_Click()
MsgBox NewCaptionTxt("54060", adrsC), vbInformation
End Sub

Private Sub txtAssum_GotFocus()
    Call GF(txtAssum)
End Sub

Private Sub txtAssum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9), True
Else
    KeyAscii = KeyPressCheck(KeyAscii, 2)
End If
End Sub

Private Sub wallCmd_Click()
ParaDialog.Filter = "Pictures Files (*.bmp;*.ico;*.dib;*.jpg;*.cur;*.ani;*.gif)|*.bmp;*.ico;*.dib;*.jpg;*.cur;*.ani;*.gif"
ParaDialog.ShowOpen
End Sub

Private Sub WallTxt_Change()
        EditSaveCmd.Enabled = True
End Sub

Private Sub WallTxt_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys Chr(9)
End Sub

Private Sub WeekBgCombo_Change()
        EditSaveCmd.Enabled = True
End Sub

Private Sub WeekBgCombo_Click()
        EditSaveCmd.Enabled = True
End Sub

Private Sub YrStartCombo_Change()
        EditSaveCmd.Enabled = True
End Sub

Private Sub YrStartCombo_Click()
        EditSaveCmd.Enabled = True
End Sub

Private Sub FillControl()
'check for the install
On Error GoTo ErrRetPara
If adRsInstall.State = 1 Then adRsInstall.Close
Set adRsInstall = ConMain.Execute("Select * from install")
If Not (adRsInstall.BOF And adRsInstall.EOF) Then      'No records
        AfterShiftMask = IIf(IsNull(adRsInstall("upto")), Format(0, "00.00"), Format(adRsInstall("upto"), "00.00"))             '' Upto
        CodeSZCombo.Text = IIf(IsNull(adRsInstall("e_codesize")), "", adRsInstall("e_codesize"))                                '' Code Size
        CardSizeCombo.Text = IIf(IsNull(adRsInstall("e_cardsize")), "", adRsInstall("e_cardsize"))                              '' Card Size
                   
        CurrYRCombo.Text = IIf(IsNull(adRsInstall("cur_year")), "", adRsInstall("cur_year"))                                    '' Current Year
        IgnPunchMask = IIf(IsNull(adRsInstall("filt_time")), Format(0, "00.00"), Format(adRsInstall("filt_time"), "00.00"))     '' Filter Time
        ShftChgLtMask = IIf(IsNull(adRsInstall("PostLt")), Format(0, "00.00"), Format(adRsInstall("postlt"), "00.00"))          '' Post Late
        ShftChgErlMask = IIf(IsNull(adRsInstall("PostErl")), Format(0, "00.00"), Format(adRsInstall("posterl"), "00.00"))       '' Post Early

        YrStartCombo.Text = IIf(IsNull(adRsInstall("yearfrom")), 0, MonthName(Month(Year_Start)))                               '' Year Start
        WeekBgCombo.Text = IIf(IsNull(adRsInstall("weekfrom")), 0, CDay(adRsInstall("weekfrom")))                               '' Week Start
        Dim path() As String
        path = Split(IIf(IsNull(adRsInstall("datpath")), "", adRsInstall("datpath")), "|")
        If UBound(path) >= 0 Then PathTxt = path(0)
        If UBound(path) >= 1 Then Pathtxt2 = path(1)
        If UBound(path) >= 2 Then Pathtxt3 = path(2)
        'PathTxt = IIf(IsNull(adRsInstall("datpath")), "", adRsInstall("datpath"))                                               '' Data Path
                
        EmailChk.Value = IIf(IsNull(adRsInstall("Email")) Or adRsInstall("Email") = 0, 0, 1)                                    '' Use Email
         If InVar.blnAssum = "1" Then
            chkAssum.Value = IIf(IsNull(adRsInstall("defincut")) Or adRsInstall("defincut") = "N", 0, 1)
            txtAssum.Text = IIf(IsNull(adRsInstall("cutdt")), 0, adRsInstall("cutdt"))
        End If
        
        chkIO.Value = IIf(IsNull(adRsInstall("IO")) Or adRsInstall("IO") = "N", 0, 1)
        chkIgnore.Value = IIf(IsNull(adRsInstall("IgnoreP")) Or adRsInstall("IgnoreP") = "N", 0, 1)
        chkAssum.Value = IIf(adRsInstall("defincut") = "Y", 1, 0)
       txtAssum.Text = IIf(adRsInstall("cutdt") = "", 0, adRsInstall("cutdt"))
       
       If Not IsNull(adRsInstall("PEND")) Then
        If adRsInstall("PEND") = 1 Then
            Pathtxt2.Visible = True
            Pathcmd3.Visible = True
            Pathcmd2.Visible = True
            Pathtxt3.Visible = True
            MultiDB = True
        Else
            MultiDB = False
        End If
        
       End If
        
End If
adRsInstall.Close
Exit Sub
ErrRetPara:
    ShowError ("FillControl :: " & Me.Caption)
    Resume Next
End Sub


Private Function ValidDecimal(ByVal Ptim As Single) As Boolean
On Error GoTo ERR_P
If Val(Val(Ptim) - Int(Val(Ptim))) > 0.59 Then
    MsgBox NewCaptionTxt("00024", adrsMod), vbExclamation
    Exit Function
End If
If Val(Ptim) > 23.59 Then
    MsgBox NewCaptionTxt("00025", adrsMod), vbExclamation
    Exit Function
End If
ValidDecimal = True
Exit Function
ERR_P:
    ShowError ("ValidDecimal :: " & Me.Caption)
    ValidDecimal = False
End Function

Private Function CDay(ByVal dy As Integer) As String
Select Case dy
    Case 1: CDay = "Sunday"
    Case 2: CDay = "Monday"
    Case 3: CDay = "Tuesday"
    Case 4: CDay = "Wednesday"
    Case 5: CDay = "Thursday"
    Case 6: CDay = "Friday"
    Case 7: CDay = "Saturday"
End Select
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then KeyCode = 0
End Sub
