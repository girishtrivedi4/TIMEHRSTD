VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_mst_Shift 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOT 
      Height          =   2340
      Left            =   1200
      TabIndex        =   44
      Top             =   4800
      Visible         =   0   'False
      Width           =   7605
      Begin VB.Frame frOTDetail 
         Height          =   1845
         Left            =   120
         TabIndex        =   46
         Top             =   450
         Width           =   4815
         Begin VB.TextBox txtOT 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3120
            TabIndex        =   55
            Top             =   720
            Width           =   555
         End
         Begin VB.TextBox txtTo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   54
            Top             =   720
            Width           =   555
         End
         Begin VB.TextBox txtFrom 
            Enabled         =   0   'False
            Height          =   285
            Left            =   600
            TabIndex        =   53
            Top             =   720
            Width           =   555
         End
         Begin VB.TextBox txtFrom2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   600
            TabIndex        =   52
            Top             =   1080
            Width           =   555
         End
         Begin VB.TextBox txtTo2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   51
            Top             =   1080
            Width           =   555
         End
         Begin VB.TextBox txtOT2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3120
            TabIndex        =   50
            Top             =   1080
            Width           =   555
         End
         Begin VB.TextBox txtFrom3 
            Enabled         =   0   'False
            Height          =   285
            Left            =   600
            TabIndex        =   49
            Top             =   1440
            Width           =   555
         End
         Begin VB.TextBox txtTo3 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   48
            Top             =   1440
            Width           =   555
         End
         Begin VB.TextBox txtOT3 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3120
            TabIndex        =   47
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label lblFactoryOT 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OT Calculated As Per factory"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2640
            TabIndex        =   59
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   1800
            TabIndex        =   58
            Top             =   480
            Width           =   195
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   720
            TabIndex        =   57
            Top             =   480
            Width           =   345
         End
         Begin VB.Label lblActualOT 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OT Range Calculated Between"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   2220
         End
      End
      Begin VB.CheckBox chkOTHrs 
         Caption         =   "Calculate OT Hrs As Per Factory"
         Height          =   255
         Left            =   600
         TabIndex        =   45
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   435
      Left            =   7980
      TabIndex        =   20
      Top             =   4410
      Width           =   1785
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command1"
      Height          =   435
      Left            =   6160
      TabIndex        =   19
      Top             =   4410
      Width           =   1785
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "&Delete"
      Height          =   435
      Left            =   4340
      TabIndex        =   18
      Top             =   4410
      Width           =   1785
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command2"
      Height          =   435
      Left            =   2520
      TabIndex        =   17
      Top             =   4410
      Width           =   1785
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   2880
      TabIndex        =   40
      Top             =   3480
      Width           =   6855
      Begin MSMask.MaskEdBox txtUpto 
         Height          =   375
         Left            =   3960
         TabIndex        =   16
         Top             =   160
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
      Begin VB.Label ParaFramLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee can work after the shift for next"
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
         Index           =   5
         Left            =   120
         TabIndex        =   42
         Top             =   227
         Width           =   3600
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
         Index           =   8
         Left            =   5040
         TabIndex        =   41
         Top             =   227
         Width           =   510
      End
   End
   Begin VB.Frame frBrk 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   5610
      TabIndex        =   30
      Top             =   1080
      Width           =   4155
      Begin VB.TextBox txtFBS 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1440
         TabIndex        =   10
         Top             =   488
         Width           =   780
      End
      Begin VB.TextBox txtSBS 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1440
         TabIndex        =   12
         Top             =   1024
         Width           =   780
      End
      Begin VB.TextBox txtTBS 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1440
         TabIndex        =   14
         Top             =   1560
         Width           =   780
      End
      Begin VB.TextBox txtTBT 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3240
         TabIndex        =   33
         Top             =   1560
         Width           =   780
      End
      Begin VB.TextBox txtTBE 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2340
         TabIndex        =   15
         Top             =   1560
         Width           =   780
      End
      Begin VB.TextBox txtSBT 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3240
         TabIndex        =   32
         Top             =   1024
         Width           =   780
      End
      Begin VB.TextBox txtSBE 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2340
         TabIndex        =   13
         Top             =   1024
         Width           =   780
      End
      Begin VB.TextBox txtFBT 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3240
         TabIndex        =   31
         Top             =   488
         Width           =   780
      End
      Begin VB.TextBox txtFBE 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2340
         TabIndex        =   11
         Top             =   488
         Width           =   780
      End
      Begin VB.Label lblE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ends at"
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
         Left            =   2333
         TabIndex        =   39
         Top             =   180
         Width           =   795
      End
      Begin VB.Label lblT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   3353
         TabIndex        =   38
         Top             =   180
         Width           =   555
      End
      Begin VB.Label lblS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Starts at"
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
         Left            =   1395
         TabIndex        =   37
         Top             =   180
         Width           =   870
      End
      Begin VB.Label lblTB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Third Break"
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
         Left            =   330
         TabIndex        =   36
         Top             =   1605
         Width           =   990
      End
      Begin VB.Label lblSB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Second Break"
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
         TabIndex        =   35
         Top             =   1076
         Width           =   1230
      End
      Begin VB.Label lblFB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Break"
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
         Left            =   360
         TabIndex        =   34
         Top             =   540
         Width           =   960
      End
   End
   Begin VB.Frame frMisc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2640
      TabIndex        =   27
      Top             =   0
      Width           =   7125
      Begin VB.CheckBox chkFlexiShf 
         Caption         =   "Flexi Shift"
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
         Left            =   5520
         TabIndex        =   5
         Top             =   548
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkNight 
         Caption         =   "This is a night shift"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1965
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   1
         Top             =   210
         Width           =   615
      End
      Begin VB.CheckBox chkBrk 
         Caption         =   "Deduct Break Hrs from Shift hrs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2400
         TabIndex        =   4
         Top             =   600
         Width           =   3135
      End
      Begin MSMask.MaskEdBox txtName 
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   180
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   49
         PromptChar      =   "_"
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Code"
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
         TabIndex        =   29
         Top             =   240
         Width           =   1080
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
         Left            =   2160
         TabIndex        =   28
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame frHours 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   2600
      TabIndex        =   21
      Top             =   1080
      Width           =   2955
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2040
         TabIndex        =   9
         Top             =   1680
         Width           =   810
      End
      Begin VB.TextBox txtEnd 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2040
         TabIndex        =   43
         Top             =   1320
         Width           =   810
      End
      Begin VB.TextBox txtSStart 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2040
         TabIndex        =   8
         Top             =   960
         Width           =   810
      End
      Begin VB.TextBox txtHStart 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2040
         TabIndex        =   7
         Top             =   600
         Width           =   810
      End
      Begin VB.TextBox txtStart 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   810
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Shift Time"
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
         Left            =   570
         TabIndex        =   26
         Top             =   1740
         Width           =   1350
      End
      Begin VB.Label lblEnd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Ends at"
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
         Left            =   795
         TabIndex        =   25
         Top             =   1380
         Width           =   1125
      End
      Begin VB.Label lblSStart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Second half starts at"
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
         TabIndex        =   24
         Top             =   960
         Width           =   1800
      End
      Begin VB.Label lblHStart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Half Ends at"
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
         Left            =   405
         TabIndex        =   23
         Top             =   630
         Width           =   1515
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Starts at"
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
         Left            =   720
         TabIndex        =   22
         Top             =   240
         Width           =   1200
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   4035
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   7117
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColorFixed  =   12632256
      WordWrap        =   -1  'True
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
Attribute VB_Name = "frm_mst_Shift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adrsC As New ADODB.Recordset

Private Sub chkBrk_Click()
    Call CommonCalc
End Sub

Private Sub chkOTHrs_Click()    ' 13-03
If chkOTHrs.Value = 1 Then
    txtFrom.Enabled = True: txtFrom2.Enabled = True: txtFrom3.Enabled = True
    txtTo.Enabled = True: txtTo2.Enabled = True: txtTo3.Enabled = True
    txtOT.Enabled = True: txtOT2.Enabled = True: txtOT3.Enabled = True
Else
    txtFrom.Enabled = False: txtFrom2.Enabled = False: txtFrom3.Enabled = False
    txtTo.Enabled = False: txtTo2.Enabled = False: txtTo3.Enabled = False
    txtOT.Enabled = False: txtOT2.Enabled = False: txtOT3.Enabled = False
End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)        '' Set the Form Icon
Call RetCaption             '' Retreive Captions
Call OpenMasterTable        '' Open Master Table
Call FillGrid               '' Fill Grid
Call GetRights              '' Gets Rights for the Operations
bytMode = 1
Call ChangeMode             '' Take Action on the Appropriate Mode
End Sub

Private Sub GetRights()     '' Gets and Sets the Appropriate Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 7)
If Mid(strTmp, 1, 1) = "1" Then AddRights = True
If Mid(strTmp, 2, 1) = "1" Then EditRights = True
If Mid(strTmp, 3, 1) = "1" Then DeleteRights = True
Exit Sub
ERR_P:
    ShowError ("Rights :: " & Me.Caption)
    AddRights = False
    EditRights = False
    DeleteRights = False
End Sub

Private Sub RetCaption()
On Error GoTo ERR_P

If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '48%'", ConMain, adOpenStatic, adLockReadOnly
Me.Caption = NewCaptionTxt("48001", adrsC)        '' Form caption
Call SetOtherCaps                           '' Set Captions for other Captions
Call SetGButtonCap(Me)                           '' Sets Appropriate Caption for the Buttons
Call CapGrid                                '' Captions For Grid
Exit Sub
ERR_P:
    ShowError ("RetCaption :: " & Me.Caption)
End Sub

Private Sub SetOtherCaps()
'' Misc
lblCode.Caption = NewCaptionTxt("48002", adrsC)
lblName.Caption = NewCaptionTxt("00048", adrsMod)
chkNight.Caption = NewCaptionTxt("48003", adrsC)
chkBrk.Caption = NewCaptionTxt("48004", adrsC)
'' Hours
frHours.Caption = NewCaptionTxt("48005", adrsC)
lblStart.Caption = NewCaptionTxt("48006", adrsC)
lblHStart.Caption = NewCaptionTxt("48007", adrsC)
lblSStart.Caption = NewCaptionTxt("48008", adrsC)
lblEnd.Caption = NewCaptionTxt("48009", adrsC)
lblTotal.Caption = NewCaptionTxt("48010", adrsC)
'' Break
frBrk.Caption = NewCaptionTxt("48011", adrsC)
lblS.Caption = NewCaptionTxt("48012", adrsC)
lblE.Caption = NewCaptionTxt("48013", adrsC)
lblT.Caption = NewCaptionTxt("48014", adrsC)
lblFB.Caption = NewCaptionTxt("48015", adrsC)
lblSB.Caption = NewCaptionTxt("48016", adrsC)
lblTB.Caption = NewCaptionTxt("48017", adrsC)
End Sub

Private Sub CapGrid()           '' Gives the Captions to the Grid
With MSF1
    '' Sets the Column Widhts
    .ColWidth(0) = .ColWidth(0) * 0.5
    .ColWidth(1) = .ColWidth(1) * 1.9
    '.ColWidth(2) = .ColWidth(2) * 1.24
    '.ColWidth(3) = .ColWidth(3) * 1.24
    '' Sets the Column Alignment
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    '.ColAlignment(2) = flexAlignLeftCenter
    '.ColAlignment(3) = flexAlignLeftCenter
    '' Sets the Appropriate Captions
    .TextMatrix(0, 0) = "Code"
    .TextMatrix(0, 1) = "Shift Name"
    '.TextMatrix(0, 2) = NewCaptionTxt("48006", adrsC) '' Shift starts at
    '.TextMatrix(0, 3) = NewCaptionTxt("48009", adrsC) '' Shift Ends at
End With
End Sub

Private Sub OpenMasterTable()             '' Open the recordset for the Display purposes
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select * from Instshft WHERE SHIFT <> '100' Order by Shift", ConMain, adOpenStatic
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
    Exit Sub
End If
MSF1.Rows = adrsDept1.RecordCount + 1   '' Sets Rows Appropriately
adrsDept1.MoveFirst
For intCounter = 1 To adrsDept1.RecordCount     '' Fills the Grid
    With MSF1
        .TextMatrix(intCounter, 0) = adrsDept1("Shift")
        .TextMatrix(intCounter, 1) = IIf(IsNull(adrsDept1("ShiftName")), "", adrsDept1("ShiftName"))
        '.TextMatrix(intCounter, 2) = Format(adrsDept1("Shf_In"), "0.00")
        '.TextMatrix(intCounter, 3) = Format(adrsDept1("Shf_Out"), "0.00")
    End With
    adrsDept1.MoveNext
Next
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
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
cmdDel.Enabled = True           '' Enable Delete Button
'' Disable Needed Controls
frMisc.Enabled = False          '' Disable Info Frmae
frHours.Enabled = False         '' Disable Hours Frame

Frame1.Enabled = False
''
frBrk.Enabled = False           '' Disable Break Frame

Call SetGButtonCap(Me)
Call Display
End Sub
Private Sub EditAction()    '' Procedure for Edit Mode
'' Enable Necessary Controls
frMisc.Enabled = True           '' Enable Info Frame
frHours.Enabled = True          '' Enable Hours Frame

Frame1.Enabled = True
''
frBrk.Enabled = True            '' Enabel Break Frame
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
txtCode.Enabled = False         '' Disable Code TextBox
txtName.SetFocus                '' Set Focus on the Name TextBox
End Sub

Private Sub AddAction()     '' Procedure for Addition Mode
'' Enable Necessary Controls
frMisc.Enabled = True           '' Enable Info Frame
frHours.Enabled = True          '' Enable Hours Frame

Frame1.Enabled = True
''
frBrk.Enabled = True            '' Enabel Break Frame
txtCode.Enabled = True
'' Disable Necessary Controls
cmdDel.Enabled = False      '' Disable Delete Button
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Clear Necessary Controls
'' Misc
txtCode.SetFocus        '' Set Focus to the Shift Code
Call clear
End Sub

Private Sub clear()
txtCode.Text = ""       '' Clear Code Control
txtName.Text = ""       '' Clear Name Control
chkNight.Value = 0      '' Reset Night Shift
chkBrk.Value = 0        '' Reset Break Hrs
'' Hours
txtStart.Text = ""      '' Shift Start
txtHStart.Text = ""     '' First Half Start
txtSStart.Text = ""     '' Second half Start
txtEnd.Text = ""        '' Shift End
txtTotal.Text = ""      '' Total Shift Time
'' Break
txtFBS.Text = ""        '' First Break Start/End/total
txtFBE.Text = ""
txtFBT.Text = ""
txtSBS.Text = ""        '' Second Break Start/End/total
txtSBE.Text = ""
txtSBT.Text = ""
txtTBS.Text = ""        '' Third Break Start/End/total
txtTBE.Text = ""
txtTBT.Text = ""
End Sub

Private Function CommonCalc()
txtTotal.Text = TimDiff(Val(txtEnd.Text), Val(txtStart.Text))
If chkBrk.Value = 1 Then
        txtTotal.Text = TimDiff(Val(txtTotal), (TimAdd((TimAdd(Val(txtFBT.Text), _
        Val(txtSBT.Text))), Val(txtTBT.Text))))
End If
End Function

Private Sub txtCode_GotFocus()
    Call GF(txtCode)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 6))))
End If
End Sub

Private Sub txtEnd_GotFocus()
    Call GF(txtEnd)
End Sub

Private Sub txtEnd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtEnd)
End If
End Sub

Private Sub txtEnd_Validate(Cancel As Boolean)
    Call CommonCalc
End Sub

Private Sub txtFBE_GotFocus()
    Call GF(txtFBE)
End Sub

Private Sub txtFBE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtFBE)
End If
End Sub

Private Sub txtFBE_LostFocus()
    Call CommonCalc
    Call CalculateBrkHrs1
End Sub

Private Sub txtFBE_Validate(Cancel As Boolean)
    Call CommonCalc
    Call CalculateBrkHrs1
End Sub

Private Sub txtFBS_GotFocus()
    Call GF(txtFBS)
End Sub

Private Sub txtFBS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtFBS)
End If
End Sub


Private Sub txtFBS_Validate(Cancel As Boolean)
    Call CommonCalc
    Call CalculateBrkHrs1
End Sub

Private Sub txtFBT_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtFrom_GotFocus()
Call GF(txtFrom)
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtFrom)
End If
End Sub
Private Sub txtFrom2_GotFocus()
Call GF(txtFrom2)
End Sub

Private Sub txtFrom2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtFrom2)
End If
End Sub
Private Sub txtFrom3_GotFocus()
Call GF(txtFrom3)
End Sub

Private Sub txtFrom3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtFrom3)
End If
End Sub
Private Sub txtHStart_GotFocus()
    Call GF(txtHStart)
End Sub

Private Sub txtHStart_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtHStart)
End If
End Sub

Private Sub txtName_GotFocus()
    Call GF(txtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 3))))
End If
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
Private Sub txtOT2_GotFocus()
Call GF(txtOT2)
End Sub

Private Sub txtOT2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtOT2)
End If
End Sub
Private Sub txtOT3_GotFocus()
Call GF(txtOT3)
End Sub

Private Sub txtOT3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtOT3)
End If
End Sub

Private Sub txtSBE_GotFocus()
    Call GF(txtSBE)
End Sub

Private Sub txtSBE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtSBE)
End If
End Sub

Private Sub txtSBE_LostFocus()
    Call CommonCalc
    Call CalculateBrkHrs2
End Sub

Private Sub txtSBE_Validate(Cancel As Boolean)
    Call CommonCalc
    Call CalculateBrkHrs2
End Sub

Private Sub txtSBS_GotFocus()
    Call GF(txtSBS)
End Sub

Private Sub txtSBS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtSBS)
End If
End Sub

Private Sub txtSBS_Validate(Cancel As Boolean)
    Call CommonCalc
    Call CalculateBrkHrs2
End Sub


Private Sub txtSBT_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtSStart_GotFocus()
    Call GF(txtSStart)
End Sub

Private Sub txtSStart_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtSStart)
End If
End Sub

Private Sub txtStart_GotFocus()
    Call GF(txtStart)
End Sub

Private Sub txtStart_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtStart)
End If
End Sub

Private Sub txtStart_Validate(Cancel As Boolean)
    Call CommonCalc
End Sub

Private Sub txtTBE_GotFocus()
    Call GF(txtTBE)
End Sub

Private Sub txtTBE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtTBE)
End If
End Sub

Private Sub txtTBE_LostFocus()
    Call CommonCalc
    Call CalculateBrkHrs3
End Sub

Private Sub txtTBE_Validate(Cancel As Boolean)
    Call CommonCalc
    Call CalculateBrkHrs3
End Sub

Private Sub txtTBS_GotFocus()
    Call GF(txtTBS)
End Sub

Private Sub txtTBS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtTBS)
End If
End Sub

Private Sub txtTBS_Validate(Cancel As Boolean)
    Call CommonCalc
    Call CalculateBrkHrs3
End Sub

Private Function ValidateAddmaster() As Boolean     '' Validate If in Addition Mode
On Error GoTo ERR_P
ValidateAddmaster = True
'' Check for Blank Code
If Trim(txtCode.Text) = "" Then
    MsgBox NewCaptionTxt("48018", adrsC), vbExclamation
    ValidateAddmaster = False
    txtCode.SetFocus
    Exit Function
End If
'' Check for Existing Code
If MSF1.Rows > 1 Then
    adrsDept1.MoveFirst
    adrsDept1.Find "Shift='" & txtCode.Text & "'"
    If Not adrsDept1.EOF Then
        MsgBox NewCaptionTxt("48019", adrsC), vbExclamation
        txtCode.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If
'' Check for Blank Name
If Trim(txtName.Text) = "" Then
    MsgBox NewCaptionTxt("48020", adrsC), vbExclamation
    ValidateAddmaster = False
    txtName.SetFocus
    Exit Function
End If
'' MakeFormat-->Put All Values in the Format of 0.00
Call MakeFormat(txtStart)
Call MakeFormat(txtHStart)
Call MakeFormat(txtSStart)
Call MakeFormat(txtEnd)
Call MakeFormat(txtTotal)
Call MakeFormat(txtFBS)
Call MakeFormat(txtFBE)
Call MakeFormat(txtFBT)
Call MakeFormat(txtSBS)
Call MakeFormat(txtSBE)
Call MakeFormat(txtSBT)
Call MakeFormat(txtTBS)
Call MakeFormat(txtTBE)
Call MakeFormat(txtTBT)

Call MakeFormat(txtUpto)
''
'' CheckZeros--> Used to Check if Required Values are not Missing
If Not CheckZeros(txtStart) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckZeros(txtHStart) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckZeros(txtSStart) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckZeros(txtEnd) Then
    ValidateAddmaster = False
    Exit Function
End If
'' CheckDecimal --> Used to Check if Decimal Values are not greater than 59
If Not CheckDecimal(txtStart) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtHStart) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtSStart) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtEnd) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtFBS) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtFBE) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtFBT) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtSBS) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtSBE) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtSBT) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtTBS) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtTBE) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtTBT) Then
    ValidateAddmaster = False
    Exit Function
End If

If Not CheckDecimal(txtUpto) Then
    ValidateAddmaster = False
    Exit Function
End If
''
'' Check24 --> used to see if the Existing Values are not Greater than 23.59
If Not Check24(txtStart) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not Check24(txtHStart) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not Check24(txtSStart) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not Check24(txtEnd) Then
    ValidateAddmaster = False
    Exit Function
End If

If Not Check24(txtUpto) Then
    ValidateAddmaster = False
    Exit Function
End If
''
'' CheckBet --> Used to find out if the time falls between the valid range
If Not CheckBet(txtHStart) Then
    ValidateAddmaster = False
    Exit Function
End If
'' Manual Check if Second Shift Start Time is Less than First Shift End Time
If Val(txtSStart.Text) < Val(txtHStart.Text) Then
    MsgBox NewCaptionTxt("48021", adrsC), vbExclamation
    txtSStart.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
'' Manual Check if Break Ends are not Greater then Break Starts
'' First Break
If Val(txtFBE.Text) < Val(txtFBS.Text) Then
    MsgBox NewCaptionTxt("48022", adrsC), vbExclamation
    txtFBE.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
'' Second Break
If Val(txtSBE.Text) < Val(txtSBS.Text) Then
    MsgBox NewCaptionTxt("48023", adrsC), vbExclamation
    txtSBE.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
'' Third Break
If Val(txtTBE.Text) < Val(txtTBS.Text) Then
    MsgBox NewCaptionTxt("48024", adrsC), vbExclamation
    txtTBE.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
'' Check if Breaks Fall Between Shift Start Time and Shift End Time
If Val(txtFBS) > 0 Then
    If Not CheckBet(txtFBS) Then
        ValidateAddmaster = False
        Exit Function
    End If
End If
If Val(txtSBS) > 0 Then
    If Not CheckBet(txtSBS) Then
        ValidateAddmaster = False
        Exit Function
    End If
End If
If Val(txtTBS) > 0 Then
    If Not CheckBet(txtTBS) Then
        ValidateAddmaster = False
        Exit Function
    End If
End If
If Val(txtSBS.Text) > 0 And Val(txtSBS.Text) < Val(txtFBE.Text) Then
    MsgBox NewCaptionTxt("48025", adrsC), vbExclamation
    txtSBS.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If Val(txtTBS.Text) > 0 And Val(txtTBS.Text) < Val(txtSBE.Text) Then
    MsgBox NewCaptionTxt("48026", adrsC), vbExclamation
    txtTBS.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
'' Check if Break End Timings Fall between the Valid Ranges
If Val(txtFBE.Text) > 0 Then
    If Not CheckBet(txtFBE) Then
        ValidateAddmaster = False
        Exit Function
    End If
End If
If Val(txtSBE.Text) > 0 Then
    If Not CheckBet(txtSBE) Then
        ValidateAddmaster = False
        Exit Function
    End If
End If
If Val(txtTBE.Text) > 0 Then
    If Not CheckBet(txtTBE) Then
        ValidateAddmaster = False
        Exit Function
    End If
End If
'' Check on Total Value
'' Zero Check
If Not CheckZeros(txtTotal) Then
    ValidateAddmaster = False
    Exit Function
End If

If Not CheckZeros(txtUpto) Then
    ValidateAddmaster = False
    Exit Function
End If
''
'' 0.59 Check
If Not CheckDecimal(txtTotal) Then
    ValidateAddmaster = False
    Exit Function
End If
'' 23.59 Check
If Not Check24(txtTotal) Then
    ValidateAddmaster = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster :: " & Me.Caption)
    ValidateAddmaster = False
End Function

Private Function Check24(ByRef txt As Object) As Boolean
If Val(txt.Text) > 23.59 And chkNight.Value <> 1 Then
    MsgBox NewCaptionTxt("00025", adrsMod), vbExclamation
    txt.SetFocus
    Check24 = False
Else
    Check24 = True
End If
End Function

Private Function CheckZeros(ByRef txt As Object) As Boolean
If Val(txt.Text) <= 0 Then
    MsgBox NewCaptionTxt("00060", adrsMod), vbExclamation
    txt.SetFocus
    CheckZeros = False
Else
    CheckZeros = True
End If
End Function

Private Function CheckBet(ByRef txt As Object) As Boolean
CheckBet = True
If Val(txt.Text) <= Val(txtStart.Text) Then
    MsgBox NewCaptionTxt("48027", adrsC), vbExclamation
    txt.SetFocus
    CheckBet = False
    Exit Function
End If
If Val(txt.Text) >= Val(txtEnd.Text) Then
    MsgBox NewCaptionTxt("48028", adrsC), vbExclamation
    txt.SetFocus
    CheckBet = False
End If
End Function

Private Function CheckDecimal(ByRef txt As Object) As Boolean
If Val(Right(txt.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod), vbExclamation
    txt.SetFocus
    CheckDecimal = False
Else
    CheckDecimal = True
End If
End Function

Private Sub MakeFormat(ByRef txt As Object)
    If txt.Text = "" Then
        txt.Text = "0.00"
    Else
        txt.Text = Format(txt.Text, "0.00")
    End If
End Sub

Private Function ValidateModMaster() As Boolean     '' Validate If in Addition Mode
On Error GoTo ERR_P
ValidateModMaster = True
'' Check for Blank Name
If Trim(txtName.Text) = "" Then
    MsgBox NewCaptionTxt("48020", adrsC), vbExclamation
    ValidateModMaster = False
    txtName.SetFocus
    Exit Function
End If
'' MakeFormat-->Put All Values in the Format of 0.00
Call MakeFormat(txtStart)
Call MakeFormat(txtHStart)
Call MakeFormat(txtSStart)
Call MakeFormat(txtEnd)
Call MakeFormat(txtTotal)
Call MakeFormat(txtFBS)
Call MakeFormat(txtFBE)
Call MakeFormat(txtFBT)
Call MakeFormat(txtSBS)
Call MakeFormat(txtSBE)
Call MakeFormat(txtSBT)
Call MakeFormat(txtTBS)
Call MakeFormat(txtTBE)
Call MakeFormat(txtTBT)

Call MakeFormat(txtUpto)
''
'' CheckZeros--> Used to Check if Required Values are not Missing
If Not CheckZeros(txtStart) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckZeros(txtHStart) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckZeros(txtSStart) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckZeros(txtEnd) Then
    ValidateModMaster = False
    Exit Function
End If
'' CheckDecimal --> Used to Check if Decimal Values are not greater than 59
If Not CheckDecimal(txtStart) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtHStart) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtSStart) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtEnd) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtFBS) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtFBE) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtFBT) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtSBS) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtSBE) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtSBT) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtTBS) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtTBE) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtTBT) Then
    ValidateModMaster = False
    Exit Function
End If

If Not CheckDecimal(txtUpto) Then
    ValidateModMaster = False
    Exit Function
End If

'' Check24 --> used to see if the Existing Values are not Greater than 23.59
If Not Check24(txtStart) Then
    ValidateModMaster = False
    Exit Function
End If
If Not Check24(txtHStart) Then
    ValidateModMaster = False
    Exit Function
End If
If Not Check24(txtSStart) Then
    ValidateModMaster = False
    Exit Function
End If
If Not Check24(txtEnd) Then
    ValidateModMaster = False
    Exit Function
End If

If Not Check24(txtUpto) Then
    ValidateModMaster = False
    Exit Function
End If
If Val(txtEnd.Text) < 24 And chkNight.Value = 1 Then
    MsgBox "Departure time cannot be less than 24 if it is a night shift.", vbExclamation
    txtEnd.SetFocus
    ValidateModMaster = False
    Exit Function
End If

'' CheckBet --> Used to find out if the time falls between the valid range
If Not CheckBet(txtHStart) Then
    ValidateModMaster = False
    Exit Function
End If
'' Manual Check if Second Shift Start Time is Less than First Shift End Time
If Val(txtSStart.Text) < Val(txtHStart.Text) Then
            
    MsgBox NewCaptionTxt("48021", adrsC), vbExclamation
    txtSStart.SetFocus
    ValidateModMaster = False
    Exit Function
End If
'' Manual Check if Break Ends are not Greater then Break Starts
'' First Break
If Val(txtFBE.Text) < Val(txtFBS.Text) Then
    MsgBox NewCaptionTxt("48022", adrsC), vbExclamation
    txtFBE.SetFocus
    ValidateModMaster = False
    Exit Function
End If
'' Second Break
If Val(txtSBE.Text) < Val(txtSBS.Text) Then
    MsgBox NewCaptionTxt("48023", adrsC) _
    , vbExclamation, App.EXEName
    txtSBE.SetFocus
    ValidateModMaster = False
    Exit Function
End If
'' Third Break
If Val(txtTBE.Text) < Val(txtTBS.Text) Then
    MsgBox NewCaptionTxt("48024", adrsC), vbExclamation
    txtTBE.SetFocus
    ValidateModMaster = False
    Exit Function
End If
'' Check if Breaks Fall Between Shift Start Time and Shift End Time
If Val(txtFBS) > 0 Then
    If Not CheckBet(txtFBS) Then
        ValidateModMaster = False
        Exit Function
    End If
End If
If Val(txtSBS) > 0 Then
    If Not CheckBet(txtSBS) Then
        ValidateModMaster = False
        Exit Function
    End If
End If
If Val(txtTBS) > 0 Then
    If Not CheckBet(txtTBS) Then
        ValidateModMaster = False
        Exit Function
    End If
End If
If Val(txtSBS.Text) > 0 And Val(txtSBS.Text) < Val(txtFBE.Text) Then
    MsgBox NewCaptionTxt("48025", adrsC), vbExclamation
    txtSBS.SetFocus
    ValidateModMaster = False
    Exit Function
End If
If Val(txtTBS.Text) > 0 And Val(txtTBS.Text) < Val(txtSBE.Text) Then
    MsgBox NewCaptionTxt("48026", adrsC), vbExclamation
    txtTBS.SetFocus
    ValidateModMaster = False
    Exit Function
End If
'' Check if Break End Timings Fall between the Valid Ranges
If Val(txtFBE.Text) > 0 Then
    If Not CheckBet(txtFBE) Then
        ValidateModMaster = False
        Exit Function
    End If
End If
If Val(txtSBE.Text) > 0 Then
    If Not CheckBet(txtSBE) Then
        ValidateModMaster = False
        Exit Function
    End If
End If
If Val(txtTBE.Text) > 0 Then
    If Not CheckBet(txtTBE) Then
        ValidateModMaster = False
        Exit Function
    End If
End If

If Not CheckZeros(txtUpto) Then
    ValidateModMaster = False
    Exit Function
End If
''
'' Check on Total Value
'' Zero Check
If Not CheckZeros(txtTotal) Then
    ValidateModMaster = False
    Exit Function
End If
'' 0.59 Check
If Not CheckDecimal(txtTotal) Then
    ValidateModMaster = False
    Exit Function
End If
'' 23.59 Check
If Not Check24(txtTotal) Then
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
SaveAddMaster = True        '' Insert

Call CalculateBrkHrs1
Call CalculateBrkHrs2
Call CalculateBrkHrs3
Call CommonCalc
Dim StrOpen As String
    StrOpen = ""
''
    ConMain.Execute "insert into Instshft Values('" & Trim(txtCode.Text) & "'," & _
    txtStart.Text & "," & txtEnd.Text & "," & txtTotal.Text & "," & txtFBE.Text & "," & _
    txtFBS.Text & "," & txtFBT.Text & "," & txtSBS.Text & "," & txtSBE.Text & "," & _
    txtSBT.Text & "," & txtTBS.Text & "," & txtTBE.Text & "," & txtTBT.Text & "," & _
    IIf(chkNight.Value = 0, 0, 1) & "," & txtHStart.Text & "," & txtSStart.Text & ",'" & _
    txtName.Text & "','" & chkBrk.Value & "'," & txtUpto.Text & " " & StrOpen & ")"
Exit Function
ERR_P:
    Select Case Err.Number
        Case -2147217900
            MsgBox NewCaptionTxt("48029", adrsC), vbExclamation
        Case Else
            ShowError ("SaveAddMaster :: " & Me.Caption)
    End Select
    SaveAddMaster = False
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
SaveModMaster = True        '' Update

Call CalculateBrkHrs1
Call CalculateBrkHrs2
Call CalculateBrkHrs3
Call CommonCalc
''
ConMain.Execute "Update Instshft Set Shf_in=" & txtStart.Text & ",Shf_Out=" & _
txtEnd.Text & ",Shf_Hrs=" & txtTotal.Text & ",Rst_Out=" & txtFBE.Text & ",Rst_In=" & _
txtFBS.Text & ",Rst_Brk=" & txtFBT.Text & ",Rst_In_2=" & txtSBS.Text & ",Rst_Out_2=" & _
txtSBE.Text & ",Rst_Brk_2=" & txtSBT.Text & ",Rst_In_3=" & txtTBS.Text & ",Rst_Out_3=" & _
txtTBE.Text & ",Rst_Brk_3=" & txtTBT.Text & ",Night=" & IIf(chkNight.Value = 0, 0, 1) & _
",hdend=" & txtHStart.Text & ",hdstart=" & txtSStart.Text & ",ShiftName='" & _
txtName.Text & "',BrkShf='" & chkBrk.Value & "',upto=" & txtUpto.Text & " Where Shift='" & txtCode.Text & "'"
Exit Function
ERR_P:
    SaveModMaster = False
    ShowError ("SaveModMaster :: " & Me.Caption)
End Function

Private Sub Display()       '' Displays the Given Master Records
On Error GoTo ERR_P
adrsDept1.Requery
If adrsDept1.EOF Then
    cmdEditCan.Enabled = False
    Exit Sub
End If
adrsDept1.MoveFirst
adrsDept1.Find "Shift='" & MSF1.TextMatrix(MSF1.Row, 0) & "'"
If Not (adrsDept1.EOF) Then
    '' Misc
    txtCode.Text = adrsDept1("Shift")
    txtName.Text = adrsDept1("ShiftName")
    chkNight.Value = IIf(IsNull(adrsDept1("Night")) Or adrsDept1("Night") = 0, 0, 1)
    chkBrk.Value = IIf(IsNull(adrsDept1("BrkShf")) Or adrsDept1("BrkShf") = "0", 0, 1)
    '' Hours
    txtStart.Text = IIf(IsNull(adrsDept1("Shf_In")), "0.00", _
                    Format(adrsDept1("Shf_In"), "0.00"))
    txtHStart.Text = IIf(IsNull(adrsDept1("Hdend")), "0.00", _
                    Format(adrsDept1("Hdend"), "0.00"))
    txtSStart.Text = IIf(IsNull(adrsDept1("Hdstart")), "0.00", _
                    Format(adrsDept1("Hdstart"), "0.00"))
    txtEnd.Text = IIf(IsNull(adrsDept1("shf_Out")), "0.00", _
                    Format(adrsDept1("shf_Out"), "0.00"))
    txtTotal.Text = IIf(IsNull(adrsDept1("Shf_Hrs")), "0.00", _
                    Format(adrsDept1("Shf_Hrs"), "0.00"))
    
    txtUpto.Text = IIf(IsNull(adrsDept1("upto")), "0.00", _
                    Format(adrsDept1("upto"), "0.00"))
    ''
    '' Break Rst_Brk
    txtFBS.Text = IIf(IsNull(adrsDept1("Rst_In")), "0.00", _
                    Format(adrsDept1("Rst_In"), "0.00"))            '' First Break
    txtFBE.Text = IIf(IsNull(adrsDept1("Rst_Out")), "0.00", _
                    Format(adrsDept1("Rst_Out"), "0.00"))
    txtFBT.Text = IIf(IsNull(adrsDept1("Rst_Brk")), "0.00", _
                    Format(adrsDept1("Rst_Brk"), "0.00"))
    txtSBS.Text = IIf(IsNull(adrsDept1("Rst_In_2")), "0.00", _
                    Format(adrsDept1("Rst_In_2"), "0.00"))          '' Second Break
    txtSBE.Text = IIf(IsNull(adrsDept1("Rst_Out_2")), "0.00", _
                    Format(adrsDept1("Rst_Out_2"), "0.00"))
    txtSBT.Text = IIf(IsNull(adrsDept1("Rst_Brk_2")), "0.00", _
                    Format(adrsDept1("Rst_Brk_2"), "0.00"))
    txtTBS.Text = IIf(IsNull(adrsDept1("Rst_In_3")), "0.00", _
                    Format(adrsDept1("Rst_In_3"), "0.00"))          '' Third Break
    txtTBE.Text = IIf(IsNull(adrsDept1("Rst_Out_3")), "0.00", _
                    Format(adrsDept1("Rst_Out_3"), "0.00"))
    txtTBT.Text = IIf(IsNull(adrsDept1("Rst_Brk_3")), "0.00", _
                    Format(adrsDept1("Rst_Brk_3"), "0.00"))
Else
    MsgBox NewCaptionTxt("48030", adrsC), vbCritical
    Exit Sub
End If
Exit Sub
ERR_P:
    ShowError ("Display :: " & Me.Caption)
End Sub

Private Sub MSF1_DblClick()
Call Display
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

Private Sub cmdDel_Click()
On Error GoTo ERR_P
'' Check for Rights
If Not DeleteRights Then
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    Exit Sub
Else
    If MsgBox(NewCaptionTxt("48031", adrsC), vbYesNo + vbQuestion, Me.Caption) = vbYes Then            '' Delete the Record
        ConMain.Execute "delete from  Instshft where Shift='" & txtCode.Text & "'"
        Call AddActivityLog(lgDelete_Action, 1, 2)  '' Delete Log
        Call AuditInfo("DELETE", Me.Caption, "Deleted Shift: " & txtCode.Text)
    End If
    
    Call FillGrid       '' Reflect the Grid
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
ERR_P:
    If Err.Number = -2147467259 Or Err.Number = -2147217900 Or Err.Number = -2147217873 Then
            MsgBox "Shift Cannot be deleted because employees belong to this Shift.", vbCritical, Me.Caption
            Exit Sub
    End If
    ShowError ("Delete Record :: " & Me.Caption)
End Sub

Private Sub cmdEditCan_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        If Not EditRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        bytMode = 3
        Call ChangeMode
    Case 2       '' Add Mode
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

Private Sub CalculateBrkHrs1()
txtFBT.Text = Format(TimDiff(Val(txtFBE.Text), Val(txtFBS.Text)), "0.00")
End Sub

Private Sub CalculateBrkHrs2()
txtSBT.Text = Format(TimDiff(Val(txtSBE.Text), Val(txtSBS.Text)), "0.00")
End Sub

Private Sub CalculateBrkHrs3()
txtTBT.Text = Format(TimDiff(Val(txtTBE.Text), Val(txtTBS.Text)), "0.00")
End Sub

Private Sub txtTBT_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtTo_GotFocus()
Call GF(txtTo)
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtTo)
End If
End Sub

Private Sub txtTotal_GotFocus()
    Call GF(txtTotal)
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtTotal)
End If
End Sub

Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo ERR_P
Call AddActivityLog(lgADD_MODE, 1, 2)     '' Add Activity
Call AuditInfo("ADD", Me.Caption, "Added Shift: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 2)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edited Shift: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then KeyCode = 0
End Sub

Private Sub txtUpto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtUpto)
End If
End Sub
