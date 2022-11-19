VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmEmpShift 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmWoRule 
      Caption         =   "Selection Rule For Extra Week Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   48
      Top             =   7560
      Width           =   7815
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Week Off Rule"
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
         TabIndex        =   50
         Top             =   300
         Width           =   1980
      End
      Begin MSForms.ComboBox cmbWORule 
         Height          =   375
         Left            =   3480
         TabIndex        =   49
         Top             =   240
         Width           =   1575
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2778;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame frmWeekOffOfMonth 
      Caption         =   "Week Off on first week of month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   45
      TabIndex        =   44
      Top             =   6840
      Width           =   7935
      Begin MSForms.ComboBox cmbWeekDays 
         Height          =   375
         Left            =   3480
         TabIndex        =   47
         Top             =   240
         Width           =   1575
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2778;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Of a first week of Month"
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
         Left            =   5490
         TabIndex        =   46
         Top             =   330
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "There is a week Off on"
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
         TabIndex        =   45
         Top             =   300
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdSetAll 
      Caption         =   "Set for more employees"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   30
      TabIndex        =   20
      Top             =   6420
      Width           =   2925
   End
   Begin VB.Frame frOther 
      Caption         =   "Details regarding Daily Processing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   0
      TabIndex        =   41
      Top             =   4230
      Width           =   7935
      Begin VB.Frame frBlank 
         Caption         =   "If Blank Shift found"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   4470
         TabIndex        =   43
         Top             =   330
         Width           =   3405
         Begin VB.OptionButton optBlank 
            Caption         =   "Assign the following Shift"
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
            Left            =   150
            TabIndex        =   18
            Top             =   630
            Width           =   3045
         End
         Begin VB.OptionButton optBlank 
            Caption         =   "Keep it Blank"
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
            Left            =   150
            TabIndex        =   17
            Top             =   300
            Width           =   3015
         End
         Begin MSForms.ComboBox cboBlank 
            Height          =   375
            Left            =   1500
            TabIndex        =   19
            Top             =   960
            Width           =   1035
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "1826;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame frWOHL 
         Caption         =   "On Weekoff / Holiday do the following"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   90
         TabIndex        =   42
         Top             =   330
         Width           =   4335
         Begin VB.CheckBox chkAuto 
            Caption         =   "Assign Auto shift if punch found"
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
            TabIndex        =   16
            Top             =   1380
            Visible         =   0   'False
            Width           =   3525
         End
         Begin VB.OptionButton optDet 
            Caption         =   "Assign the following Shift"
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
            Left            =   90
            TabIndex        =   14
            Top             =   1020
            Width           =   2985
         End
         Begin VB.OptionButton optDet 
            Caption         =   "Assign Next Day Shift (Schedule)"
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
            Left            =   90
            TabIndex        =   13
            Top             =   690
            Width           =   4065
         End
         Begin VB.OptionButton optDet 
            Caption         =   "Assign Previous Day Shift (Transaction)"
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
            Left            =   90
            TabIndex        =   12
            Top             =   330
            Width           =   4095
         End
         Begin MSForms.ComboBox cboShift 
            Height          =   375
            Left            =   3240
            TabIndex        =   15
            Top             =   960
            Width           =   945
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "1667;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
   End
   Begin VB.Frame frMain 
      Height          =   4275
      Left            =   0
      TabIndex        =   23
      Top             =   -60
      Width           =   7935
      Begin VB.Frame frGen 
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   45
         TabIndex        =   24
         Top             =   135
         Width           =   7830
         Begin VB.TextBox txtCode 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   1695
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
            Height          =   195
            Left            =   4710
            TabIndex        =   28
            Top             =   330
            Width           =   75
         End
         Begin VB.Label lblCode 
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
            Left            =   135
            TabIndex        =   25
            Top             =   315
            Width           =   1710
         End
         Begin VB.Label lblNameCap 
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
            Left            =   3870
            TabIndex        =   27
            Top             =   315
            Width           =   510
         End
         Begin MSForms.ComboBox cboCode 
            Height          =   375
            Left            =   2070
            TabIndex        =   26
            Top             =   240
            Width           =   1695
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "2990;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame frInfo 
         Caption         =   "Shift Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   45
         TabIndex        =   29
         Top             =   900
         Width           =   7815
         Begin VB.TextBox txtDate 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "D"
            Text            =   " "
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label lblType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shift Type"
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
            Left            =   315
            TabIndex        =   30
            Top             =   300
            Width           =   1470
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Starting Date"
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
            Left            =   315
            TabIndex        =   31
            Top             =   675
            Width           =   1410
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shift Code"
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
            Left            =   3480
            TabIndex        =   33
            Top             =   720
            Width           =   1800
         End
         Begin VB.Label lblRot 
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
            Left            =   3480
            TabIndex        =   32
            Top             =   330
            Width           =   1860
         End
         Begin MSForms.ComboBox cboType 
            Height          =   375
            Left            =   2010
            TabIndex        =   1
            Top             =   240
            Width           =   1335
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "2355;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboFix 
            Height          =   375
            Left            =   5460
            TabIndex        =   4
            Top             =   630
            Width           =   1815
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "3201;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboRot 
            Height          =   375
            Left            =   5460
            TabIndex        =   2
            Top             =   210
            Width           =   1815
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "3201;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame frWO 
         Caption         =   "Week Off"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   45
         TabIndex        =   34
         Top             =   2010
         Width           =   7830
         Begin VB.Label lblWO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "There is a week Off on every "
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
            TabIndex        =   35
            Top             =   300
            Width           =   3435
         End
         Begin VB.Label lblWO1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Of a week"
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
            Left            =   5490
            TabIndex        =   36
            Top             =   330
            Width           =   1380
         End
         Begin MSForms.ComboBox cboWO 
            Height          =   375
            Left            =   3510
            TabIndex        =   5
            Top             =   270
            Width           =   1575
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "2778;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame frAWO 
         Caption         =   "Additional Week-Offs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         Left            =   60
         TabIndex        =   37
         Top             =   2760
         Width           =   7815
         Begin VB.CheckBox chkAWO 
            Caption         =   "There is a week Off every"
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
            Left            =   135
            TabIndex        =   6
            Top             =   300
            Width           =   4305
         End
         Begin VB.CheckBox chkAWO 
            Caption         =   "There is a week Off on the first  && Third"
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
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   675
            Width           =   4200
         End
         Begin VB.CheckBox chkAWO 
            Caption         =   "There is a week off on the Second && Fourth"
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
            Left            =   120
            TabIndex        =   10
            Top             =   1155
            Width           =   4365
         End
         Begin VB.Label lblAWO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Of a week"
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
            Left            =   6510
            TabIndex        =   38
            Top             =   270
            Width           =   1230
         End
         Begin VB.Label lblAWO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Of a week"
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
            Left            =   6480
            TabIndex        =   39
            Top             =   705
            Width           =   1230
         End
         Begin VB.Label lblAWO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Of a week"
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
            Left            =   6480
            TabIndex        =   40
            Top             =   1125
            Width           =   1230
         End
         Begin MSForms.ComboBox cboAWO 
            Height          =   375
            Index           =   0
            Left            =   4770
            TabIndex        =   7
            Top             =   210
            Width           =   1335
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "2355;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboAWO 
            Height          =   375
            Index           =   1
            Left            =   4770
            TabIndex        =   9
            Top             =   630
            Width           =   1335
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "2355;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboAWO 
            Height          =   375
            Index           =   2
            Left            =   4770
            TabIndex        =   11
            Top             =   1050
            Width           =   1335
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "2355;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
   End
   Begin VB.CommandButton cmdEditSave 
      Caption         =   " "
      Height          =   405
      Left            =   5160
      TabIndex        =   21
      Top             =   6420
      Width           =   1335
   End
   Begin VB.CommandButton cmdExitCan 
      Caption         =   " "
      Height          =   405
      Left            =   6570
      TabIndex        =   22
      Top             =   6420
      Width           =   1335
   End
End
Attribute VB_Name = "frmEmpShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bytShiftModeFrm As Byte
''
Dim adrsC As New ADODB.Recordset

Private Sub cboCode_Click()
On Error GoTo ERR_P
If cboCode.Text = "" Then Exit Sub
lblName.Caption = cboCode.List(cboCode.ListIndex, 1)
Call GetEmpDetails
Exit Sub
ERR_P:
    ShowError ("Code :: " & Me.Caption)
End Sub

Private Sub cboType_Click()
    Call AdjustType
End Sub

Private Sub chkAWO_Click(Index As Integer)
If chkAWO(Index).Value = 1 Then
    cboAWO(Index).Enabled = True
Else
    cboAWO(Index).Enabled = False
    cboAWO(Index).ListIndex = cboAWO(Index).ListCount - 1
End If
End Sub

Private Sub cmdEditSave_Click()
On Error GoTo ERR_P
Select Case bytShiftModeFrm
    Case 1  '' No Mode
        If bytShfMode = 2 And cboCode.Text = "" Then Exit Sub
        bytShiftModeFrm = 2
        Call ChangeMode
    Case 2  '' Edit Mode
        If Not ValidateModMaster Then Exit Sub
        If Not SaveModMaster Then Exit Sub
        If bytShfMode <> 1 And bytShfMode <> 3 Then adrsPaid.Requery
        bytShiftModeFrm = 1
        Call ChangeMode
        ''  on mansi's request
'        frmShiftCr.Show vbModal
End Select
Exit Sub
ERR_P:
    ShowError ("Edit/Save :: " & Me.Caption)
End Sub

Private Sub cmdExitCan_Click()
On Error GoTo ERR_P
Select Case bytShiftModeFrm
    Case 1
        Unload Me
    Case 2
        bytShiftModeFrm = 1
        Call ChangeMode
End Select
Exit Sub
ERR_P:
    ShowError ("Exit / Cancel :: " & Me.Caption)
End Sub

Private Sub cmdSetAll_Click()
If Trim(cboCode.Text) = "" Then
    MsgBox "Please Select Employee Code"
    cboCode.SetFocus
    Exit Sub
End If
'' Valid Shift Selected
Shft.ShiftType = cboType.Text
If Shft.ShiftType = "F" Then
    Shft.ShiftCode = cboFix.Text    '' Shift Code
Else
    Shft.ShiftCode = cboRot.Text    '' Rotation Code
End If
Shft.startdate = DateCompDate(txtDate.Text)       '' Get the Shift Date
'' Week Off Selected
Shft.WO = Left(cboWO.Text, 2)       '' Week Off
'' Additional Week Offs
'' 0
Shft.WO1 = Left(cboAWO(0).Text, 2)
'' 1
Shft.WO2 = Left(cboAWO(1).Text, 2)
'' 2
Shft.WO3 = Left(cboAWO(2).Text, 2)
'' For Details Regarding Daily Processing
If optDet(0).Value = True Then Shft.WOHLAction = 0
If optDet(1).Value = True Then Shft.WOHLAction = 1
If optDet(2).Value = True Then Shft.WOHLAction = 2
Shft.Action3Shift = Trim(cboShift.Text)
Shft.AutoOnPunch = IIf(chkAuto.Value = 1, True, False)
If optBlank(0).Value = True Then
    Shft.ActionBlank = ""
Else
    Shft.ActionBlank = Trim(cboBlank.Text)
End If
frmShiftAll.Show vbModal
If bytShfMode <> 1 And bytShfMode <> 3 Then adrsPaid.Requery
bytShiftModeFrm = 1
Call ChangeMode
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me, True)                                '' Sets the Form Icon
Call SetToolTipText(Me)                             '' Sets the ToolTip Texts
Call RetCaptions                                    '' Gets and Stes the Form Captions
Call GetRights                                      '' Gets and Sets the Rights
Call FillCombos                                     '' Fills the Combo Boxes
bytShiftModeFrm = 4                                 '' Set the Mode to Load
If bytShfMode = 1 Or bytShfMode = 3 Then
    bytShiftModeFrm = 3                             '' If From Other Forms
    cmdSetAll.Visible = False
End If
Call ChangeMode                                     '' Take Action Accordingly
If bytShiftModeFrm <> 2 Then bytShiftModeFrm = 1    '' Set the Mode to Normal

End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '24%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("24001", adrsC)                 '' Forms Caption
'' General Frame
frGen.Caption = NewCaptionTxt("24002", adrsC)               '' General
lblCode.Caption = NewCaptionTxt("00061", adrsMod)             '' Employee Code
lblNameCap.Caption = NewCaptionTxt("00048", adrsMod)          '' Name
'' Shift Frame
frInfo.Caption = NewCaptionTxt("24003", adrsC)              '' Shift Info
lblType.Caption = NewCaptionTxt("24004", adrsC)             '' Shift Type
lblDate.Caption = NewCaptionTxt("24005", adrsC)             '' Starting Date
lblRot.Caption = NewCaptionTxt("24006", adrsC)              '' Rotation Code
lblFix.Caption = NewCaptionTxt("24007", adrsC)              '' Shift Code
'' WeekOff Frame
frWO.Caption = NewCaptionTxt("24008", adrsC)                '' Week Off
lblWO.Caption = NewCaptionTxt("24009", adrsC)               '' There is a ...
lblWO1.Caption = NewCaptionTxt("24010", adrsC)              '' Of a Week
'' Additional Week Off Frame
frAWO.Caption = NewCaptionTxt("24011", adrsC)               '' Addditional Week Offs
chkAWO(0).Caption = NewCaptionTxt("24012", adrsC)           '' There is a ...
chkAWO(1).Caption = NewCaptionTxt("24013", adrsC)           '' There is a ... 1 _ 3
chkAWO(2).Caption = NewCaptionTxt("24014", adrsC)           '' There is a ... 2 _ 4
lblAWO(0).Caption = NewCaptionTxt("24010", adrsC)           '' Of a Week
lblAWO(1).Caption = NewCaptionTxt("24010", adrsC)           '' Of a Week
lblAWO(2).Caption = NewCaptionTxt("24010", adrsC)           '' Of a Week
'' Details regarding Daily Processing
frOther.Caption = NewCaptionTxt("24024", adrsC)             '' Details regarding....
frWOHL.Caption = NewCaptionTxt("24025", adrsC)              '' on Weekoff / Holiday....
optDet(0).Caption = NewCaptionTxt("24026", adrsC)           '' Assign Previous Day Shift
optDet(1).Caption = NewCaptionTxt("24027", adrsC)           '' Assign Next Day Shift
optDet(2).Caption = NewCaptionTxt("24028", adrsC)           '' Assign the following Shift
chkAuto.Caption = NewCaptionTxt("24029", adrsC)             '' Assign Auto shift if....
frBlank.Caption = NewCaptionTxt("24030", adrsC)             '' If Blank Shift found
optBlank(0).Caption = NewCaptionTxt("24031", adrsC)         '' Keep it Blank
optBlank(1).Caption = NewCaptionTxt("24032", adrsC)         '' Assign the following Shift

Call SetButtonCap
End Sub

Private Sub SetButtonCap(Optional bytFlgCap As Byte = 1)
Select Case bytFlgCap           '' Sets Captions for the Buttons
    Case 1
        cmdEditSave.Caption = "Update"
        cmdExitCan.Caption = "Exit"
        cmdExitCan.Cancel = True
        cmdSetAll.Enabled = True
        
    Case 2
        cmdEditSave.Caption = "Save"
        cmdExitCan.Caption = "Cancel"
        cmdExitCan.Cancel = False
        cmdSetAll.Enabled = False
End Select
End Sub

Private Sub GetRights()         '' Gets and Sets Rights for the Form
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 11, 1)
If Mid(strTmp, 2, 1) <> "1" Then
    cmdEditSave.Enabled = False
    cmdSetAll.Visible = False
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
End If
Exit Sub
ERR_P:
    ShowError ("GetRights :: " & Me.Caption)
   AddRights = False
    EditRights = False
   DeleteRights = False
End Sub

Private Sub FillCombos()        '' Fills ComboBoxes
On Error GoTo ERR_P
Dim bytTmp As Byte
'' Employee Combo
Call FillEmpCombo
'' Rotation Combo
Call FillRotCombo
'' Shift Combo
Call FillShiftCombo
'' Shift Type
cboType.AddItem "F"
cboType.AddItem "R"
'' Week Off
For bytTmp = 1 To 8
    cboWO.AddItem Choose(bytTmp, "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "")
Next
'' Additional Week Off's
'' 0
For bytTmp = 1 To 8
    cboAWO(0).AddItem Choose(bytTmp, "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "")
Next
'' 1
For bytTmp = 1 To 8
    cboAWO(1).AddItem Choose(bytTmp, "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "")
Next
'' 2
For bytTmp = 1 To 8
    cboAWO(2).AddItem Choose(bytTmp, "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "")
Next
Exit Sub
ERR_P:
    ShowError ("Fill Combos " & Me.Caption)
End Sub

Private Sub FillEmpCombo()      '' Fills Employee ComboBox
On Error GoTo ERR_P
Call ComboFill(cboCode, 16, 2)
Exit Sub
ERR_P:
    ShowError ("FillEmpCombo :: " & Me.Caption)
End Sub

Private Sub FillRotCombo()      '' Fills Rotation ComboBox
On Error GoTo ERR_P
Dim strArrTmp() As String, bytTmp As Byte
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select SCode,Mon_Oth,Skp,Pattern from Ro_Shift where scode <> '100' order by SCode", _
ConMain, adOpenStatic
If Not (adrsDept1.BOF And adrsDept1.EOF) Then
    ReDim strArrTmp(adrsDept1.RecordCount - 1, 3)
    cboRot.ColumnCount = 4
    cboRot.ListWidth = "6 cm"
    cboRot.ColumnWidths = "1cm;1 cm;2 cm; 2 cm"
    For bytTmp = 0 To adrsDept1.RecordCount - 1
        strArrTmp(bytTmp, 0) = adrsDept1("SCode")           '' Shift Code
        strArrTmp(bytTmp, 1) = adrsDept1("Mon_Oth")         '' Type
        strArrTmp(bytTmp, 2) = adrsDept1("Skp")             '' Skip type
        strArrTmp(bytTmp, 3) = adrsDept1("Pattern")         '' Shift Pattern
        adrsDept1.MoveNext
    Next
    cboRot.List = strArrTmp
    Erase strArrTmp
End If
Exit Sub
ERR_P:
    ShowError ("FillRotCombo :: " & Me.Caption)
End Sub

Private Sub FillShiftCombo()        '' Fills Shift ComboBox
On Error GoTo ERR_P
Dim strArrTmp() As String, bytTmp As Byte
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select Shift,Shf_In,Shf_Out from Instshft where shift <> '100' Order by Shift", _
ConMain, adOpenStatic
If Not (adrsDept1.BOF And adrsDept1.EOF) Then
    cboFix.ColumnCount = 3
    cboFix.ListWidth = "5.5 cm"
    cboFix.ColumnWidths = "1.5 cm;2  cm;2 cm"
    ReDim Preserve strArrTmp(adrsDept1.RecordCount - 1, 2)
    For bytTmp = 0 To adrsDept1.RecordCount - 1
        strArrTmp(bytTmp, 0) = adrsDept1("Shift")                       '' Shift Code
        strArrTmp(bytTmp, 1) = Format(adrsDept1("Shf_In"), "00.00")     '' Shift In Time
        strArrTmp(bytTmp, 2) = Format(adrsDept1("Shf_Out"), "00.00")    '' Shift Out Time
        adrsDept1.MoveNext
    Next
    cboFix.List = strArrTmp
    '' For Details regarding Daily Processing.
    cboShift.List = strArrTmp
    cboShift.AddItem ""
    cboBlank.List = strArrTmp
    cboBlank.AddItem ""
    Erase strArrTmp
End If
Exit Sub
ERR_P:
    ShowError ("FillShiftCombo :: " & Me.Caption)
End Sub

Private Sub ChangeMode()        '' Takes Action Based on the Byte Mode
Select Case bytShiftModeFrm
    Case 1
        Call ViewAction         '' When no Mode
    Case 2
        Call EditAction         '' When in Edit Mode
    Case 3
        Call EmpAction          '' When in Employee Master Mode
    Case 4
        Call ViewAction         '' Load Mode
        cboCode.Visible = True: txtCode.Visible = False
        If adrsPaid.State = 1 Then adrsPaid.Close
        If Not GetFlagStatus("WO") Then
            adrsPaid.Open "Select EmpCode,Name,JoinDate,Shf_Date,Cat," & _
            "" & strKOff & ",Off2,STyp,F_Shf,SCode,WO_1_3,WO_2_4,WOHLAction,Action3Shift,AutoForPunch" & _
            ",ActionBlank from Empmst Order by EmpCode", ConMain, adOpenStatic
        Else
            adrsPaid.Open "Select EmpCode,Name,JoinDate,Shf_Date,Cat," & _
            "" & strKOff & ",Off2,STyp,F_Shf,SCode,WO_1_3,WO_2_4,WOHLAction,Action3Shift,AutoForPunch" & _
            ",ActionBlank,WeekOffRule  from Empmst Order by EmpCode", ConMain, adOpenStatic
        End If
End Select
End Sub

Private Sub EmpAction()         '' Action Taken in the Form Load event When the
cboCode.Visible = False         '' bytShiftModeFrm is 3 i.e from Employee Master Form
txtCode.Visible = True
txtDate.Text = DateDisp(Shft.startdate)
txtCode.Text = Shft.Empcode
lblName.Caption = strRotPass
'' Shift Type
cboType.Text = Shft.ShiftType
Call AdjustType
'' Shift Code
If cboType.Text = "F" Then
    If Shft.ShiftCode <> "" Then cboFix.Text = Trim(Shft.ShiftCode)     ''Add Trim
Else
    If Shft.ShiftCode <> "" Then cboRot.Text = Trim(Shft.ShiftCode)     ''Add Trim By  08-11-08
End If
'' Week Offs
If Not IsNull(Shft.WO) And Shft.WO <> "" Then
    cboWO.Value = RetDayText(Shft.WO)
End If
'' Additional Week Offs
'' 0
If Not IsNull(Shft.WO1) And Shft.WO1 <> "" Then
    chkAWO(0).Value = 1
    cboAWO(0).Value = RetDayText(Shft.WO1)
Else
    cboAWO(0).Enabled = False
End If
'' 1
If Not IsNull(Shft.WO2) And Shft.WO2 <> "" Then
    chkAWO(1).Value = 1
    cboAWO(1).Value = RetDayText(Shft.WO2)
Else
    cboAWO(1).Enabled = False
End If
'' 2
If Not IsNull(Shft.WO3) And Shft.WO3 <> "" Then
    chkAWO(2).Value = 1
    cboAWO(2).Value = RetDayText(Shft.WO3)
Else
    cboAWO(2).Enabled = False
End If
''
On Error Resume Next
optDet(Shft.WOHLAction).Value = True
If Shft.WOHLAction = 2 Then cboShift.Text = Shft.Action3Shift
chkAuto.Value = IIf(Shft.AutoOnPunch, 1, 0)
If Trim(Shft.ActionBlank) = "" Then
    optBlank(0).Value = True
Else
    optBlank(1).Value = True
    cboBlank.Text = Shft.ActionBlank
End If
''
On Error GoTo 0
If cmdEditSave.Enabled = False Then
    bytShiftModeFrm = 1
    Call ChangeMode
    Exit Sub
End If
Call SetButtonCap(2)
bytShiftModeFrm = 2

End Sub

Private Sub ViewAction()    '' Action Taken when there is no Specific Mode
'' Disable Necessary Controls
frInfo.Enabled = False
frWO.Enabled = False
frAWO.Enabled = False
frOther.Enabled = False

frmWeekOffOfMonth.Enabled = False
frmWoRule.Enabled = False
'' Set Necessary Captions
Call SetButtonCap
End Sub

Private Sub EditAction()    '' Action Taken When Edit Mode
'' Disable Necessary Controls
frInfo.Enabled = True
frWO.Enabled = True
frAWO.Enabled = True
frOther.Enabled = True

frmWeekOffOfMonth.Enabled = True
frmWoRule.Enabled = True
'' Set Necessary Captions
Call SetButtonCap(2)
End Sub

Private Function ValidateModMaster() As Boolean     '' Checks for the Validations
On Error GoTo ERR_P
ValidateModMaster = True
'' Valid date
If Trim(txtDate.Text) = "" Then
    MsgBox NewCaptionTxt("24015", adrsC), vbExclamation
    txtDate.Text = ""
    txtDate.SetFocus
    ValidateModMaster = False
    Exit Function
Else
    Shft.startdate = DateCompDate(txtDate.Text)       '' Get the Shift Date
End If
If bytShfMode <> 1 And bytShfMode <> 3 Then
    If DateCompDate(Shft.startdate) < adrsPaid("Joindate") Then
        MsgBox NewCaptionTxt("24016", adrsC), vbExclamation
        txtDate.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
End If
'' Valid Shift Selected
Select Case cboType.Text
    Case ""
        MsgBox NewCaptionTxt("24017", adrsC), vbExclamation
        cboType.SetFocus
        ValidateModMaster = False
        Exit Function
    Case "F"
        If cboFix.Text = "" Then
            MsgBox NewCaptionTxt("24018", adrsC), vbExclamation
            cboFix.SetFocus
            ValidateModMaster = False
            Exit Function
        Else
            Shft.ShiftType = cboType.Text
            Shft.ShiftCode = cboFix.Text    '' Shift Code
        End If
    Case "R"
        If cboRot.Text = "" Then
            MsgBox NewCaptionTxt("24019", adrsC), vbExclamation
            cboRot.SetFocus
            ValidateModMaster = False
            Exit Function
        Else
            Shft.ShiftType = cboType.Text
            Shft.ShiftCode = cboRot.Text    '' Shift Code
        End If
End Select
'' Week Off Selected
If cboWO.Text = "" And (cboAWO(0).Text <> "" Or cboAWO(1).Text <> "" Or _
cboAWO(2).Text <> "") Then
    MsgBox NewCaptionTxt("24020", adrsC), vbExclamation
    cboWO.SetFocus
    ValidateModMaster = False
    Exit Function
End If
Shft.WO = Left(cboWO.Text, 2)       '' Week Off
'' Additional Week Offs
'' 0
If chkAWO(0).Value = 1 And cboAWO(0).Text = "" Then
    MsgBox NewCaptionTxt("24021", adrsC), vbExclamation
    cboAWO(0).SetFocus
    ValidateModMaster = False
    Exit Function
Else
    Shft.WO1 = Left(cboAWO(0).Text, 2)
End If
'' 1
If chkAWO(1).Value = 1 And cboAWO(1).Text = "" Then
    MsgBox NewCaptionTxt("24022", adrsC), vbExclamation
    cboAWO(1).SetFocus
    ValidateModMaster = False
    Exit Function
Else
    Shft.WO2 = Left(cboAWO(1).Text, 2)
End If
'' 2
If chkAWO(2).Value = 1 And cboAWO(2).Text = "" Then
    MsgBox NewCaptionTxt("24023", adrsC), vbExclamation
    cboAWO(2).SetFocus
    ValidateModMaster = False
    Exit Function
Else
    Shft.WO3 = Left(cboAWO(2).Text, 2)
End If

'' For Details Regarding Daily Processing
If optDet(2).Value = True And Trim(cboShift.Text) = "" Then
    MsgBox NewCaptionTxt("24033", adrsC), vbExclamation
    cboShift.SetFocus
    ValidateModMaster = False
    Exit Function
End If
If optBlank(1).Value = True And Trim(cboBlank.Text) = "" Then
    MsgBox NewCaptionTxt("24034", adrsC), vbExclamation
    cboBlank.SetFocus
    ValidateModMaster = False
    Exit Function
End If
'' Fill in Shift Values
If optDet(0).Value = True Then Shft.WOHLAction = 0
If optDet(1).Value = True Then Shft.WOHLAction = 1
If optDet(2).Value = True Then Shft.WOHLAction = 2
Shft.Action3Shift = Trim(cboShift.Text)
Shft.AutoOnPunch = IIf(chkAuto.Value = 1, True, False)
If optBlank(0).Value = True Then
    Shft.ActionBlank = ""
Else
    Shft.ActionBlank = Trim(cboBlank.Text)
End If
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
    ValidateModMaster = False
End Function

Private Function SaveModMaster() As Boolean     '' Saves Data After Validations
On Error GoTo ERR_P
SaveModMaster = True
Select Case bytShfMode
    Case 1      '' New employee is Added
    Case 2      '' Employee Code Comes from cboCode.Text
        Call SaveModLog                         '' Save the Edit Log
        If Shft.ShiftType = "F" Then
            ConMain.Execute "Update EmpMst Set STyp='F',F_Shf='" & _
            Shft.ShiftCode & "',SCode='100',Shf_Date=" & strDTEnc & DateCompStr(Shft.startdate) & _
            strDTEnc & "," & strKOff & "='" & Shft.WO & "',Off2='" & Shft.WO1 & "',WO_1_3='" & _
            Shft.WO2 & "',WO_2_4='" & Shft.WO3 & "' Where Empcode='" & cboCode.Text & "'"
        Else
            ConMain.Execute "Update EmpMst Set STyp='R',F_Shf=''" & _
            ",SCode='" & Shft.ShiftCode & "',Shf_Date=" & strDTEnc & DateCompStr(Shft.startdate) & _
            strDTEnc & "," & strKOff & "='" & Shft.WO & "',Off2='" & Shft.WO1 & "',WO_1_3='" & _
            Shft.WO2 & "',WO_2_4='" & Shft.WO3 & "' Where Empcode='" & cboCode.Text & "'"
        End If
        '' For Details Regarding Daily Processing
        ConMain.Execute "Update EmpMst Set WOHLAction=" & Shft.WOHLAction & _
        ",Action3Shift='" & Shft.Action3Shift & "',AutoForPunch=" & _
        IIf(Shft.AutoOnPunch, 1, 0) & ",ActionBlank='" & Shft.ActionBlank & "' Where " & _
        "Empcode='" & cboCode.Text & "'"
        ''
    Case 3     '' Employee Code Comes from txtCode.Text
        Call SaveModSchLog                         '' Save the Edit Log
        If Shft.ShiftType = "F" Then
            ConMain.Execute "Update EmpMst Set STyp='F',F_Shf='" & _
            Shft.ShiftCode & "',SCode='100',Shf_Date=" & strDTEnc & DateCompStr(Shft.startdate) & _
            strDTEnc & "," & strKOff & "='" & Shft.WO & "',Off2='" & Shft.WO1 & "',WO_1_3='" & _
            Shft.WO2 & "',WO_2_4='" & Shft.WO3 & "' Where Empcode='" & txtCode.Text & "'"
        Else
            ConMain.Execute "Update EmpMst Set STyp='R',F_Shf=''" & _
            ",SCode='" & Shft.ShiftCode & "',Shf_Date=" & strDTEnc & DateCompStr(Shft.startdate) & _
            strDTEnc & "," & strKOff & "='" & Shft.WO & "',Off2='" & Shft.WO1 & "',WO_1_3='" & _
            Shft.WO2 & "',WO_2_4='" & Shft.WO3 & "' Where Empcode='" & txtCode.Text & "'"
        End If
        '' For Details Regarding Daily Processing
        ConMain.Execute "Update EmpMst Set WOHLAction=" & Shft.WOHLAction & _
        ",Action3Shift='" & Shft.Action3Shift & "',AutoForPunch=" & _
        IIf(Shft.AutoOnPunch, 1, 0) & ",ActionBlank='" & Shft.ActionBlank & "' Where " & _
        "Empcode='" & txtCode.Text & "'"
        EmailSend = True
   
        ''
End Select
Exit Function
ERR_P:
SaveModMaster = False
End Function

Private Sub GetEmpDetails()         '' Gets the Employee Details
On Error GoTo ERR_P
Dim bytTmp As Byte
adrsPaid.MoveFirst
adrsPaid.Find "EmpCode='" & cboCode.Text & "'"
If Not adrsPaid.EOF Then
    cboType.Text = adrsPaid("STyp")
    If cboType.Text = "F" Then
        cboFix.Text = Trim(GetShiftEmp(adrsPaid("F_Shf")))        '' Fixed
    Else
        cboRot.Text = Trim(GetShiftEmp(adrsPaid("SCode"), 2))     '' Rotational ''Add Trim By
    End If
    Call AdjustType
    '' Shift Date
    txtDate.Text = DateDisp(CStr(adrsPaid("Shf_Date")))
    '' Week Off
    If Trim(adrsPaid("Off")) <> "" Then
        cboWO.Text = RetDayText(adrsPaid("Off"))
    End If
    '' Additional Week Off
    '' 0
    If Trim(adrsPaid("Off2")) <> "" Then
        cboAWO(0).Text = RetDayText(adrsPaid("Off2"))
        chkAWO(0).Value = 1
    Else
        chkAWO(0).Value = 0
    End If
    '' 1
    If Trim(adrsPaid("WO_1_3")) <> "" Then
        cboAWO(1).Text = RetDayText(adrsPaid("WO_1_3"))
        chkAWO(1).Value = 1
    Else
        chkAWO(1).Value = 0
    End If
    '' 2
    If Trim(adrsPaid("WO_2_4")) <> "" Then
        cboAWO(2).Text = RetDayText(adrsPaid("WO_2_4"))
        chkAWO(2).Value = 1
    Else
        chkAWO(2).Value = 0
    End If
    '' For Details regarding Daily Processing
    If Not IsNull(adrsPaid("WOHLAction")) Then
        bytTmp = adrsPaid("WOHLAction")
    Else
        bytTmp = 0
    End If
    optDet(bytTmp).Value = True
    If bytTmp = 2 Then
        Err.clear
        On Error Resume Next
        cboShift.Value = adrsPaid("Action3Shift")
        If Err.Number <> 0 Then cboShift.Value = ""
        On Error GoTo ERR_P
    End If
    chkAuto.Value = IIf(adrsPaid("AutoForPunch") = 1, 1, 0)
    If Not IsNull(adrsPaid("ActionBlank")) Then
        If Trim(adrsPaid("ActionBlank")) = "" Then
            optBlank(0).Value = True
        Else
            Err.clear
            On Error Resume Next
            optBlank(1).Value = True
            cboBlank.Value = adrsPaid("ActionBlank")
            If Err.Number <> 0 Then
                optBlank(0).Value = True
            End If
            On Error GoTo ERR_P
        End If
    Else
        optBlank(0).Value = True
    End If
End If
Exit Sub
ERR_P:
    ShowError ("Get Employee Details :: " & Me.Caption)
    Resume Next
End Sub

Private Function GetShiftEmp(ByVal strShiftTmp As String, _
Optional bytFlagTmp As Byte = 1) As String      '' Returns Valid Shifts
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close
Select Case bytFlagTmp
    Case 1      '' Fixed
        adrsDept1.Open "Select Count(*) from InstShft where Shift='" & strShiftTmp & "'", _
        ConMain
    Case 2      '' Rotational
        adrsDept1.Open "Select Count(*) from Ro_Shift where SCode='" & strShiftTmp & "'", _
        ConMain
End Select
Select Case adrsDept1(0)
    Case Null               '' Returns NULL
        GetShiftEmp = ""
    Case Empty              '' Returns Empty
        GetShiftEmp = ""
    Case 0                  '' No Shift is Found
        GetShiftEmp = ""
    Case Else               '' Shift is Found
        GetShiftEmp = strShiftTmp
End Select
Exit Function
ERR_P:
    ShowError ("get Shift Employee :: " & Me.Caption)
    GetShiftEmp = ""
End Function

Private Sub AdjustType()    '' Adjusts the Visibility of the Rotational and Fixed Shift
If cboType.Text = "F" Then  '' Controls
    lblFix.Visible = True
    cboFix.Visible = True
    lblRot.Visible = False
    cboRot.Visible = False
Else
    lblRot.Visible = True
    cboRot.Visible = True
    lblFix.Visible = False
    cboFix.Visible = False
End If
End Sub

Private Sub optBlank_Click(Index As Integer)
On Error GoTo ERR_P
Select Case Index
    Case 1
        cboBlank.Enabled = True
    Case Else
        cboBlank.Enabled = False
        cboBlank.Value = ""
End Select
Exit Sub
ERR_P:
    ShowError ("Option Blank ::" & Me.Caption)
    Resume Next
End Sub

Private Sub optDet_Click(Index As Integer)
On Error GoTo ERR_P
Select Case Index
    Case 2
        cboShift.Enabled = True
    Case Else
        cboShift.Enabled = False
        cboShift.Value = ""
End Select
Exit Sub
ERR_P:
    ShowError ("Option Details ::" & Me.Caption)
    Resume Next
End Sub

Private Sub txtDate_Click()
varCalDt = ""
varCalDt = Trim(txtDate.Text)
txtDate.Text = ""
Call ShowCalendar
End Sub

Private Sub txtDate_GotFocus()
    Call GF(txtDate)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    Call CDK(txtDate, KeyAscii)
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
    If Not ValidDate(txtDate) Then txtDate.SetFocus: Cancel = True
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 15)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edit Shift Schdule Of Employee " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModSchLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 3, 15)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edit Shift Schdule Of Employee " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Function RetDayText(ByVal str1 As String)
Select Case UCase(str1)
    Case "SU": RetDayText = "Sun"
    Case "MO": RetDayText = "Mon"
    Case "TU": RetDayText = "Tue"
    Case "WE": RetDayText = "Wed"
    Case "TH": RetDayText = "Thu"
    Case "FR": RetDayText = "Fri"
    Case "SA": RetDayText = "Sat"
End Select
End Function
