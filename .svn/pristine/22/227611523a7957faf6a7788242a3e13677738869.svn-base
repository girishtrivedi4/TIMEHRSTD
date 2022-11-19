VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reports"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   147
      Top             =   0
      Width           =   3015
      Begin VB.ComboBox cboSelectReport 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label19 
         Caption         =   "Select Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   148
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      Height          =   855
      Index           =   0
      Left            =   6120
      TabIndex        =   143
      Top             =   5520
      Width           =   3255
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   1800
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frSel 
      Caption         =   "Selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   45
      TabIndex        =   61
      Top             =   960
      Width           =   9285
      Begin VB.Frame frDesig 
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5940
         TabIndex        =   140
         Top             =   990
         Visible         =   0   'False
         Width           =   1455
         Begin VB.Label lblToDesig 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   142
            Top             =   720
            Width           =   225
         End
         Begin VB.Label lblFrDesig 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FROM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   45
            TabIndex        =   141
            Top             =   360
            Width           =   465
         End
         Begin MSForms.ComboBox cboFrDesigSel 
            Height          =   330
            Left            =   540
            TabIndex        =   15
            Top             =   240
            Width           =   885
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1561;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboToDesigSel 
            Height          =   330
            Left            =   540
            TabIndex        =   16
            Top             =   600
            Width           =   885
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1561;582"
            ListWidth       =   6000
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
            Object.Width           =   "1411;5291"
         End
      End
      Begin VB.Frame frCatList 
         Height          =   1095
         Left            =   10320
         TabIndex        =   91
         Top             =   4080
         Visible         =   0   'False
         Width           =   1935
         Begin MSForms.ListBox catlist 
            Height          =   615
            Left            =   0
            TabIndex        =   92
            Top             =   405
            Width           =   1935
            BorderStyle     =   1
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "3413;1085"
            MatchEntry      =   0
            MultiSelect     =   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox catall 
            Height          =   255
            Left            =   375
            TabIndex        =   93
            Top             =   150
            Width           =   1335
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2355;450"
            Value           =   "0"
            Caption         =   "CHECK ALL"
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame frDeptList 
         Height          =   1095
         Left            =   9600
         TabIndex        =   88
         Top             =   3840
         Visible         =   0   'False
         Width           =   1935
         Begin MSForms.ListBox cbodept 
            Height          =   615
            Left            =   0
            TabIndex        =   90
            Top             =   405
            Width           =   1935
            BorderStyle     =   1
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "3413;1085"
            MatchEntry      =   0
            MultiSelect     =   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox deptall 
            Height          =   255
            Left            =   375
            TabIndex        =   89
            Top             =   150
            Width           =   1335
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2355;450"
            Value           =   "0"
            Caption         =   "CHECK ALL"
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame frshift 
         Caption         =   "Shift"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   5895
         TabIndex        =   137
         Top             =   2070
         Visible         =   0   'False
         Width           =   1500
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   139
            Top             =   630
            Width           =   225
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FROM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   30
            TabIndex        =   138
            Top             =   360
            Width           =   465
         End
         Begin MSForms.ComboBox cmbFrShiftSel 
            Height          =   330
            Left            =   495
            TabIndex        =   26
            Top             =   240
            Width           =   930
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1640;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbToShiftSel 
            Height          =   330
            Left            =   495
            TabIndex        =   27
            Top             =   600
            Width           =   930
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1640;582"
            ListWidth       =   6000
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
            Object.Width           =   "1411;5291"
         End
      End
      Begin VB.Frame frGrade 
         Caption         =   "Grade"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5940
         TabIndex        =   132
         Top             =   990
         Visible         =   0   'False
         Width           =   1455
         Begin MSForms.ComboBox cmbToGrSel 
            Height          =   330
            Left            =   540
            TabIndex        =   136
            Top             =   600
            Width           =   885
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1561;582"
            ListWidth       =   6000
            ColumnCount     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbFrGrSel 
            Height          =   330
            Left            =   540
            TabIndex        =   135
            Top             =   240
            Width           =   885
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1561;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FROM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   45
            TabIndex        =   134
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   133
            Top             =   720
            Width           =   225
         End
      End
      Begin MSMAPI.MAPIMessages MAPIMessages1 
         Left            =   6000
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         AddressEditFieldCount=   1
         AddressModifiable=   0   'False
         AddressResolveUI=   0   'False
         FetchSorted     =   0   'False
         FetchUnreadOnly =   0   'False
      End
      Begin VB.Frame frDept 
         Height          =   1095
         Left            =   2040
         TabIndex        =   85
         Top             =   480
         Width           =   1935
         Begin MSForms.ComboBox cmbFrDepSel 
            Height          =   330
            Left            =   720
            TabIndex        =   8
            Top             =   240
            Width           =   1080
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1905;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbToDepSel 
            Height          =   330
            Left            =   720
            TabIndex        =   9
            Top             =   600
            Width           =   1080
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1905;582"
            ListWidth       =   6000
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
            Object.Width           =   "1411;5291"
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FROM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   86
            Top             =   720
            Width           =   225
         End
      End
      Begin VB.Frame frEmpList 
         Height          =   1095
         Left            =   2040
         TabIndex        =   115
         Top             =   5040
         Visible         =   0   'False
         Width           =   1935
         Begin MSForms.CheckBox chkEmp 
            Height          =   255
            Left            =   360
            TabIndex        =   117
            Top             =   150
            Width           =   1335
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2355;450"
            Value           =   "0"
            Caption         =   "CHECK ALL"
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin MSForms.ListBox cboEmp 
            Height          =   615
            Left            =   0
            TabIndex        =   116
            Top             =   405
            Width           =   1935
            BorderStyle     =   1
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "3413;1085"
            MatchEntry      =   0
            MultiSelect     =   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame frEmp 
         Height          =   1095
         Left            =   120
         TabIndex        =   112
         Top             =   480
         Width           =   1935
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   114
            Top             =   720
            Width           =   225
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FROM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   110
            TabIndex        =   113
            Top             =   360
            Width           =   465
         End
         Begin MSForms.ComboBox cmbToEmpSel 
            Height          =   330
            Left            =   600
            TabIndex        =   6
            Top             =   600
            Width           =   1215
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "2143;582"
            ListWidth       =   6000
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
            Object.Width           =   "1411;5291"
         End
         Begin MSForms.ComboBox cmbFrEmpSel 
            Height          =   330
            Left            =   600
            TabIndex        =   5
            Top             =   240
            Width           =   1215
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "2143;582"
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame frGrp 
         Height          =   1095
         Left            =   120
         TabIndex        =   100
         Top             =   2025
         Width           =   1935
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   102
            Top             =   720
            Width           =   225
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FROM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   101
            Top             =   360
            Width           =   465
         End
         Begin MSForms.ComboBox cmbToGrpSel 
            Height          =   330
            Left            =   720
            TabIndex        =   18
            Top             =   600
            Width           =   1095
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1931;582"
            ListWidth       =   6000
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
            Object.Width           =   "1411;5291"
         End
         Begin MSForms.ComboBox cmbFrGrpSel 
            Height          =   330
            Left            =   720
            TabIndex        =   17
            Top             =   240
            Width           =   1095
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1931;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame frDivList 
         Height          =   1095
         Left            =   10680
         TabIndex        =   97
         Top             =   3600
         Visible         =   0   'False
         Width           =   1935
         Begin MSForms.ListBox cboDiv 
            Height          =   615
            Left            =   0
            TabIndex        =   99
            Top             =   405
            Width           =   1935
            BorderStyle     =   1
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "3413;1085"
            MatchEntry      =   0
            MultiSelect     =   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkDiv 
            Height          =   255
            Left            =   375
            TabIndex        =   98
            Top             =   150
            Width           =   1335
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2355;450"
            Value           =   "0"
            Caption         =   "CHECK ALL"
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
      End
      Begin VB.OptionButton optGrpDiv1 
         Caption         =   "Division"
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
         Index           =   6
         Left            =   8760
         TabIndex        =   77
         Top             =   7080
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.OptionButton optGrpLoc1 
         Caption         =   "Location"
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
         Index           =   5
         Left            =   9000
         TabIndex        =   74
         Top             =   7320
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.OptionButton optGrpDC1 
         Caption         =   "Department / Category"
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
         Index           =   4
         Left            =   9600
         TabIndex        =   73
         Top             =   6960
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.OptionButton optGrpGrp1 
         Caption         =   "Group"
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
         Left            =   10800
         TabIndex        =   72
         Top             =   4680
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optGrpCat1 
         Caption         =   "Category"
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
         Left            =   8520
         TabIndex        =   71
         Top             =   6960
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.OptionButton optGrpDep1 
         Caption         =   "Department"
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
         Left            =   8520
         TabIndex        =   70
         Top             =   5400
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton optGrpEmp1 
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
         Index           =   0
         Left            =   8400
         TabIndex        =   69
         Top             =   7440
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Frame frLoc 
         Height          =   1095
         Left            =   2040
         TabIndex        =   106
         Top             =   2025
         Width           =   1935
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   108
            Top             =   720
            Width           =   225
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FROM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   107
            Top             =   360
            Width           =   465
         End
         Begin MSForms.ComboBox cmbToLocSel 
            Height          =   330
            Left            =   720
            TabIndex        =   21
            Top             =   600
            Width           =   1095
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1931;582"
            ListWidth       =   6000
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
            Object.Width           =   "1411;5291"
         End
         Begin MSForms.ComboBox cmbFrLocSel 
            Height          =   330
            Left            =   720
            TabIndex        =   20
            Top             =   240
            Width           =   1095
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1931;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame frDiv 
         Height          =   1095
         Left            =   3960
         TabIndex        =   94
         Top             =   2025
         Width           =   1935
         Begin MSForms.ComboBox cmbToDivSel 
            Height          =   330
            Left            =   720
            TabIndex        =   24
            Top             =   600
            Width           =   1095
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1931;582"
            ListWidth       =   6000
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
            Object.Width           =   "1411;5291"
         End
         Begin MSForms.ComboBox cmbFrDivSel 
            Height          =   330
            Left            =   720
            TabIndex        =   23
            Top             =   240
            Width           =   1095
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1931;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FROM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   96
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   95
            Top             =   720
            Width           =   225
         End
      End
      Begin VB.Frame frmCat 
         Height          =   1095
         Left            =   3960
         TabIndex        =   82
         Top             =   480
         Width           =   1935
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   84
            Top             =   720
            Width           =   225
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FROM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   360
            Width           =   465
         End
         Begin MSForms.ComboBox cmbFrCatSel 
            Height          =   330
            Left            =   720
            TabIndex        =   11
            Top             =   240
            Width           =   1095
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1931;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbToCatSel 
            Height          =   330
            Left            =   720
            TabIndex        =   12
            Top             =   600
            Width           =   1095
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1931;582"
            ListWidth       =   6000
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
            Object.Width           =   "1411;5291"
         End
      End
      Begin VB.Frame frGrpList 
         Height          =   1095
         Left            =   10920
         TabIndex        =   103
         Top             =   3720
         Visible         =   0   'False
         Width           =   1935
         Begin MSForms.CheckBox chkGroup 
            Height          =   255
            Left            =   375
            TabIndex        =   105
            Top             =   150
            Width           =   1335
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2355;450"
            Value           =   "0"
            Caption         =   "CHECK ALL"
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin MSForms.ListBox cboGroup 
            Height          =   615
            Left            =   0
            TabIndex        =   104
            Top             =   405
            Width           =   1935
            BorderStyle     =   1
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "3413;1085"
            MatchEntry      =   0
            MultiSelect     =   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame frLoclist 
         Height          =   1095
         Left            =   10560
         TabIndex        =   109
         Top             =   3600
         Visible         =   0   'False
         Width           =   1935
         Begin MSForms.ListBox cboLoc 
            Height          =   615
            Left            =   0
            TabIndex        =   111
            Top             =   405
            Width           =   1935
            BorderStyle     =   1
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "3413;1085"
            MatchEntry      =   0
            MultiSelect     =   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkLoc 
            Height          =   255
            Left            =   375
            TabIndex        =   110
            Top             =   150
            Width           =   1335
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2355;450"
            Value           =   "0"
            Caption         =   "CHECK ALL"
            SpecialEffect   =   0
            FontName        =   "Arial"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Label lblGrp 
         Height          =   255
         Index           =   7
         Left            =   7530
         TabIndex        =   124
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label lblGrp 
         Height          =   255
         Index           =   6
         Left            =   7530
         TabIndex        =   123
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblGrp 
         Height          =   255
         Index           =   5
         Left            =   7530
         TabIndex        =   122
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblGrp 
         Height          =   255
         Index           =   4
         Left            =   7530
         TabIndex        =   121
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblGrp 
         Height          =   255
         Index           =   3
         Left            =   7530
         TabIndex        =   120
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblGrp 
         Height          =   255
         Index           =   2
         Left            =   7530
         TabIndex        =   119
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblGrp 
         Height          =   255
         Index           =   1
         Left            =   7530
         TabIndex        =   118
         Top             =   600
         Width           =   255
      End
      Begin MSForms.CheckBox chkEmpSel 
         Height          =   255
         Left            =   330
         TabIndex        =   7
         Top             =   1560
         Width           =   1695
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2990;450"
         Value           =   "0"
         Caption         =   "Random Selection"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkLocSel 
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   3120
         Width           =   1695
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2990;450"
         Value           =   "0"
         Caption         =   "Random Selection"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkGrpSel 
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3120
         Width           =   1695
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2990;450"
         Value           =   "0"
         Caption         =   "Random Selection"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkDivSel 
         Height          =   255
         Left            =   4200
         TabIndex        =   25
         Top             =   3120
         Width           =   1695
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2990;450"
         Value           =   "0"
         Caption         =   "Random Selection"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkCatsel 
         Height          =   255
         Left            =   4200
         TabIndex        =   13
         Top             =   1560
         Width           =   1695
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2990;450"
         Value           =   "0"
         Caption         =   "Random Selection"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkDeptSel 
         Height          =   255
         Left            =   2250
         TabIndex        =   10
         Top             =   1560
         Width           =   1695
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2990;450"
         Value           =   "0"
         Caption         =   "Random Selection"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox optGrpDep2 
         Height          =   375
         Index           =   7
         Left            =   8520
         TabIndex        =   81
         Top             =   7080
         Visible         =   0   'False
         Width           =   1335
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2355;661"
         Value           =   "0"
         Caption         =   "Department"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   195
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox optGrp 
         Height          =   375
         Index           =   6
         Left            =   7890
         TabIndex        =   36
         Top             =   2280
         Width           =   1215
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2143;661"
         Value           =   "0"
         Caption         =   "Division"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   195
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox optGrp 
         Height          =   375
         Index           =   4
         Left            =   7890
         TabIndex        =   34
         Top             =   1560
         Width           =   1215
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2143;661"
         Value           =   "0"
         Caption         =   "Location"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   195
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox optGrp 
         Height          =   375
         Index           =   5
         Left            =   7890
         TabIndex        =   35
         Top             =   1920
         Width           =   1215
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2143;661"
         Value           =   "0"
         Caption         =   "Company"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   195
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox optGrp 
         Height          =   375
         Index           =   7
         Left            =   7890
         TabIndex        =   37
         Top             =   2640
         Width           =   1215
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2143;661"
         Value           =   "0"
         Caption         =   "Group"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   195
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox optGrp 
         Height          =   375
         Index           =   3
         Left            =   7890
         TabIndex        =   33
         Top             =   1200
         Width           =   1215
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2143;661"
         Value           =   "0"
         Caption         =   "Category"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   195
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox optGrp 
         Height          =   375
         Index           =   2
         Left            =   7890
         TabIndex        =   32
         Top             =   840
         Width           =   1335
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2355;661"
         Value           =   "0"
         Caption         =   "Department"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   195
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox optGrp 
         Height          =   375
         Index           =   1
         Left            =   7890
         TabIndex        =   31
         Top             =   480
         Width           =   1365
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2408;661"
         Value           =   "0"
         Caption         =   "Employee"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   195
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "S. Group"
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
         Index           =   0
         Left            =   6000
         TabIndex        =   79
         Top             =   1800
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSForms.ComboBox cboMain 
         Height          =   315
         Left            =   6000
         TabIndex        =   78
         Top             =   2160
         Visible         =   0   'False
         Width           =   1065
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1879;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   178
         FontPitchAndFamily=   2
         Object.Width           =   "1500;4500"
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   7395
         X2              =   7395
         Y1              =   150
         Y2              =   3600
      End
      Begin MSForms.ComboBox cmbFrComSel 
         Height          =   315
         Left            =   6000
         TabIndex        =   14
         Top             =   600
         Width           =   1305
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2302;556"
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   178
         FontPitchAndFamily=   2
         Object.Width           =   "1411;5291"
      End
      Begin VB.Label lblGroupBy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group By"
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
         Left            =   7710
         TabIndex        =   68
         Top             =   135
         Width           =   885
      End
      Begin VB.Label lblComSel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
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
         Left            =   6120
         TabIndex        =   67
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblGrpSel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
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
         Left            =   720
         TabIndex        =   65
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label lblCatSel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         Left            =   4440
         TabIndex        =   64
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblDepSel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Left            =   2400
         TabIndex        =   63
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label lblEmpSel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
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
         Left            =   480
         TabIndex        =   62
         Top             =   240
         Width           =   945
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   7425
         X2              =   7425
         Y1              =   150
         Y2              =   3600
      End
      Begin VB.Label lblLocSel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Left            =   2520
         TabIndex        =   66
         Top             =   1800
         Width           =   825
      End
      Begin VB.Label lblDivSel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
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
         Left            =   4560
         TabIndex        =   76
         Top             =   1800
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   3285
      TabIndex        =   130
      Top             =   1680
      Visible         =   0   'False
      Width           =   2760
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Please Wait...."
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -15
         TabIndex        =   131
         Top             =   105
         Width           =   2745
      End
   End
   Begin VB.Frame frmCboOpt 
      Height          =   975
      Left            =   120
      TabIndex        =   125
      Top             =   4440
      Width           =   9195
      Begin VB.ComboBox cboPaperSize 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6390
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   480
         Width           =   2070
      End
      Begin VB.ComboBox cboPaperOrientation 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3450
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   480
         Width           =   2070
      End
      Begin VB.ComboBox cboPrinterDuplex 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   510
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   495
         Width           =   2070
      End
      Begin MSForms.CheckBox chkOne 
         Height          =   255
         Left            =   5280
         TabIndex        =   129
         Top             =   240
         Width           =   1575
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;450"
         Value           =   "0"
         Caption         =   "One Page at time"
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Duplexing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1005
         TabIndex        =   128
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Orientation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   3720
         TabIndex        =   127
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   8640
         TabIndex        =   126
         Top             =   960
         Width           =   915
      End
   End
   Begin VB.CheckBox chkCat 
      Caption         =   "All Category"
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
      Left            =   14400
      TabIndex        =   80
      Top             =   8880
      Visible         =   0   'False
      Width           =   1395
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   285
      Left            =   120
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   6480
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   503
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   12648447
      ForeColor       =   4194304
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      Appearance      =   0
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
   Begin MSMAPI.MAPISession ReportSession 
      Left            =   12120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages ReportMessage 
      Left            =   12240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.Frame frMonth 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3225
      TabIndex        =   50
      Top             =   0
      Width           =   6195
      Begin VB.ComboBox cmbMonYear 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   975
      End
      Begin VB.ComboBox cmbMonth 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   1425
      End
      Begin VB.Label lblMonYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   4260
         TabIndex        =   52
         Top             =   360
         Width           =   435
      End
      Begin VB.Label lblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report for the month of"
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
         Left            =   90
         TabIndex        =   51
         Top             =   360
         Width           =   2205
      End
   End
   Begin VB.Frame frYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   144
      Top             =   0
      Width           =   3315
      Begin VB.ComboBox cmbYear 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   145
         Top             =   300
         Width           =   915
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report for the Year"
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
         Left            =   240
         TabIndex        =   146
         Top             =   360
         Width           =   1980
      End
   End
   Begin VB.Frame frPeri 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3225
      TabIndex        =   56
      Top             =   0
      Width           =   5835
      Begin VB.TextBox txtFrPeri 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   58
         Tag             =   "D"
         Top             =   300
         Width           =   1125
      End
      Begin VB.TextBox txtToPeri 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   60
         Tag             =   "D"
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lblFrPeri 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report for the period from"
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
         Left            =   120
         TabIndex        =   57
         Top             =   360
         Width           =   2670
      End
      Begin VB.Label lblToPeri 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   4080
         TabIndex        =   59
         Top             =   360
         Width           =   300
      End
   End
   Begin VB.Frame frDly 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3225
      TabIndex        =   30
      Top             =   0
      Width           =   5835
      Begin VB.OptionButton optDly 
         Caption         =   "Unauthorized OT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   3600
         TabIndex        =   42
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtDaily 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   44
         Tag             =   "D"
         Top             =   210
         Width           =   1125
      End
      Begin VB.OptionButton optDly 
         Caption         =   "Performance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   480
         TabIndex        =   41
         Top             =   1740
         Width           =   2115
      End
      Begin MSForms.ComboBox cboShift 
         Height          =   345
         Left            =   4200
         TabIndex        =   46
         Top             =   218
         Width           =   1395
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "2461;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblShf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
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
         Left            =   3600
         TabIndex        =   45
         Top             =   270
         Width           =   375
      End
      Begin VB.Label lblDaily 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report for the Date"
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
         Left            =   120
         TabIndex        =   43
         Top             =   270
         Width           =   1980
      End
   End
   Begin VB.Frame frWeek 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3225
      TabIndex        =   47
      Top             =   0
      Width           =   4875
      Begin VB.TextBox txtWeek 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   49
         Tag             =   "D"
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label lblWeek 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report for the Week starting from "
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
         Left            =   240
         TabIndex        =   48
         Top             =   330
         Width           =   3480
      End
   End
   Begin VB.Frame frMast 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3225
      TabIndex        =   2
      Top             =   0
      Width           =   4275
      Begin VB.TextBox txtMastTo 
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
         Left            =   2940
         MaxLength       =   10
         TabIndex        =   55
         Tag             =   "D"
         Top             =   293
         Width           =   1095
      End
      Begin VB.TextBox txtMastFr 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "D"
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblMastTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   2430
         TabIndex        =   54
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblMastFr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
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
         Left            =   90
         TabIndex        =   53
         Top             =   360
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  '' Used for storing commandtext temporary
Private strlastdatem As String, strlastdatem1 As String, strFirstDateM As String, strAttachPath As String
Dim rpgroup As String
 '*****
Public clsInOperetor As New clsInOperetor
 '*****
''
'Dim Crystal As CRAXDRT.Application
'Dim Report As CRAXDRT.Report
Dim adrsC As New ADODB.Recordset
Dim strFileName As String ' 18-04-09
Dim SMnth As Integer, EMnth As Integer, SYr As Integer, EYr As Integer
Dim WithEvents oSectionGH As CRAXDRT.Section
Attribute oSectionGH.VB_VarHelpID = -1
Public tmpblnAssum As String

Private Sub catall_Click()
Dim i As Integer
If catall.Value = True Then
    If catlist.ListCount > 0 Then
        For i = 0 To catlist.ListCount - 1
            catlist.Selected(i) = True
        Next i
    End If
Else
        If catlist.ListCount > 0 Then
        For i = 0 To catlist.ListCount - 1
            catlist.Selected(i) = False
        Next i
    End If
    If catlist.ListCount > 0 Then
       catlist.Selected(0) = True
    End If
End If
End Sub
'******
Private Sub catlist_Click()
Call catlist_GotFocus
End Sub
Private Sub catlist_GotFocus()
catlist.Height = 1500
frCatList.Height = 2000
End Sub

Private Sub catlist_LostFocus()
catlist.Height = 615
frCatList.Height = 1095
End Sub

'*******
Private Sub cboDept_Click()
Call cboDept_GotFocus
End Sub

Private Sub cboDept_GotFocus()
cbodept.Height = 1500
frDeptList.Height = 2000
End Sub

Private Sub cboDept_LostFocus()
cbodept.Height = 615
frDeptList.Height = 1095
End Sub

Private Sub cboDiv_Click()
Call cboDiv_GotFocus
End Sub

Private Sub cboDiv_GotFocus()
cboDiv.Height = 900
frDivList.Height = 1300
End Sub

Private Sub cboDiv_LostFocus()
cboDiv.Height = 615
frDivList.Height = 1095
End Sub

Private Sub cboEmp_Click()
Call cboEmp_GotFocus
End Sub

Private Sub cboEmp_GotFocus()
cboEmp.Height = 1500
frEmpList.Height = 2000
End Sub

Private Sub cboEmp_LostFocus()
cboEmp.Height = 615
frEmpList.Height = 1095
End Sub

Private Sub cboFrDesigSel_Click()
    If cboFrDesigSel.ListIndex >= 0 Then cboToDesigSel.ListIndex = cboFrDesigSel.ListIndex
End Sub


Private Sub cboGroup_Click()
Call cboGroup_GotFocus
End Sub

Private Sub cboGroup_GotFocus()
cboGroup.Height = 900
frGrpList.Height = 1300
End Sub

Private Sub cboGroup_LostFocus()
cboGroup.Height = 615
frGrpList.Height = 1095
End Sub

Private Sub cboLoc_Click()
Call cboLoc_GotFocus
End Sub
Private Sub cboLoc_GotFocus()
cboLoc.Height = 900
frLoclist.Height = 1300
End Sub

Private Sub cboLoc_LostFocus()
cboLoc.Height = 615
frLoclist.Height = 1095
End Sub

Private Sub cboSelectReport_Click()

    Select Case ReportType
    Case "Daily"
        optDly_Click cboSelectReport.ItemData(cboSelectReport.ListIndex)
    Case "Weekly"
        optWee_Click cboSelectReport.ItemData(cboSelectReport.ListIndex)
    Case "Monthly"
        optMon_Click cboSelectReport.ItemData(cboSelectReport.ListIndex)
    Case "Yearly"
        optYea_Click cboSelectReport.ItemData(cboSelectReport.ListIndex)
    Case "Masters"
        optMas_Click cboSelectReport.ItemData(cboSelectReport.ListIndex)
    Case "Periodic"
        optPer_Click cboSelectReport.ItemData(cboSelectReport.ListIndex)
    End Select
End Sub

'
Private Sub chkCat_Click()
Call FrSelFill
End Sub
'******
Private Sub chkCatsel_Click()
If chkCatsel.Value = True Then
        frCatList.Visible = True
        frmCat.Visible = False
        catall.Value = False
        frCatList.Top = 480
        frCatList.Left = frmCat.Left    ' 4680
Else
        frCatList.Visible = False
        frmCat.Visible = True
End If
End Sub


Private Sub chkDeptSel_Click()
If chkDeptSel.Value = True Then
        frDeptList.Visible = True
        frDept.Visible = False
        deptall.Value = False
        frDeptList.Top = 480
        frDeptList.Left = frDept.Left   ' 2400
Else
        frDeptList.Visible = False
        frDept.Visible = True
End If
End Sub
''*******
Private Sub chkDiv_Click()
Dim i As Integer
If chkDiv.Value = True Then
    If cboDiv.ListCount > 0 Then
        For i = 0 To cboDiv.ListCount - 1
            cboDiv.Selected(i) = True
        Next i
    End If
Else
        If cboDiv.ListCount > 0 Then
             For i = 0 To cboDiv.ListCount - 1
                    cboDiv.Selected(i) = False
             Next i
        End If
        If cboDiv.ListCount > 0 Then
                 cboDiv.Selected(0) = True
         End If
             
End If
End Sub
'*******
Private Sub chkDivSel_Click()
If chkDivSel.Value = True Then
        frDivList.Visible = True
        frDiv.Visible = False
        chkDiv.Value = False
        frDivList.Top = 2150
        frDivList.Left = frDiv.Left ' 4685
Else
        frDivList.Visible = False
        frDiv.Visible = True
End If
End Sub


Private Sub chkEmp_Click()
Dim i As Integer
If chkEmp.Value = True Then
    If cboEmp.ListCount > 0 Then
        For i = 0 To cboEmp.ListCount - 1
            cboEmp.Selected(i) = True
        Next i
    End If
Else
        If cboEmp.ListCount > 0 Then
             For i = 0 To cboEmp.ListCount - 1
                    cboEmp.Selected(i) = False
             Next i
        End If
        If cboEmp.ListCount > 0 Then
                 cboEmp.Selected(0) = True
         End If
End If
End Sub

'******
Private Sub chkEmpSel_Click()
If chkEmpSel.Value = True Then
        frEmpList.Visible = True
        frEmp.Visible = False
        chkEmp.Value = False
        frEmpList.Top = 480
        frEmpList.Left = frEmp.Left ' 120
Else
        frEmpList.Visible = False
        frEmp.Visible = True
End If
End Sub

'*******
Private Sub chkGroup_Click()
Dim i As Integer
If chkGroup.Value = True Then
    If cboGroup.ListCount > 0 Then
        For i = 0 To cboGroup.ListCount - 1
            cboGroup.Selected(i) = True
        Next i
    End If
Else
        If cboGroup.ListCount > 0 Then
             For i = 0 To cboGroup.ListCount - 1
                    cboGroup.Selected(i) = False
             Next i
        End If
        If cboGroup.ListCount > 0 Then
                 cboGroup.Selected(0) = True
         End If
End If
End Sub
'*****
Private Sub chkGrpSel_Click()
If chkGrpSel.Value = True Then
        frGrpList.Visible = True
        frGrp.Visible = False
        chkGroup.Value = False
        frGrpList.Top = 2150
        frGrpList.Left = frGrp.Left ' 120
Else
        frGrpList.Visible = False
        frGrp.Visible = True
End If
End Sub

Private Sub chkLoc_Click()
Dim i As Integer
If chkLoc.Value = True Then
    If cboLoc.ListCount > 0 Then
        For i = 0 To cboLoc.ListCount - 1
            cboLoc.Selected(i) = True
        Next i
    End If
Else
        If cboLoc.ListCount > 0 Then
             For i = 0 To cboLoc.ListCount - 1
                    cboLoc.Selected(i) = False
             Next i
        End If
        If cboLoc.ListCount > 0 Then
                 cboLoc.Selected(0) = True
         End If
             
End If
End Sub

'''*******
Private Sub chkLocSel_Click()
If chkLocSel.Value = True Then
        frLoclist.Visible = True
        frLoc.Visible = False
        chkLoc.Value = False
        frLoclist.Top = 2150
        frLoclist.Left = frLoc.Left ' 2400
Else
        frLoclist.Visible = False
        frLoc.Visible = True
End If
End Sub

Private Sub cmbFrCatSel_Click()
    If cmbFrCatSel.ListIndex >= 0 Then cmbToCatSel.ListIndex = cmbFrCatSel.ListIndex
End Sub

Private Sub cmbFrDepSel_Click()
    If cmbFrDepSel.ListIndex >= 0 Then cmbToDepSel.ListIndex = cmbFrDepSel.ListIndex
End Sub

Private Sub cmbFrEmpSel_Click()
    If cmbFrEmpSel.ListIndex >= 0 Then cmbToEmpSel.ListIndex = cmbFrEmpSel.ListIndex
End Sub

Private Sub cmbFrGrpSel_Click()
    If cmbFrGrpSel.ListIndex >= 0 Then cmbToGrpSel.ListIndex = cmbFrGrpSel.ListIndex
End Sub

Private Sub cmbFrGrSel_Click()
    If cmbFrGrSel.ListIndex >= 0 Then cmbToGrSel.ListIndex = cmbFrGrSel.ListIndex
End Sub

Private Sub cmbFrLocSel_Click()
    If cmbFrLocSel.ListIndex >= 0 Then cmbToLocSel.ListIndex = cmbFrLocSel.ListIndex
End Sub
Private Sub cmbFrDivSel_Click()
    If cmbFrDivSel.ListIndex >= 0 Then cmbToDivSel.ListIndex = cmbFrDivSel.ListIndex
End Sub

Private Sub cmbMonth_Click()    '' Checks if the trasaction file for the selected month
If bytRepMode <> 7 Then         '' is available or not(Monthly Reports)
 If cmbMonth.Text <> "" And cmbMonYear.Text <> "" Then
     If Not FindTable(Left(cmbMonth.Text, 3) & Right(cmbMonYear.Text, 2) & "trn") Then
         MsgBox "Transaction File Not Found for the Month of  " & cmbMonth.Text & Space(1) & _
            cmbMonYear.Text, vbExclamation
         Exit Sub
     End If
 End If
End If
End Sub

Private Sub cmbMonYear_Click()  '' Checks if the trasaction file for the selected year
If bytRepMode <> 1 Then         '' and month is available or not.(Monthly Reports)
    If cmbMonth.Text <> "" And cmbMonYear.Text <> "" Then
        If Not FindTable(Left(cmbMonth.Text, 3) & Right(cmbMonYear.Text, 2) & "trn") Then
            MsgBox NewCaptionTxt("40065", adrsC) & cmbMonth.Text & Space(1) & _
            cmbMonYear.Text, vbExclamation
            Exit Sub
        End If
    End If
End If
End Sub

Private Sub cmbYear_Click()     '' Checks if the leave Transation file for the selected year
If bytRepMode <> 7 Then         '' is available or not.
    If cmbYear.Text <> "" Then
        If Not FindTable("lvtrn" & Right(Trim(cmbYear.Text), 2)) Then
            MsgBox NewCaptionTxt("00054", adrsMod) & cmbYear.Text & NewCaptionTxt("00055", adrsMod)
            Exit Sub
        End If
    End If
End If
End Sub

Public Sub cmdPreview_Click()  '' Gives Preview of the selected Report
'add by
Call EnableDisablComm(False)
Frame2.Visible = True
bytPrint = 2
Call ReportsMod
Call SetVarEmpty
'add by
Call EnableDisablComm(True)
Frame2.Visible = False
End Sub



'add by
Private Function EnableDisablComm(blnI As Boolean)
    cmdPreview.Enabled = blnI
    
  
'    cmdFile.Enabled = blnI
'    cmdSend.Enabled = blnI
'    cmdSelect.Enabled = blnI
End Function


Private Sub cmdExit_Click()
'added by  for leave balance option
If typOptIdx.bytMon = 9 Then
    frmLvBalOption.optMnthWise.Value = True
    Unload frmLvBalOption
    typDlyLvBal.typFdate = ""
    typDlyLvBal.typLdate = ""
End If
    Unload Me
End Sub


'*****
Private Sub deptall_Click()
Dim i As Integer
If deptall.Value = True Then
    If cbodept.ListCount > 0 Then
        For i = 0 To cbodept.ListCount - 1
            cbodept.Selected(i) = True
        Next i
    End If
Else
        If cbodept.ListCount > 0 Then
             For i = 0 To cbodept.ListCount - 1
                    cbodept.Selected(i) = False
             Next i
        End If
        If cbodept.ListCount > 0 Then
                 cbodept.Selected(0) = True
         End If
             
End If
End Sub

Private Sub Form_Activate()
'' Enable Date TextBoxes
txtDaily.Enabled = True
txtWeek.Enabled = True
txtFrPeri.Enabled = True
txtToPeri.Enabled = True
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
CatFlag = True
Dim lngLeft As Long
Dim lngwidth As Long
Dim lngTop As Long
Dim lngheight As Long
Call SetFormIcon(Me)            '' Set the Forms Icon
'Call ReportCon
'*****************************************************
'' Disable Date TextBoxes
'bytRepMode = 1
txtDaily.Enabled = False
txtWeek.Enabled = False
txtFrPeri.Enabled = False
txtToPeri.Enabled = False
bytAction = 0 ''No Action defined
RptChk = 0
MSF1.ColWidth(0) = MSF1.Width - 10

Call SetRepVars                 '' SETS DEFAULT VALUES OF TEXTBOXES AND RADIO BUTTONS
Call RetCaptions                '' RETRIEVES CAPTIONS FOR ALL LABELS
Call LoadSpecifics              '' procedure to Perform Other Actions on Load
If adrsDSR.State = 1 Then adrsDSR.Close
' changed for new logic
adrsDSR.Open "Select * from NewCaptions Where ID Like 'D%' or ID Like '00%'", ConMain, adOpenStatic

rpgroup = "groupmst.grupdesc ,catdesc.cat ,deptdesc.dept," & _
" CatDesc." & strKDesc & " as catdescdesc ,deptdesc." & strKDesc & _
" as deptdescdesc ,Location.Location,Location.LocDesc,Division.Div,Division.DivDesc,Company.Company,Company.CName"
    rpgroup1 = rpgroup & ",grade.gradecode "
   Call Call_Daily
   RptExp = 0
   ExpBl = True

Exit Sub
ERR_P:
    ShowError ("Load :: " & Me.Caption)
  Resume Next
End Sub

'---------------------------------------------------------------------------------------

Private Sub ReportSetting()
On Error GoTo ERR_P:
Call GroupRange
Call ShowPrinterDuplex
Call ShowPaperOrientation
Call ShowPaperSize
Exit Sub
ERR_P:
    ShowError ("Report Setting :: " & Me.Caption)
'Resume Next
End Sub
' intialising report first time.
Private Sub Call_Daily()
On Error GoTo ERR_P:

Call FrVisible
Call ShowGroupEx

Exit Sub
ERR_P:
    ShowError ("Call_Daily ::  " & Me.Caption)
End Sub

Private Sub RetCaptions()
On Error GoTo ERR_P
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '40%'", ConMain, adOpenStatic, adLockReadOnly
'form

Exit Sub
ERR_P:
    ShowError ("RetCaptions :: " & Me.Caption)
    Resume Next
End Sub

Private Sub FrVisible()
   frDly.Visible = False                       '' DEPENDING UPON THE CURRENT TAB SELECTED
    frWeek.Visible = False                      '' VISIBLES APPROPRIATE FRAME
    frMonth.Visible = False
    frYear.Visible = False
    frMast.Visible = False
    frPeri.Visible = False
    frSel.Enabled = True
    Select Case ReportType
        Case "Daily"
            frDly.Visible = True
        Case "Weekly"
            frWeek.Visible = True
        Case "Monthly"
            frMonth.Visible = True
        Case "Yearly"
            frYear.Visible = True
        Case "Periodic"
            frPeri.Visible = True
        Case "Masters"
            frMast.Visible = True
    End Select
  
End Sub

Private Function MonthYearFill()
On Error GoTo ERR_P
Dim i As Byte                       '' FILLS REQUIRED MONTH AND YEAR COMBOS
For i = 1 To 12
    cmbMonth.AddItem Choose(i, "January", "February", "March", "April", "May", _
    "June", "July", "August", "September", "October", "November", "December")
Next i

For i = 0 To 99
    cmbYear.AddItem (1997 + i)
    cmbMonYear.AddItem (1997 + i)
Next i

cmbYear.Text = Year(Date)
cmbMonth.Text = MonthName(Month(Date))
cmbMonYear.Text = Year(Date)
Exit Function
ERR_P:
    ShowError ("MonthYearFill :: " & Me.Caption)
End Function

Private Sub LoadSpecifics()
On Error GoTo ERR_P
Call SetToolTipText(Me)         '' Sets the ToolTipText for Date Text Boxes
Call MonthYearFill              '' FILLS MONTH AND YEAR COMBOS
Call FrSelFill                  '' FILLS SELECTION FRAME COMBOS
'Call FrchkFill                  '' FILLS CHECKBOX FRAME
'Call PutZeros                   '' SELECTS FIRST OPTION BUTTONS FOR ALL TABS
Call GetRights                  '' Check for Rights
Call FillReportType
Exit Sub
ERR_P:
    ShowError ("Load Specifics :: " & Me.Caption)
End Sub


Private Sub FillReportType()
    Select Case ReportType
    Case "Daily"
        bytRepMode = 1
        cboSelectReport.AddItem "Physical Arrival"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 0
        cboSelectReport.AddItem "Absent"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 1
'        cboSelectReport.AddItem ""
        cboSelectReport.AddItem "Late Arrival"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 3
        cboSelectReport.AddItem "Early Departure"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 4
        cboSelectReport.AddItem "Performance"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 5
        cboSelectReport.AddItem "Irregular"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 6
        cboSelectReport.AddItem "Authorized OT"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 7
        cboSelectReport.AddItem "Entries"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 8
        cboSelectReport.AddItem "Shift Arrangement"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 9
        cboSelectReport.AddItem "Manpower"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 10
        cboSelectReport.AddItem "Outdoor Duty"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 11
        cboSelectReport.AddItem "Summary"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 12
        cboSelectReport.AddItem "Unauthorized OT"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 13
    Case "Weekly"
        bytRepMode = 2
        cboSelectReport.AddItem "Performance"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 0
        cboSelectReport.AddItem "Absent"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 1
        cboSelectReport.AddItem "Attendance"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 2
        cboSelectReport.AddItem "Late Arrival"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 3
        cboSelectReport.AddItem "Early Departure"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 4
        cboSelectReport.AddItem "Overtime"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 5
        cboSelectReport.AddItem "Shift Schedule"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 6
        cboSelectReport.AddItem "Irregular"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 7
'        cboSelectReport.AddItem "Absentism Report"
'        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 8
    Case "Monthly"
        bytRepMode = 3
        cboSelectReport.AddItem ("Performance")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 0
        cboSelectReport.AddItem ("Attendance")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 1
        cboSelectReport.AddItem ("Muster Report")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 2
        cboSelectReport.AddItem ("Monthly Present")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 3
        cboSelectReport.AddItem ("Monthly Absent")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 4
        cboSelectReport.AddItem ("Overtime")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 5
        cboSelectReport.AddItem ("Overtime Paid")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 6
        cboSelectReport.AddItem ("Absent Memo")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 7
        cboSelectReport.AddItem ("Absent /Late /Early")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 8
        cboSelectReport.AddItem ("Leave Balance")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 9
        cboSelectReport.AddItem ("Late Arrival")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 10
        cboSelectReport.AddItem ("Early Departure")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 11
        cboSelectReport.AddItem ("Late Arrival Memo")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 12
        cboSelectReport.AddItem ("Early Departure Memo")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 13
        cboSelectReport.AddItem ("Leave Consumption")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 14
        cboSelectReport.AddItem ("Total Lates")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 15
        cboSelectReport.AddItem ("Total Earlys")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 16
        cboSelectReport.AddItem ("Shift schedule")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 17
        cboSelectReport.AddItem ("WO on Holiday")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 18
    
    Case "Yearly"
        bytRepMode = 4
        cboSelectReport.AddItem ("Absent")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 0
        cboSelectReport.AddItem ("Mandays")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 1
        cboSelectReport.AddItem ("Performance")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 2
        cboSelectReport.AddItem ("Present")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 3
        cboSelectReport.AddItem ("Leave Information")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 4
        cboSelectReport.AddItem ("Leave Balance")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 5
    Case "Periodic"
        bytRepMode = 6
        cboSelectReport.AddItem ("Performance")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 0
        cboSelectReport.AddItem ("Muster Report")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 1
        cboSelectReport.AddItem ("Overtime")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 2
        cboSelectReport.AddItem ("Late Arrival ")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 3
        cboSelectReport.AddItem ("Early Departure")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 4
        cboSelectReport.AddItem ("Continuous Absent")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 5
        cboSelectReport.AddItem ("Summary")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 6
        cboSelectReport.AddItem ("Physical Absent") ', 16)
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 16
        If GetFlagStatus("Jivdhani") Then
            cboSelectReport.AddItem ("Attendance Report")
            cboSelectReport.ItemData(cboSelectReport.NewIndex) = 7
        End If
    Case "Masters"
        bytRepMode = 5
        cboSelectReport.AddItem ("Employee List")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 0
        cboSelectReport.AddItem ("Employee Details")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 1
        cboSelectReport.AddItem ("Left Employee")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 2
        cboSelectReport.AddItem ("Leave")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 3
        cboSelectReport.AddItem ("Shift")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 4
        cboSelectReport.AddItem ("Rotational Shift")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 5
        cboSelectReport.AddItem ("Holiday")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 6
        cboSelectReport.AddItem "Department"
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 7
        cboSelectReport.AddItem ("Category")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 8
        cboSelectReport.AddItem ("Group")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 9
        cboSelectReport.AddItem ("Designation")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 10
        cboSelectReport.AddItem ("Division")
        cboSelectReport.ItemData(cboSelectReport.NewIndex) = 12
    End Select
    cboSelectReport.ListIndex = 0
    Call FrVisible
End Sub

Private Sub GetRights()
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(4, 16, 6, 1)
If strTmp = "1" Then
    cmdPreview.Enabled = True
Else
    cmdPreview.Enabled = False
End If
Exit Sub
ERR_P:
    ShowError ("GetRights::" & Me.Caption)
    cmdPreview.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If typOptIdx.bytPer = 7 Then typOptIdx.bytPer = 0
CatFlag = False
'If rptCon.State = 1 Then rptCon.Close
End Sub

Private Sub optDly_Click(Index As Integer)
lblShf.Visible = True
cboShift.Visible = True
If Index = 12 Or Index = 10 Or Index = 30 Then
     lblShf.Visible = False
    cboShift.Visible = False
End If




If Index = 12 Then

        frshift.Visible = False
        frGrade.Visible = False
      

End If

Select Case typOptIdx.bytDly
'10,1 7 9 12
    Case 0, 3, 4, 5, 6, 7, 8, 9, 10, 12, 32  '  25-05-09
        If optGrp(1).Value = True Then optGrp(1).Value = False
        optGrp(1).Enabled = False
End Select


typOptIdx.bytDly = Index
Call ReportSetting
End Sub

'**********************************
Public Sub optGrp_Click(Index As Integer)
If optGrp(Index).Value = True Then
        RptChk = RptChk + 1
        Dim i As Integer, k As Integer
        Dim j As Integer
        lblGrp(Index).Caption = RptChk
ElseIf optGrp(Index).Value = False Then
        j = Val(lblGrp(Index).Caption)
        lblGrp(Index).Caption = ""
        RptChk = RptChk - 1
        For i = 1 To 7
                If Val(lblGrp(i).Caption) > j Then
                     lblGrp(i).Caption = Val(lblGrp(i).Caption) - 1
                End If
        Next
End If
End Sub
'**********************************


Private Sub optWee_Click(Index As Integer)
typOptIdx.bytWek = Index
SetRepVars (2)
Call UnchkGrp
If Index = 9 Then
    For i = 1 To 7
        optGrp(i).Enabled = False
    Next
Else
    For i = 1 To 7
        optGrp(i).Enabled = True
    Next
End If
    
Select Case typOptIdx.bytWek
 Case 0, 7
         If optGrp(1).Value = True Then optGrp(1).Value = False
        optGrp(1).Enabled = False
 Case 9
 Case Else
 optGrp(1).Enabled = True
 End Select
Call ReportSetting
End Sub

Private Sub optMon_Click(Index As Integer)
typOptIdx.bytMon = Index
'add by
    Dim i As Byte
If typOptIdx.bytMon = 34 Then   ' 25-06
    Call UnchkGrp
    Call ShowGroupEx(False)
Else
   Call ShowGroupEx(True)
End If
If typOptIdx.bytMon = 22 Then
    For i = 1 To 7
        optGrp(i).Enabled = False
    Next
    'optGrp(2).Enabled = True
    optGrp(2).Value = vbChecked
    Exit Sub
Else
    For i = 1 To 7
        optGrp(i).Enabled = True
    Next
End If
''
If typOptIdx.bytMon = 41 Then
    For i = 1 To 7
        optGrp(i).Enabled = False
    Next
End If
Call SetRepVars(3)
Select Case typOptIdx.bytMon
    Case 1, 14, 2, 3, 4, 7, 12, 13, 17, 0, 5, 10, 11, 6 '32  25-05-09
        If optGrp(1).Value = True Then optGrp(1).Value = False

    Case 28 'For absenteeism  28-03
        lblGroupBy.Enabled = False
   Case Else
    optGrp(1).Enabled = True
End Select

Select Case typOptIdx.bytMon
    Case 1, 6, 15, 16, 35     'FOR ATTENDANCE, OT PAID, TOTAL LATES ,TOTAL EARLYS REPORTS
        MsgBox NewCaptionTxt("40070", adrsC) & vbCrLf & vbTab & _
        NewCaptionTxt("40071", adrsC), vbInformation
    Case 7, 12, 13          'FOR ABSENT,LATE,EARLY MEMO REPORTS
        bytMode = Index
        frmMemo.Show vbModal
    Case 9      'added by  for leave balance option
        Load frmLvBalOption
        frmLvBalOption.Show vbModal
    
    Case 51
        MsgBox ("If Daily and Monthly Process for Selected month and next 2 months is not done then do it first")
    
End Select
''''''''
''''''''
Call ReportSetting
'WriteLog ("End reports setings function")
End Sub


Private Sub optYea_Click(Index As Integer)
    typOptIdx.bytYer = Index
    Call SetRepVars(4)
    If Index = 11 Then
        Call UnchkGrp
         For i = 1 To 7
            optGrp(i).Enabled = False
        Next i
    Else
        Call UnchkGrp
        For i = 1 To 7
            optGrp(i).Enabled = True
        Next i
    End If
    Select Case Index


    Case 4, 1, 2, 3
        
       If optGrp(1).Value = True Then optGrp(1).Value = False
        optGrp(1).Enabled = False
    
    Case Else
        optGrp(1).Enabled = True
          optGrp(1).Value = vbUnchecked
        optGrp(1).Enabled = True
    End Select
    Call ReportSetting
End Sub

Private Sub optMas_Click(Index As Integer)
typOptIdx.bytMst = Index
Call SetRepVars(5)

    lblMastFr.Enabled = False
    lblMastTo.Enabled = False
    txtMastFr.Enabled = False
    txtMastTo.Enabled = False
Select Case Index
    Case 0, 1, 13, 14, 15                              'FOR EMPLOYEE REPORT
       If optGrp(1).Value = True Then optGrp(1).Value = False
        optGrp(1).Enabled = False
        cbodept.Enabled = True
        deptall.Enabled = True
        catall.Enabled = True
        catlist.Enabled = True
     '*****
        If Index = 0 Then
            lblMastFr.Enabled = True
            lblMastTo.Enabled = True
            txtMastFr.Enabled = True
            txtMastTo.Enabled = True
            frMast.Caption = "Joining"
            txtMastFr.Text = ""
            txtMastTo.Text = ""
        Else
            frMast.Caption = ""
        End If
        Call ShowGroupEx
    Case 2                                  'FOR LEFT EMPLOYEE REPORT
        lblMastFr.Enabled = True
        lblMastTo.Enabled = True
        txtMastFr.Enabled = True
        txtMastTo.Enabled = True
    '*****
        cbodept.Enabled = True
        deptall.Enabled = True
        catall.Enabled = True
        catlist.Enabled = True
     '*****
        Call ShowGroupEx
        frMast.Visible = True
        txtMastFr.SetFocus
    Case Else                               'FOR OTHER REPORTS
    '*****
        cbodept.Enabled = False
        deptall.Enabled = False
        catall.Enabled = False
        catlist.Enabled = False
     '*****
         Call UnchkGrp
        Call ShowGroupEx(False)
End Select
Call ReportSetting
'add by  for left employee display in combo box because for TGL hide from combo
If GetFlagStatus("leftemployee") Then
    If Index = 2 Then
        Call FillEmpCombo(1)
    Else
        Call FillEmpCombo(17)
    End If
End If
End Sub

Private Sub optPer_Click(Index As Integer)
typOptIdx.bytPer = Index
Call SetRepVars(6)

If typOptIdx.bytPer = 40 Then           ' for Nestle samalkha
   For i = 1 To 7
            optGrp(i).Enabled = False
        Next
End If
Select Case Index
  
  
    Case 7 ''Meal Allowance
     '*****
         cbodept.Enabled = False
        deptall.Enabled = False
        catall.Enabled = False
        catlist.Enabled = False
     '*****
        'Call ShowGroup
    Case Else
     '*****
        cbodept.Enabled = True
        deptall.Enabled = True
        catall.Enabled = True
        catlist.Enabled = True
     '*****
        'Call ShowGroup
End Select

     Select Case typOptIdx.bytPer
            Case 0, 1, 2, 3, 4, 5, 6, 7, 16
            optGrp(1).Enabled = False
            Case Else
            optGrp(1).Enabled = True
            End Select
    Call ReportSetting
End Sub

Private Sub UnchkGrp()
    For i = 1 To 7
        optGrp(i).Value = 0
    Next
    Erase strAGrp
    Erase strAlbl
End Sub

''**********************************
Private Sub GroupRange()
Dim i As Integer
Dim strGrpChk As String, strlblChk As String

For i = 1 To 7
 If lblGrp(i).Caption <> "" Then
     If optGrp(i).Caption = "Employee" Then
             strGrpChk = "Empcode"
             strlblChk = "'Employee :  ' +    {cmd.Empcode}  "
             If bytRepMode = 6 Then     ' 15-01
                strGrpChk = "empcode"
                strlblChk = "'Employee :  ' +    {cmd.empcode}  "
             End If
     ElseIf optGrp(i).Caption = "Category" Then
             strGrpChk = "cat"
             strlblChk = "'Category : '  +  {cmd.catdescdesc}"
      ElseIf optGrp(i).Caption = "Department" Then
             strGrpChk = "dept"
             strlblChk = "'Department : ' + {cmd.deptdescdesc}"
      ElseIf optGrp(i).Caption = "Location" Then
             strGrpChk = "Location"
             strlblChk = "'Location  : ' +  {cmd.LocDesc}"
      ElseIf optGrp(i).Caption = "Division" Then
             strGrpChk = "Div"
             strlblChk = "'Division  :  '+  {cmd.DivDesc}"
      ElseIf optGrp(i).Caption = "Group" Then
             strGrpChk = "Group"
             strlblChk = "'Group  : ' +  {cmd.grupdesc}"
      ElseIf optGrp(i).Caption = "Company" Then
             strGrpChk = "Company"
             strlblChk = "'Company  : ' +  {cmd.CName}"
     End If
     
    strAGrp(Val(lblGrp(i).Caption)) = strGrpChk
    strAlbl(Val(lblGrp(i).Caption)) = strlblChk
    strAhead(Val(lblGrp(i).Caption)) = optGrp(i).Caption & ":"
  End If
 Next i
End Sub

'******
Private Sub OrderRange()
Dim i As Integer
Dim strord As String
strord = ""
strOrderBy = ""
For i = 1 To 7
 If lblGrp(i).Caption <> "" Then
     If optGrp(i).Caption = "Employee" Then
             strord = "Empcode" & ","
      ElseIf optGrp(i).Caption = "Category" Then
             strord = "cat" & ","
      ElseIf optGrp(i).Caption = "Department" Then
             strord = "dept" & ","
      ElseIf optGrp(i).Caption = "Location" Then
             strord = "Location" & ","
      ElseIf optGrp(i).Caption = "Division" Then
             strord = "Div" & ","
      ElseIf optGrp(i).Caption = "Group" Then
             strord = "Group" & ","
      ElseIf optGrp(i).Caption = "Company" Then
             strord = "Company" & ","
    End If
 strOrderBy = strOrderBy + strord
 End If
 Next i
 If strOrderBy <> "" Then
     strOrderBy = Left(strOrderBy, Len(strOrderBy) - 1)
     'strOrderBy = Right(strOrderBy, Len(strOrderBy) - 1)
 End If
End Sub

Public Sub ReportsMod()

On Error GoTo RepErr
Call RetValues 'Retrieves values selected or entered by user in selection frame
''**********************************
     Call GroupRange
     ''**********************************
Select Case bytRepMode
    Case 1                                      'DAILY TAB
        If Not dlyValid Then Call SetMSF1Cap(1): Exit Sub
        Call SetRepVars(1)
        If Not dlyCreateFiles Then Call SetMSF1Cap(1): Exit Sub 'CREATES TEMPORARY FILE
        If Not dlyReportsMod Then Call SetMSF1Cap(1): Exit Sub  'DUMPS VALUES IN TEMP FILE
        If Not dlySetEmpstr3 Then Call SetMSF1Cap(1): Exit Sub  'SETS QUERY FOR DSR
        'If Not dlyTotalCalc Then Call SetMSF1Cap(1): Exit Sub   'CALCULATES VALUES REQ. BY DSR
        If Not SetRepName Then Call SetMSF1Cap(1): Exit Sub     'SETS NAME OF THE REPORT,COMMAND ETC.
        ' Call ReportSetting
        If Not RecordsFound Then                'CHECKS FOR AVAILABILITY OF REQ.ED RECORDS
            'Call ChkRepFile                     'DELETES TEMPORARY FILE
            Call SetMSF1Cap(0)
            Exit Sub
        End If
        Call SetMSF1Cap(1)
        'If Not ChkPrinter(repname, bytPoLa) Then Exit Sub 'CHECKS FOR REQUIRED PAPERSIZE
    Case 2                                      'WEEKLY TAB
        If Not WkValid Then Call SetMSF1Cap(2): Exit Sub
        Call SetRepVars(2)
        If Not WkCreateFiles Then Call SetMSF1Cap(2): Exit Sub
        If Not wkReportsMod Then Call SetMSF1Cap(2): Exit Sub
        If Not WkSetEmpstr3 Then Call SetMSF1Cap(2): Exit Sub
        If Not SetRepName Then Call SetMSF1Cap(2): Exit Sub
        If Not RecordsFound Then
            Call ChkRepFile                     'DELETES TEMPORARY FILE
            Call SetMSF1Cap(2)
            Exit Sub
        End If
        Call SetMSF1Cap(2)
       ' If Not ChkPrinter(repname, bytPoLa) Then Call SetMSF1Cap(2): Exit Sub
    Case 3                                      'MONTHLY TAB
        ''
        
        ''*********************************
        Call GroupRange
        '**********************************
        
        If Not monValid Then Call SetMSF1Cap(3): Exit Sub
        Call SetRepVars(3)
        If Not monCreateFiles Then Call SetMSF1Cap(3): Exit Sub
        If Not monReportsMod Then Call SetMSF1Cap(3): Exit Sub
        If Not monSetEmpstr3 Then Call SetMSF1Cap(3): Exit Sub
        'Debug.Print empstr3
        
        If Not SetRepName Then Call SetMSF1Cap(3): Exit Sub
        If Not RecordsFound Then
        Call ChkRepFile                     'DELETES TEMPORARY FILE
        Call SetMSF1Cap(3)
            Exit Sub
        End If
        
        Call SetMSF1Cap(3)
       ' If Not ChkPrinter(repname, bytPoLa) Then Call SetMSF1Cap(3): Exit Sub
    Case 4
        Call GroupRange

          ''
        If Not yrValid Then Call SetMSF1Cap(4): Exit Sub
        Call SetRepVars(4)
        If Not yrCreateFiles Then Call SetMSF1Cap(4): Exit Sub
        If Not yrReportsMod Then Call SetMSF1Cap(4): Exit Sub
        If Not yrSetEmpstr3 Then Call SetMSF1Cap(4): Exit Sub
        If Not SetRepName Then Call SetMSF1Cap(4): Exit Sub
        If Not RecordsFound Then
             Call ChkRepFile                    'DELETES TEMPORARY FILE
             Call SetMSF1Cap(4)
             Exit Sub
        End If
        Call SetMSF1Cap(4)
       ' If Not ChkPrinter(repname, bytPoLa) Then Call SetMSF1Cap(4): Exit Sub
    Case 5                                      'MASTER TAB
        If Not maValid Then Call SetMSF1Cap(5): Exit Sub
        Call SetRepVars(5)
        If Not maReportsMod Then Call SetMSF1Cap(5): Exit Sub
        If Not maSetEmpstr3 Then Call SetMSF1Cap(5): Exit Sub
        If Not SetRepName Then Call SetMSF1Cap(5): Exit Sub
        If Not RecordsFound Then
            Call SetMSF1Cap(5)
            Exit Sub
        End If
        Call SetMSF1Cap(5)
        'If Not ChkPrinter(repname, bytPoLa) Then Call SetMSF1Cap(5): Exit Sub
    Case 6                      'PERIODIC TAB
            '
            '
        Call GroupRange
        If Not PeValid Then Call SetMSF1Cap(6): Exit Sub
        Call SetRepVars(6)
        If Not peCreateFiles Then Call SetMSF1Cap(6): Exit Sub
        If Not peReportsMod Then Call SetMSF1Cap(6): Exit Sub
        If Not peSetEmpstr3 Then Call SetMSF1Cap(6): Exit Sub
        If Not SetRepName Then Call SetMSF1Cap(6): Exit Sub
        If Not RecordsFound Then
            Call ChkRepFile                     'DELETES TEMPORARY FILE
            Call SetMSF1Cap(6)
            Exit Sub
        End If
        Call SetMSF1Cap(6)
        'If Not ChkPrinter(repname, bytPoLa) Then Call SetMSF1Cap(6): Exit Sub
End Select
frmExportReport.ReportModComplete = True
Exit Sub
Resume Next
RepErr:
    Select Case Err.Number
    
        Case -2147217865
            Call SetMSF1Cap(10)
            MsgBox NewCaptionTxt("40072", adrsC) & "ERL:" & Erl, vbInformation
            Call SetMSF1Cap(bytRepMode)
        Case 8542
            Call SetMSF1Cap(10)
            MsgBox NewCaptionTxt("40073", adrsC) & vbCrLf & NewCaptionTxt("40074", adrsC) & _
                vbCrLf & vbTab & NewCaptionTxt("40075", adrsC), vbInformation
            Call SetMSF1Cap(bytRepMode)
        Case Is <> 0
            'Call SetMSF1Cap(0)
            MsgBox Err.Description, vbCritical
            Call SetMSF1Cap(bytRepMode)
           ' Resume Next
    End Select
        If strRepFile <> "" Then Call ChkRepFile
        Call SetVarEmpty
End Sub

Private Sub RetValues()             '' RETRIEVES SELECTIONS MADE BY USER
On Error GoTo ERR_P
Dim strTmp As String
Dim catArray As String
Dim cnt As Integer

rpTables = "empmst,catdesc,deptdesc,groupmst,company,Location,Division"
                                                                                                                 

If cmbFrComSel.ListIndex = cmbFrComSel.ListCount - 1 Then
    strCName = InVar.strCOM
    strTmp = ""
Else
    strTmp = " and empmst.company =" & Trim(cmbFrComSel.Text)
    If cmbFrComSel.ListIndex >= 0 Then strCName = cmbFrComSel.List(cmbFrComSel.ListIndex, 1)
End If

strSql = ""
sqlStr = ""
headGrp = ""

Call OrderRange
'**** For Dept and Cat list box in Report Form*****
If chkEmpSel.Value = False Then
    If strCurrentUserType = HOD Then
        For cnt = 0 To cmbFrEmpSel.ListCount - 1
            If cnt = 0 Then clsInOperetor.strEmpSel = "("
            clsInOperetor.flgEmpSel = True
            clsInOperetor.strEmpSel = clsInOperetor.strEmpSel & "'" & Trim(Left(cmbFrEmpSel.List(cnt), pVStar.CodeSize)) & "',"
            If cmbFrEmpSel.List(cnt) = cmbToEmpSel.Text Then Exit For
         Next
         clsInOperetor.strEmpSel = Mid(clsInOperetor.strEmpSel, 1, Len(clsInOperetor.strEmpSel) - 1) & ")"
        strSql = "and empmst.Empcode IN " & clsInOperetor.strEmpSel
    Else
      strSql = " and empmst.Empcode between '" & Trim(cmbFrEmpSel.Text) & "' and '" & _
               Trim(cmbToEmpSel.Text) & "'"
    End If
    
    
Else
    clsInOperetor.strEmpSel = ""
    clsInOperetor.flgEmpSel = False
    For cnt = 0 To cboEmp.ListCount - 1
        If cnt = 0 Then clsInOperetor.strEmpSel = "("
        If cboEmp.Selected(cnt) = True Then
            clsInOperetor.flgEmpSel = True
            cboEmp.ListIndex = cnt
            
            clsInOperetor.strEmpSel = clsInOperetor.strEmpSel & "'" & Trim(Left(cboEmp.List(cnt), pVStar.CodeSize)) & "',"
            ''
        End If
        If cnt = cboEmp.ListCount - 1 Then clsInOperetor.strEmpSel = Mid(clsInOperetor.strEmpSel, 1, Len(clsInOperetor.strEmpSel) - 1) & ")"
    Next
    
    strSql = "and empmst.Empcode IN " & clsInOperetor.strEmpSel
End If


 If chkDeptSel.Value = False Then
    
    'this flag add by  for department datatype MIS2007DF026
    If blnFlagForDept = True Then
        'if true means datatype should be text
        strSql = strSql + " and  deptdesc.dept between '" & Trim(cmbFrDepSel.Text) & _
         "' and '" & Trim(cmbToDepSel.Text) & "'"
    Else
        'if true means datatype should be numeric
        strSql = strSql + " and  deptdesc.dept between " & Trim(cmbFrDepSel.Text) & _
         " and " & Trim(cmbToDepSel.Text) & ""
    End If
        ''
 Else
    clsInOperetor.strdeptSel = ""
    clsInOperetor.flgDeptSel = False
    For cnt = 0 To cbodept.ListCount - 1
        If cnt = 0 Then clsInOperetor.strdeptSel = "("
        If cbodept.Selected(cnt) = True Then
            clsInOperetor.flgDeptSel = True
            cbodept.ListIndex = cnt
            
            clsInOperetor.strdeptSel = clsInOperetor.strdeptSel & " " & Trim(Left(cbodept.List(cnt), InStr(1, cbodept.List(cnt), " ") - 1)) & ","
            ''
        End If
        If cnt = cbodept.ListCount - 1 Then clsInOperetor.strdeptSel = Mid(clsInOperetor.strdeptSel, 1, Len(clsInOperetor.strdeptSel) - 1) & ")"
    Next
    strSql = strSql + " and deptdesc.dept IN " & clsInOperetor.strdeptSel
End If
'**FOR CATEGORY****

If chkCatsel.Value = False Then
    strSql = strSql + " and catdesc.cat  between '" & _
         Trim(cmbFrCatSel.Text) & "' and '" & Trim(cmbToCatSel.Text) & "'"
Else
 clsInOperetor.strcatSel = ""
   clsInOperetor.flgcatSel = False
        For cnt = 0 To catlist.ListCount - 1
           If cnt = 0 Then clsInOperetor.strcatSel = "("
            If catlist.Selected(cnt) = True Then
                clsInOperetor.flgcatSel = True
                catlist.ListIndex = cnt
                clsInOperetor.strcatSel = clsInOperetor.strcatSel & "'" & Trim(Left(catlist.List(cnt), InStr(1, catlist.List(cnt), " ") - 1)) & "',"
            End If
            If cnt = catlist.ListCount - 1 Then clsInOperetor.strcatSel = Mid(clsInOperetor.strcatSel, 1, Len(clsInOperetor.strcatSel) - 1) & ")"
        Next
        strSql = strSql + " and catdesc.cat  IN " & clsInOperetor.strcatSel
End If
        

'***Group List ***
If chkGrpSel.Value = False Then
    strSql = strSql + " and groupmst." & strKGroup & " between " & Trim(cmbFrGrpSel.Text) & " and  " & _
             Trim(cmbToGrpSel.Text) & " " & strTmp
Else
    clsInOperetor.strGrpSel = ""
    clsInOperetor.flgGrpSel = False
    For cnt = 0 To cboGroup.ListCount - 1
        If cnt = 0 Then clsInOperetor.strGrpSel = "("
        If cboGroup.Selected(cnt) = True Then
            clsInOperetor.flgGrpSel = True
            cboGroup.ListIndex = cnt
            
            clsInOperetor.strGrpSel = clsInOperetor.strGrpSel & " " & Trim(Left(cboGroup.List(cnt), InStr(1, cboGroup.List(cnt), " ") - 1)) & ","
            ''
        End If
        If cnt = cboGroup.ListCount - 1 Then clsInOperetor.strGrpSel = Mid(clsInOperetor.strGrpSel, 1, Len(clsInOperetor.strGrpSel) - 1) & ")"
    Next
    strSql = strSql + " and groupmst." & strKGroup & " IN " & clsInOperetor.strGrpSel
End If
 
'***Location List ***
If chkLocSel.Value = False Then
strSql = strSql + " and Location.Location between " & _
         Trim(cmbFrLocSel.Text) & " and " & Trim(cmbToLocSel.Text)
Else
 clsInOperetor.strLocSel = ""
 clsInOperetor.flgLocSel = False
        For cnt = 0 To cboLoc.ListCount - 1
           If cnt = 0 Then clsInOperetor.strLocSel = "("
            If cboLoc.Selected(cnt) = True Then
                clsInOperetor.flgLocSel = True
                cboLoc.ListIndex = cnt
                
                clsInOperetor.strLocSel = clsInOperetor.strLocSel & Trim(Left(cboLoc.List(cnt), InStr(1, cboLoc.List(cnt), " ") - 1)) & ","
                ''
            End If
            If cnt = cboLoc.ListCount - 1 Then clsInOperetor.strLocSel = Mid(clsInOperetor.strLocSel, 1, Len(clsInOperetor.strLocSel) - 1) & ")"
        Next
        strSql = strSql + " and Location.Location IN " & clsInOperetor.strLocSel
End If

'***Division List****
If chkDivSel.Value = False Then
    strSql = strSql + " and Division.Div between " & Trim(cmbFrDivSel.Text) & " and " & _
             Trim(cmbToDivSel.Text)
Else

clsInOperetor.strDivSel = ""
clsInOperetor.flgDivSel = False
        For cnt = 0 To cboDiv.ListCount - 1
           If cnt = 0 Then clsInOperetor.strDivSel = "("
            If cboDiv.Selected(cnt) = True Then
                clsInOperetor.flgDivSel = True
                cboDiv.ListIndex = cnt
                
                clsInOperetor.strDivSel = clsInOperetor.strDivSel & Trim(Left(cboDiv.List(cnt), InStr(1, cboDiv.List(cnt), " ") - 1)) & ","
                ''
            End If
            If cnt = cboDiv.ListCount - 1 Then clsInOperetor.strDivSel = Mid(clsInOperetor.strDivSel, 1, Len(clsInOperetor.strDivSel) - 1) & ")"
        Next
        strSql = strSql + " and Division.Div IN " & clsInOperetor.strDivSel
End If



        strSql = strSql + " and empmst.dept = deptdesc.dept and " & _
         " empmst.cat = catdesc.cat and empmst." & strKGroup & " = groupmst." & strKGroup & " and " & _
         "empmst.company = company.company and empmst.Location = Location.Location" & _
         " and empmst.Div = Division.Div" + strTmp


''For Mauritius 20-08-2003
If UCase(Trim(cboMain.Text)) <> "ALL" Then
    strSql = strSql & " And Empmst.qualf = '" & cboMain.Text & "'"
End If
''
Exit Sub
ERR_P:
    ShowError ("RetVALUES :: Reports Form")
    Resume Next
End Sub

Private Sub FrSelFill()
On Error GoTo ERR_P
'filling all the combos in selection frame
'this if condition add by  for TGL
If GetFlagStatus("leftemployee") Then
    Call ComboFill(cmbFrEmpSel, 17, 2)      'emp
Else
    Call ComboFill(cmbFrEmpSel, 16, 2)      'emp
End If

Call ComboFill(cboFrDesigSel, 20, 2)        'Designation
cboToDesigSel.List = cboFrDesigSel.List

cmbToEmpSel.List = cmbFrEmpSel.List
Call ComboFill(cmbFrDepSel, 22, 2)      'dept
cmbToDepSel.List = cmbFrDepSel.List
Call ComboFill(cmbFrCatSel, 23, 2)      'cat
cmbToCatSel.List = cmbFrCatSel.List
Call ComboFill(cmbFrGrpSel, 24, 2)      'group
cmbToGrpSel.List = cmbFrGrpSel.List
Call ComboFill(cmbFrLocSel, 25, 2)     'Location
cmbToLocSel.List = cmbFrLocSel.List
Call ComboFill(cmbFrDivSel, 26, 2)     'Division
cmbToDivSel.List = cmbFrDivSel.List

Call ComboFill(cmbFrComSel, 21, 2)      'company
Call FillShiftCombo                    'shift
cmbFrComSel.AddItem "All"
''For Mauritius 20-08-2003
Call FillMainCombo                      ''Maintain-nonMaintain

''fill list cate
 '*****
 If adrsASC.State = 1 Then adrsASC.Close
    adrsASC.Open "select Empcode,name from empmst order by Empcode", ConMain, adOpenKeyset, adLockReadOnly
    Dim cnt As Integer
    If adrsASC.RecordCount >= 1 Then
        For cnt = 0 To adrsASC.RecordCount - 1
            cboEmp.AddItem adrsASC(0) & "   " & adrsASC(1)
            adrsASC.MoveNext
        Next
 End If
     
If adrsASC.State = 1 Then adrsASC.Close
    adrsASC.Open "select * from catdesc where cat <> '100' order by cat", ConMain, adOpenKeyset, adLockReadOnly
    'Dim cnt As Integer
    If adrsASC.RecordCount >= 1 Then
        For cnt = 0 To adrsASC.RecordCount - 1
            catlist.AddItem adrsASC(0) & "   " & adrsASC(1)
            adrsASC.MoveNext
        Next
     End If

If adrsC.State = 1 Then adrsASC.Close
     adrsASC.Open "select* from deptdesc order by dept", ConMain, adOpenKeyset, adLockReadOnly
       If adrsASC.RecordCount >= 1 Then
        For cnt = 0 To adrsASC.RecordCount - 1
            cbodept.AddItem adrsASC(0) & "   " & adrsASC(1)
            adrsASC.MoveNext
        Next
     End If
     
If adrsASC.State = 1 Then adrsASC.Close
     adrsASC.Open "select* from groupmst order by " & strKGroup & "", ConMain, adOpenKeyset, adLockReadOnly
       If adrsASC.RecordCount >= 1 Then
        For cnt = 0 To adrsASC.RecordCount - 1
            cboGroup.AddItem adrsASC(0) & "   " & adrsASC(1)
            adrsASC.MoveNext
        Next
End If

If adrsASC.State = 1 Then adrsASC.Close
     adrsASC.Open "select* from Location order by Location", ConMain, adOpenKeyset, adLockReadOnly
       If adrsASC.RecordCount >= 1 Then
        For cnt = 0 To adrsASC.RecordCount - 1
            cboLoc.AddItem adrsASC(0) & "   " & adrsASC(1)
            adrsASC.MoveNext
        Next
End If

If adrsASC.State = 1 Then adrsASC.Close
     adrsASC.Open "select* from division order by div", ConMain, adOpenKeyset, adLockReadOnly
       If adrsASC.RecordCount >= 1 Then
        For cnt = 0 To adrsASC.RecordCount - 1
            cboDiv.AddItem adrsASC(0) & "   " & adrsASC(1)
            adrsASC.MoveNext
        Next
End If

If cboEmp.ListCount > 0 Then cboEmp.Selected(0) = True
If catlist.ListCount > 0 Then catlist.Selected(0) = True
If cbodept.ListCount > 0 Then cbodept.Selected(0) = True
If cboGroup.ListCount > 0 Then cboGroup.Selected(0) = True
If cboLoc.ListCount > 0 Then cboLoc.Selected(0) = True
If cboDiv.ListCount > 0 Then cboDiv.Selected(0) = True

 '*****
'if employeemaster empty then skip
If cmbFrEmpSel.ListCount <> 0 Then cmbFrEmpSel.Text = cmbFrEmpSel.List(0)
If cmbFrDepSel.ListCount <> 0 Then cmbFrDepSel.Text = cmbFrDepSel.List(0)
If cmbFrCatSel.ListCount <> 0 Then cmbFrCatSel.Text = cmbFrCatSel.List(0)

If cmbFrGrpSel.ListCount <> 0 Then cmbFrGrpSel.Text = cmbFrGrpSel.List(0)
If cmbFrLocSel.ListCount <> 0 Then cmbFrLocSel.Text = cmbFrLocSel.List(0)
If cmbFrDivSel.ListCount <> 0 Then cmbFrDivSel.Text = cmbFrDivSel.List(0)
If cmbFrComSel.ListCount <> 0 Then cmbFrComSel.Text = cmbFrComSel.List(cmbFrComSel.ListCount - 1)
If cmbToEmpSel.ListCount <> 0 Then cmbToEmpSel.Text = cmbToEmpSel.List(cmbToEmpSel.ListCount - 1)
If cmbToDepSel.ListCount <> 0 Then cmbToDepSel.Text = cmbToDepSel.List(cmbToDepSel.ListCount - 1)
If cmbToCatSel.ListCount <> 0 Then cmbToCatSel.Text = cmbToCatSel.List(cmbToCatSel.ListCount - 1)
If cmbToGrpSel.ListCount <> 0 Then cmbToGrpSel.Text = cmbToGrpSel.List(cmbToGrpSel.ListCount - 1)
If cmbToLocSel.ListCount <> 0 Then cmbToLocSel.Text = cmbToLocSel.List(cmbToLocSel.ListCount - 1)
If cmbToDivSel.ListCount <> 0 Then cmbToDivSel.Text = cmbToDivSel.List(cmbToDivSel.ListCount - 1)
If cboFrDesigSel.ListCount <> 0 Then cboFrDesigSel.Text = cboFrDesigSel.List(0)
If cboToDesigSel.ListCount <> 0 Then cboToDesigSel.Text = cboToDesigSel.List(cboToDesigSel.ListCount - 1)

'optGrpEmp(0).Value = True
Exit Sub
ERR_P:
    ShowError ("Combos Selects :: Reports Form")
    Resume Next
End Sub

Public Sub SetRepVars(Optional bytSetRep As Byte = 7)
On Error GoTo ERR_P
typRep.strDlyDate = ""                  '' RETREIVING VALUES FROM
typRep.strLeftFr = ""                   '' TABS TO RELATED TYPE VARIABLES
typRep.strLeftTo = ""
typRep.strMonMth = ""
typRep.strMonYear = ""
typRep.strPeriFr = ""
typRep.strPeriTo = ""
typRep.strWkDate = ""
typRep.strYear = ""
Select Case bytSetRep
    Case 1      '' Daily
        typRep.strDlyDate = txtDaily.Text
    Case 2      '' Weekly
        typRep.strWkDate = txtWeek.Text
    Case 3      '' Monthly
        typRep.strMonMth = cmbMonth.Text
        typRep.strMonYear = cmbMonYear.Text
    Case 4      '' Yearly
        typRep.strYear = cmbYear.Text
    Case 5      '' Masters
            typRep.strLeftFr = txtMastFr.Text
            typRep.strLeftTo = txtMastTo.Text
    Case 6      '' Periodic
        typRep.strPeriFr = txtFrPeri.Text
        typRep.strPeriTo = txtToPeri.Text
    Case 7
        txtDaily.Text = DateDisp(Date)
        txtWeek.Text = DateDisp(Date)
        txtMastFr.Text = DateDisp(Date)
        txtMastTo.Text = DateDisp(Date)
        txtFrPeri.Text = DateDisp(Date)
        txtToPeri.Text = DateDisp(Date)
End Select
Exit Sub
ERR_P:
    ShowError ("SetRepVars :: Reports Form")
End Sub

Private Sub txtDaily_Click()
varCalDt = ""
varCalDt = Trim(txtDaily.Text)
txtDaily.Text = ""
Call ShowCalendar
End Sub

Private Sub txtDaily_GotFocus()
    Call GF(txtDaily)
End Sub

Private Sub txtDaily_KeyPress(KeyAscii As Integer)
    Call CDK(txtDaily, KeyAscii)
End Sub

Private Sub txtDaily_Validate(Cancel As Boolean)
If Not ValidDate(txtDaily) Then
    txtDaily.SetFocus
    Cancel = True
End If
End Sub


Private Sub txtWeek_Click()
varCalDt = ""
varCalDt = Trim(txtWeek.Text)
txtWeek.Text = ""
Call ShowCalendar
End Sub

Private Sub txtWeek_GotFocus()
    Call GF(txtWeek)
End Sub

Private Sub txtWeek_KeyPress(KeyAscii As Integer)
    Call CDK(txtWeek, KeyAscii)
End Sub

Private Sub txtWeek_Validate(Cancel As Boolean)
If Not ValidDate(txtWeek) Then
    txtWeek.SetFocus
    Cancel = True
End If
End Sub

Private Sub txtMastFr_Click()
varCalDt = ""
varCalDt = Trim(txtMastFr.Text)
txtMastFr.Text = ""
Call ShowCalendar
End Sub

Private Sub txtMastFr_GotFocus()
    Call GF(txtMastFr)
End Sub

Private Sub txtMastFr_KeyPress(KeyAscii As Integer)
    Call CDK(txtMastFr, KeyAscii)
End Sub

Private Sub txtMastFr_Validate(Cancel As Boolean)
If Not ValidDate(txtMastFr) Then
    txtMastFr.SetFocus
    Cancel = True
End If
End Sub

Private Sub txtMastTo_Click()
varCalDt = ""
varCalDt = Trim(txtMastTo.Text)
txtMastTo.Text = ""
Call ShowCalendar
End Sub

Private Sub txtMastTo_GotFocus()
    Call GF(txtMastTo)
End Sub

Private Sub txtMastTo_KeyPress(KeyAscii As Integer)
    Call CDK(txtMastTo, KeyAscii)
End Sub

Private Sub txtMastTo_Validate(Cancel As Boolean)
If Not ValidDate(txtMastTo) Then
    txtMastTo.SetFocus
    Cancel = True
End If
End Sub

Private Sub txtFrPeri_Click()
varCalDt = ""
varCalDt = Trim(txtFrPeri.Text)
txtFrPeri.Text = ""
Call ShowCalendar
End Sub

Private Sub txtFrPeri_GotFocus()
    Call GF(txtFrPeri)
End Sub

Private Sub txtFrPeri_KeyPress(KeyAscii As Integer)
    Call CDK(txtFrPeri, KeyAscii)
End Sub

Private Sub txtFrPeri_Validate(Cancel As Boolean)
If Not ValidDate(txtFrPeri) Then
    txtFrPeri.SetFocus
    Cancel = True
End If
End Sub

Private Sub txtToPeri_Click()
varCalDt = ""
varCalDt = Trim(txtToPeri.Text)
txtToPeri.Text = ""
Call ShowCalendar
End Sub

Private Sub txtToPeri_GotFocus()
    Call GF(txtToPeri)
End Sub

Private Sub txtToPeri_KeyPress(KeyAscii As Integer)
    Call CDK(txtToPeri, KeyAscii)
End Sub

Private Sub txtToPeri_Validate(Cancel As Boolean)
If Not ValidDate(txtToPeri) Then
    txtToPeri.SetFocus
    Cancel = True
End If
End Sub

Private Function RecordsFound() As Boolean  '' CHECKS IF THE REQUIRED RECORDS ARE
On Error GoTo ERR_P                         '' AVAILABLE OR NOT
RecordsFound = True
Call SetMSF1Cap(9)
If bytPrint = 1 Then
    If blnIntz = True Then
        CRV.ViewReport
        Load frmCRV
        DoEvents: DoEvents: DoEvents:
        frmCRV.CRV.PrintReport
    Else
        Call SetMSF1Cap(10)
        RecordsFound = False
        Exit Function
    End If
ElseIf bytPrint = 2 Then
    If blnIntz = True Then
        CRV.ViewReport
        Do While CRV.IsBusy              'ZOOM METHOD DOES NOT WORK WHILE
            DoEvents                          'REPORT IS LOADING, SO WE MUST PAUSE
        Loop
        frmCRV.Show vbModal
    Else
        Call SetMSF1Cap(10)
        RecordsFound = False
        Exit Function
    End If
ElseIf bytPrint = 3 Then
    If blnIntz = True Then
        CRV.ViewReport
        Report.Export
        DoEvents: DoEvents: DoEvents
    End If
ElseIf bytPrint = 4 Then
    If blnIntz = True Then
        CRV.ViewReport
      
        DoEvents: DoEvents: DoEvents
     End If
End If
Exit Function
ERR_P:
    ShowError ("Records Found :: " & Me.Caption)
    Set Report = Nothing
    'With Report.ExportOptions
    RecordsFound = False
End Function

Private Function dlyReportsMod() As Boolean
On Error GoTo ERR_P
dlyReportsMod = False                   '' FUNCTION FOR DAILY REPORTS
'Call SetRepVars(1)
'' Adjust Shift Inclusion Statements based on the Report & Shift Selected
If cboShift.ListIndex <> -1 Then
    If cboShift.ListIndex <> cboShift.ListCount - 1 Then
        Select Case typOptIdx.bytDly
            Case 10, 12 ''manpower,summary
                strSql = strSql & " and " & strMon_Trn & ".Shift='" & _
                cboShift.Text & "'"
      
            Case Else
                If typOptIdx.bytDly <> 9 Then
                    strSql = strSql & " and " & strMon_Trn & ".Shift='" & _
                    cboShift.Text & "'"
                    ''supriya dated 25/05/05 for removin err in ODReport
                    'If typOptIdx.bytDly <> 10 Then _
                        'rpTables = rpTables & "," & strMon_Trn
                Else
                    strSql = strSql & " and " & strMon_Trn & ".D" & _
                        Day(typRep.strDlyDate) & "='" & cboShift.Text & "'"
                End If
                
        End Select
    End If
End If


Call SetMSF1Cap(8)
Select Case typOptIdx.bytDly
    Case 2 'Continuous Absent
        If Not dlyContAbs(CStr(DateCompDate(typRep.strDlyDate) - (bytNoDay - 1)), _
        typRep.strDlyDate) Then Exit Function
        dlyReportsMod = True
    Case 0, 1, 3, 4, 5, 6, 7, 9, 11, 13
        dlyReportsMod = True
    Case 8
        typDT.dtFrom = typRep.strDlyDate
        typDT.dtTo = typRep.strDlyDate
        If Not AppendDataFile(frmReports) Then Exit Function
        If Not DlyEntries Then Exit Function
         dlyReportsMod = True
    Case 10 'Manpower
        If Not DlyManpower Then Exit Function
        dlyReportsMod = True
    Case 12 ''Summary
         If Not Fuc_NewSummary Then Exit Function
         dlyReportsMod = True
   Case Else
        dlyReportsMod = False
End Select
Exit Function
ERR_P:
    ShowError ("DailyReportsMod :: " & Me.Caption)
   Resume Next
End Function

Private Function dlyValid() As Boolean
On Error GoTo ERR_P
dlyValid = True                                 '' FUNCTION FOR DAILY REPORT VALIDATIONS
Call SetMSF1Cap(7)
If txtDaily.Text = "" Then
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00072", adrsMod)
    txtDaily.SetFocus
    dlyValid = False
    Exit Function
End If
If typOptIdx.bytDly = 2 Then 'CHECK FOR CONTINUES ABSENT REPORT
    If bytNoDay > 1 Then
        If Not FindTable(MakeName(MonthName(Month(DateCompDate(txtDaily.Text) - (bytNoDay - 1))), _
             Year(DateCompDate(txtDaily.Text) - (bytNoDay - 1)), "trn")) Then
            Call SetMSF1Cap(10)
            MsgBox NewCaptionTxt("40065", adrsC) & _
                MonthName(Month(DateCompDate(txtDaily.Text) - (bytNoDay - 1))), vbExclamation
            txtDaily.SetFocus
            optDly(2).Value = False
            typOptIdx.bytDly = 14
            dlyValid = False
            Exit Function
        End If
    End If
End If  ''End of Cont abs Reports check
If typOptIdx.bytDly = 12 Then
    Dim cntchk As Integer
    For i = 1 To 7
      If optGrp(i).Value = True Then cntchk = cntchk + 1
      
    Next i
 
    If cntchk = 0 Then
       MsgBox "Selection of atleast one  group is mandatory for this  Report!", vbInformation
       dlyValid = False
       Exit Function
    End If
 End If
'*****
Dim cnt As Integer
 clsInOperetor.strdeptSel = ""
 clsInOperetor.flgDeptSel = False
        For cnt = 0 To catlist.ListCount - 1
           If cnt = 0 Then clsInOperetor.strdeptSel = "("
            If catlist.Selected(cnt) = True Then
                clsInOperetor.flgDeptSel = True
                catlist.ListIndex = cnt
                clsInOperetor.strdeptSel = clsInOperetor.strdeptSel & "'" & catlist.List(cnt) & "',"
            End If
            If cnt = catlist.ListCount - 1 Then clsInOperetor.strdeptSel = Mid(clsInOperetor.strdeptSel, 1, Len(clsInOperetor.strdeptSel) - 1) & ")"
        Next
        If Not clsInOperetor.flgDeptSel Then
             MsgBox "Category Not Selected ", vbCritical, "Wrong Value"
            dlyValid = False
            catlist.SetFocus
            Exit Function
        Else
        catsel = clsInOperetor.strdeptSel
        End If
    
    
 clsInOperetor.strcatSel = ""
 clsInOperetor.flgcatSel = False
        For cnt = 0 To cbodept.ListCount - 1
           If cnt = 0 Then clsInOperetor.strcatSel = "("
            If cbodept.Selected(cnt) = True Then
                clsInOperetor.flgcatSel = True
                cbodept.ListIndex = cnt
                clsInOperetor.strcatSel = clsInOperetor.strcatSel & Left(cbodept.List(cnt), InStr(1, cbodept.List(cnt), " ") - 1) & ","
            End If
            If cnt = cbodept.ListCount - 1 Then clsInOperetor.strcatSel = Mid(clsInOperetor.strdeptSel, 1, Len(clsInOperetor.strdeptSel) - 1) & ")"
        Next
        If Not clsInOperetor.flgcatSel Then
             MsgBox "Dept Not Selected", vbCritical, "Wrong Value"
            dlyValid = False
            catlist.SetFocus
            Exit Function
             
        Else
            deptsel = clsInOperetor.strdeptSel
        End If
'*****
strMon_Trn = MakeName(MonthName(Month(DateCompDate(txtDaily.Text))), _
Year(DateCompDate(txtDaily.Text)), "trn")
If Not FindTable(strMon_Trn) Then
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("40065", adrsC) & _
        MonthName(Month(DateCompDate(txtDaily.Text))), vbExclamation
    txtDaily.SetFocus
    dlyValid = False
    Exit Function
End If



Exit Function
ERR_P:
    ShowError ("dlyValid :: Reportsfrm")
    dlyValid = False
    Resume Next
End Function

Private Function WkValid() As Boolean
On Error GoTo ERR_P
WkValid = True                          '' FUNCTION FOR WEEKLY REPORT VALIDATIONS
Call SetMSF1Cap(7)
If txtWeek.Text = "" Then
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00072", adrsMod)
    WkValid = False
    txtWeek.SetFocus
    Exit Function
End If
If typOptIdx.bytWek = 6 Then
    If Not FindTable(Left(MonthName(Month(CDate(txtWeek.Text))), 3) & _
    Right(Year(CDate(txtWeek.Text)), 2) & "shf") Then
        Call SetMSF1Cap(10)
        MsgBox NewCaptionTxt("40076", adrsC) & _
            MonthName(Month(CDate(txtWeek.Text))), vbExclamation
        WkValid = False
        txtWeek.SetFocus
        Exit Function
    End If
End If
If Not FindTable(Left(MonthName(Month(CDate(txtWeek.Text))), 3) & _
Right(Year(CDate(txtWeek.Text)), 2) & "trn") Then
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("40065", adrsC) & _
        MonthName(Month(CDate(txtWeek.Text))), vbExclamation
    WkValid = False
    txtWeek.SetFocus
    Exit Function
End If
If Not FindTable(Left(MonthName(Month(CDate(txtWeek.Text) + 6)), 3) & _
Right(Year(CDate(txtWeek.Text) + 6), 2) & "trn") Then
    Call SetMSF1Cap(10)
    MsgBox "Monthly File Not found for the Month of   " & _
        MonthName(Month(CDate(txtWeek.Text) + 6)), vbExclamation
    WkValid = False
    txtWeek.SetFocus
    Exit Function
End If
 '*****
Dim cnt As Integer
 clsInOperetor.strdeptSel = ""
 clsInOperetor.flgDeptSel = False
        For cnt = 0 To catlist.ListCount - 1
           If cnt = 0 Then clsInOperetor.strdeptSel = "("
            If catlist.Selected(cnt) = True Then
                clsInOperetor.flgDeptSel = True
                catlist.ListIndex = cnt
                clsInOperetor.strdeptSel = clsInOperetor.strdeptSel & "'" & catlist.List(cnt) & "',"
            End If
            If cnt = catlist.ListCount - 1 Then clsInOperetor.strdeptSel = Mid(clsInOperetor.strdeptSel, 1, Len(clsInOperetor.strdeptSel) - 1) & ")"
        Next
        If Not clsInOperetor.flgDeptSel Then
             MsgBox "Category Not Selected ", vbCritical, "Wrong Value"
            WkValid = False
            catlist.SetFocus
            Exit Function
        Else
        catsel = clsInOperetor.strdeptSel
        End If
    
    
 clsInOperetor.strcatSel = ""
 clsInOperetor.flgcatSel = False
        For cnt = 0 To cbodept.ListCount - 1
           If cnt = 0 Then clsInOperetor.strcatSel = "("
            If cbodept.Selected(cnt) = True Then
                clsInOperetor.flgcatSel = True
                cbodept.ListIndex = cnt
                clsInOperetor.strcatSel = clsInOperetor.strcatSel & Left(cbodept.List(cnt), 4) & ","
            End If
            'If cnt = cboDept.ListCount - 1 Then clsInOperetor.strcatSel = Mid(clsInOperetor.strdeptSel1, 1, Len(clsInOperetor.strdeptSel1) - 1) & ")"
        Next
        If Not clsInOperetor.flgcatSel Then
             MsgBox "Dept Not Selected", vbCritical, "Wrong Value"
            WkValid = False
            catlist.SetFocus
            Exit Function
             
        Else
            'deptsel = clsInOperetor.strdeptSel1
        End If
 '*****
Exit Function
ERR_P:
    ShowError ("wkValid :: Reportsfrm")
    WkValid = False
End Function

Private Function monValid() As Boolean
On Error GoTo ERR_P
monValid = True                             '' FUNCTION FOR MONTHLY REPORT VALIDATIONS
Call SetMSF1Cap(7)
''No check required for
''Leave consumption Report,Leave Balance ,OT paid hrs.
'21 index add by
Select Case typOptIdx.bytMon
    Case 6, 9, 14: Exit Function
End Select
''Check if Monthly Transaction file is available or not.
If cmbMonth.Text <> "" And cmbMonYear.Text <> "" Then
    If typOptIdx.bytMon = 17 Then
        If Not FindTable(Left(cmbMonth.Text, 3) & Right(cmbMonYear.Text, 2) & "shf") Then
            Call SetMSF1Cap(10)
            MsgBox NewCaptionTxt("40076", adrsC) & " " & _
                cmbMonth.Text & Space(1) & cmbMonYear.Text, vbExclamation
            cmbMonth.SetFocus
            monValid = False
            Exit Function
        End If
    Else
        If Not FindTable(Left(cmbMonth.Text, 3) & Right(cmbMonYear.Text, 2) & "trn") Then
            Call SetMSF1Cap(10)
            MsgBox "Transaction FIle Not found for the Month of    " & cmbMonth.Text & _
                Space(1) & cmbMonYear.Text, vbExclamation
            cmbMonth.SetFocus
            monValid = False
            Exit Function
        End If
    End If
ElseIf cmbMonth.Text = "" Then
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("40077", adrsC)
    cmbMonth.SetFocus
    monValid = False
    Exit Function
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("40078", adrsC)
    cmbMonYear.SetFocus
    monValid = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("monValid :: Reportsfrm")
   ' Resume Next
    monValid = False
End Function

Private Function yrValid() As Boolean
On Error GoTo ERR_P
yrValid = True                                  '' FUNCTION FOR YEARLY REPORT VALIDATIONS
Call SetMSF1Cap(7)
If cmbFrEmpSel.ListCount <= 0 Then
    MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation
    yrValid = False
    Exit Function
End If
If cmbYear.Text <> "" Then
    If Not FindTable("lvtrn" & Right(Trim(cmbYear.Text), 2)) Then
        Call SetMSF1Cap(10)
        MsgBox NewCaptionTxt("00054", adrsMod) & cmbYear.Text & NewCaptionTxt("00055", adrsMod)
        yrValid = False
        cmbYear.SetFocus
        Exit Function
    End If
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("40078", adrsC)
    yrValid = False
    cmbYear.SetFocus
    Exit Function
End If
'add by
If typOptIdx.bytYer = 5 Then
    If Not FindTable("lvbal" & Right(Trim(cmbYear.Text), 2)) Then
        Call SetMSF1Cap(10)
        MsgBox cmbYear.Text & " " & "year's leave balance table not found", vbInformation
        yrValid = False
        cmbYear.SetFocus
        Exit Function
    End If
End If
'add by
Exit Function
ERR_P:
    ShowError ("yrValid :: Reportsfrm")
    yrValid = False
End Function

Private Function maValid() As Boolean
On Error GoTo ERR_P
maValid = True                              '' FUNCTION FOR MASTER REPORT VALIDATIONS
Call SetMSF1Cap(7)
Select Case typOptIdx.bytMst
    Case 0, 1, 2
        If cmbFrEmpSel.ListCount <= 0 Then
            MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation
            maValid = False
            Exit Function
        End If
End Select
If typOptIdx.bytMst = 2 Then
    If txtMastFr.Text = "" Then
        Call SetMSF1Cap(10)
        MsgBox NewCaptionTxt("00016", adrsMod)
        maValid = False
        txtMastFr.SetFocus
        Exit Function
    End If
    If txtMastTo.Text = "" Then
        Call SetMSF1Cap(10)
        MsgBox NewCaptionTxt("00017", adrsMod)
        maValid = False
        txtMastTo.SetFocus
        Exit Function
    End If
    If CDate(txtMastFr.Text) > CDate(txtMastTo.Text) Then
        Call SetMSF1Cap(10)
        MsgBox NewCaptionTxt("00018", adrsMod), vbInformation
        txtMastTo.SetFocus
        maValid = False
        Exit Function
    End If
End If
'*****
''
Dim cnt As Integer
 clsInOperetor.strdeptSel = ""
 clsInOperetor.flgDeptSel = False
        For cnt = 0 To catlist.ListCount - 1
           If cnt = 0 Then clsInOperetor.strdeptSel = "("
            If catlist.Selected(cnt) = True Then
                clsInOperetor.flgDeptSel = True
                catlist.ListIndex = cnt
                clsInOperetor.strdeptSel = clsInOperetor.strdeptSel & "'" & catlist.List(cnt) & "',"
            End If
            If cnt = catlist.ListCount - 1 Then clsInOperetor.strdeptSel = Mid(clsInOperetor.strdeptSel, 1, Len(clsInOperetor.strdeptSel) - 1) & ")"
        Next
        If Not clsInOperetor.flgDeptSel Then
             MsgBox "Category Not Selected ", vbCritical, "Wrong Value"
            maValid = False
            catlist.SetFocus
            Exit Function
        Else
        catsel = clsInOperetor.strdeptSel
        End If

 clsInOperetor.strcatSel = ""
 clsInOperetor.flgcatSel = False
        For cnt = 0 To cbodept.ListCount - 1
           If cnt = 0 Then clsInOperetor.strcatSel = "("
            If cbodept.Selected(cnt) = True Then
                clsInOperetor.flgcatSel = True
                cbodept.ListIndex = cnt
                clsInOperetor.strcatSel = clsInOperetor.strcatSel & Left(cbodept.List(cnt), 4) & ","
            End If
            If cnt = cbodept.ListCount - 1 Then clsInOperetor.strcatSel = Mid(clsInOperetor.strcatSel, 1, Len(clsInOperetor.strcatSel) - 1) & ")"
        Next
        If Not clsInOperetor.flgcatSel Then
             MsgBox "Dept Not Selected", vbCritical, "Wrong Value"
            maValid = False
            catlist.SetFocus
            Exit Function
        Else
            deptsel = clsInOperetor.strcatSel
        End If
'*****
Exit Function
ERR_P:
    ShowError ("maValid :: Reportsfrm")
    maValid = False
    'Resume Next
End Function

Private Function PeValid() As Boolean
On Error GoTo ERR_P
PeValid = True                              '' FUNCTION FOR PERIODIC REPORT VALIDATIONS
Call SetMSF1Cap(7)
If txtFrPeri.Text = "" Then
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00016", adrsMod), vbInformation
    txtFrPeri.SetFocus
    PeValid = False
    Exit Function
End If
If txtToPeri.Text = "" Then
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00017", adrsMod), vbInformation
    txtToPeri.SetFocus
    PeValid = False
    Exit Function
End If
If CDate(txtFrPeri.Text) > CDate(txtToPeri.Text) Then
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00018", adrsMod), vbInformation
    txtFrPeri.SetFocus
    PeValid = False
    Exit Function
End If

If typOptIdx.bytPer = 16 Then ' 24-08
Else
    If (CDate(txtFrPeri.Text) + 31) < CDate(txtToPeri.Text) Then
        Call SetMSF1Cap(10)
        MsgBox NewCaptionTxt("40079", adrsC), vbInformation
        If typOptIdx.bytPer <> 17 Then txtFrPeri.SetFocus
        PeValid = False
        Exit Function
    End If
    If Month(CDate(txtToPeri.Text)) - Month(CDate(txtFrPeri.Text)) > 1 Then
        Call SetMSF1Cap(10)
        MsgBox NewCaptionTxt("40080", adrsC), vbInformation
        If typOptIdx.bytPer <> 17 Then txtFrPeri.SetFocus
        PeValid = False
        Exit Function
    End If
End If
''For Mauritius 05-08-2003
If typOptIdx.bytPer = 10 Then
    strMon_Trn = "lvinfo" & Right(Year(txtFrPeri.Text), 2)
    If Not FindTable(strMon_Trn) Then
        MsgBox NewCaptionTxt("40082", adrsC), vbInformation
        PeValid = False
        txtFrPeri.SetFocus
        Exit Function
    End If
    Exit Function '' No need to check transaction file for future leave availment Report.
End If
If typOptIdx.bytPer = 6 Then
Dim cntchk As Integer
 For i = 1 To 7
   If optGrp(i).Value = True Then cntchk = cntchk + 1
 Next i
 If cntchk = 0 Then
 MsgBox "Selection of atleast one  group is mandatory for this  Report!", vbInformation
 PeValid = False
 Exit Function
 End If
 End If
''

    If Not FindTable(Left(MonthName(Month(CDate(txtFrPeri.Text))), 3) & _
    Right(Year(CDate(txtFrPeri.Text)), 2) & "trn") Then
        Call SetMSF1Cap(10)
        MsgBox "Monthly Transaction File is Not Found for the  Month of   " & _
            MonthName(Month(CDate(txtFrPeri.Text))), vbExclamation
        PeValid = False
        txtFrPeri.SetFocus
        Exit Function
    End If
    If Not FindTable(Left(MonthName(Month(CDate(txtToPeri.Text))), 3) & _
    Right(Year(CDate(txtToPeri.Text)), 2) & "trn") Then
        Call SetMSF1Cap(10)
        MsgBox NewCaptionTxt("40065", adrsC) & _
        MonthName(Month(CDate(txtToPeri.Text))), vbExclamation
        PeValid = False
        txtToPeri.SetFocus
        Exit Function
    End If

Exit Function
ERR_P:
    ShowError ("peValid :: Reportsfrm")
    PeValid = False
    'Resume Next
End Function

Private Function wkReportsMod() As Boolean
On Error GoTo ERR_P
wkReportsMod = False                    '' FUNCTION FOR WEEKLY REPORT
Call SetRepVars(2)
Call SetMSF1Cap(8)
Select Case typOptIdx.bytWek
Case 0, 7, 8 'Irregular 'Performance '8 add by  03-04
'    If Not WkPerfo(typRep.strWkDate, CStr(DateCompDate(typRep.strWkDate) + 6)) Then Exit Function
'           wkReportsMod = True
'    ''For Mauriitus 11-08-2003
    If Not WKPerfOvt Then Exit Function
    wkReportsMod = True
Case 1, 2, 3, 4, 5
    If Not WkOtherRep Then Exit Function
    wkReportsMod = True
Case 6 'Shift schedule
    If Not WkShiftRep() Then Exit Function
    wkReportsMod = True
Case 9  ' 14-04
    If Not WKFormJ Then Exit Function
    wkReportsMod = True
Case Else
    wkReportsMod = False
End Select
Exit Function
ERR_P:
    ShowError ("wkReportsMod :: Reportsfrm")
End Function
': Final Changes done by removiving logical errors. and for customization of Crystal Reports.
Private Function dlySetEmpstr3() As Boolean
On Error GoTo RepErr                    '' FUNCTION TO CREATE COMMANDTEXT QUERIES FOR DAILY

dlySetEmpstr3 = True
empstr3 = ""
Select Case typOptIdx.bytDly
'artim to entry changed by  on 16th sep 2008 for viraj and standard also.
'MIS2007DF027
Case 0 'Physical Arrival
    empstr3 = "SELECT " & strMon_Trn & ".Empcode," & strMon_Trn & ".shift," & strMon_Trn & ".date," & _
    rpgroup & ",empmst.Name,empmst.card,Empmst." & strKGroup & ",groupmst.grupdesc," & strMon_Trn & ".arrtim," & strMon_Trn & ".deptim, " & strMon_Trn & ".latehrs, " & _
    strMon_Trn & ".presabs, " & strMon_Trn & ".remarks FROM " & strMon_Trn & "," & rpTables & _
    " WHERE " & strMon_Trn & _
        ".entry >0 and  empmst.Empcode = " & strMon_Trn & ".Empcode AND " & _
    strMon_Trn & "." & strKDate & "=" & strDTEnc & Format(DateCompDate(typRep.strDlyDate), "DD-MMM-YYYY") & strDTEnc & _
    " " & strSql & " order by " & strMon_Trn & ".Empcode"
Case 1 'Absent
        'Query Changed By :
    empstr3 = "SELECT " & strMon_Trn & ".Empcode, " & strMon_Trn & ".Shift, " & strMon_Trn & ".Remarks," & rpgroup & _
    ",empmst.Name,Empmst." & strKGroup & ",groupmst.grupdesc," & strMon_Trn & ".presabs FROM " & strMon_Trn & "," & rpTables & _
    " WHERE " & strMon_Trn & "." & strKDate & " = " & strDTEnc & Format(DateCompDate(typRep.strDlyDate), "DD-MMM-YYYY") & _
    strDTEnc & " " & " and (" & LeftStr(strMon_Trn & ".presabs") & " = '" & pVStar.AbsCode & _
    "' or " & RightStr(strMon_Trn & ".presabs") & "='" & pVStar.AbsCode & "')" & _
    " and " & strMon_Trn & ".Empcode = empmst.Empcode " & strSql & " order by " & _
    strMon_Trn & ".Empcode"
Case 2 'Cont Absent
    empstr3 = "SHAPE {SELECT " & strRepFile & ".PresAbsStr," & strRepFile & ".Empcode," & _
    "empmst.Name," & rpgroup & " FROM " & strRepFile & "," & rpTables & " WHERE " & _
     strRepFile & ".Empcode=empmst.Empcode " & strSql & " order by " & strRepFile & _
     ".Empcode} as weekReport compuTE weekReport BY '" & sqlStr & "','" & headGrp & "'"
Case 3 'Late Arrival
    empstr3 = "SELECT " & strMon_Trn & ".Empcode," & strMon_Trn & ".shift," & strMon_Trn & ".Remarks," & _
    rpgroup & ",empmst.Name,Empmst." & strKGroup & ",groupmst.grupdesc," & strMon_Trn & ".arrtim, " & strMon_Trn & ".latehrs, " & _
    strMon_Trn & ".presabs, " & strMon_Trn & ".remarks FROM " & strMon_Trn & "," & rpTables & _
    " WHERE " & strMon_Trn & _
        ".latehrs >0 and  empmst.Empcode = " & strMon_Trn & ".Empcode AND " & _
    strMon_Trn & "." & strKDate & "=" & strDTEnc & Format(DateCompDate(typRep.strDlyDate), "DD-MMM-YYYY") & strDTEnc & _
    " " & strSql & " order by " & strMon_Trn & ".Empcode"
Case 4 'Early Dep
    empstr3 = "SELECT " & strMon_Trn & ".Empcode," & strMon_Trn & ".shift," & _
    rpgroup & ",empmst.Name," & strMon_Trn & ".deptim, " & strMon_Trn & ".earlhrs,Empmst." & strKGroup & "," & _
    strMon_Trn & ".presabs, " & strMon_Trn & ".remarks,groupmst.grupdesc FROM " & strMon_Trn & "," & _
    rpTables & " WHERE " & strMon_Trn & ".Empcode = empmst.Empcode AND " & strMon_Trn & _
    "." & strKDate & "=" & strDTEnc & Format(DateCompDate(typRep.strDlyDate), "DD-MMM-YYYY") & strDTEnc & " and " & _
    strMon_Trn & ".earlhrs>0" & " " & strSql & " order by " & strMon_Trn & ".Empcode"
Case 5 'Perf
    empstr3 = "SELECT " & strMon_Trn & ".Empcode," & strMon_Trn & ".shift," & _
    rpgroup & ",empmst.Name," & strMon_Trn & ".arrtim," & strMon_Trn & ".latehrs," & _
    strMon_Trn & ".actrt_o," & strMon_Trn & ".actrt_i, " & strMon_Trn & ".time5, " & _
    strMon_Trn & ".time6, " & strMon_Trn & ".time7," & strMon_Trn & ".time8," & strMon_Trn & ".deptim, " & strMon_Trn & ".earlhrs, " & _
    strMon_Trn & ".wrkhrs, " & strMon_Trn & ".presabs, " & strMon_Trn & ".remarks," & _
    strMon_Trn & ".ovtim," & strMon_Trn & ".OTConf,Groupmst.grupdesc,empmst." & strKGroup & " FROM " & strMon_Trn & "," & rpTables & " WHERE " & strMon_Trn & _
    ".Empcode = empmst.Empcode AND " & strMon_Trn & "." & strKDate & "=" & strDTEnc & _
    Format(DateCompDate(typRep.strDlyDate), "DD-MMM-YYYY") & strDTEnc & strSql & " order by " & strMon_Trn & _
    ".Empcode"
Case 6 'Irreg
' Changed the Query as it contain logical error

  empstr3 = "SELECT " & strMon_Trn & ".Empcode," & rpgroup & ",empmst.Name,Empmst." & strKGroup & "," & _
    strMon_Trn & ".arrtim," & strMon_Trn & ".latehrs," & strMon_Trn & ".actrt_o," & _
    strMon_Trn & ".actrt_i," & strMon_Trn & ".od_from," & strMon_Trn & ".od_to," & _
    strMon_Trn & ".deptim," & strMon_Trn & ".wrkhrs,Groupmst.grupdesc," & strMon_Trn & ".Presabs," & strMon_Trn & ".earlhrs," & strMon_Trn & ".shift," & strMon_Trn & ".entry as punches FROM " & strMon_Trn & "," & rpTables & _
    " WHERE " & strMon_Trn & ".Empcode = empmst.Empcode AND " & strMon_Trn & ".entry In (1,3,5,7) AND " & strMon_Trn & "." & strKDate & " = " & _
    strDTEnc & Format(DateCompDate(typRep.strDlyDate), "DD-MMM-YYYY") & strDTEnc & " " & _
    strSql & " order by " & strMon_Trn & ".Empcode"
Case 7, 13 'authorized / unauthorized OT
    empstr3 = "SELECT " & strMon_Trn & ".Empcode," & strMon_Trn & ".shift," & _
    rpgroup & ",empmst.Name,Empmst." & strKGroup & "," & strMon_Trn & ".arrtim," & strMon_Trn & ".latehrs," & _
    strMon_Trn & ".actrt_o, " & strMon_Trn & ".actrt_i, " & strMon_Trn & ".deptim, " & _
    strMon_Trn & ".earlhrs, " & strMon_Trn & ".wrkhrs, " & strMon_Trn & ".ovtim," & _
    strMon_Trn & ".OTRem as remarks,Groupmst.grupdesc FROM " & strMon_Trn & "," & rpTables & " WHERE " & strMon_Trn & "." & _
    strKDate & " = " & strDTEnc & Format(DateCompDate(typRep.strDlyDate), "DD-MMM-YYYY") & strDTEnc & " and " & _
    strMon_Trn & ".OTConf = '" & IIf(typOptIdx.bytDly = 7, "Y", "N") & "' " & _
    " and " & _
    strMon_Trn & ".ovtim>0 AND " & strMon_Trn & ".Empcode = empmst.Empcode AND " & _
    strMon_Trn & "." & strKDate & "=" & strDTEnc & Format(DateCompDate(typRep.strDlyDate), "DD-MMM-YYYY") & strDTEnc & _
    strSql & " order by " & strMon_Trn & ".Empcode "
Case 8 'Entries
    empstr3 = "SELECT " & strRepFile & ".Empcode," & rpgroup & ",empmst.Name,Empmst." & strKGroup & "," & _
    strRepFile & ".punches from " & strRepFile & "," & rpTables & " where " & strRepFile & _
    ".Empcode =  empmst.Empcode " & strSql & " order by " & strRepFile & ".Empcode"
Case 9 'Shift Arrangement
    strMon_Trn = Left(MonthName(Month(DateCompDate(typRep.strDlyDate))), 3) & _
            Right(Year(DateCompDate(typRep.strDlyDate)), 2) & "shf"
        If Not FindTable(strMon_Trn) Then
            MsgBox NewCaptionTxt("M7001", adrsMod) & MonthName(Month(DateCompDate(typRep.strDlyDate))) & _
            NewCaptionTxt("00055", adrsMod), vbInformation
            Exit Function
        End If
    empstr3 = "SELECT " & strMon_Trn & ".Empcode," & strMon_Trn & ".d" & _
    Day(DateCompDate(typRep.strDlyDate)) & " as shift, " & rpgroup & ",Groupmst.grupdesc,empmst.Name,Empmst." & strKGroup & " FROM " & _
    strMon_Trn & "," & rpTables & " WHERE (" & strMon_Trn & ".d" & Day(DateCompDate(typRep.strDlyDate)) & _
    " <> '' OR " & strMon_Trn & ".d" & Day(DateCompDate(typRep.strDlyDate)) & " IS NOT NULL ) AND " & _
    strMon_Trn & ".Empcode = empmst.Empcode " & strSql & " order by " & strMon_Trn & _
    ".Empcode "
Case 10 'Manpower
    
    empstr3 = "SELECT " & strRepFile & ".srno," & strRepFile & ".Empcode ," & _
    rpgroup & " ,empmst.Name,Empmst." & strKGroup & "," & strRepFile & ".present," & strRepFile & ".absent," & _
    strRepFile & ".offs," & strMon_Trn & ".ovtim,absentT,presentT,offT FROM " & strRepFile & "," & strMon_Trn & _
    "," & rpTables & " WHERE " & strRepFile & ".Empcode = empmst.Empcode AND " & strRepFile & _
    ".Empcode = " & strMon_Trn & ".Empcode AND empmst.Empcode = " & strMon_Trn & ".Empcode AND " & strMon_Trn & "." & strKDate & " = " & strDTEnc & _
    Format(DateCompDate(typRep.strDlyDate), "DD-MMM-YYYY") & strDTEnc & strSql & " order by " & strRepFile & _
    ".Empcode"

Case 11 'OutDoor
 '   Adjusment done  for daily Report
    ' for Fairfield on 21/02/2005 to display outdoor in daily Report
    strLvloc = ""
    strLvloc = Left(MonthName(Month(typRep.strDlyDate)), 3) & Right(Year(typRep.strDlyDate), 2) & "trn"
    If InVar.strSer <> 3 Then
    empstr3 = "SELECT " & strLvloc & ".Empcode," & strLvloc & ".shift," & _
    rpgroup & ",empmst.Name,Empmst." & strKGroup & "," & strLvloc & "." & strKDate & "," & strLvloc & ".presabs FROM " & _
    strLvloc & "," & rpTables & " WHERE " & strLvloc & ".Empcode = empmst.Empcode " & _
    "AND " & strLvloc & "." & strKDate & "=" & strDTEnc & Format(DateCompDate(typRep.strDlyDate), "DD-MMM-YYYY") & strDTEnc & _
    " and (left(" & strLvloc & ".presabs,2) = 'OD' or right(" & strLvloc & ".presabs,2) = 'OD')" & strSql & " order by " & _
    strLvloc & ".Empcode"
     Else
     ' For Oracle
    empstr3 = "SELECT " & strLvloc & ".Empcode," & strLvloc & ".shift," & _
    rpgroup & ",empmst.Name,Empmst." & strKGroup & "," & strLvloc & "." & strKDate & "," & strLvloc & ".presabs FROM " & _
    strLvloc & "," & rpTables & " WHERE " & strLvloc & ".Empcode = empmst.Empcode " & _
    "AND " & strLvloc & "." & strKDate & "=" & strDTEnc & Format(DateCompDate(typRep.strDlyDate), "DD-MMM-YYYY") & strDTEnc & _
    " and (Lpad(" & strLvloc & ".presabs,2) = 'OD' or Rpad(" & strLvloc & ".presabs,2) = 'OD')" & strSql & " order by " & _
    strLvloc & ".Empcode"
    End If
Case 12 'Summary
   If GetFlagStatus("IMAGE") Then
   empstr3 = " Select " & strRepFile & ".* ," & strRepFile & ".empname as name ," & strRepFile & ".Cname as CName,company.img  From " & strRepFile & ",company"
   Else
   empstr3 = " Select " & strRepFile & ".* ," & strRepFile & ".empname as name ," & strRepFile & ".Cname as CName From " & strRepFile & ""
   End If
End Select
'Call CadSetting
Exit Function
RepErr:
    dlySetEmpstr3 = False
    ShowError ("Dlysetempstr3 :: " & Me.Caption)
End Function

Private Function WkSetEmpstr3() As Boolean
On Error GoTo RepErr                '' FUNCTION TO CREATE COMMANDTEXT QUERIES FOR WEEKLY

                                    '' FUNCTION TO CREATE COMMANDTEXT QUERIES FOR WEEKLY
WkSetEmpstr3 = True
empstr3 = ""
Select Case typOptIdx.bytWek
   Case 0, 7 'Performance,Irregular
        empstr3 = "SELECT empmst.Name," & strRepFile & _
        ".* ,empmst." & strKGroup & "," & rpgroup & " FROM " & strRepFile & "," & _
         rpTables & " WHERE  " & strRepFile & ".Empcode = empmst.Empcode " & strSql & _
        " ORDER BY " & strRepFile & ".Empcode"
   Case 1, 2, 3, 4, 5 'Absent,Attendance,Late Arrival,Early Departure,Overtime
         empstr3 = "select " & strRepFile & ".Empcode," & rpgroup & ",empmst.Name," & _
         strRepFile & ".frw," & strRepFile & ".secw," & strRepFile & ".thw," & strRepFile & _
         ".fow," & strRepFile & ".fiw," & strRepFile & ".siw," & strRepFile & ".sevw ,empmst." & strKGroup & " from " & _
        strRepFile & "," & rpTables & " where empmst.Empcode=" & strRepFile & ".Empcode " & _
        strSql & " ORDER BY " & strRepFile & ".Empcode"
   Case 6 'Shift Schedule
        empstr3 = "select " & strRepFile & ".Empcode," & rpgroup & ",empmst.Name," & _
        strRepFile & ".frw," & strRepFile & ".secw," & strRepFile & ".thw," & strRepFile & _
        ".fow," & strRepFile & ".fiw," & strRepFile & ".siw," & strRepFile & ".sevw ,empmst." & strKGroup & " from " & _
        strRepFile & "," & rpTables & " where empmst.Empcode=" & strRepFile & ".Empcode " & _
        strSql
   Case 8   ' 19-03
        empstr3 = "select Empmst.Empcode," & strRepFile & ".*, Empmst.Name, Empmst.Sex, Empmst.birth_dt, GroupMst.[Group]," & rpgroup & " from " & strRepFile & ", " & rpTables & " Where " & strRepFile & ".Empcode = empmst.Empcode  " & strSql & " ORDER BY EmpMst.Empcode"
  
End Select

Exit Function
RepErr:
    WkSetEmpstr3 = False
    ShowError ("WkSetEmpstr3 :: " & Me.Caption)
End Function


Private Function monSetEmpstr3() As Boolean
On Error GoTo RepErr               '' FUNCTION TO CREATE COMMANDTEXT QUERIES FOR MONTHLY
monSetEmpstr3 = True
empstr3 = ""
Select Case typOptIdx.bytMon
'***:******************
   Case 0, 5, 10, 11  'Performance ':  Modified according to the new logic
  
  'Select Case InVar.strSer
         
   '     Case 2 ''ms-access
        'change by
'        If InVar.strSer = 3 Then
            empstr3 = "SELECT " & strRepFile & ".ArrStr," & strRepFile & ".DepStr," & _
            strRepFile & ".EarlStr, " & strRepFile & ".LateStr, " & strRepFile & ".OTStr," & _
            strRepFile & ".PresAbsStr, " & strRepFile & ".ShfStr, " & strRepFile & ".WorkStr," & _
            strRepFile & "." & strKDate & ", " & strRepFile & ".Empcode, empmst.name," & rpgroup & _
            ", " & strRepFile & ".sumlate," & strRepFile & ".sumearly," & strRepFile & _
            ".sumwork," & strRepFile & ".sumOT,Groupmst." & strKGroup & "  FROM " & strRepFile & "," & rpTables & " WHERE " & _
            strRepFile & ".Empcode = empmst.Empcode " & strSql & " ORDER BY " & strRepFile & _
            ".Empcode"
'        Else
'             Dim strLeaveFile1 As String
'            If CByte(pVStar.Yearstart) > MonthNumber(typRep.strMonMth) Then     ' 27-01
'            strLeaveFile1 = "Lvtrn" & Right((DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear - 1, "l"))), 2)
'            Else
'            strLeaveFile1 = "Lvtrn" & Right((DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "l"))), 2)
'            End If
'
'            empstr3 = "SELECT " & strRepFile & ".ArrStr," & strRepFile & ".DepStr," & _
'            strRepFile & ".EarlStr, " & strRepFile & ".LateStr, " & strRepFile & ".OTStr," & _
'            strRepFile & ".PresAbsStr, " & strRepFile & ".ShfStr, " & strRepFile & ".WorkStr," & _
'            strRepFile & "." & strKDate & ", " & strRepFile & ".Empcode," & strLeaveFile1 & ".P," & strLeaveFile1 & ".A," & strLeaveFile1 & ".WO," & strLeaveFile1 & ".paiddays,empmst.[name]," & rpgroup & _
'            ", " & strRepFile & ".sumlate," & strRepFile & ".sumearly," & strRepFile & _
'            ".sumwork," & strRepFile & ".sumOT,Groupmst." & strKGroup & " FROM " & strLeaveFile1 & "," & strRepFile & "," & rpTables & " WHERE " & _
'            strLeaveFile1 & ".empcode=" & strRepFile & ".empcode and " & strLeaveFile1 & ".lst_date=" & strDTEnc & Format(DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "l")), "DD/MMM/YY") & strDTEnc & " and " & strLeaveFile1 & ".empcode=empmst.empcode and " & strRepFile & ".Empcode = empmst.Empcode " & strSql & " ORDER BY " & strRepFile & _
'            ".Empcode"
'
'        End If

    Case 2, 3, 4, 7, 12, 13, 17 '' : Muster Reports,Present Report,Absent Report,Memo Report,Late arrival memo
        
        empstr3 = "Select " & strRepMfile & ".*, empmst.name,empmst." & strKGroup & "," & rpgroup & " from " _
        & strRepMfile & "," & rpTables & " WHERE " & _
        strRepMfile & ".Empcode = empmst.Empcode  " & strSql & " ORDER BY " & strRepMfile & _
        ".Empcode"

     
   Case 1
        empstr3 = "Select " & strRepMfile & ".*," & strMon_Trn & ".*,groupmst." & strKGroup & ",empmst.name," & rpgroup & " from " _
        & strRepMfile & "," & strMon_Trn & "," & rpTables & " WHERE " & strMon_Trn & ".LST_DATE = " & strDTEnc & _
        Format(DateCompDate(strlastdatem), "DD-MMM-YYYY") & strDTEnc & " AND empmst.Empcode = " & strMon_Trn & ".Empcode and " & _
        strRepMfile & ".Empcode = empmst.Empcode " & strSql & " ORDER BY " & strRepMfile & _
        ".Empcode"
        Sleep (2000)

   Case 6
        empstr3 = "SELECT " & strMon_Trn & ".Empcode," & strMon_Trn & ".ot_hrs," & _
         strMon_Trn & ".otpd_hrs," & strMon_Trn & ".lst_date, empmst.Name,groupmst." & strKGroup & ", groupmst.grupdesc, " & rpgroup & "  FROM " & strMon_Trn & _
         "," & rpTables & " WHERE " & strMon_Trn & ".LST_DATE = " & strDTEnc & _
         Format(DateCompDate(strlastdatem), "DD-MMM-YYYY") & strDTEnc & " AND " & strMon_Trn & ".otpd_hrs > 0 AND " & _
         "empmst.Empcode = " & strMon_Trn & ".Empcode " & strSql & " "

   Case 8, 15, 16 '  Absent/Late/Early
        empstr3 = "SELECT Distinct " & strRepFile & ".*,Empmst.Name,empmst." & strKGroup & "," & rpgroup & " FROM " & strRepFile & "," & rpTables & " WHERE " & _
        "Empmst.Empcode = " & strRepFile & ".Empcode  " & strSql & ""
   Case 9 '  Leave Balance
        empstr3 = "SELECT " & strRepFile & ".*,empmst.Name,empmst." & strKGroup & "," & rpgroup & " FROM " & strRepFile & "," & _
        rpTables & " WHERE " & strRepFile & ".Empcode = Empmst.Empcode " & strSql & _
        " ORDER BY " & strRepFile & ".Empcode"
   Case 14 ' Leave Consumption

           empstr3 = "select distinct " & strRepFile & ".Empcode,lcode,fromdate,todate,days," & _
            "leave,trcd, absent,name,groupmst." & strKGroup & ",groupmst.grupdesc, " & rpgroup & " from " & strRepFile & ",leavdesc," & rpTables & " WHERE " & strRepFile & _
            ".Empcode = empmst.Empcode AND " & strRepFile & ".lcode = Leavdesc.lvcode AND " & _
            "leavdesc.cat = catdesc.cat " & strSql & " order by " & strRepFile & ".Empcode"
            strMon_Trn = ""
        
   Case 18 ' WO On Holiday
        empstr3 = "SELECT " & strMon_Trn & ".Empcode," & strMon_Trn & "." & strKDate & "," & _
        strMon_Trn & ".presabs,groupmst." & strKGroup & ", empmst.Name," & rpgroup & " From " & strMon_Trn & "," & _
        rpTables & " WHERE " & "(" & LeftStr(strlastdatem) & " = empmst." & strKOff & " OR " & _
        LeftStr(strlastdatem) & " = empmst.off2 OR " & LeftStr(strlastdatem) & " " & _
        "= empmst.wo_1_3 OR " & LeftStr(strlastdatem) & " = empmst.wo_2_4) AND " & _
        strMon_Trn & ".presabs = '" & pVStar.HlsCode & pVStar.HlsCode & "' AND " & _
        "empmst.Empcode = " & strMon_Trn & ".Empcode " & strSql
        
End Select
'Call CadSetting
Exit Function
RepErr:
    monSetEmpstr3 = False
    ShowError ("monSetEmpstr3 :: " & Me.Caption)
End Function

Private Function yrSetEmpstr3() As Boolean
On Error GoTo RepErr            '' FUNCTION TO CREATE COMMANDTEXT QUERIES FOR YEARLY
yrSetEmpstr3 = True
empstr3 = ""
Select Case typOptIdx.bytYer
    Case 0, 3
        empstr3 = "select " & strRepFile & ".*,name,empmst." & strKGroup & "," & rpgroup & " from " & strRepFile & "," & rpTables & " where " & _
        strRepFile & ".Empcode=empmst.Empcode " & strSql & ""
    Case 1, 2 'Mandays,Performance
        ''empstr3 = "select distinct " & strMon_Trn & ".* ,name,empmst." & strKGroup & "," & rpGroup & " from " & strMon_Trn & "," & rpTables & " where " & _
        strMon_Trn & ".Empcode=empmst.Empcode " & strSql & ""
        empstr3 = "select " & strRepFile & ".empcode,NAME," & strRepFile & ".ystr," & strRepFile & ".yvalstr," & _
        strRepFile & ".pddaysstr," & strRepFile & ".wrkstr," & strRepFile & ".nightstr," & _
        strRepFile & ".ltno," & strRepFile & ".latehrs," & strRepFile & ".erno," & _
        strRepFile & ".earlhrs," & strRepFile & ".counter," & rpgroup & IIf(GetFlagStatus("IMAGE"), ",company.img", "") & " FROM " & _
        strRepFile & "," & rpTables & " WHERE Empmst.empcode = " & strRepFile & ".Empcode " & _
        strSql & " ORDER BY " & strRepFile & ".EMPCODE," & strRepFile & ".counter"
    Case 4 'Leave Information
        
        empstr3 = "SELECT " & strRepFile & ".Empcode, " & strRepFile & ".yvalstr," & strRepFile & ".lcode," & strRepFile & ".FromLv ," & _
        strRepFile & ".todate, " & strRepFile & ".AvailLv, " & strRepFile & ".CreditLv,name, " & rpgroup & " " & _
        "FROM " & strRepFile & "," & rpTables & " WHERE empmst.Empcode = " & strRepFile & _
        ".Empcode " & strSql & ""
        ''
    Case 5
        Select Case bytBackEnd
            Case 1 'SQL
                empstr3 = "SELECT TmpLeavBalH.strT as HEADERS,TmpLeavBal.strT AS VALS,empmst.Empcode," & rpgroup & " FROM TmpLeavBalH,TmpLeavBal," & rpTables & _
                " WHERE SUBSTRING(TmpLeavBal.strT,1," & pVStar.CodeSize & ") = empmst.empcode " & strSql & ""
            Case 2 'Access
            empstr3 = "SELECT TmpLeavBalH.strT as HEADERS,TmpLeavBal.strT AS VALS,empmst.Empcode," & rpgroup & " FROM TmpLeavBalH,TmpLeavBal," & rpTables & _
                " WHERE MID(TmpLeavBal.strT,1," & pVStar.CodeSize & ") = empmst.empcode " & strSql & ""
            Case 3 'oracle
            empstr3 = "SELECT TmpLeavBalH.strT as HEADERS,TmpLeavBal.strT AS VALS,empmst.Empcode," & rpgroup & " FROM TmpLeavBalH,TmpLeavBal," & rpTables & _
                " WHERE SUBSTR(TmpLeavBal.strT,1," & pVStar.CodeSize & ") = empmst.empcode " & strSql & ""
        End Select
    End Select
'Call CadSetting
Exit Function
RepErr:
    yrSetEmpstr3 = False
    ShowError ("yrSetEmpstr3 :: " & Me.Caption)
End Function

Private Function maSetEmpstr3() As Boolean
On Error GoTo RepErr                '' FUNCTION TO CREATE COMMANDTEXT QUERIES FOR MASTERS
maSetEmpstr3 = True
Dim Empstr1 As String, Empstr2 As String, strSql1 As String


    Empstr1 = "SELECT Empmst.name,Empmst.Empcode,Empmst.card, frmDesignation.DesigName," & _
            rpgroup & ",Empmst.joindate,Empmst." & strKGroup & ",Empmst.styp," & _
            "Empmst.entry,Empmst.birth_dt,Empmst.salary,Empmst.resadd1,Empmst.city," & _
            "Empmst.phone,Empmst.pin,Empmst.sex,Empmst.bg,Empmst.udf1,Empmst.udf2," & _
            "Empmst.udf3,Empmst.udf4,Empmst.udf5,Empmst.udf6,Empmst.udf7,Empmst.udf9," & _
                "Empmst.udf10,Empmst.qualf,Empmst.email_id,Empmst.name2,Empmst.leavdate FROM " & rpTables & ", frmDesignation WHERE "
                '"Empmst.udf10,Empmst.leavdate FROM " & rpTables & " WHERE "

Empstr2 = " And frmDesignation.DesigCode = empmst.designatn ORDER BY Empmst.Empcode,catdesc.cat ,deptdesc.dept"
empstr3 = ""
'******
Dim strTmp As String

Select Case typOptIdx.bytMst
    Case 0, 1
        
        If frMast.Caption = "Joining" Then
            If IsDate(txtMastFr.Text) And IsDate(txtMastTo.Text) Then
            Empstr1 = Empstr1 + "Empmst.joindate >= " & strDTEnc & Format(DateCompDate(txtMastFr.Text), "DD/MMM/YYYY") & strDTEnc & " And Empmst.joindate <=" & strDTEnc & Format(DateCompDate(txtMastTo.Text), "DD/MMM/YYYY") & strDTEnc & " AND "
            End If
            
        End If
        
        
         empstr3 = Empstr1 & " (empmst.leavdate is null) " & strSql & _
                             Empstr2 & ""

        
    Case 2 'Left Employee
        ''Supriya dated 22/08/05(For Crystal Rep)
        empstr3 = Empstr1 & " empmst.leavdate is not NULL and " & _
        "empmst.leavdate between " & strDTEnc & Format(DateCompDate(typRep.strLeftFr), "DD-MMM-YYYY") & strDTEnc & _
        " and " & strDTEnc & Format(DateCompDate(typRep.strLeftTo), "DD-MMM-YYYY") & strDTEnc & " " & strSql & _
        " " & Empstr2
    Case 3 'Leave Master

        empstr3 = "select * from leavdesc where isitleave = 'Y'"
   
    Case 4  ''Shift master
        If InVar.strSer = 1 Then
            empstr3 = "SELECT *,  shiftdd= " & _
                         "CASE " & _
                        "WHEN  isnumeric(shift) =0  and len(shift)=1 THEN ascii(substring(shift,1,1))" & _
                        "WHEN  isnumeric(shift) =0  and len(shift)=2 THEN ascii(substring(shift,1,1))" & _
                        " +  ascii(substring(shift,2,1))" & _
                        "WHEN  isnumeric(shift) =1   THEN Shift " & _
                        "End " & _
                        " From instshft where shift<>'100'" & _
                        "order by shiftdd"
        ElseIf GetFlagStatus("IMAGE") Then
           empstr3 = "Select hdend,hdstart,rst_in,rst_out,shf_hrs,shf_in,shf_out,shift,shiftname,rst_brk,company.img from instshft,company where shift <> '100' ORDER BY shift"
        Else
           empstr3 = "Select hdend,hdstart,rst_in,rst_out,shf_hrs,shf_in,shf_out,shift,shiftname,rst_brk from instshft where shift <> '100' ORDER BY shift"
        End If
    Case 5  ''Rotation Shift Master
        If InVar.strSer = 1 Then
            empstr3 = "select scode,Name,mon_oth,left(skp,len(skp)-1) as Skp1,left(pattern,len(pattern)-1) as Pattern1 from ro_shift where Scode <> '100' order by scode"
        ElseIf InVar.strSer = 2 Then
            'empstr3 = "select scode,Name,mon_oth,left(skp,len(skp)-1) as Skp,left(pattern,len(pattern)-1) as Pattern from ro_shift where Scode <> '100' order by scode"
              empstr3 = "select scode,Name,mon_oth,left(skp,len(skp)-1) as Skp1,left(pattern,len(pattern)-1) as Pattern1 from ro_shift,company where Scode <> '100' order by scode"
        Else
                empstr3 = "select scode,Name,mon_oth,lpad(skp,length(skp)-1) as skp1 ,lpad(pattern,length(pattern)-1) as pattern1  from ro_shift where Scode <> '100' order by scode"
        End If
    Case 6 'Holiday

            empstr3 = "select Catdesc." & strKDesc & " as cat," & strKDate & ",Holiday." & strKDesc & " from holiday,Catdesc where holiday.cat=catdesc.cat and " & _
            "holiday." & strKDate & " between " & _
            strDTEnc & Format(DateCompDate(FdtLdt(CByte(pVStar.Yearstart), pVStar.YearSel, "f")), "DD-MMM-YYYY") & strDTEnc & " and " & strDTEnc & Format(FdtLdt(CByte(pVStar.Yearstart) - 1, _
            IIf(pVStar.Yearstart = "1", pVStar.YearSel, CStr(Val(pVStar.YearSel) + 1)), "l"), "DD-MMM-YYYY") & _
            strDTEnc & " ORDER BY holiday." & strKDate & ",holiday.cat"

    Case 7  ''Department

            empstr3 = "SELECT Count(Empmst.empcode) AS Strength, deptdesc.dept, deptdesc.desc FROM deptdesc INNER JOIN Empmst ON deptdesc.dept = Empmst.dept Group by  deptdesc.dept, deptdesc.desc"

    
    Case 8  ''Category

        empstr3 = "Select * from Catdesc  where cat <> '100' order by cat"
    
    Case 9  ''Group

            empstr3 = "select groupmst." & strKGroup & " as dept,grupdesc as " & strKDesc & ", " & _
                      " count(empmst." & strKGroup & ") as stre from company,groupmst,empmst where " & _
                      "empmst." & strKGroup & " = groupmst." & strKGroup & "  group by groupmst." & strKGroup & ",grupdesc order by groupmst." & strKGroup & ""
      
    Case 10

             empstr3 = "Select * From frmDesignation order by DesigCode"
            
    Case 11 ''Company
   
            empstr3 = "select company.company as dept,cname as " & strKDesc & ", " & _
                      " count(empmst.company) as stre from company left outer join empmst on " & _
                      "empmst.company = company.company  group by company.company,cname"

    Case 12 ''Division

            empstr3 = "select division.div as dept,divdesc as " & strKDesc & ",count(empmst.div) " & _
            " as stre from company,division,empmst where empmst.div = division.div  group by division.div,divdesc order by division.div "
 
End Select
Exit Function
RepErr:
    maSetEmpstr3 = False
    ShowError ("maSetEmpstr3 :: " & Me.Caption)
End Function

Private Function peSetEmpstr3() As Boolean
On Error GoTo RepErr                    '' FUNCTION TO CREATE COMMANDTEXT QUERIES FOR PERIODIC
peSetEmpstr3 = True
empstr3 = ""
Select Case typOptIdx.bytPer
    Case 0, 2, 17

                empstr3 = "SELECT  #" & Format(txtFrPeri.Text, "dd/mmm/yy") & "# as dt, #" & Format(txtToPeri.Text, "dd/mmm/yy") & "# as dtto, " & strRepFile & ".Empcode,empmst.Name," & strRepFile & _
                "." & strKDate & "," & strRepFile & ".ArrStr, " & strRepFile & ".DepStr," & strRepFile & _
                ".EarlStr, " & strRepFile & ".LateStr, " & strRepFile & ".OTStr," & strRepFile & _
                ".PresAbsStr, " & strRepFile & ".ShfStr, " & strRepFile & ".WorkStr," & strRepFile & _
                ".sumlate," & strRepFile & ".sumearly," & strRepFile & ".sumwork," & strRepFile & ".sumOT," & strRepFile & ".paid," & strRepFile & ".HL," & strRepFile & ".WO," & strRepFile & ".EL," & strRepFile & ".PL," & strRepFile & ".PR," & strRepFile & ".CO," & strRepFile & ".A," & strRepFile & ".P," & strRepFile & ".OD," & strRepFile & ".WP,groupmst." & strKGroup & "," & rpgroup & _
                " FROM " & strRepFile & "," & rpTables & " WHERE " & _
                strRepFile & ".Empcode=empmst.Empcode " & strSql & ""
   

   Case 1 ' pe muster
         empstr3 = "SELECT " & strRepMfile & ".*,empmst.Name,empmst." & strKGroup & "," & rpgroup & _
         " FROM " & strRepMfile & "," & rpTables & " WHERE " & _
         strRepMfile & ".Empcode=empmst.Empcode " & strSql & ""

   
   Case 3, 4 'Late Arrival Early Departure
'   Select Case InVar.strSer
'   Case 2 'ms-access
        empstr3 = "SELECT " & strRepFile & ".Empcode,empmst.Name," & strRepFile & _
        ".ArrStr, " & strRepFile & ".DepStr," & strRepFile & ".EarlStr, " & strRepFile & _
        ".LateStr, " & strRepFile & ".OTStr," & strRepFile & ".PresAbsStr, " & strRepFile & _
        ".ShfStr, " & strRepFile & ".WorkStr," & strRepFile & "." & strKDate & "," & strRepFile & _
        ".sumlate," & strRepFile & ".sumearly,groupmst." & strKGroup & ", " & rpgroup & " FROM " & strRepFile & "," & _
        rpTables & " WHERE " & strRepFile & ".Empcode=empmst.Empcode " & strSql & ""

   
    Case 5       '
         empstr3 = "SELECT " & strRepMfile & ".*,empmst.Name,empmst." & strKGroup & "," & rpgroup & _
         " FROM " & strRepMfile & "," & rpTables & " WHERE " & _
         strRepMfile & ".Empcode=empmst.Empcode " & strSql & ""
  
    Case 6  ''Summary
       empstr3 = " Select " & strRepFile & ".* from " & strRepFile & ",company "
    Case 16           '' 04-03
        empstr3 = "SELECT " & strRepFile & ".Empcode,empmst.Name," & strRepFile & _
        "." & strKDate & "," & strRepFile & ".ArrStr, " & strRepFile & ".DepStr," & strRepFile & _
        ".EarlStr, " & strRepFile & ".LateStr, " & strRepFile & ".OTStr," & strRepFile & _
        ".PresAbsStr, " & strRepFile & ".ShfStr, " & strRepFile & ".WorkStr," & strRepFile & _
        ".sumlate," & strRepFile & ".sumearly," & strRepFile & ".sumwork, groupmst." & strKGroup & "," & rpgroup & _
        " FROM " & strRepFile & "," & rpTables & " WHERE " & _
        strRepFile & ".Empcode=empmst.Empcode " & strSql & ""
    Case 7
        empstr3 = "SELECT " & strRepMfile & ".*, " & strMon_Trn & ".* ,empmst.Empcode AS Empcode, empmst.Name,empmst." & strKGroup & "," & rpgroup & _
         " FROM " & strMon_Trn & ", " & strRepMfile & "," & rpTables & " WHERE  Empmst.Empcode = " & strMon_Trn & ".Empcode And " & _
         strRepMfile & ".Empcode=empmst.Empcode " & strSql & ""

End Select
'Call CadSetting
Exit Function
RepErr:
    peSetEmpstr3 = False
    ShowError ("peSetEmpstr3 :: " & Me.Caption)
End Function
'


Private Function maReportsMod() As Boolean
maReportsMod = True         '' FUNCTION FOR MASTER REPORTS
Call SetRepVars(5)
Call SetMSF1Cap(8)
End Function

Private Function peReportsMod() As Boolean
peReportsMod = False        '' FUNCTION FOR PERIODIC REPORTS
Call SetRepVars(6)
Call SetMSF1Cap(8)
Select Case typOptIdx.bytPer
    Case 0, 2, 17
'        Select Case InVar.strSer
'        Case 2 'MS-access
        If Not pePerfOvt() Then Exit Function
        peReportsMod = True
   Case 1, 7 'Muster Report
        If typOptIdx.bytPer = 7 Then
        strMon_Trn = "prAbsentLvt"
        TruncateTable (strMon_Trn)
       Dim dtfromdate As Date
       Dim dttodate As Date
       dtfromdate = DateCompDate(typRep.strPeriFr)
        dttodate = DateCompDate(typRep.strPeriTo)
       
        frmMonthly.cmdProcess_Click
        Unload frmMonthly
'        strlastdatem = FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "L")
'        If CByte(pVStar.Yearstart) > MonthNumber(dtfromdate) Then
'            strMon_Trn = "lvtrn" & Right(CStr(CInt(typRep.strMonYear) - 1), 2)
'        Else
'            strMon_Trn = "lvtrn" & Right(typRep.strMonYear, 2)
'        End If
        
        If Not FindTable(strMon_Trn) Then
            MsgBox NewCaptionTxt("40081", adrsC), vbInformation
            Exit Function
        End If

        End If
   
        If Not pePerfCryst() Then Exit Function
        peReportsMod = True
   Case 3, 4 ''Late arrival, Early Departure
      ' Select Case InVar.strSer
        'Case 2 Ms''-acces
           If Not peLateEarl Then Exit Function
        peReportsMod = True
   Case 5

        If Not PerContAbs Then Exit Function
   
        peReportsMod = True
    ''For Mauritius 11-07-2003
    Case 6 'Summary
         If Not Fuc_NewSummary Then Exit Function
         peReportsMod = True
    Case 16
        If Not pePerfPhysicalAbsent Then Exit Function
        peReportsMod = True
    End Select
End Function

Private Function yrReportsMod() As Boolean
yrReportsMod = False            '' FUNCTION FOR YEARLY REPORTS
Call SetRepVars(4)
Call SetMSF1Cap(8)
Select Case typOptIdx.bytYer
    Case 0, 3, 16, 17, 18 'Absent,Present
        If Not yrAbsPrs1 Then Exit Function
        yrReportsMod = True
    Case 1, 2 'Mandays,Perofrmance,LeaveWages
        
        ''
        If Not yrManPerf Then Exit Function
        yrReportsMod = True
    Case 4 'Leave Info
        
            ''
        If Not yrLeaveInfo Then Exit Function
        yrReportsMod = True

    Case 5
        If Not yrLeaveBAl Then Exit Function
        yrReportsMod = True
    
    End Select
End Function

Private Function monReportsMod() As Boolean
On Error GoTo ERR_P
monReportsMod = False               '' FUNCTION FOR MONTHLY REPORTS
Call SetRepVars(3)
Call SetMSF1Cap(8)
Select Case typOptIdx.bytMon
       
    Case 0, 5, 10, 11
         
 'Select Case InVar.strSer
    
  '     Case 2 ''ms-access
        If Not monPerfOt Then Exit Function

               monReportsMod = True
             
 'End Select
    Case 10
    
       
    '***
    Case 2, 3, 4
        If Not monMuPACry Then Exit Function
        monReportsMod = True
    
    '***
    Case 1, 7
        If Not monMuPACry Then Exit Function
        strMon_Trn = ""
        strlastdatem = FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "L")
        If CByte(pVStar.Yearstart) > MonthNumber(typRep.strMonMth) Then
            strMon_Trn = "lvtrn" & Right(CStr(CInt(typRep.strMonYear) - 1), 2)
        Else
            strMon_Trn = "lvtrn" & Right(typRep.strMonYear, 2)
        End If
        If Not FindTable(strMon_Trn) Then
            MsgBox NewCaptionTxt("40081", adrsC), vbInformation
            Exit Function
        End If
        monReportsMod = True
    Case 6
        strMon_Trn = ""
            strlastdatem = FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "L")
        If CByte(pVStar.Yearstart) > MonthNumber(typRep.strMonMth) Then
            strMon_Trn = "lvtrn" & Right(CStr(CInt(typRep.strMonYear) - 1), 2)
        Else
            strMon_Trn = "lvtrn" & Right(CStr(CInt(typRep.strMonYear)), 2)
        End If
        If Not FindTable(strMon_Trn) Then
            MsgBox NewCaptionTxt("40081", adrsC), vbInformation
            Exit Function
        End If
        monReportsMod = True
    Case 7, 12, 13 'Absent Memo,Late Arrival Memo,Early Departure Memo
        If Not CFuncMonMemo Then Exit Function
        'If Not monALEMemo Then Exit Function
        monReportsMod = True
    Case 8 'Total Absent/Late/Early
        If Not monALERep Then Exit Function
        monReportsMod = True
    Case 9 'Leave Balance
        ''If Not monLvBalCry Then Exit Function
        If Not monLeaveBal Then Exit Function
            monReportsMod = True
    Case 14 'Leave Consumption
        strMon_Trn = ""
        If Val(pVStar.Yearstart) > MonthNumber(typRep.strMonMth) Then
             strMon_Trn = "lvinfo" & Right(CStr(CInt(typRep.strMonYear) - 1), 2)
        Else
             strMon_Trn = "lvinfo" & Right(typRep.strMonYear, 2)
        End If
        If Not FindTable(strMon_Trn) Then
            MsgBox NewCaptionTxt("40082", adrsC), vbInformation
''            Exit Function
        End If
        strFirstDateM = FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "f")
        strlastdatem = FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "l")

'              '' Previous Code 06-11
            ConMain.Execute " insert into " & strRepFile & "(Empcode,lcode," & _
            "fromdate,todate,days,trcd,Absent ) select Empcode,lcode,fromdate,todate,days,trcd,fordate  from " & _
            strMon_Trn & " where trcd in(4,6,7) and (fromdate between " & strDTEnc & DateCompStr(strFirstDateM) & _
            strDTEnc & " and" & strDTEnc & DateCompStr(strlastdatem) & strDTEnc & " or " & _
            "todate between " & strDTEnc & DateCompStr(strFirstDateM) & strDTEnc & " and " & _
            strDTEnc & DateCompStr(strlastdatem) & strDTEnc & ")"
'        End If
'
        ConMain.Execute "update " & strRepFile & " set trcd= ' ' where trcd ='4'"
        ConMain.Execute "update " & strRepFile & " set trcd= 'Late Cut' where trcd ='6'"
        ConMain.Execute "update " & strRepFile & " set trcd= 'Early Cut' where trcd ='7'"
        
        ' 17-04-09
        strrepfile1 = strRepFile
                                                                                                                 
        monReportsMod = True
                                                                                                                     
    Case 15, 16 'Total Lates,Total Earlys
        If Not monTotLtEr Then Exit Function
        monReportsMod = True
    Case 17 'Shift schedule
'        If Not monShiftSch Then Exit Function
        If Not CFuncMonShift Then Exit Function
        monReportsMod = True
    Case 18 'WO on Holiday
        strMon_Trn = ""
        strMon_Trn = Left(typRep.strMonMth, 3) & Right(typRep.strMonYear, 2) & "trn"
        Select Case bytBackEnd
            Case 1  ''SQL SERVER
                strlastdatem = "datename(dw," & strMon_Trn & "." & strKDate & ")"
            Case 2  ''MS-Access
                strlastdatem = "format(" & strMon_Trn & "." & strKDate & ",'dddd')"
            Case 3  ''Oracle
                strlastdatem = "TO_CHAR(" & strMon_Trn & "." & strKDate & ",'Day')"
        End Select
        monReportsMod = True

End Select
Exit Function
ERR_P:
    ShowError ("monReportsMod :: " & Me.Caption)
    'Resume Next
End Function

Private Sub SetVarEmpty()
On Error GoTo ERR_P
If strRepFile <> "" Then Call ChkRepFile        '' INITIALIZES REPORT'S ALL GLOBAL VARIABLES
empstr3 = "": sqlStr = "": headGrp = "": strSql = ""
Set RsName = Nothing
Set Report = Nothing
Set crxApp = Nothing
'CRV.ViewReport
Set CRV = Nothing
Dim i As Integer

    For i = 1 To 7
        strAGrp(i) = ""
         strAlbl(i) = ""
        strAhead(i) = ""
    Next
    
strRepName = "": strMon_Trn = "": bytPoLa = 0
DateStr = ""
Select Case bytRepMode
    Case 1 'daily
        typRep.strDlyDate = ""
    Case 2 'weekly
        typRep.strWkDate = ""
    Case 3 'monthly
        typRep.strMonMth = ""
        typRep.strMonYear = ""
    Case 4 'yearly
        typRep.strYear = ""
    Case 5 'masters
        typRep.strLeftFr = ""
        typRep.strLeftTo = ""
    Case 6 'periodic
        typRep.strPeriFr = ""
        typRep.strPeriTo = ""
End Select
Exit Sub
ERR_P:
    Select Case Err.Number
        Case 91
        Case Else
            ShowError ("SetVarEmpty :: Reports")
    End Select
    'Resume Next
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
    Erase strArrTmp
End If
cboShift.AddItem "All"
cboShift.ListIndex = cboShift.ListCount - 1
Exit Sub
ERR_P:
    ShowError ("FillShiftCombo :: " & Me.Caption)
End Sub

Private Sub FillMainCombo()
On Error GoTo ERR_P
Dim strArrTmp() As String, bytTmp As Byte
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select distinct qualf from empmst where qualf is not null Order by qualf", ConMain, adOpenStatic
If Not (adrsDept1.BOF And adrsDept1.EOF) Then
    cboMain.ColumnCount = 1
    cboMain.ListWidth = "4.5 cm"
    cboMain.ColumnWidths = "1.5 cm"
    ReDim Preserve strArrTmp(adrsDept1.RecordCount - 1)
    For bytTmp = 0 To adrsDept1.RecordCount - 1
        strArrTmp(bytTmp) = adrsDept1("Qualf")                       '' Qualf
        adrsDept1.MoveNext
    Next
    cboMain.List = strArrTmp
    Erase strArrTmp
End If
cboMain.AddItem "All"
cboMain.ListIndex = cboMain.ListCount - 1
Exit Sub
ERR_P:
    ShowError ("FillMainCombo :: " & Me.Caption)
End Sub

Private Sub ShowGroupEx(Optional blnShow As Boolean = True)
    frSel.Enabled = blnShow
    lblGroupBy.Enabled = blnShow
    cmbFrEmpSel.Enabled = blnShow
    cmbToEmpSel.Enabled = blnShow
    cmbFrDepSel.Enabled = blnShow
    cmbToDepSel.Enabled = blnShow
    cmbFrCatSel.Enabled = blnShow
    cmbToCatSel.Enabled = blnShow
    cmbFrGrpSel.Enabled = blnShow
    cmbToGrpSel.Enabled = blnShow
    cmbFrLocSel.Enabled = blnShow
    cmbToLocSel.Enabled = blnShow
    cmbFrComSel.Enabled = blnShow
    cmbFrDivSel.Enabled = blnShow
    cmbToDivSel.Enabled = blnShow
    cbodept.Enabled = blnShow
    deptall.Enabled = blnShow
    catall.Enabled = blnShow
    catlist.Enabled = blnShow
'*****
'    chkNewP.Enabled = blnShow
    
End Sub

'
Private Sub ShowPrinterDuplex()
    Dim i As Integer
    Dim PrinterDuplex As CRPrinterDuplexType
    If cboPrinterDuplex.ListCount = 0 Then
    Addcbo cboPrinterDuplex, "Simplex", crPRDPSimplex
    Addcbo cboPrinterDuplex, "Horizontal", crPRDPHorizontal
    Addcbo cboPrinterDuplex, "Vertical", crPRDPVertical
     End If
     ' hard coded for the setting in which report was made but user can change it run time.
    With cboPrinterDuplex
        For i = 0 To .ListCount - 1
            If .ItemData(i) = crPRDPSimplex Then .ListIndex = i
        Next i
    End With
End Sub

'
Private Sub ShowPaperOrientation()
    Dim i As Integer
    Dim PaperOrientation As CRPaperOrientation
     If cboPaperOrientation.ListCount = 0 Then
    Addcbo cboPaperOrientation, "Portrait", crPortrait
    Addcbo cboPaperOrientation, "Landscape", crLandscape
       End If
       ' hard coded for the setting in which report was made but user can change it run time.
    With cboPaperOrientation
        For i = 0 To .ListCount - 1
            If .ItemData(i) = crLandscape Then .ListIndex = i
        Next i
    End With
End Sub

'
Private Sub ShowPaperSize()
    Dim i As Integer                            ' Counter
    Dim PaperSize As CRPaperSize
       If cboPaperSize.ListCount = 0 Then
    Addcbo cboPaperSize, "Default", crDefaultPaperSize
    Addcbo cboPaperSize, "Letter", crPaperLetter
    Addcbo cboPaperSize, "Small Letter", crPaperLetterSmall
    Addcbo cboPaperSize, "Legal", crPaperLegal
    Addcbo cboPaperSize, "10x14", crPaper10x14
    Addcbo cboPaperSize, "11x17", crPaper11x17
    Addcbo cboPaperSize, "A3", crPaperA3
    Addcbo cboPaperSize, "A4", crPaperA4
    Addcbo cboPaperSize, "A4 Small", crPaperA4Small
    Addcbo cboPaperSize, "A5", crPaperA5
    Addcbo cboPaperSize, "B4", crPaperB4
    Addcbo cboPaperSize, "B5", crPaperB5
    Addcbo cboPaperSize, "C Sheet", crPaperCsheet
    Addcbo cboPaperSize, "D Sheet", crPaperDsheet
    Addcbo cboPaperSize, "Envelope 9", crPaperEnvelope9
    Addcbo cboPaperSize, "Envelope 10", crPaperEnvelope10
    Addcbo cboPaperSize, "Envelope 11", crPaperEnvelope11
    Addcbo cboPaperSize, "Envelope 12", crPaperEnvelope12
    Addcbo cboPaperSize, "Envelope 14", crPaperEnvelope14
    Addcbo cboPaperSize, "Envelope B4", crPaperEnvelopeB4
    Addcbo cboPaperSize, "Envelope B5", crPaperEnvelopeB5
    Addcbo cboPaperSize, "Envelope B6", crPaperEnvelopeB6
    Addcbo cboPaperSize, "Envelope C3", crPaperEnvelopeC3
    Addcbo cboPaperSize, "Envelope C4", crPaperEnvelopeC4
    Addcbo cboPaperSize, "Envelope C5", crPaperEnvelopeC5
    Addcbo cboPaperSize, "Envelope C6", crPaperEnvelopeC6
    Addcbo cboPaperSize, "Envelope C65", crPaperEnvelopeC65
    Addcbo cboPaperSize, "Envelope DL", crPaperEnvelopeDL
    Addcbo cboPaperSize, "Envelope Italy", crPaperEnvelopeItaly
    Addcbo cboPaperSize, "Envelope Monarch", crPaperEnvelopeMonarch
    Addcbo cboPaperSize, "Envelope Personal", crPaperEnvelopePersonal
    Addcbo cboPaperSize, "E Sheet", crPaperEsheet
    Addcbo cboPaperSize, "Executive", crPaperExecutive
    Addcbo cboPaperSize, "Fanfold Legal German", crPaperFanfoldLegalGerman
    Addcbo cboPaperSize, "Fanfold Standard German", crPaperFanfoldStdGerman
    Addcbo cboPaperSize, "Fanfold US", crPaperFanfoldUS
    Addcbo cboPaperSize, "FanFold 8.5 * 12", 119
    Addcbo cboPaperSize, "Folio", crPaperFolio
    Addcbo cboPaperSize, "Ledger", crPaperLedger
    Addcbo cboPaperSize, "Note", crPaperNote
    Addcbo cboPaperSize, "Quarto", crPaperQuarto
    Addcbo cboPaperSize, "Statement", crPaperStatement
    Addcbo cboPaperSize, "Tabloid", crPaperTabloid
    End If
    ' hard coded for the setting in which report was made but user can change it run time.
    With cboPaperSize
        For i = 0 To .ListCount - 1
            If .ItemData(i) = crPaperA4 Then .ListIndex = i
        Next i
    End With
End Sub

' *************************************************************
' A small helper function for the ShowPrinterOption functions that
' helps reduce the amount of code to write
'   Addcbo format:   <combo name to add item to>, <item caption>, <.itemdata(.listindex) to assign>
Private Sub Addcbo(cbo As ComboBox, name As String, Index As Integer)
    cbo.AddItem name                        ' Add the name of the item to the combo box
    cbo.ItemData(cbo.NewIndex) = Index      ' Set the .itemdata(.listindex) for later retrieval
End Sub
'******************************************************finished here***************
'


Private Sub ClearGroups()
Dim i As Integer
For i = 1 To 7
optGrp(i).Value = False
Next
End Sub
Public Function peWrkhrs()
On Error GoTo Err
Dim rptRs As New ADODB.Recordset
Dim totwrkhrs As Double
Dim dtfromdate As Date
Dim dttodate As Date
Dim strTrnFile, strTrnFile2 As String
Dim ECode As Variant
Dim EName As String

peWrkhrs = False
dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)

strTrnFile = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strTrnFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")
        
empstr3 = "SELECT empmst.empcode," & strTrnFile & ".wrkhrs,empmst.Name,empmst." & strKGroup & "," & rpgroup & _
        " FROM " & strTrnFile & "," & rpTables & " WHERE " & _
        strTrnFile & ".Empcode=empmst.Empcode " & strSql & " and " & strKDate & " >=  '" & Format(dtfromdate, "dd-mmm-yy") & "'" & _
        "Union " & _
        "SELECT empmst.empcode," & strTrnFile2 & ".wrkhrs,empmst.Name,empmst." & strKGroup & "," & rpgroup & _
        " FROM " & strTrnFile2 & "," & rpTables & " WHERE " & _
        strTrnFile2 & ".Empcode=empmst.Empcode " & strSql & " and " & strKDate & " <='" & Format(dttodate, "dd-mmm-yy") & "'  order by empmst.empcode"
        
ConMain.Execute "Truncate table RptWrkHrs"

If rptRs.State = 1 Then rptRs.Close
rptRs.Open empstr3, ConMain, adOpenForwardOnly, adLockReadOnly

Do While Not rptRs.EOF
ECode = rptRs.Fields("empcode")

    totwrkhrs = TimAdd(IIf(IsNull(totwrkhrs), 0, totwrkhrs), IIf(IsNull(rptRs.Fields("wrkHrs")), 0, rptRs.Fields("wrkHrs")))
    rptRs.MoveNext
    totwrkhrs = Round(totwrkhrs, 2)
    If ECode <> rptRs.Fields("empcode") Then
        ConMain.Execute "Insert into RptWrkHrs values('" & ECode & "'," & totwrkhrs & ")"
    End If
Loop
Exit Function
Err:
    ConMain.Execute "Insert into RptWrkHrs values('" & ECode & "'," & totwrkhrs & ")"
    peWrkhrs = True
End Function
'add by
Public Function pePerfAbsent() As Boolean
10    On Error GoTo ERR_P
20    pePerfAbsent = True
      Dim p_str As String
      Dim strGP As String
      Dim dtfromdate As Date, dttodate As Date, dtFirstDate As Date
      Dim DTESTR As String, STRECODE As String
      Dim strfile1 As String, strFile2 As String
      Dim dblTotalAbsent As Double
30    dtFirstDate = DateCompDate(typRep.strPeriFr)
40    dtfromdate = DateCompDate(typRep.strPeriFr)
50    dttodate = DateCompDate(typRep.strPeriTo)
60    strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
70    strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")
80    DTESTR = ""

90    If strfile1 = strFile2 Then
100       strGP = "select " & strfile1 & ".Empcode," & strKDate & ", " & _
          "presabs from " & strfile1 & "," & rpTables & " where " & _
          strfile1 & ".Empcode = empmst.Empcode  and " & strKDate & ">=" & strDTEnc & DateCompDate(typRep.strPeriFr) & _
          strDTEnc & " and " & strKDate & "<=" & strDTEnc & Format(DateCompDate(typRep.strPeriTo), "DD/MMM/YYYY") & strDTEnc & " " & strSql
110   Else
120       strGP = "select " & strfile1 & ".Empcode," & strKDate & "," & _
          "presabs from " & strfile1 & "," & rpTables & " where " & _
          strfile1 & ".Empcode = empmst.Empcode and  " & strKDate & ">=" & strDTEnc & _
          Format(DateCompDate(typRep.strPeriFr), "DD/MMM/YYYY") & _
          strDTEnc & strSql & " union select " & strFile2 & ".Empcode," & strKDate & "," & _
          "presabs from " & _
          strFile2 & "," & rpTables & " where " & strFile2 & ".Empcode = empmst.Empcode " & _
          " and " & strKDate & "<=" & strDTEnc & Format(DateCompDate(typRep.strPeriTo), "DD/MMM/YYYY") & strDTEnc & " " & strSql
130   End If
140   Select Case bytBackEnd
          Case 1, 2 ''SQLServer,MS-Access
150           strGP = strGP & " order by " & strfile1 & ".Empcode," & strfile1 & "." & strKDate & ""
160       Case 3    '' ORACLE
170           strGP = strGP & " order by Empcode," & strKDate
180   End Select
190   dtfromdate = dtFirstDate
200   If adrsTemp.State = 1 Then adrsTemp.Close
210   adrsTemp.Open strGP, ConMain, adOpenStatic
220   If Not (adrsTemp.BOF And adrsTemp.EOF) Then    'adrsTemp
230       If Not (IsNull(adrsTemp(0)) Or IsEmpty(adrsTemp(0))) Then
240           adrsTemp.MoveFirst
250           Do While Not (adrsTemp.EOF) And dtfromdate <= dttodate
260               STRECODE = adrsTemp!Empcode
270               dtfromdate = dtFirstDate
280               p_str = ""
290               DTESTR = ""
300               Do While dtfromdate <= dttodate
310                   If adrsTemp.EOF Then Exit Do
320                   If adrsTemp!Empcode = STRECODE Then
330                       If adrsTemp!Date = dtfromdate Then
340                           If IIf(Not IsNull(adrsTemp!presabs), _
                                  adrsTemp!presabs, "") = pVStar.AbsCode & _
                                  pVStar.AbsCode Then
350                                 If Len(Day(dtfromdate)) = 1 Then
360                                     DTESTR = DTESTR & Day(dtfromdate) & Space(5)
370                                     p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Space(2)
380                                 Else
390                                     DTESTR = DTESTR & Day(dtfromdate) & Space(4)
400                                     p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Space(2)
410                                 End If
420                                 dblTotalAbsent = dblTotalAbsent + 1
430                           ElseIf Left(IIf(Not IsNull(adrsTemp!presabs), _
                                  adrsTemp!presabs, ""), 2) = pVStar.AbsCode Then
440                                 If Len(Day(dtfromdate)) = 1 Then
450                                     DTESTR = DTESTR & Day(dtfromdate) & Space(5)
460                                     p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Space(2)
470                                 Else
480                                     DTESTR = DTESTR & Day(dtfromdate) & Space(4)
490                                     p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Space(2)
500                                 End If
510                                 dblTotalAbsent = dblTotalAbsent + 0.5
520                           ElseIf Right(IIf(Not IsNull(adrsTemp!presabs), _
                                  adrsTemp!presabs, ""), 2) = pVStar.AbsCode Then
530                                 If Len(Day(dtfromdate)) = 1 Then
540                                     DTESTR = DTESTR & Day(dtfromdate) & Space(5)
550                                     p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Space(2)
560                                 Else
570                                     DTESTR = DTESTR & Day(dtfromdate) & Space(4)
580                                     p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Space(2)
590                                 End If
600                                 dblTotalAbsent = dblTotalAbsent + 0.5
610                           End If
620                           ElseIf adrsTemp!Date <> dtfromdate Then
630                           p_str = p_str & Space(6)
640                           DTESTR = DTESTR & Spaces(6)
650                       End If
660                   Else
670                       Exit Do
680                   End If
690                   If dtfromdate = adrsTemp!Date Then adrsTemp.MoveNext
700                   dtfromdate = DateAdd("d", 1, dtfromdate)
710               Loop 'END OF DATE LOOP
'570               p_str = p_str & "Total number of absent:" & dblTotalAbsent
720               If Trim(p_str) = "" Then
730               Else
                      Debug.Print DTESTR
                      Debug.Print p_str
                       'here sumOt column is use for total number of absent
740                   ConMain.Execute "insert into " & strRepFile & "" & _
                      "(Empcode," & strKDate & _
                      ",presabsstr,sumwork)  values('" & STRECODE & _
                      "','" & DTESTR & "','" & p_str & "'," & dblTotalAbsent & ")"
750               End If
760               dblTotalAbsent = 0
770               dtfromdate = dtFirstDate
780           Loop 'END OF EMPLOYEE LOOP
790       End If
800   Else
810       Call SetMSF1Cap(10)
820       MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
830       pePerfAbsent = False
840   End If 'adrsTemp.eof
850   Exit Function
ERR_P:
860       ShowError ("Periodic Absent :: Reports Line:" & Erl & "And Empcode:" & STRECODE)
870       pePerfAbsent = False
End Function
'add by
Public Function pePerfPresent() As Boolean
10    On Error GoTo ERR_P
20    pePerfPresent = True
      Dim p_str As String
      Dim strGP As String
      Dim dtfromdate As Date, dttodate As Date, dtFirstDate As Date
      Dim DTESTR As String, STRECODE As String
      Dim strfile1 As String, strFile2 As String
      Dim dblTotalAbsent As Double
30    dtFirstDate = DateCompDate(typRep.strPeriFr)
40    dtfromdate = DateCompDate(typRep.strPeriFr)
50    dttodate = DateCompDate(typRep.strPeriTo)
60    strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
70    strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")
80    DTESTR = ""

90    If strfile1 = strFile2 Then
100       strGP = "select " & strfile1 & ".Empcode," & strKDate & ", " & _
          "presabs from " & strfile1 & "," & rpTables & " where " & _
          strfile1 & ".Empcode = empmst.Empcode  and " & strKDate & ">=" & strDTEnc & DateCompDate(typRep.strPeriFr) & _
          strDTEnc & " and " & strKDate & "<=" & strDTEnc & Format(DateCompDate(typRep.strPeriTo), "DD/MMM/YYYY") & strDTEnc & " " & strSql
110   Else
120       strGP = "select " & strfile1 & ".Empcode," & strKDate & "," & _
          "presabs from " & strfile1 & "," & rpTables & " where " & _
          strfile1 & ".Empcode = empmst.Empcode and  " & strKDate & ">=" & strDTEnc & _
          Format(DateCompDate(typRep.strPeriFr), "DD/MMM/YYYY") & _
          strDTEnc & strSql & " union select " & strFile2 & ".Empcode," & strKDate & "," & _
          "presabs from " & _
          strFile2 & "," & rpTables & " where " & strFile2 & ".Empcode = empmst.Empcode " & _
          " and " & strKDate & "<=" & strDTEnc & Format(DateCompDate(typRep.strPeriTo), "DD/MMM/YYYY") & strDTEnc & " " & strSql
130   End If
140   Select Case bytBackEnd
          Case 1, 2 ''SQLServer,MS-Access
150           strGP = strGP & " order by " & strfile1 & ".Empcode," & strfile1 & "." & strKDate & ""
160       Case 3    '' ORACLE
170           strGP = strGP & " order by Empcode," & strKDate
180   End Select
190   dtfromdate = dtFirstDate
200   If adrsTemp.State = 1 Then adrsTemp.Close
210   adrsTemp.Open strGP, ConMain, adOpenStatic
220   If Not (adrsTemp.BOF And adrsTemp.EOF) Then    'adrsTemp
230       If Not (IsNull(adrsTemp(0)) Or IsEmpty(adrsTemp(0))) Then
240           adrsTemp.MoveFirst
250           Do While Not (adrsTemp.EOF) And dtfromdate <= dttodate
260               STRECODE = adrsTemp!Empcode
270               dtfromdate = dtFirstDate
280               p_str = ""
290               DTESTR = ""
300               Do While dtfromdate <= dttodate
310                   If adrsTemp.EOF Then Exit Do
320                   If adrsTemp!Empcode = STRECODE Then
330                       If adrsTemp!Date = dtfromdate Then
340                           If IIf(Not IsNull(adrsTemp!presabs), _
                                  adrsTemp!presabs, "") = pVStar.PrsCode & _
                                  pVStar.PrsCode Then
350                                 If Len(Day(dtfromdate)) = 1 Then
360                                     DTESTR = DTESTR & Day(dtfromdate) & Space(5)
370                                     p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Space(2)
380                                 Else
390                                     DTESTR = DTESTR & Day(dtfromdate) & Space(4)
400                                     p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Space(2)
410                                 End If
420                                 dblTotalAbsent = dblTotalAbsent + 1
430                           ElseIf Left(IIf(Not IsNull(adrsTemp!presabs), _
                                  adrsTemp!presabs, ""), 2) = pVStar.PrsCode Then
440                                 If Len(Day(dtfromdate)) = 1 Then
450                                     DTESTR = DTESTR & Day(dtfromdate) & Space(5)
460                                     p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Space(2)
470                                 Else
480                                     DTESTR = DTESTR & Day(dtfromdate) & Space(4)
490                                     p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Space(2)
500                                 End If
510                                 dblTotalAbsent = dblTotalAbsent + 0.5
520                           ElseIf Right(IIf(Not IsNull(adrsTemp!presabs), _
                                  adrsTemp!presabs, ""), 2) = pVStar.PrsCode Then
530                                 If Len(Day(dtfromdate)) = 1 Then
540                                     DTESTR = DTESTR & Day(dtfromdate) & Space(5)
550                                     p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Space(2)
560                                 Else
570                                     DTESTR = DTESTR & Day(dtfromdate) & Space(4)
580                                     p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Space(2)
590                                 End If
600                                 dblTotalAbsent = dblTotalAbsent + 0.5
610                           End If
620                           ElseIf adrsTemp!Date <> dtfromdate Then
630                           p_str = p_str & Space(6)
640                           DTESTR = DTESTR & Spaces(6)
650                       End If
660                   Else
670                       Exit Do
680                   End If
690                   If dtfromdate = adrsTemp!Date Then adrsTemp.MoveNext
700                   dtfromdate = DateAdd("d", 1, dtfromdate)
710               Loop 'END OF DATE LOOP
'570               p_str = p_str & "Total number of absent:" & dblTotalAbsent
720               If Trim(p_str) = "" Then
730               Else
                      Debug.Print DTESTR
                      Debug.Print p_str
                       'here sumOt column is use for total number of absent
740                   ConMain.Execute "insert into " & strRepFile & "" & _
                      "(Empcode," & strKDate & _
                      ",presabsstr,sumwork)  values('" & STRECODE & _
                      "','" & DTESTR & "','" & p_str & "'," & dblTotalAbsent & ")"
750               End If
760               dblTotalAbsent = 0
770               dtfromdate = dtFirstDate
780           Loop 'END OF EMPLOYEE LOOP
790       End If
800   Else
810       Call SetMSF1Cap(10)
820       MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
830       pePerfPresent = False
840   End If 'adrsTemp.eof
850   Exit Function
ERR_P:
860       ShowError ("Periodic Absent :: Reports Line:" & Erl & "And Empcode:" & STRECODE)
870       pePerfPresent = False
End Function
'add by
Public Function pePerfPhysicalAbsent() As Boolean
On Error GoTo ERR_P
pePerfPhysicalAbsent = True
Dim p_str As String
Dim strGP As String
Dim dtfromdate As Date, dttodate As Date, dtFirstDate As Date
Dim DTESTR As String, STRECODE As String
Dim strfile1 As String, strFile2 As String
Dim dblTotalAbsent As Double
dtFirstDate = DateCompDate(typRep.strPeriFr)
dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)
strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")
DTESTR = ""


    If strfile1 = strFile2 Then

            strGP = "select " & strfile1 & ".Empcode," & strKDate & ", " & _
            "presabs from " & strfile1 & "," & rpTables & " where (LEFT(presabs,2)<>'" & _
            pVStar.PrsCode & "' OR RIGHT(presabs,2)<>'" & pVStar.PrsCode & "') AND " & _
            strfile1 & ".Empcode = empmst.Empcode  and " & strKDate & ">=" & strDTEnc & DateCompDate(typRep.strPeriFr) & _
            strDTEnc & " and " & strKDate & "<=" & strDTEnc & Format(DateCompDate(typRep.strPeriTo), "DD/MMM/YYYY") & strDTEnc & " " & strSql
    Else

            strGP = "select " & strfile1 & ".Empcode," & strKDate & "," & _
            "presabs from " & strfile1 & "," & rpTables & " where (LEFT(presabs,2)<>'" & _
            pVStar.PrsCode & "' OR RIGHT(presabs,2)<>'" & pVStar.PrsCode & "') AND " & _
            strfile1 & ".Empcode = empmst.Empcode and  " & strKDate & ">=" & strDTEnc & _
            Format(DateCompDate(typRep.strPeriFr), "DD/MMM/YYYY") & _
            strDTEnc & strSql & " union select " & strFile2 & ".Empcode," & strKDate & "," & _
            "presabs from " & _
            strFile2 & "," & rpTables & " where (LEFT(presabs,2)<>'" & _
            pVStar.PrsCode & "' OR RIGHT(presabs,2)<>'" & pVStar.PrsCode & "') AND " & strFile2 & ".Empcode = empmst.Empcode " & _
            " and " & strKDate & "<=" & strDTEnc & Format(DateCompDate(typRep.strPeriTo), "DD/MMM/YYYY") & strDTEnc & " " & strSql
    End If



Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        strGP = strGP & " order by " & strfile1 & ".Empcode," & strfile1 & "." & strKDate & ""
    Case 3    '' ORACLE
        strGP = strGP & " order by Empcode," & strKDate
End Select
dtfromdate = dtFirstDate
DTESTR = ""
Do While dtfromdate <= dttodate
    If Len(Day(dtfromdate)) = 1 Then
        DTESTR = DTESTR & Day(dtfromdate) & Space(5)
    Else
        DTESTR = DTESTR & Day(dtfromdate) & Space(4)
    End If
    dtfromdate = DateAdd("d", 1, dtfromdate)
Loop
dtfromdate = dtFirstDate
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strGP, ConMain, adOpenStatic
If Not (adrsTemp.BOF And adrsTemp.EOF) Then    'adrsTemp
    If Not (IsNull(adrsTemp(0)) Or IsEmpty(adrsTemp(0))) Then
        adrsTemp.MoveFirst
        Do While Not (adrsTemp.EOF) And dtfromdate <= dttodate
            STRECODE = adrsTemp!Empcode
            dtfromdate = dtFirstDate
            p_str = ""
            'DTESTR = ""
            Do While dtfromdate <= dttodate
                If adrsTemp.EOF Then Exit Do
                If adrsTemp!Empcode = STRECODE Then
                    If adrsTemp!Date = dtfromdate Then
                        If Len(Day(dtfromdate)) = 1 Then
                            'DTESTR = DTESTR & Day(dtfromdate) & Space(5)
                            p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Space(2)
                        Else
                            'DTESTR = DTESTR & Day(dtfromdate) & Space(4)
                            p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Space(2)
                        End If
                    ElseIf adrsTemp!Date <> dtfromdate Then
                        p_str = p_str & Space(6)
                        'DTESTR = DTESTR & Spaces(6)
                    End If
                Else
                    Exit Do
                End If
                If dtfromdate = adrsTemp!Date Then adrsTemp.MoveNext
                dtfromdate = DateAdd("d", 1, dtfromdate)
            Loop 'END OF DATE LOOP
'570               p_str = p_str & "Total number of absent:" & dblTotalAbsent
            If Trim(p_str) = "" Then
            Else
                Debug.Print DTESTR
                Debug.Print p_str
                 'here sumOt column is use for total number of absent
                ConMain.Execute "insert into " & strRepFile & "" & _
                "(Empcode," & strKDate & _
                ",presabsstr)  values('" & STRECODE & _
                "','" & DTESTR & "','" & p_str & "')"
            End If
            dblTotalAbsent = 0
            dtfromdate = dtFirstDate
        Loop 'END OF EMPLOYEE LOOP
    End If
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
    pePerfPhysicalAbsent = False
End If 'adrsTemp.eof
Exit Function
ERR_P:
670       ShowError ("Periodic Absent :: Reports Line:" & Erl & "And Empcode:" & STRECODE)
680       pePerfPhysicalAbsent = False
End Function

Private Sub FillEmpCombo(bytI As Byte)
    Call ComboFill(cmbFrEmpSel, bytI, 2)       'emp
    cmbToEmpSel.List = cmbFrEmpSel.List
    If cmbFrEmpSel.ListCount <> 0 Then
        cmbFrEmpSel.Text = cmbFrEmpSel.List(0)
        cmbToEmpSel.Text = cmbFrEmpSel.List(cmbFrEmpSel.ListCount - 1)
    End If
End Sub

