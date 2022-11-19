VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLeaves 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   6480
      TabIndex        =   28
      Top             =   6000
      Width           =   2235
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4320
      TabIndex        =   27
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2160
      TabIndex        =   26
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command1"
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   6000
      Width           =   2175
   End
   Begin TabDlg.SSTab TB1 
      Height          =   5970
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   10530
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
      TabPicture(0)   =   "frmLeaves.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSF1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Details"
      TabPicture(1)   =   "frmLeaves.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frDetails"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cboCatCode"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.TextBox cboCatCode 
         Height          =   405
         Left            =   -67320
         TabIndex        =   62
         Text            =   "code"
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame frDetails 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5550
         Left            =   -74925
         TabIndex        =   31
         Top             =   330
         Width           =   8580
         Begin VB.Frame FrBusnessRule 
            Caption         =   "Business Rule"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   2790
            TabIndex        =   61
            Top             =   4395
            Width           =   1935
            Begin VB.OptionButton optBRStrick 
               Caption         =   "Strictly"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   135
               TabIndex        =   11
               Top             =   315
               Width           =   1230
            End
            Begin VB.OptionButton optBRFlexi 
               Caption         =   "Flexible"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   135
               TabIndex        =   12
               Top             =   675
               Width           =   1320
            End
         End
         Begin VB.CheckBox chkMonthly 
            Caption         =   "Monthly Basis"
            Height          =   495
            Left            =   3510
            TabIndex        =   4
            Top             =   1485
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame frwhile 
            Caption         =   "For yearly leave update"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   90
            TabIndex        =   53
            Top             =   4395
            Width           =   2655
            Begin VB.OptionButton optProp 
               Caption         =   "Credit in Proportion"
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
               TabIndex        =   10
               Top             =   630
               Width           =   2040
            End
            Begin VB.OptionButton optFull 
               Caption         =   "Credit Leaves full"
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
               TabIndex        =   9
               Top             =   270
               Width           =   2385
            End
         End
         Begin VB.Frame frMisc 
            Caption         =   "Leave"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   45
            TabIndex        =   32
            Top             =   180
            Width           =   5280
            Begin MSMask.MaskEdBox txtName 
               Height          =   420
               Left            =   2955
               TabIndex        =   1
               Top             =   240
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   741
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   19
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
               Height          =   390
               Left            =   720
               TabIndex        =   0
               Top             =   240
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   688
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   2
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
            Begin VB.Label lblCode 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Code"
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
               TabIndex        =   33
               Top             =   300
               Width           =   450
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
               Left            =   1935
               TabIndex        =   34
               Top             =   330
               Width           =   510
            End
         End
         Begin VB.Frame frDef 
            Caption         =   "Definition"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   60
            TabIndex        =   37
            Top             =   990
            Width           =   3330
            Begin VB.CheckBox chkPayable 
               Caption         =   "Count this leave in payable days"
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
               Left            =   120
               TabIndex        =   2
               Top             =   285
               Width           =   3120
            End
            Begin VB.CheckBox chkBal 
               Caption         =   "Keep balance for this leave"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   120
               TabIndex        =   3
               Top             =   525
               Width           =   3105
            End
         End
         Begin VB.Frame frNewEmp 
            Caption         =   "Crediting for new Employees"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   45
            TabIndex        =   45
            Top             =   3390
            Width           =   4665
            Begin VB.OptionButton optImd 
               Caption         =   "Credit Immediately"
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
               Left            =   210
               TabIndex        =   7
               Top             =   285
               Width           =   1980
            End
            Begin VB.OptionButton optNext 
               Caption         =   "Credit next year"
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
               Left            =   180
               TabIndex        =   8
               Top             =   585
               Width           =   1725
            End
         End
         Begin VB.Frame frMark 
            Caption         =   "Mark Leaves"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1260
            Left            =   5055
            TabIndex        =   43
            Top             =   1200
            Width           =   3495
            Begin VB.OptionButton optIncl 
               Caption         =   "Including WeekOff / Holidays"
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
               Left            =   120
               TabIndex        =   15
               Top             =   330
               Width           =   3120
            End
            Begin VB.OptionButton optExcl 
               Caption         =   "Excluding WeekOff / Holidays"
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
               Top             =   630
               Width           =   3345
            End
            Begin VB.OptionButton optDecide 
               Caption         =   "Decide while entering leave"
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
               Left            =   135
               TabIndex        =   17
               Top             =   900
               Width           =   2805
            End
         End
         Begin VB.Frame frEnd 
            Caption         =   "At the End of the year"
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
            Left            =   5040
            TabIndex        =   44
            Top             =   2520
            Width           =   3480
            Begin VB.CheckBox chkCarry 
               Caption         =   "Carry Forward Balance Leaves"
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
               TabIndex        =   18
               Top             =   270
               Width           =   3435
            End
            Begin VB.CheckBox chkEncash 
               Caption         =   "Encash Balance Leaves"
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
               TabIndex        =   19
               Top             =   525
               Width           =   3375
            End
         End
         Begin VB.Frame frRules 
            Caption         =   "Check following rules"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1995
            Left            =   4830
            TabIndex        =   46
            Top             =   3465
            Width           =   3705
            Begin VB.OptionButton OptFlexible 
               Caption         =   "Flexible"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1125
               TabIndex        =   21
               Top             =   360
               Visible         =   0   'False
               Width           =   1320
            End
            Begin VB.OptionButton OptStrictly 
               Caption         =   "Strictly"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   135
               TabIndex        =   20
               Top             =   360
               Visible         =   0   'False
               Width           =   1230
            End
            Begin MSMask.MaskEdBox txtAllow 
               Height          =   375
               Left            =   1110
               TabIndex        =   22
               Top             =   675
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   661
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               HideSelection   =   0   'False
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
            Begin MSMask.MaskEdBox txtMin 
               Height          =   375
               Left            =   1110
               TabIndex        =   24
               Top             =   1395
               Width           =   795
               _ExtentX        =   1402
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
               Format          =   "0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtMax 
               Height          =   375
               Left            =   1110
               TabIndex        =   23
               Top             =   1035
               Width           =   795
               _ExtentX        =   1402
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
               Format          =   "0.00"
               PromptChar      =   " "
            End
            Begin VB.Label lblAllow 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Allow"
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
               TabIndex        =   47
               Top             =   675
               Width           =   465
            End
            Begin VB.Label lblMax 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Maximum"
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
               TabIndex        =   49
               Top             =   1065
               Width           =   855
            End
            Begin VB.Label lblMin 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Minimum"
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
               TabIndex        =   51
               Top             =   1425
               Width           =   795
            End
            Begin VB.Label lblAllowTimes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Times in a Year"
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
               Left            =   2040
               TabIndex        =   48
               Top             =   675
               Width           =   1365
            End
            Begin VB.Label lblMaxdays 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "days at a time"
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
               Left            =   2040
               TabIndex        =   50
               Top             =   1050
               Width           =   1245
            End
            Begin VB.Label lblMinDays 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "days at a time"
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
               Left            =   2040
               TabIndex        =   52
               Top             =   1425
               Width           =   1245
            End
         End
         Begin VB.Frame frCat 
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
            Height          =   975
            Left            =   5355
            TabIndex        =   35
            Top             =   165
            Width           =   3180
            Begin VB.CheckBox chkCat 
               Caption         =   "All Categories"
               Height          =   195
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label lblCat 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Specific"
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
               Left            =   105
               TabIndex        =   36
               Top             =   525
               Width           =   690
            End
            Begin MSForms.ComboBox cboCat 
               Height          =   375
               Left            =   960
               TabIndex        =   14
               Top             =   480
               Width           =   1935
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3413;661"
               MatchEntry      =   0
               ShowDropButtonWhen=   2
               Value           =   " "
               SpecialEffect   =   0
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.Frame frCYear 
            Caption         =   "For the Current Year"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1410
            Left            =   30
            TabIndex        =   38
            Top             =   1950
            Width           =   4710
            Begin MSMask.MaskEdBox txtAllowMax 
               Height          =   345
               Left            =   1830
               TabIndex        =   6
               Top             =   750
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtCredit 
               Height          =   345
               Left            =   1830
               TabIndex        =   5
               Top             =   300
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "0.00"
               PromptChar      =   " "
            End
            Begin VB.Label lblCredit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Credit"
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
               Left            =   180
               TabIndex        =   39
               Top             =   270
               Width           =   510
            End
            Begin VB.Label lblAllowMax 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Allow Maximium"
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
               TabIndex        =   41
               Top             =   810
               Width           =   1425
            End
            Begin VB.Label lblAllowMaxdays 
               BackStyle       =   0  'Transparent
               Caption         =   "days leave "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   660
               Left            =   2820
               TabIndex        =   42
               Top             =   750
               Width           =   1815
            End
            Begin VB.Label lblCreditDays 
               BackStyle       =   0  'Transparent
               Caption         =   "days leaves for    consumption "
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
               Left            =   2850
               TabIndex        =   40
               Top             =   240
               Width           =   1770
            End
         End
         Begin VB.Frame FraCMonth 
            Caption         =   "For the Current Month"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1290
            Left            =   30
            TabIndex        =   54
            Top             =   1950
            Width           =   4710
            Begin MSMask.MaskEdBox txtAllowMaxM 
               Height          =   345
               Left            =   1830
               TabIndex        =   56
               Top             =   750
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtCreditM 
               Height          =   345
               Left            =   1830
               TabIndex        =   55
               Top             =   300
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "0.00"
               PromptChar      =   " "
            End
            Begin VB.Label lblCons 
               BackStyle       =   0  'Transparent
               Caption         =   "days leaves for    consumption "
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
               Left            =   2850
               TabIndex        =   60
               Top             =   240
               Width           =   1770
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "days leave to be Accumulated"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   2820
               TabIndex        =   59
               Top             =   750
               Width           =   1815
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Allow Maximium"
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
               TabIndex        =   58
               Top             =   810
               Width           =   1425
            End
            Begin VB.Label lblcreditM 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Credit"
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
               Left            =   180
               TabIndex        =   57
               Top             =   270
               Width           =   510
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   5055
         Left            =   450
         TabIndex        =   30
         Top             =   390
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   8916
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   12632256
         AllowBigSelection=   0   'False
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
Attribute VB_Name = "frmLeaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset
''
Dim adrslv As New ADODB.Recordset
Dim ans

Private Sub cboCat_Click()
    If cboCat.ListIndex = -1 Then Exit Sub
    cboCatCode.Text = cboCat.List(cboCat.ListIndex, 1)
End Sub

Private Sub chkCat_Click()
If chkCat.Value = 1 Then
    cboCat.Value = ""
    cboCat.Enabled = False
Else
    cboCat.Enabled = True
End If
If bytMode = 2 Or bytMode = 3 Then SendKeys Chr(9)
End Sub

Private Sub chkBal_Click()
    Call AdjustActions
End Sub

Private Sub chkMonthly_Click()
    If chkMonthly.Value = 1 Then
        FraCMonth.Visible = True
        frCYear.Visible = False
    Else
        FraCMonth.Visible = False
        frCYear.Visible = True
    End If
End Sub

Private Sub chkPayable_Click()
    Call AdjustActions
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    OptStrictly.Visible = False
    OptFlexible.Visible = False
    FrBusnessRule.Visible = False
Call SetFormIcon(Me)        '' Set the Form Icon
Call RetCaptions            '' Retreive Captions
Call OpenMasterTable        '' Open Master Table
Call FillGrid               '' Fill Grid
Call FillCatCombo           '' Fill Category ComboBox
FraCMonth.Visible = False   ' 15-06
TB1.Tab = 0                 '' Set the Tab to List
Call GetRights              '' Gets Rights for the Operations
bytMode = 1
Call ChangeMode             '' Take Action on the Appropriate Mode
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '31%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("31001", adrsC)              '' Form caption
TB1.TabCaption(0) = NewCaptionTxt("00013", adrsMod)       '' List
TB1.TabCaption(1) = NewCaptionTxt("00014", adrsMod)       '' Details
Call SetOtherCaps                           '' Set Captions to the Other Controls
Call SetGButtonCap(Me)                           '' Sets Appropriate Caption for the Buttons
Call CapGrid                                '' Captions For Grid
End Sub

Private Sub SetOtherCaps()
'' Frame Misc
frMisc.Caption = NewCaptionTxt("00063", adrsMod)          '' Leave
lblCode.Caption = NewCaptionTxt("00047", adrsMod)         '' Code
lblName.Caption = NewCaptionTxt("00048", adrsMod)         '' Name
'' Frame Cat
frCat.Caption = NewCaptionTxt("00051", adrsMod)           '' Category
chkCat.Caption = NewCaptionTxt("31020", adrsC)          '' All Categories
lblCat.Caption = NewCaptionTxt("31021", adrsC)          '' Specific
'' Frame Definition
frDef.Caption = NewCaptionTxt("31006", adrsC)           '' Definition
chkPayable.Caption = NewCaptionTxt("31007", adrsC)      '' Count This Leave...
chkBal.Caption = NewCaptionTxt("31008", adrsC)          '' Keep Balance ...
'' Frame Current Year
frCYear.Caption = NewCaptionTxt("31009", adrsC)         '' For the Current Year
lblCredit.Caption = NewCaptionTxt("31010", adrsC)       '' Credit
lblCreditDays.Caption = NewCaptionTxt("31011", adrsC)   '' Days Leaves for...
lblAllowMax.Caption = NewCaptionTxt("31012", adrsC)     '' Allow Maximum
lblAllowMaxdays.Caption = NewCaptionTxt("31013", adrsC) '' Days Leaves to ...
'' Frame Mark Leaves
frMark.Caption = NewCaptionTxt("31022", adrsC)          '' Mark Leaves
optIncl.Caption = NewCaptionTxt("31023", adrsC)         '' Including Week Off...
optExcl.Caption = NewCaptionTxt("31024", adrsC)         '' Excluding WeekOff
optDecide.Caption = NewCaptionTxt("31025", adrsC)       '' Decide While ...
'' Frame End of the Year
frEnd.Caption = NewCaptionTxt("31026", adrsC)           '' At the End ...
chkCarry.Caption = NewCaptionTxt("31027", adrsC)        '' Carry Forward
chkEncash.Caption = NewCaptionTxt("31028", adrsC)       '' Encassh Balance
'' Frame Rules
frRules.Caption = NewCaptionTxt("31029", adrsC)         '' Check Following Rules
lblAllow.Caption = NewCaptionTxt("31030", adrsC)        '' Allow
lblAllowTimes.Caption = NewCaptionTxt("31031", adrsC)   '' Times in a ...
lblMax.Caption = NewCaptionTxt("31032", adrsC)          '' Maximum
lblMaxdays.Caption = NewCaptionTxt("31033", adrsC)      '' Days at a ...
lblMin.Caption = NewCaptionTxt("31034", adrsC)          '' Minimum
lblMinDays.Caption = NewCaptionTxt("31033", adrsC)      '' Days at a ...
'' Frame New Employee
frNewEmp.Caption = NewCaptionTxt("31014", adrsC)        '' Crediting for New ...
optImd.Caption = NewCaptionTxt("31015", adrsC)          '' Credit Immidiately
optNext.Caption = NewCaptionTxt("31016", adrsC)         '' Credit Next Year
'' Frame While Crediting

frwhile.Caption = "For yearly leave update" ''NewCaptionTxt("31017", adrsC)       '' While Crediting
optFull.Caption = NewCaptionTxt("31018", adrsC)         '' Credit leaves Full
optProp.Caption = NewCaptionTxt("31019", adrsC)         '' Credit in Proportion
End Sub

Private Sub CapGrid()           '' Gives the Captions to the Grid
With MSF1
    '' Sets the Column Widhts
    .ColWidth(0) = .ColWidth(0) * 1.5
    .ColWidth(1) = .ColWidth(1) * 1.25
    .ColWidth(2) = .ColWidth(2) * 2.2
    .ColWidth(3) = .ColWidth(3) * 2
    '' Sets the Column Alignment
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    '' Sets the Appropriate Captions
    .TextMatrix(0, 0) = NewCaptionTxt("31003", adrsC)   '' Leave Code
    .TextMatrix(0, 1) = NewCaptionTxt("00051", adrsMod)   '' Category
    .TextMatrix(0, 2) = NewCaptionTxt("31004", adrsC)   '' Name of Leave
    .TextMatrix(0, 3) = NewCaptionTxt("31005", adrsC)   '' Leave balance
End With
End Sub

Private Sub OpenMasterTable()             '' Open the Recordset for the Display purposes
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select LeavDesc.*, CatDesc.[desc] From LeavDesc, catdesc Where LeavDesc.cat = catdesc.cat and IsItLeave='Y' and LvCode Not in('" & _
pVStar.AbsCode & "','" & pVStar.PrsCode & "','" & pVStar.WosCode & "','" & _
pVStar.HlsCode & "') Order by LvCode,catdesc.Cat", ConMain, adOpenStatic
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
        .TextMatrix(intCounter, 0) = adrsDept1("LvCode")
        .TextMatrix(intCounter, 1) = IIf(IsNull(adrsDept1("Cat")), "", adrsDept1("Cat"))
        .TextMatrix(intCounter, 2) = IIf(IsNull(adrsDept1("Leave")), "", adrsDept1("Leave"))
        .TextMatrix(intCounter, 3) = IIf(IsNull(adrsDept1("Type")), "", adrsDept1("Type"))
        .TextMatrix(intCounter, 4) = IIf(IsNull(adrsDept1("desc")), "", adrsDept1("desc"))
  
    End With
    adrsDept1.MoveNext
Next
MSF1.ColWidth(4) = 0
TB1.TabEnabled(1) = True        '' If Records Found Enables the TAB 1
TB1.Tab = 0
Exit Sub
ERR_P:
    'Resume Next
End Sub

Private Sub GetRights()     '' Gets and Sets the Appropriate Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 14)
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

Private Sub ViewAction()        '' Procedure for Viewing Mode
'' Enable Needed Controls
cmdAddSave.Enabled = True       '' Enable ADD/SAVE Button
cmdEditCan.Enabled = True       '' Enable Edit/Cancel Button
cmdDel.Enabled = True           '' Enable Disable Button
'' Disable Needed Controls
txtCode.Enabled = False         '' Disable Code TextBox
frMisc.Enabled = False          '' Diable Miscellaneous Frame
frCat.Enabled = False           '' Disable Category Frame
frDef.Enabled = False           '' Disable Definition Frame
frCYear.Enabled = False         '' Disable Current Year Frame
frMark.Enabled = False          '' Disable Mark Leaves Frame
frEnd.Enabled = False           '' Disable End of the Year Frame
frRules.Enabled = False         '' Disable Rules Frame
frNewEmp.Enabled = False        '' Disable New Employee Frame
FraCMonth.Enabled = False
chkMonthly.Enabled = False
FrBusnessRule.Enabled = False
'' Give Captions to the Needed Controls
Call SetGButtonCap(Me)
TB1.Tab = 0     '' Set the Tab to the First Tab
End Sub

Private Sub AddAction()     '' Procedure for Addition Mode
If TB1.TabEnabled(1) = False Then TB1.TabEnabled(1) = True
TB1.Tab = 1
'' Enable Necessary Controls
txtCode.Enabled = True          '' Enable Code TextBox
frMisc.Enabled = True           '' Enable Miscellaneous Frame
frCat.Enabled = True            '' Enable Category Frame
frDef.Enabled = True            '' Enable Definition Frame
frMark.Enabled = True           '' Enable Mark Leaves Frame
frEnd.Enabled = True            '' Enable End of the Year Frame
frRules.Enabled = True          '' Enable Rules Frame
frNewEmp.Enabled = True         '' Enable New Employee Frame
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Clear Necessary Controls
txtCode.Text = ""               '' Clear Code TextBox
txtName.Text = ""               '' Clear Name TextBox
cboCat.Value = ""               '' Clear Category ComboBox
txtCredit.Text = ""             '' Clear Frame Current Year / Credit TextBox
txtAllowMax.Text = ""           '' Clear Frame Current Year / Allow Maximum textBox
txtAllowMaxM = ""
txtAllow.Text = ""              '' Clear Frame Rules / Allow TextBox
txtCreditM.Text = ""
txtMax.Text = ""                '' Clear Frame Rules / Allow Maximum TextBox
txtMin.Text = ""                '' Clear Frame Rules / Allow Minimum TextBox
chkPayable.Value = 0            '' Reset Payable to 0
chkBal.Value = 0                '' Reset Balance to 0
chkMonthly.Value = 0            '' Reset monthly check to 0 '
chkMonthly.Enabled = False
optIncl.Value = True            '' Reset Including to TRUE
optImd.Value = True             '' Reset Immidiately to TRUE

FrBusnessRule.Enabled = True
txtCode.SetFocus                '' Set Focus to the Code
End Sub

Private Sub EditAction()    '' Procedure for Edit Mode
'' Enable Necessary Controls
txtCode.Enabled = False         '' Enable Code TextBox
frMisc.Enabled = True           '' Enable Miscellaneous Frame
frDef.Enabled = True            '' Enable Definition Frame
frMark.Enabled = True           '' Enable Mark Leaves Frame
frEnd.Enabled = True            '' Enable End of the Year Frame
frRules.Enabled = True          '' Enable Rules Frame
frNewEmp.Enabled = True         '' Enable New Employee Frame
If chkBal.Value = 1 Then
    frCYear.Enabled = True
    FraCMonth.Enabled = True
    chkMonthly.Enabled = True
Else
    frCYear.Enabled = False
    FraCMonth.Enabled = False
    chkMonthly.Enabled = False
End If
If optImd.Value = True Then
    frwhile.Enabled = False
Else
    frwhile.Enabled = True
End If
FrBusnessRule.Enabled = True

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

Private Sub optImd_Click()
    frwhile.Enabled = False
End Sub

Private Sub optNext_Click()
If bytMode = 2 Or bytMode = 3 Then frwhile.Enabled = True
End Sub

Private Sub TB1_Click(PreviousTab As Integer)
If bytMode = 2 Then Exit Sub
If PreviousTab = 1 Then Exit Sub
MSF1.Col = 0
If MSF1.Text = NewCaptionTxt("31003", adrsC) Then Exit Sub
Call Display
End Sub

Private Sub Display()       '' Displays the Given Master Records
On Error GoTo ERR_P
txtCode.Text = MSF1.TextMatrix(MSF1.Row, 0)     '' Leave Code
cboCat.Value = MSF1.TextMatrix(MSF1.Row, 4)     '' Category

adrsDept1.MoveFirst
Do
    adrsDept1.Find "LvCode='" & txtCode.Text & "' "
    
    If adrsDept1("Cat") <> cboCatCode.Text Then adrsDept1.MoveNext
    If adrsDept1.EOF Then Exit Do
Loop While Not adrsDept1("Cat") = cboCatCode.Text
If adrsDept1.EOF Then
    bytMode = 1
    Call ChangeMode
Else
    txtName.Text = adrsDept1("Leave")
    chkBal.Value = IIf(adrsDept1("Type") = "N", 0, 1)
    chkPayable.Value = IIf(adrsDept1("Paid") = "N", 0, 1)
    Select Case adrsDept1("Run_Wrk")
        Case "R": optIncl.Value = True
        Case "W": optExcl.Value = True
        Case "O": optDecide.Value = True
    End Select
    chkCarry.Value = IIf(adrsDept1("Lv_Cof") = "N", 0, 1)
    chkEncash.Value = IIf(adrsDept1("EnCase") = "N", 0, 1)
    txtCredit.Text = Format(adrsDept1("Lv_Qty"), "0.00")
    txtAllowMax.Text = Format(adrsDept1("Lv_Acumul"), "0.00")
    optImd.Value = IIf(adrsDept1("CreditNow") = "N", 0, 1)
    optNext.Value = IIf(adrsDept1("CreditNow") = "N", 1, 0)
    optFull.Value = IIf(adrsDept1("FulCredit") = "N", 0, 1)
    optProp.Value = IIf(adrsDept1("FulCredit") = "N", 1, 0)
    txtAllow.Text = adrsDept1("No_OfTimes")
    txtMax.Text = Format(adrsDept1("AllowDays"), "0.00")
    txtMin.Text = Format(adrsDept1("MinAllowdays"), "0.00")
End If

    chkMonthly.Value = IIf((adrsDept1("CrMonthly") = "N") Or (IsNull(adrsDept1("CrMonthly")) = True), 0, 1)
    If chkMonthly.Value = 0 Then
        FraCMonth.Visible = False
        frCYear.Visible = True
    ElseIf chkMonthly.Value = 1 Then
        FraCMonth.Visible = True
        frCYear.Visible = False
    Else
        FraCMonth.Visible = True
        frCYear.Visible = False
    End If
    txtCreditM.Text = txtCredit.Text
    txtAllowMaxM = txtAllowMax.Text

Exit Sub
ERR_P:
    'Resume Next
End Sub

Private Sub MSF1_DblClick()
If TB1.TabEnabled(1) = False Then Exit Sub
TB1.Tab = 1
End Sub

Private Function ValidateAddmaster() As Boolean     '' Validate If in Addition Mode
On Error GoTo ERR_P
Dim bytTmp As Integer, bytTmpCnt As Integer
ValidateAddmaster = True
'' Check for Empty Code
If Len(Trim(txtCode.Text)) < 2 Then
    MsgBox NewCaptionTxt("31038", adrsC), vbExclamation
    txtCode.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
Select Case UCase(txtCode.Text)
    Case pVStar.AbsCode, pVStar.PrsCode, pVStar.WosCode, pVStar.HlsCode
        MsgBox NewCaptionTxt("31039", adrsC), vbInformation
        txtCode.SetFocus
        ValidateAddmaster = False
        Exit Function
End Select
If cboCat.ListCount <= 0 Then
    MsgBox NewCaptionTxt("31040", adrsC), vbExclamation
    ValidateAddmaster = False
    Exit Function
End If
If chkCat.Value = 0 And cboCat.Value = "" Then
    MsgBox NewCaptionTxt("31041", adrsC), vbExclamation
    cboCat.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
'' Check for Duplicate Leaves
If MSF1.Rows > 1 Then
    bytTmpCnt = 0
    If chkCat.Value = 1 Then
        For bytTmp = 1 To MSF1.Rows - 1
            If MSF1.TextMatrix(bytTmp, 0) = txtCode.Text Then bytTmpCnt = bytTmpCnt + 1
        Next
    Else
        For bytTmp = 1 To MSF1.Rows - 1
            If MSF1.TextMatrix(bytTmp, 0) = txtCode.Text And _
            MSF1.TextMatrix(bytTmp, 1) = cboCat.Value Then bytTmpCnt = bytTmpCnt + 1
        Next
    End If
Else
    bytTmpCnt = 0
End If
If bytTmpCnt > 0 Then
    MsgBox NewCaptionTxt("31042", adrsC), vbExclamation
    txtCode.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
'' Check for Empty Name
If Trim(txtName.Text) = "" Then
    MsgBox NewCaptionTxt("31043", adrsC), vbExclamation
    txtName.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
'' Check for Invalid Accumulation Amount
If chkCarry.Value = 1 Then
    If Val(txtCredit.Text) > 0 And Val(txtCredit.Text) > Val(txtAllowMax.Text) Then
        MsgBox NewCaptionTxt("31044", adrsC), vbExclamation
        If txtAllowMax.Enabled = True Then txtAllowMax.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If
If optNext.Value = True Then
    If optFull.Value = False And optProp.Value = False Then
        MsgBox NewCaptionTxt("31045", adrsC), vbExclamation
        If optFull.Enabled = True Then optFull.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If
txtAllow.Text = IIf(txtAllow.Text = "", "0", txtAllow.Text)
txtMax.Text = IIf(Trim(txtMax.Text) = "", "0.00", Format(txtMax.Text, "0.00"))
txtMin.Text = IIf(Trim(txtMin.Text) = "", "0.00", Format(txtMin.Text, "0.00"))
If Right(txtMax.Text, 2) <> "00" And Right(txtMax.Text, 2) <> "50" Then
    MsgBox NewCaptionTxt("00056", adrsMod), vbExclamation
    txtMax.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If Right(txtMin.Text, 2) <> "00" And Right(txtMin.Text, 2) <> "50" Then
    MsgBox NewCaptionTxt("00056", adrsMod), vbExclamation
    txtMin.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If chkMonthly.Value = 1 Then    ' 15-07
    txtAllowMaxM.Text = IIf(Trim(txtAllowMaxM.Text) = "", "0.00", Format(txtAllowMaxM.Text, "0.00"))
    txtCreditM.Text = IIf(Trim(txtCreditM.Text) = "", "0.00", Format(txtCreditM.Text, "0.00"))
    If Right(txtAllowMaxM.Text, 2) <> "00" And Right(txtAllowMaxM.Text, 2) <> "50" Then
        MsgBox NewCaptionTxt("00056", adrsMod), vbExclamation
        txtAllowMaxM.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
    If Right(txtCreditM.Text, 2) <> "00" And Right(txtCreditM.Text, 2) <> "50" Then
        MsgBox NewCaptionTxt("00056", adrsMod), vbExclamation
        txtCreditM.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster :: " & Me.Caption)
    ValidateAddmaster = False
End Function

Private Function ValidateModMaster() As Boolean     '' Validate If in Addition Mode
On Error GoTo ERR_P
ValidateModMaster = True
'' Check for Empty Name
If Trim(txtName.Text) = "" Then
    MsgBox NewCaptionTxt("31043", adrsC), vbExclamation
    txtName.SetFocus
    ValidateModMaster = False
    Exit Function
End If
If chkCarry.Value = 1 Then
    If Val(txtCredit.Text) > 0 And Val(txtCredit.Text) > Val(txtAllowMax.Text) Then
        MsgBox NewCaptionTxt("31044", adrsC)
        If txtAllowMax.Enabled = True Then txtAllowMax.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
End If
If optNext.Value = True Then
    If optFull.Value = False And optProp.Value = False Then
        MsgBox NewCaptionTxt("31045", adrsC), vbExclamation
        If optFull.Enabled = True Then optFull.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
End If
txtAllow.Text = IIf(txtAllow.Text = "", "0", txtAllow.Text)
txtMax.Text = IIf(Trim(txtMax.Text) = "", "0.00", Format(txtMax.Text, "0.00"))
txtMin.Text = IIf(Trim(txtMin.Text) = "", "0.00", Format(txtMin.Text, "0.00"))
If Right(txtMax.Text, 2) <> "00" And Right(txtMax.Text, 2) <> "50" Then
    MsgBox NewCaptionTxt("00056", adrsMod), vbExclamation
    txtMax.SetFocus
    ValidateModMaster = False
    Exit Function
End If
If Right(txtMin.Text, 2) <> "00" And Right(txtMin.Text, 2) <> "50" Then
    MsgBox NewCaptionTxt("00056", adrsMod), vbExclamation
    txtMin.SetFocus
    ValidateModMaster = False
    Exit Function
End If
If chkMonthly.Value = 1 Then    ' 15-07
    txtAllowMaxM.Text = IIf(Trim(txtAllowMaxM.Text) = "", "0.00", Format(txtAllowMaxM.Text, "0.00"))
    txtCreditM.Text = IIf(Trim(txtCreditM.Text) = "", "0.00", Format(txtCreditM.Text, "0.00"))
    If Right(txtAllowMaxM.Text, 2) <> "00" And Right(txtAllowMaxM.Text, 2) <> "50" Then
        MsgBox NewCaptionTxt("00056", adrsMod), vbExclamation
        txtAllowMaxM.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
    If Right(txtCreditM.Text, 2) <> "00" And Right(txtCreditM.Text, 2) <> "50" Then
        MsgBox NewCaptionTxt("00056", adrsMod), vbExclamation
        txtCreditM.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
End If
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
    ValidateModMaster = False
End Function

Private Function SaveAddMaster() As Boolean
Dim WebStr As String

On Error GoTo ERR_P
SaveAddMaster = True
Dim strRW As String, bytTmp As Byte
If optIncl.Value = True Then strRW = "R"
If optExcl.Value = True Then strRW = "W"
If optDecide.Value = True Then strRW = "O"
If chkCat.Value = 1 Then        '' Insert for All Categories
    For bytTmp = 0 To cboCat.ListCount - 1
        ConMain.Execute "insert into LeavDesc Values('" & txtName.Text & _
        "','" & txtCode.Text & "','" & IIf(chkBal.Value = 1, "Y", "N") & "','" & _
        IIf(chkPayable.Value = 1, "Y", "N") & "','" & IIf(chkEncash.Value = 1, "Y", "N") & "','" & _
        IIf(chkCarry.Value = 1, "Y", "N") & "'," & IIf(Trim(txtCredit.Text) = "", _
        "0.00", Format(txtCredit.Text, "0.00")) & "," & IIf(Trim(txtAllowMax.Text) = "", _
        "0.00", Format(txtAllowMax.Text, "0.00")) & _
        ",'" & strRW & "','Y','" & cboCat.List(bytTmp, 1) & "','" & _
        IIf(optImd.Value = True, "Y", "N") & "','" & IIf(optFull.Value = True, "Y", "N") & "'," & _
        txtAllow.Text & "," & txtMax.Text & "," & txtMin.Text & ",''" & IIf(GetFlagStatus("NOMIS2010"), "", IIf(chkMonthly.Value = 1, ",'Y'", ",'N'")) & _
        " " & WebStr & ")"
        If SubLeaveFlag = 1 Then ' 15-10
            If txtCode.Text = "EL" Then
                If adrsASC.State = 1 Then adrsASC.Close
                adrsASC.Open "select sum(Lv_Qty),sum(Lv_Acumul) from LeavDesc where LvCode in ('EN','NE') and Cat='" & cboCat.List(bytTmp, 1) & "'", ConMain, adOpenStatic
                ConMain.Execute "Update LeavDesc Set Lv_Qty=" & Format(adrsASC.Fields(0), "0.00") & ",Lv_Acumul=" & Format(adrsASC.Fields(1), "0.00") & " Where LvCode='EL' AND Cat='" & cboCat.List(bytTmp, 1) & "'"
            End If
        End If
    Next
Else
    ConMain.Execute "insert into LeavDesc Values('" & txtName.Text & _
    "','" & txtCode.Text & "','" & IIf(chkBal.Value = 1, "Y", "N") & "','" & _
    IIf(chkPayable.Value = 1, "Y", "N") & "','" & IIf(chkEncash.Value = 1, "Y", "N") & "','" & _
    IIf(chkCarry.Value = 1, "Y", "N") & "'," & IIf(Trim(txtCredit.Text) = "", _
    "0.00", Format(txtCredit.Text, "0.00")) & "," & IIf(Trim(txtAllowMax.Text) = "", _
    "0.00", Format(txtAllowMax.Text, "0.00")) & ",'" & strRW & "','Y','" & cboCatCode.Text & _
    "','" & IIf(optImd.Value = True, "Y", "N") & "','" & _
    IIf(optFull.Value = True, "Y", "N") & "'," & txtAllow.Text & "," & txtMax.Text & _
    "," & txtMin.Text & ",''" & IIf(GetFlagStatus("NOMIS2010"), "", IIf(chkMonthly.Value = 1, ",'Y'", ",'N'")) & "" & WebStr & ")"
    If SubLeaveFlag = 1 Then ' 15-10
        If txtCode.Text = "EL" Then
            If adrsASC.State = 1 Then adrsASC.Close
            adrsASC.Open "select sum(Lv_Qty),sum(Lv_Acumul) from LeavDesc where LvCode in ('EN','NE') and Cat='" & cboCatCode.Text & "'", ConMain, adOpenStatic
            ConMain.Execute "Update LeavDesc Set Lv_Qty=" & Format(adrsASC.Fields(0), "0.00") & ",Lv_Acumul=" & Format(adrsASC.Fields(1), "0.00") & " Where LvCode='EL' AND Cat='" & cboCatCode.Text & "'"
        End If
    End If
End If
Exit Function
ERR_P:
    SaveAddMaster = False
    ShowError ("SaveAddMaster :: " & Me.Caption)
''Resume Next
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
Dim WebStr As String
SaveModMaster = True        '' Update
Dim strRW As String, bytTmp As String
If optIncl.Value = True Then strRW = "R"
If optExcl.Value = True Then strRW = "W"
If optDecide.Value = True Then strRW = "O"
ConMain.Execute "Update LeavDesc Set Leave='" & txtName.Text & "',Type='" & _
IIf(chkBal.Value = 1, "Y", "N") & "',Paid='" & IIf(chkPayable.Value = 1, "Y", "N") & _
"',Encase='" & IIf(chkEncash.Value = 1, "Y", "N") & "',Lv_Cof='" & _
IIf(chkCarry.Value = 1, "Y", "N") & "',Lv_Qty=" & IIf(Trim(txtCredit.Text) = "", _
"0.00", Format(txtCredit.Text, "0.00")) & ",Lv_Acumul=" & _
IIf(Trim(txtAllowMax.Text) = "", "0.00", Format(txtAllowMax.Text, "0.00")) & ",Run_Wrk='" & _
strRW & "',CreditNow='" & IIf(optImd.Value = True, "Y", "N") & _
"',FulCredit='" & IIf(optFull.Value = True, "Y", "N") & "',No_OfTimes=" & txtAllow.Text & _
",AllowDays=" & txtMax.Text & ",MinAllowDays=" & txtMin.Text & IIf(GetFlagStatus("NOMIS2010"), "", ",CrMonthly=" & IIf(chkMonthly.Value = 1, "'Y'", "'N'")) & " " & WebStr & " Where LvCode='" & _
txtCode.Text & "' and Cat='" & cboCatCode.Text & "'"
If SubLeaveFlag = 1 Then  ' 15-10
    If (txtCode.Text = "EN" Or txtCode.Text = "NE") Then
        If adrsASC.State = 1 Then adrsASC.Close
        adrsASC.Open "select sum(Lv_Qty),sum(Lv_Acumul) from LeavDesc where LvCode in ('EN','NE') and Cat='" & cboCatCode.Text & "'", ConMain, adOpenStatic
        ConMain.Execute "Update LeavDesc Set Lv_Qty=" & Format(adrsASC.Fields(0), "0.00") & ",Lv_Acumul=" & Format(adrsASC.Fields(1), "0.00") & " Where LvCode='EL' AND Cat='" & cboCatCode.Text & "'"
    End If
End If
Exit Function
ERR_P:
    SaveModMaster = False
    ShowError ("SaveModMaster :: " & Me.Caption)
End Function

Private Sub txtAllow_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = KeyPressCheck(KeyAscii, 2)
End Select
End Sub

Private Sub txtAllowMaxM_Change()
txtAllowMax.Text = txtAllowMaxM.Text
End Sub

Private Sub txtAllowMaxM_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = KeyDecimal3(KeyAscii, txtAllowMaxM)
End Select
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 1))))
End Select
End Sub

Private Sub txtCredit_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = KeyDecimal3(KeyAscii, txtCredit)
End Select
End Sub

Private Sub txtAllowMax_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = KeyDecimal3(KeyAscii, txtAllowMax)
End Select
End Sub

Private Sub txtCreditM_Change()
txtCredit.Text = txtCreditM.Text
End Sub

Private Sub txtCreditM_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = KeyDecimal3(KeyAscii, txtCreditM)
End Select
End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = keycheck(KeyAscii, txtMax)
End Select
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = keycheck(KeyAscii, txtMin)
End Select
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
    If MsgBox(NewCaptionTxt("00015", adrsMod) & vbCrLf & _
    NewCaptionTxt("31046", adrsC), vbYesNo + vbQuestion) = vbYes Then            '' Delete the Record
        ''
        If adrslv.State = 1 Then adrslv.Close
        If FindTable("LvBal" & Right(Year(Date), 2)) Then
        If FieldExists("LvBaL" & Right(Year(Date), 2), txtCode.Text) Then
            adrslv.Open "select " & txtCode.Text & " from LvBal" & Right(Year(Date), 2) & ", empmst where empmst.Empcode =LvBal" & Right(Year(Date), 2) & ".Empcode and empmst.cat='" & cboCatCode.Text & "' order by lvbal" & Right(Year(Date), 2) & ".Empcode", ConMain, adOpenStatic, adLockReadOnly
                If adrslv.EOF = False Then
                    adrslv.MoveFirst
                    Do While adrslv.EOF = False
                    If adrslv.Fields(txtCode.Text) > 0 Then
                    '
                        ans = ""
                        MsgBox "Leaves cannot be deleted because Balance Leaves are there "
                        ans = MsgBox("Do you realy want to delete this leave?", vbYesNo)
                        If ans = vbYes Then
                            ConMain.Execute "update lvbal" & Right(Year(Date), 2) & " set " & txtCode.Text & " = 0 where LvBal" & Right(Year(Date), 2) & ".Empcode in (select empcode from empmst where empmst.cat='" & cboCatCode.Text & "')"
                            If SubLeaveFlag = 1 Then  ' 15-10
                                Dim strSubLeave As String, strLv As String
                                Dim strquery As String
                                If (txtCode.Text = "EN" Or txtCode.Text = "NE") Then
                                    strLv = "EL"
                                    If FieldExists("LvBaL" & Right(Year(Date), 2), "EN") Then strSubLeave = ",EN"
                                    If FieldExists("LvBaL" & Right(Year(Date), 2), "NE") Then strSubLeave = strSubLeave & ",NE"
                                End If
                                If strSubLeave <> "" Then
                                    strSubLeave = Right(strSubLeave, Len(strSubLeave) - 1)
                                    strquery = "select " & strSubLeave & ",lvbal" & Right(pVStar.YearSel, 2) & ".empcode from lvbal" & Right(pVStar.YearSel, 2) & " where lvbal" & Right(pVStar.YearSel, 2) & ".Empcode in (select empcode from empmst where empmst.cat='" & cboCatCode.Text & "')"
                                    Call UpDateSubLeave("lvbal" & Right(pVStar.YearSel, 2), strSubLeave, strquery, strLv)
                                End If
                                

                            End If
                
                            Call AddActivityLog(lgDelete_Action, 1, 5)
                            ConMain.Execute "Delete from LeavDesc where LvCode='" & _
                            txtCode.Text & "' and Cat='" & cboCatCode.Text & "'"
                            If SubLeaveFlag = 1 Then ' 15-10
                                If txtCode.Text = "CM" Or txtCode.Text = "HP" Then
                                    Call DelSubLeave("'CM','HP'", "'SL'")
                                Else
                                    Call DelSubLeave("'EN','NE'", "'EL'")
                                End If
                            End If
                            Call AddActivityLog(lgDelete_Action, 1, 5)
                            Call AuditInfo("DELETE", Me.Caption, "Deleted Leave: " & txtCode.Text)
                            Exit Do
                        Else
                            Exit Do
                        End If
                                          
                    End If
                    adrslv.MoveNext
                    If adrslv.EOF = True Then
                    ConMain.Execute "Delete from LeavDesc where LvCode='" & _
                        txtCode.Text & "' and Cat='" & cboCatCode.Text & "'"
                    If SubLeaveFlag = 1 Then ' 15-10
                        If txtCode.Text = "CM" Or txtCode.Text = "HP" Then
                            Call DelSubLeave("'CM','HP'", "'SL'")
                        Else
                            Call DelSubLeave("'EN','NE'", "'EL'")
                        End If
                    End If
                        Call AddActivityLog(lgDelete_Action, 1, 5)
                        Call AuditInfo("DELETE", Me.Caption, "Deleted Leave: " & txtCode.Text)
                    Exit Do
                    End If
                    Loop
               Else
                     ConMain.Execute "Delete from LeavDesc where LvCode='" & _
                        txtCode.Text & "' and Cat='" & cboCatCode.Text & "'"
                        If SubLeaveFlag = 1 Then ' 15-10
                            If txtCode.Text = "CM" Or txtCode.Text = "HP" Then
                                Call DelSubLeave("'CM','HP'", "'SL'")
                            Else
                                Call DelSubLeave("'EN','NE'", "'EL'")
                            End If
                        End If
                        Call AddActivityLog(lgDelete_Action, 1, 5)
                        Call AuditInfo("DELETE", Me.Caption, "Deleted Leave: " & txtCode.Text)
                End If
        Else
        ConMain.Execute "Delete from LeavDesc where LvCode='" & _
                txtCode.Text & "' and Cat='" & cboCatCode.Text & "'"
                If SubLeaveFlag = 1 Then ' 15-10
                    If txtCode.Text = "CM" Or txtCode.Text = "HP" Then
                        Call DelSubLeave("'CM','HP'", "'SL'")
                    Else
                        Call DelSubLeave("'EN','NE'", "'EL'")
                    End If
                 End If
                Call AddActivityLog(lgDelete_Action, 1, 5)
                Call AuditInfo("DELETE", Me.Caption, "Deleted Leave: " & txtCode.Text)
            End If
        End If
    End If
    Call FillGrid       '' Reflect the Grid
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
ERR_P:
    
    ShowError ("Delete Record :: " & Me.Caption)
End Sub
Private Sub DelSubLeave(ByVal strSubLv As String, ByVal strLv As String)    ' 15-10
    If adrsASC.State = 1 Then adrsASC.Close
    adrsASC.Open "select LvCode from LeavDesc where Cat='" & cboCatCode.Text & "' and LvCode in (" & strSubLv & ")", ConMain, adOpenStatic
    If (adrsASC.EOF And adrsASC.BOF) Then
        ConMain.Execute "Delete from LeavDesc where LvCode=" & strLv & " and Cat='" & cboCatCode.Text & "'"
    End If
End Sub
Private Sub FillCatCombo()
On Error GoTo ERR_P
Call ComboFill(cboCat, 3, 2)
Exit Sub
ERR_P:
    ShowError ("FillCatCombo :: " & Me.Caption)
End Sub

Private Sub AdjustActions()
If bytMode = 1 Then Exit Sub
'' Check on Payable days
If chkPayable.Value = 0 Then
    chkEncash.Value = 0
    chkCarry.Value = 0
    frEnd.Enabled = False
    frCYear.Enabled = False
    FraCMonth.Enabled = False
    chkMonthly.Enabled = False
End If
'' Check on Keep Balance
If chkBal.Value = 0 Then
    chkEncash.Value = 0
    chkCarry.Value = 0
    frEnd.Enabled = False
    frCYear.Enabled = False
    FraCMonth.Enabled = False
    chkMonthly.Enabled = False
End If
If chkPayable.Value = 1 And chkBal.Value = 1 Then
    frEnd.Enabled = True
    frCYear.Enabled = True
    FraCMonth.Enabled = True
'    txtCredit.SetFocus
    chkMonthly.Enabled = True
    If chkMonthly.Value = 0 Then
        FraCMonth.Visible = False
        frCYear.Visible = True
        txtCredit.SetFocus
    ElseIf chkMonthly.Value = 1 Then
        FraCMonth.Visible = True
        frCYear.Visible = False
        txtCreditM.SetFocus
    Else
        FraCMonth.Visible = True
        frCYear.Visible = False
        txtCredit.SetFocus
    End If
End If
End Sub

Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo ERR_P
Call AddActivityLog(lgADD_MODE, 1, 5)     '' Add Activity
Call AuditInfo("ADD", Me.Caption, "Added Leave: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 5)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edited Leave: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub
