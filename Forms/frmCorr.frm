VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCorr 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frEmp 
      Height          =   1035
      Left            =   0
      TabIndex        =   79
      Top             =   0
      Width           =   9585
      Begin VB.TextBox txtEmpCode 
         Height          =   375
         Left            =   6720
         TabIndex        =   84
         Text            =   "txtEmpCode"
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   780
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   5610
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   143
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Search ECN"
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
         Left            =   7440
         TabIndex        =   83
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblDeptCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   360
         Left            =   240
         TabIndex        =   40
         Top             =   120
         Width           =   1050
      End
      Begin MSForms.ComboBox cboDept 
         Height          =   300
         Left            =   1320
         TabIndex        =   0
         Top             =   150
         Width           =   3435
         VariousPropertyBits=   612390939
         DisplayStyle    =   3
         Size            =   "6059;529"
         TextColumn      =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   5040
         TabIndex        =   43
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month "
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
         Left            =   4980
         TabIndex        =   42
         Top             =   180
         Width           =   600
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   360
         TabIndex        =   41
         Top             =   525
         Width           =   870
      End
      Begin MSForms.ComboBox cboEmp 
         Height          =   300
         Left            =   1320
         TabIndex        =   1
         Top             =   495
         Width           =   3405
         VariousPropertyBits=   612390939
         DisplayStyle    =   3
         Size            =   "6006;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin TabDlg.SSTab TB1 
      Height          =   6855
      Left            =   0
      TabIndex        =   39
      Top             =   1065
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   12091
      _Version        =   393216
      TabOrientation  =   2
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
      TabCaption(0)   =   "Attendance Records"
      TabPicture(0)   =   "frmCorr.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSF1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Attendance Details"
      TabPicture(1)   =   "frmCorr.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frAll"
      Tab(1).Control(1)=   "cmdShift"
      Tab(1).Control(2)=   "cmdRec"
      Tab(1).Control(3)=   "cmdStatus"
      Tab(1).Control(4)=   "cmdOn"
      Tab(1).Control(5)=   "cmdOff"
      Tab(1).Control(6)=   "cmdOTSave"
      Tab(1).Control(7)=   "cmdTimeCan"
      Tab(1).Control(8)=   "cmdExitIn"
      Tab(1).Control(9)=   "chkPermanentCorrection"
      Tab(1).ControlCount=   10
      Begin VB.CheckBox chkPermanentCorrection 
         Caption         =   "Click For Permanent Correction of this Records"
         Height          =   615
         Left            =   -74280
         TabIndex        =   80
         Top             =   5760
         Width           =   1815
      End
      Begin VB.CommandButton cmdExitIn 
         Caption         =   "Command1"
         Height          =   375
         Left            =   -66840
         TabIndex        =   38
         Top             =   5040
         Width           =   1125
      End
      Begin VB.CommandButton cmdTimeCan 
         Caption         =   "Command1"
         Height          =   375
         Left            =   -67920
         TabIndex        =   37
         Top             =   5040
         Width           =   1125
      End
      Begin VB.CommandButton cmdOTSave 
         Caption         =   "Command1"
         Height          =   375
         Left            =   -69000
         TabIndex        =   36
         Top             =   5040
         Width           =   1125
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Command1"
         Height          =   375
         Left            =   -70080
         TabIndex        =   35
         Top             =   5040
         Width           =   1125
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "Command1"
         Height          =   375
         Left            =   -71160
         TabIndex        =   34
         Top             =   5040
         Width           =   1125
      End
      Begin VB.CommandButton cmdStatus 
         Caption         =   "Command1"
         Height          =   375
         Left            =   -72240
         TabIndex        =   33
         Top             =   5040
         Width           =   1125
      End
      Begin VB.CommandButton cmdRec 
         Caption         =   "Command1"
         Height          =   375
         Left            =   -73320
         TabIndex        =   32
         Top             =   5040
         Width           =   1125
      End
      Begin VB.CommandButton cmdShift 
         Caption         =   "Command1"
         Height          =   375
         Left            =   -74400
         TabIndex        =   31
         Top             =   5040
         Width           =   1125
      End
      Begin VB.Frame frAll 
         Height          =   4815
         Left            =   -74520
         TabIndex        =   45
         Top             =   120
         Width           =   8955
         Begin VB.Frame frDSE 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1635
            Left            =   30
            TabIndex        =   46
            Top             =   120
            Width           =   2400
            Begin VB.Label lblEntry 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Entry"
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
               TabIndex        =   50
               Top             =   1200
               Width           =   465
            End
            Begin VB.Label lblShift 
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
               Left            =   90
               TabIndex        =   49
               Top             =   720
               Width           =   390
            End
            Begin VB.Label lblDateCap 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
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
               TabIndex        =   47
               Top             =   240
               Width           =   405
            End
            Begin VB.Label lblDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "dsda"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   1320
               TabIndex        =   48
               Top             =   240
               Width           =   420
            End
            Begin MSForms.ComboBox cboShift 
               Height          =   345
               Left            =   1380
               TabIndex        =   4
               Top             =   720
               Width           =   975
               VariousPropertyBits=   748701723
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "1720;609"
               ListWidth       =   5500
               ColumnCount     =   3
               cColumnInfo     =   3
               MatchEntry      =   0
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
               Object.Width           =   "1500;2000;2000"
            End
            Begin MSForms.ComboBox cboEntry 
               Height          =   345
               Left            =   1380
               TabIndex        =   5
               Top             =   1200
               Width           =   975
               VariousPropertyBits=   748701723
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "1720;609"
               MatchEntry      =   0
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.Frame frPerm 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   6630
            TabIndex        =   76
            Top             =   2460
            Visible         =   0   'False
            Width           =   2280
            Begin MSMask.MaskEdBox txtPEarly 
               Height          =   375
               Left            =   1290
               TabIndex        =   30
               Top             =   720
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   1
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
            Begin MSMask.MaskEdBox txtPLate 
               Height          =   375
               Left            =   1290
               TabIndex        =   29
               Top             =   240
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   1
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
            Begin VB.Label lblPEarly 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Early Card"
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
               TabIndex        =   78
               Top             =   720
               Width           =   915
            End
            Begin VB.Label lblPLate 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Late Card"
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
               TabIndex        =   77
               Top             =   360
               Width           =   840
            End
         End
         Begin VB.Frame frOff 
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Left            =   6660
            TabIndex        =   73
            Top             =   1320
            Width           =   2265
            Begin MSMask.MaskEdBox txtOffTo 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   1080
               TabIndex        =   15
               Top             =   660
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               AutoTab         =   -1  'True
               Enabled         =   0   'False
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
            Begin MSMask.MaskEdBox txtOffFrom 
               CausesValidation=   0   'False
               Height          =   345
               Left            =   1080
               TabIndex        =   14
               Top             =   240
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               AutoTab         =   -1  'True
               Enabled         =   0   'False
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
            Begin VB.Label lblOffTo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "To"
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
               TabIndex        =   75
               Top             =   660
               Width           =   210
            End
            Begin VB.Label lblOffFrom 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "From"
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
               TabIndex        =   74
               Top             =   285
               Width           =   450
            End
         End
         Begin VB.Frame frOn 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   6660
            TabIndex        =   70
            Top             =   120
            Width           =   2265
            Begin MSMask.MaskEdBox txtOnTo 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   1065
               TabIndex        =   13
               Top             =   690
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               AutoTab         =   -1  'True
               Enabled         =   0   'False
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
            Begin MSMask.MaskEdBox txtOnFrom 
               CausesValidation=   0   'False
               Height          =   330
               Left            =   1050
               TabIndex        =   12
               Top             =   240
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               AutoTab         =   -1  'True
               Enabled         =   0   'False
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
            Begin VB.Label lblOnTo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "To"
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
               TabIndex        =   72
               Top             =   720
               Width           =   210
            End
            Begin VB.Label lblOnFrom 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "From"
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
               TabIndex        =   71
               Top             =   330
               Width           =   450
            End
         End
         Begin VB.Frame frIrr 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1860
            Left            =   2490
            TabIndex        =   63
            Top             =   1830
            Width           =   4125
            Begin VB.CheckBox chkIrr 
               BackColor       =   &H8000000B&
               Caption         =   "Irregular Interval"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   240
               TabIndex        =   22
               Top             =   0
               Width           =   2280
            End
            Begin MSMask.MaskEdBox txtT2 
               CausesValidation=   0   'False
               Height          =   345
               Left            =   1080
               TabIndex        =   23
               Top             =   360
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
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
            Begin MSMask.MaskEdBox txtT4 
               CausesValidation=   0   'False
               Height          =   345
               Left            =   1080
               TabIndex        =   24
               Top             =   840
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
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
            Begin MSMask.MaskEdBox txtT6 
               CausesValidation=   0   'False
               Height          =   345
               Left            =   1080
               TabIndex        =   25
               Top             =   1320
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
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
            Begin MSMask.MaskEdBox txtT3 
               Height          =   345
               Left            =   3150
               TabIndex        =   26
               Top             =   360
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
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
            Begin MSMask.MaskEdBox txtT5 
               CausesValidation=   0   'False
               Height          =   345
               Left            =   3150
               TabIndex        =   27
               Top             =   840
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
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
            Begin MSMask.MaskEdBox txtT7 
               CausesValidation=   0   'False
               Height          =   345
               Left            =   3150
               TabIndex        =   28
               Top             =   1320
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
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
            Begin VB.Label lblT7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "7th"
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
               Left            =   2460
               TabIndex        =   69
               Top             =   1350
               Width           =   270
            End
            Begin VB.Label lblT5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "3rd"
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
               Left            =   2460
               TabIndex        =   65
               Top             =   450
               Width           =   270
            End
            Begin VB.Label lblT3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "4th"
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
               TabIndex        =   66
               Top             =   870
               Width           =   270
            End
            Begin VB.Label lblT6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "5th"
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
               Left            =   2460
               TabIndex        =   67
               Top             =   900
               Width           =   270
            End
            Begin VB.Label lblT4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "6th"
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
               TabIndex        =   68
               Top             =   1380
               Width           =   270
            End
            Begin VB.Label lblT2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "2nd"
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
               Left            =   345
               TabIndex        =   64
               Top             =   360
               Width           =   315
            End
         End
         Begin VB.Frame frTime 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1605
            Left            =   2460
            TabIndex        =   56
            Top             =   120
            Width           =   4185
            Begin MSMask.MaskEdBox txtDept 
               CausesValidation=   0   'False
               Height          =   345
               Left            =   1170
               TabIndex        =   7
               Top             =   720
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   609
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
            Begin MSMask.MaskEdBox txtArr 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   1170
               TabIndex        =   6
               Top             =   240
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   556
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
            Begin MSMask.MaskEdBox txtWork 
               CausesValidation=   0   'False
               Height          =   345
               Left            =   1200
               TabIndex        =   8
               Top             =   1200
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   609
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
            Begin MSMask.MaskEdBox txtLate 
               CausesValidation=   0   'False
               Height          =   345
               Left            =   3240
               TabIndex        =   9
               Top             =   240
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   609
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
               Format          =   "00.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtEarly 
               CausesValidation=   0   'False
               Height          =   345
               Left            =   3240
               TabIndex        =   10
               Top             =   720
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   609
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
               Format          =   "00.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtOT 
               CausesValidation=   0   'False
               Height          =   345
               Left            =   3240
               TabIndex        =   11
               Top             =   1200
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   609
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
            Begin VB.Label lblOT 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Overtime"
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
               Left            =   2280
               TabIndex        =   62
               Top             =   1260
               Width           =   765
            End
            Begin VB.Label lblEarly 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Early"
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
               Left            =   2280
               TabIndex        =   61
               Top             =   765
               Width           =   450
            End
            Begin VB.Label lblLate 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Late"
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
               Left            =   2280
               TabIndex        =   60
               Top             =   240
               Width           =   375
            End
            Begin VB.Label lblWork 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Work Hrs."
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
               TabIndex        =   59
               Top             =   1260
               Width           =   885
            End
            Begin VB.Label lblDept 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Departure"
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
               Top             =   720
               Width           =   840
            End
            Begin VB.Label lblArr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Arrival "
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
               TabIndex        =   57
               Top             =   255
               Width           =   585
            End
         End
         Begin VB.Frame frMisc 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Left            =   30
            TabIndex        =   51
            Top             =   1740
            Width           =   2445
            Begin VB.TextBox txtlate2 
               Height          =   315
               Left            =   1560
               TabIndex        =   21
               Top             =   2280
               Width           =   750
            End
            Begin VB.TextBox txtlate1 
               Height          =   315
               Left            =   1560
               TabIndex        =   20
               Top             =   1920
               Width           =   750
            End
            Begin MSMask.MaskEdBox txtRest 
               Height          =   315
               Left            =   1500
               TabIndex        =   18
               Top             =   1110
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   556
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
            Begin VB.Label lbllate2 
               Caption         =   "2.Lunch Late Hrs"
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   2400
               Width           =   1335
            End
            Begin VB.Label lbllate1 
               Caption         =   "Lunch Late"
               Height          =   255
               Left            =   120
               TabIndex        =   81
               Top             =   2040
               Width           =   1335
            End
            Begin VB.Label lblCO 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CO Days"
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
               TabIndex        =   55
               Top             =   1590
               Width           =   795
            End
            Begin VB.Label lblRest 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rest Hrs."
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
               TabIndex        =   54
               Top             =   1140
               Width           =   825
            End
            Begin VB.Label lblPdays 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Present Days"
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
               TabIndex        =   53
               Top             =   675
               Width           =   1185
            End
            Begin VB.Label lblStatus 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Status"
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
               TabIndex        =   52
               Top             =   240
               Width           =   570
            End
            Begin MSForms.ComboBox cboStatus 
               Height          =   345
               Left            =   1500
               TabIndex        =   16
               Top             =   240
               Width           =   900
               VariousPropertyBits=   748701723
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "1587;609"
               ListWidth       =   6000
               ColumnCount     =   2
               cColumnInfo     =   2
               MatchEntry      =   0
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
               Object.Width           =   "1500;4500"
            End
            Begin MSForms.ComboBox cboPresent 
               Height          =   285
               Left            =   1500
               TabIndex        =   17
               Top             =   690
               Width           =   840
               VariousPropertyBits=   748701723
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "1482;503"
               ColumnCount     =   2
               MatchEntry      =   0
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboCO 
               Height          =   345
               Left            =   1500
               TabIndex        =   19
               Top             =   1560
               Width           =   915
               VariousPropertyBits=   748701723
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "1614;609"
               MatchEntry      =   0
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   6825
         Left            =   360
         TabIndex        =   44
         Top             =   60
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   12039
         _Version        =   393216
         Rows            =   16
         Cols            =   11
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   4194368
         ForeColorFixed  =   8454143
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
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
Attribute VB_Name = "frmCorr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnRits(2 To 8) As Boolean      '' Boolean Array of Rights for Buttons
Private strFileName As String * 9       '' String File Name
Private blnFileFound As Boolean         '' Boolean for Invalid Transaction
Private bytRowVal As Byte
''
Dim adrsC As New ADODB.Recordset
Public OldArrtim As String, OldDeptim As String

Public Enum CorrectionStatus
    Corrected = 1
    UnCorrected = 2
End Enum

Private Sub cboCO_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then SendKeys Chr(9)
End Sub

Private Sub cboDept_Click()
On Error GoTo ERR_P
If cboDept.ListIndex < 0 Then Exit Sub               '' If No Department
If cboDept.Text = "ALL" Then
    Call ComboFill(cboEmp, 1, 2, 17)
Else
    Call ComboFill(cboEmp, 12, 2, cboDept.List(cboDept.ListIndex, 1))
End If
MSF1.Rows = 1
TB1.TabEnabled(1) = False
TB1.Tab = 0
Exit Sub
ERR_P:
    ShowError ("Department::" & Me.Caption)
End Sub

Private Sub cboEmp_Click()
If cboEmp.ListIndex < 0 Then Exit Sub               '' If No Employee
If blnFileFound = False Then Exit Sub               '' If File is Not Found
TB1.Tab = 0                                         '' Make the First Tab Visible
Call FillGrid                                      '' Fill the Grid with Employees Record
End Sub

Private Sub cboEntry_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then SendKeys Chr(9)
End Sub

Private Sub cboMonth_Click()
If cboMonth.Text = "" Then Exit Sub
Call ValidMonthYear
If blnFileFound = False Then        '' If no File is Found
    MSF1.Rows = 1
    TB1.TabEnabled(1) = False
    TB1.Tab = 0
Else
    If bytMode <> 0 Then            '' If File Found and not Load Mode
        If cboEmp.Text = "" Then Exit Sub
        Call FillGrid
    End If
End If
End Sub

Private Sub cboPresent_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then SendKeys Chr(9)
End Sub

Private Sub cboShift_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then SendKeys Chr(9)
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then SendKeys Chr(9)
End Sub

Private Sub cboYear_Click()
    If cboYear.Text = "" Then Exit Sub
    Call ValidMonthYear
    If blnFileFound = False Then
        MSF1.Rows = 1
        TB1.TabEnabled(1) = False
        TB1.Tab = 0
    Else
        If bytMode <> 0 Then
            If cboEmp.Text = "" Then Exit Sub
            Call FillGrid
        End If
    End If
End Sub


Private Sub cmdExitIn_Click()
   Unload Me
End Sub

Private Sub cmdOff_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1
        '' Check for Rights
        If blnRits(6) = False Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
        Else
            '' Process
            bytMode = 6     '' Off Duty
            Call ChangeMode
        End If
    Case Else
        '' No Action
End Select
Exit Sub
ERR_P:
    ShowError ("Off Duty :: " & Me.Caption)
End Sub

Private Sub cmdOn_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1
        '' Check for Rights
        If blnRits(5) = False Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
        Else
            '' Process
            bytMode = 5     '' On Duty
            Call ChangeMode
        End If
    Case Else
        '' No Action
End Select
Exit Sub
ERR_P:
    ShowError ("On Duty :: " & Me.Caption)
End Sub

Private Sub cmdOTSave_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1
        '' Check for Rights
        If blnRits(7) = False Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        If typCOVars.bytCOCode = 100 Then
            MsgBox NewCaptionTxt("13070", adrsC), vbExclamation
            Exit Sub
        End If
        '' Process
        bytMode = 7     '' OT
        Call ChangeMode
    Case 2              '' Shift
        If Not ValidateModMaster Then Exit Sub
        If Not SaveShift Then Exit Sub
        Call AddActivityLog(lgShift_MODE, 3, 23)        '' Shift Log
        Call AuditInfo("UPDATE", Me.Caption, "Update Shift Of Employee " & cboEmp.Text & " For " & cboMonth.Text & " " & cboYear.Text)
        Call FillGrid
        bytMode = 1
        Call ChangeMode
        Call Display(lblDate.Caption)
    Case 3              '' Record
        If Not ValidateModMaster Then Exit Sub
        If Not SaveRecord Then Exit Sub
        Call AddActivityLog(lgRecord_MODE, 3, 23)       '' Record Log
        Call AuditInfo("UPDATE", Me.Caption, "Update Record Of Employee " & cboEmp.Text & " For " & cboMonth.Text & " " & cboYear.Text)
        Call FillGrid
        bytMode = 1
        Call ChangeMode
        Call Display(lblDate.Caption)
    Case 4              '' Status
        If Not ValidateModMaster Then Exit Sub
        If Not SaveOthers Then Exit Sub
        Call AddActivityLog(lgStatus_MODE, 3, 23)       '' Status Log
        Call AuditInfo("UPDATE", Me.Caption, "Update Status Of Employee " & cboEmp.Text & " For " & cboMonth.Text & " " & cboYear.Text)
        Call FillGrid
        bytMode = 1
        Call ChangeMode
        Call Display(lblDate.Caption)
    Case 5              '' On Duty
        If Not ValidateModMaster Then Exit Sub
        If Not SaveOthers Then Exit Sub
        Call AddActivityLog(lgOnDuty_MODE, 3, 23)       '' On Duty Log
        Call AuditInfo("UPDATE", Me.Caption, "Update On Duty Of Employee " & cboEmp.Text & " For " & cboMonth.Text & " " & cboYear.Text)
        Call FillGrid
        bytMode = 1
        Call ChangeMode
        Call Display(lblDate.Caption)
    Case 6              '' Off Duty
        If Not ValidateModMaster Then Exit Sub
        If Not SaveOthers Then Exit Sub
        Call AddActivityLog(lgOffDuty_MODE, 3, 23)      '' Off Duty Log
        Call AuditInfo("UPDATE", Me.Caption, "Update Off Duty Of Employee " & cboEmp.Text & " For " & cboMonth.Text & " " & cboYear.Text)
        Call FillGrid
        bytMode = 1
        Call ChangeMode
        Call Display(lblDate.Caption)
    Case 7              '' OT
        ''If typEmp.blnOT Then
        ''    If Not ValidateModMaster Then Exit Sub
        ''End If
        If Not SaveOthers Then Exit Sub
        Call AddActivityLog(lgOT_MODE, 3, 23)           '' OT Log
        Call AuditInfo("UPDATE", Me.Caption, "Update OT Of Employee " & cboEmp.Text & " For " & cboMonth.Text & " " & cboYear.Text)
        Call FillGrid
        bytMode = 1
        Call ChangeMode
        Call Display(lblDate.Caption)
    Case 8              '' Time
        If Not ValidateModMaster Then Exit Sub
        If Not SaveOthers Then Exit Sub
        Call AddActivityLog(lgTime_MODE, 3, 23)         '' Time Log
        Call AuditInfo("UPDATE", Me.Caption, "Update Time Of Employee " & cboEmp.Text & " For " & cboMonth.Text & " " & cboYear.Text & "", , strarr1, strdep1, strarr2, strdep2, dte1)
        Call FillGrid
        bytMode = 1
        Call ChangeMode
        Call Display(lblDate.Caption)
    Case Else
        '' No Action
End Select

Exit Sub
ERR_P:
    ShowError ("OT / Save :: " & Me.Caption)
    'Resume Next
End Sub

Private Sub cmdRec_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1
        '' Check for Rights
        If blnRits(3) = False Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
        Else
            '' Process
            bytMode = 3     '' Record
            Call ChangeMode
        End If
    Case Else
        '' No Action
End Select
Exit Sub
ERR_P:
    ShowError ("Record :: " & Me.Caption)
End Sub

Private Sub cmdShift_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1
        '' Check for Rights
        If blnRits(2) = False Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
        Else
            '' Process
            bytMode = 2     '' Shift
            Call ChangeMode
        End If
    Case Else
        '' No Action
End Select
Exit Sub
ERR_P:
    ShowError ("Shift :: " & Me.Caption)
End Sub

Private Sub cmdStatus_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1
        '' Check for Rights
        If blnRits(4) = False Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
        Else
            '' Process
            bytMode = 4     '' Status
            Call ChangeMode
        End If
    Case Else
        '' No Action
End Select
Exit Sub
ERR_P:
    ShowError ("Status :: " & Me.Caption)
End Sub

Private Sub cmdTimeCan_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1
        '' Check for Rights
        If blnRits(8) = False Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
        Else
            '' Process
            bytMode = 8     '' Time
            Call ChangeMode
        End If
    Case Else               '' Cancel
        bytMode = 1
        Call ChangeMode
        Call Display(lblDate.Caption)
End Select
Exit Sub
ERR_P:
    ShowError ("Time / Cancel :: " & Me.Caption)
End Sub

Private Sub Form_Load()

frmCorr.txtlate1.Visible = False
frmCorr.txtlate2.Visible = False
frmCorr.lbllate1.Visible = False
frmCorr.lbllate2.Visible = False

Call SetFormIcon(Me)        '' Sets the Form Icon
Call RetCaptions            '' Sets the Captions
Call FillCombos             '' Fills All the Combos on the Form
Call GetRights              '' Gets and Sets the Rights
Call LoadSpecifics          '' Procedure to be Executed During Load

chkPermanentCorrection.Visible = False
If strCurrentUserType <> HOD Then cboDept.Text = "ALL"
End Sub

Private Sub RetCaptions()                    '' Procedure to Retreive and Set Captions
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '13%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("13001", adrsC)              '' Sets the Forms Captions
TB1.TabCaption(0) = NewCaptionTxt("13004", adrsC)       '' Attendance Records
TB1.TabCaption(1) = NewCaptionTxt("13005", adrsC)       '' Attendance Details
Call SetOtherCaps                           '' Sets the Other Captions
Call SetButtonCap                           '' Sets the Captions to the Buttons
Call CapGrid                                '' Sets the Captions to the Grid
End Sub

Private Sub SetOtherCaps()                          '' Set Captions for the Other Controls
'' Employee Details
Call SetCritLabel(lblDeptCap)
lblMonth.Caption = NewCaptionTxt("00028", adrsMod)                '' Month
lblYear.Caption = NewCaptionTxt("00027", adrsMod)                 '' Year
'' Date,Shift,Entry Details
frDSE.Caption = NewCaptionTxt("13008", adrsC)                   '' Details
lblDateCap.Caption = NewCaptionTxt("00030", adrsMod)              '' Date
lblShift.Caption = NewCaptionTxt("00031", adrsMod)                '' Shift
lblEntry.Caption = NewCaptionTxt("00032", adrsMod)                '' Entry
'' Miscellaneous Details
frMisc.Caption = NewCaptionTxt("13009", adrsC)                  '' Misc.
lblStatus.Caption = NewCaptionTxt("00033", adrsMod)               '' Status
lblPdays.Caption = NewCaptionTxt("13010", adrsC, 0)             '' Present Days
lblRest.Caption = NewCaptionTxt("13011", adrsC)                 '' Rest Hrs
lblCO.Caption = NewCaptionTxt("13012", adrsC)                   '' CO Days
'' Time Details
frTime.Caption = NewCaptionTxt("13013", adrsC)                  '' Time
lblArr.Caption = NewCaptionTxt("00034", adrsMod)                  '' Arrival
lblDept.Caption = NewCaptionTxt("00036", adrsMod)                 '' Departure
lblWork.Caption = NewCaptionTxt("13006", adrsC)                 '' Work Hrs.
lblLate.Caption = NewCaptionTxt("00035", adrsMod)                 '' Late
lblEarly.Caption = NewCaptionTxt("00037", adrsMod)                '' Early
lblOT.Caption = NewCaptionTxt("00038", adrsMod)                   '' Overtime
'' Irregular Details
chkIrr.Caption = NewCaptionTxt("13014", adrsC)                  '' Irregular Entries
lblT2.Caption = NewCaptionTxt("13015", adrsC)                   '' 2nd
lblT3.Caption = NewCaptionTxt("13016", adrsC)                   '' 3rd
lblT4.Caption = NewCaptionTxt("13017", adrsC)                   '' 4th
lblT5.Caption = NewCaptionTxt("13018", adrsC)                   '' 5th
lblT6.Caption = NewCaptionTxt("13019", adrsC)                   '' 6th
lblT7.Caption = NewCaptionTxt("13020", adrsC)                   '' 7th
'' On Duty
frOn.Caption = NewCaptionTxt("13021", adrsC)                    '' On Duty
lblOnFrom.Caption = NewCaptionTxt("00010", adrsMod)               '' From
lblOnTo.Caption = NewCaptionTxt("00011", adrsMod)                 '' To
'' Off Duty
frOff.Caption = NewCaptionTxt("13022", adrsC)                   '' Off Duty
lblOffFrom.Caption = NewCaptionTxt("00010", adrsMod)              '' From
lblOffTo.Caption = NewCaptionTxt("00011", adrsMod)                '' To
'' Permission Cards
frPerm.Caption = NewCaptionTxt("13023", adrsC)                  '' Permission
lblPLate.Caption = NewCaptionTxt("13024", adrsC)                '' Late Card
lblPEarly.Caption = NewCaptionTxt("13025", adrsC)               '' Early Card
End Sub

Private Sub SetButtonCap(Optional bytFlgCap As Byte = 1)    '' Sets Captions to the Main
If bytFlgCap = 1 Then                                       '' Buttons
    cmdShift.Caption = "Shift"
    cmdRec.Caption = "Record"
    cmdStatus.Caption = "Status"
    cmdOn.Caption = "On Duty"
    cmdOff.Caption = "Off Duty"
    cmdOTSave.Caption = "Overtime"
    cmdTimeCan.Caption = "Time"
    cmdExitIn.Caption = "Exit"
Else
    cmdOTSave.Caption = "Save"
    cmdTimeCan.Caption = "Cancel"
End If
End Sub

Private Sub CapGrid()           '' Sets Captions to the Grid and Does Aligning & Sizing.
With MSF1                       '' Captions
    .TextMatrix(0, 0) = NewCaptionTxt("00030", adrsMod)       '' Date
    .TextMatrix(0, 1) = NewCaptionTxt("00031", adrsMod)       '' Shift
    .TextMatrix(0, 2) = NewCaptionTxt("00032", adrsMod)       '' Entry
    .TextMatrix(0, 3) = NewCaptionTxt("00033", adrsMod)       '' Status
    .TextMatrix(0, 4) = NewCaptionTxt("00034", adrsMod)       '' Arrival
    .TextMatrix(0, 5) = NewCaptionTxt("00035", adrsMod)       '' Late
    .TextMatrix(0, 6) = NewCaptionTxt("00036", adrsMod)       '' Departure
    .TextMatrix(0, 7) = NewCaptionTxt("00037", adrsMod)       '' Early
    .TextMatrix(0, 8) = NewCaptionTxt("13006", adrsC)       '' Work Hours
    .TextMatrix(0, 9) = NewCaptionTxt("00038", adrsMod)       '' Overtime
    .TextMatrix(0, 10) = NewCaptionTxt("13007", adrsC)      '' Present
End With
With MSF1                       '' Sizing
    .ColWidth(0) = .ColWidth(0) * 1
    .ColWidth(1) = .ColWidth(1) * 0.55
    .ColWidth(2) = .ColWidth(2) * 0.5
    .ColWidth(3) = .ColWidth(3) * 0.72
    .ColWidth(4) = .ColWidth(4) * 0.75
    .ColWidth(5) = .ColWidth(5) * 0.7
    .ColWidth(6) = .ColWidth(6) * 0.9
    .ColWidth(7) = .ColWidth(7) * 0.68
    .ColWidth(8) = .ColWidth(8) * 0.9
    .ColWidth(10) = .ColWidth(10) * 0.7
End With
With MSF1                       '' Alignment
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    .ColAlignment(4) = flexAlignLeftCenter
    .ColAlignment(5) = flexAlignLeftCenter
    .ColAlignment(6) = flexAlignLeftCenter
    .ColAlignment(7) = flexAlignLeftCenter
    .ColAlignment(8) = flexAlignLeftCenter
    .ColAlignment(9) = flexAlignLeftCenter
    .ColAlignment(10) = flexAlignLeftCenter
End With
End Sub

Private Sub FillCombos()            '' Fills All the ComboBoxes on the Form
On Error GoTo ERR_P
Dim intTmp As Integer
'' Employee Combo
Call FillEmpCombo
'' Shift Combo
Call FillShifts
'' Month Combo
For intTmp = 1 To 12
    cboMonth.AddItem Choose(intTmp, "January", "February", "March", "April", "May", "June" _
    , "July", "August", "September", "October", "November", "December")
Next
'' Year Combo
For intTmp = 1997 To 2096
    cboYear.AddItem CStr(intTmp)
Next
'' Entry Combo
For intTmp = 0 To 8
    cboEntry.AddItem CStr(intTmp)
Next
'' Status Combo
Call FillStatus
'' Present Days

cboPresent.AddItem "0.5"
cboPresent.AddItem "1.0"

'' CO Days
cboCO.AddItem "0.0"
cboCO.AddItem "0.5"
cboCO.AddItem "1.0"
Exit Sub
ERR_P:
    ShowError ("FillCombos :: ") & Me.Caption
End Sub

Private Sub FillEmpCombo()      '' Fills the Employee Combo
''On Error GoTo ERR_P
Call SetCritCombos(cboDept)
'If strCurrentUserType <> HOD Then cbodept.AddItem "ALL"
Exit Sub
ERR_P:
    ShowError ("FillEmpCombo :: ") & Me.Caption
End Sub

Private Sub FillShifts()        '' Fills the Shift Combo
On Error GoTo ERR_P
Dim bytTmp As Byte, strArrTmpShift() As String, bytCntTmp As Byte
bytTmp = 0
If adrsRits.State = 1 Then adrsRits.Close
adrsRits.Open "select shift,shf_in,Shf_out from instshft where shift <> '100' order by Shf_In", _
ConMain, adOpenStatic
If Not (adrsRits.BOF And adrsRits.EOF) Then
    bytTmp = adrsRits.RecordCount
    adrsRits.MoveFirst
    ReDim Preserve strArrTmpShift(bytTmp - 1, 2)
    For bytCntTmp = 0 To bytTmp - 1
        strArrTmpShift(bytCntTmp, 0) = adrsRits("Shift")
        strArrTmpShift(bytCntTmp, 1) = Format(adrsRits("Shf_In"), "00.00")
        strArrTmpShift(bytCntTmp, 2) = Format(adrsRits("Shf_Out"), "00.00")
        adrsRits.MoveNext
    Next
    cboShift.List = strArrTmpShift
    Erase strArrTmpShift
End If
Exit Sub
ERR_P:
    ShowError ("FillShifts :: ") & Me.Caption
End Sub

Private Sub FillStatus()        '' Fills the Status Combo
On Error GoTo ERR_P
Dim strStatusArray(8, 1) As String
strStatusArray(0, 0) = typVar.strAbsCode: strStatusArray(0, 1) = "Absent Days"
strStatusArray(1, 0) = "AF": strStatusArray(1, 1) = "Absent 1st Half"
strStatusArray(2, 0) = "AS": strStatusArray(2, 1) = "Absent 2nd Half"
strStatusArray(3, 0) = typVar.strHlsCode: strStatusArray(3, 1) = "Holiday Days"
strStatusArray(4, 0) = typVar.strPrsCode: strStatusArray(4, 1) = "Present Days"
strStatusArray(5, 0) = "PF": strStatusArray(5, 1) = "Present 1st Half"
strStatusArray(6, 0) = "PS": strStatusArray(6, 1) = "Present 2nd Half"
strStatusArray(7, 0) = typVar.strWosCode: strStatusArray(7, 1) = "Weekly Off"
strStatusArray(8, 0) = "": strStatusArray(8, 1) = ""
cboStatus.List = strStatusArray
Exit Sub
ERR_P:
    ShowError ("FillStatus :: " & Me.Caption)
End Sub

Private Sub ChangeState(ByRef Obj As Object, Optional bytCS As Byte = 1)
On Error GoTo ERR_P
Select Case bytCS           '' Changes the State as well as Color
    Case 1  '' Enable
        Obj.Enabled = True
        Obj.BackColor = &H80000005
    Case 2  '' Disable
        Obj.Enabled = False
        Obj.BackColor = &H8000000B
End Select
Exit Sub
ERR_P:
    ShowError ("ChangeState :: " & Me.Caption)
End Sub

Private Sub GetRights()     '' Gets and Sets the Rights for Operations
On Error GoTo ERR_P
Dim strTmp As String
blnRits(2) = False
blnRits(3) = False
blnRits(4) = False
blnRits(5) = False
blnRits(6) = False
blnRits(7) = False
blnRits(8) = False
strTmp = RetRights(4, 8, 3, 8)
'' Shift
If Mid(strTmp, 1, 1) = "1" Then
    blnRits(2) = True
End If
'' Record
If Mid(strTmp, 2, 1) = "1" Then
    blnRits(3) = True
End If
'' Status
If Mid(strTmp, 3, 1) = "1" Then
    blnRits(4) = True
End If
'' On Duty
If Mid(strTmp, 4, 1) = "1" Then
    blnRits(5) = True
End If
'' Off Duty
If Mid(strTmp, 5, 1) = "1" Then
    blnRits(6) = True
End If
'' Edit CO
If Mid(strTmp, 7, 1) = "1" Then
    blnRits(7) = True
End If
'' Time
If Mid(strTmp, 8, 1) = "1" Then
    blnRits(8) = True
End If
Exit Sub
ERR_P:
    ShowError ("GetRights :: " & Me.Caption)
    blnRits(2) = False           '' Shift
    blnRits(3) = False           '' Record
    blnRits(4) = False           '' Status
    blnRits(5) = False           '' On Duty
    blnRits(6) = False           '' Off Duty
    blnRits(7) = False           '' OT
    blnRits(8) = False           '' Time
End Sub

Private Sub ViewAction()        '' Normal / View Mode Action
'' Frame DSE
Call ChangeState(cboShift, 2)       '' Shift
Call ChangeState(cboEntry, 2)       '' Entry
'' Frame Miscellaneous.
Call ChangeState(cboStatus, 2)      '' Entry
Call ChangeState(cboPresent, 2)     '' Present Days
Call ChangeState(txtRest, 2)        '' Rest Hrs.
Call ChangeState(cboCO, 2)          '' CO Days.
'' Frame  Time
Call ChangeState(txtArr, 2)         '' Arrival
Call ChangeState(txtDept, 2)        '' Departure
Call ChangeState(txtWork, 2)        '' Work Hrs.
Call ChangeState(txtLate, 2)        '' Late
Call ChangeState(txtEarly, 2)       '' Late
Call ChangeState(txtOT, 2)          '' OT
'' Frame Irregular
chkIrr.Enabled = False
Call AdjustIrregular(2)             '' Adjust the Irregular Frame
'' Frame On
Call ChangeState(txtOnFrom, 2)      '' On From
Call ChangeState(txtOnTo, 2)        '' On To
'' Frame Off
Call ChangeState(txtOffFrom, 2)     '' Off From
Call ChangeState(txtOffTo, 2)       '' Off To
'' Frame Permission
Call ChangeState(txtPLate, 2)       '' Late card
Call ChangeState(txtPEarly, 2)      '' Early Card
Call ToggleOthers(2)                '' Enable All the Buttons

Call SetButtonCap                   '' Set Captions to the Buttons
End Sub

Private Sub ShiftAction(Optional bytCS As Byte = 1)     '' Shift Action
Call ChangeState(cboShift, bytCS)       '' Shift
Call ToggleOthers                       '' Disable Other Buttons
Call SetButtonCap(2)                    '' Change Button Captions
cboShift.SetFocus
End Sub

Private Sub StatusAction(Optional bytCS As Byte = 1)    '' Status Action
Call ChangeState(cboStatus, bytCS)          '' Status
Call ChangeState(cboPresent, bytCS)         '' Present Display
Call ToggleOthers                           '' Disable Other Buttons
Call SetButtonCap(2)                        '' Change Button Captions
cboStatus.SetFocus
End Sub

Private Sub OnDutyAction(Optional bytCS As Byte = 1)    '' On Duty Action
Call ChangeState(txtWork, bytCS)        '' Work Hours
''If typEmp.blnOT Then Call ChangeState(txtOT, bytCS)           '' OT
Call ChangeState(txtOnFrom, bytCS)      '' On From
Call ChangeState(txtOnTo, bytCS)        '' On To
Call ToggleOthers                       '' Disable Other Buttons
Call SetButtonCap(2)                    '' Change Button Captions
txtOnFrom.SetFocus
End Sub

Private Sub OffDutyAction(Optional bytCS As Byte = 1)   '' Off Duty Action
Call ChangeState(txtWork, bytCS)        '' Work Hours
''If typEmp.blnOT Then Call ChangeState(txtOT, bytCS)           '' OT
Call ChangeState(txtOffFrom, bytCS)     '' Off From
Call ChangeState(txtOffTo, bytCS)       '' Off To
Call ToggleOthers                       '' Disable Other Buttons
Call SetButtonCap(2)                    '' Change Button Captions
txtOffFrom.SetFocus
End Sub

Private Sub OTAction(Optional bytCS As Byte = 1)        '' OT Action
Call ChangeState(cboCO, bytCS)
Call ToggleOthers                           '' Disable Other Buttons
Call SetButtonCap(2)                        '' Change Button Captions
cboCO.SetFocus
End Sub

Private Sub TimeAction(Optional bytCS As Byte = 1)      '' Time Action
Call ChangeState(cboEntry, bytCS)       '' Entry
Call ChangeState(txtArr, bytCS)         '' Arrival
Call ChangeState(txtDept, bytCS)        '' Departure
Call ToggleOthers                       '' Disable Other Buttons
Call SetButtonCap(2)                    '' Change Button Captions
Call AdjustIrregular                    '' Adjust Irregular Frame
cboEntry.SetFocus
End Sub

Private Sub MSF1_DblClick()
If MSF1.Rows = 1 Then Exit Sub
TB1.Tab = 1
End Sub

Private Sub MSF1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call MSF1_DblClick
End Sub

Private Sub TB1_Click(PreviousTab As Integer)
If TB1.Tab = 0 Then Exit Sub
MSF1.Col = 0
If MSF1.Text = NewCaptionTxt("00030", adrsMod) Then Exit Sub
bytRowVal = MSF1.row
Call Display(MSF1.Text)         '' Display the Record
End Sub

Private Sub txtArr_GotFocus()
    Call GF(txtArr)
End Sub

Private Sub txtArr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtArr)
End If
End Sub

Private Sub txtDept_GotFocus()
    Call GF(txtDept)
End Sub

Private Sub txtDept_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtDept)
End If
End Sub

Private Sub txtEarly_GotFocus()
    Call GF(txtEarly)
End Sub

Private Sub txtEarly_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtEarly)
End If
End Sub

Private Sub txtEmpCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        FillGrid (txtEmpCode.Text)
    End If
End Sub

Private Sub txtLate_GotFocus()
    Call GF(txtLate)
End Sub

Private Sub txtLate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtLate)
End If
End Sub

Private Sub txtOffFrom_GotFocus()
    Call GF(txtOffFrom)
End Sub

Private Sub txtOffFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtOffFrom)
End If
End Sub

Private Sub txtOffTo_GotFocus()
    Call GF(txtOffTo)
End Sub

Private Sub txtOffTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtOffTo)
End If
End Sub

Private Sub txtOnFrom_GotFocus()
    Call GF(txtOnFrom)
End Sub

Private Sub txtOnFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtOnFrom)
End If
End Sub

Private Sub txtOnTo_GotFocus()
    Call GF(txtOnTo)
End Sub

Private Sub txtOnTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtOnTo)
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

Private Sub txtRest_GotFocus()
    Call GF(txtRest)
End Sub

Private Sub txtRest_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtRest)
End If
End Sub

Private Sub txtT2_GotFocus()
    Call GF(txtT2)
End Sub

Private Sub txtT3_GotFocus()
    Call GF(txtT3)
End Sub

Private Sub txtT4_GotFocus()
    Call GF(txtT4)
End Sub

Private Sub txtT5_GotFocus()
    Call GF(txtT5)
End Sub

Private Sub txtT6_GotFocus()
    Call GF(txtT6)
End Sub

Private Sub txtT7_GotFocus()
    Call GF(txtT7)
End Sub

Private Sub txtT2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtT2)
End If
End Sub

Private Sub txtT3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtT3)
End If
End Sub

Private Sub txtT4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtT4)
End If
End Sub

Private Sub txtT5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtT5)
End If
End Sub

Private Sub txtT6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtT6)
End If
End Sub

Private Sub txtT7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtT7)
End If
End Sub

Private Sub txtWork_GotFocus()
    Call GF(txtWork)
End Sub

Private Sub txtWork_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtWork)
End If
End Sub

Private Sub AdjustIrregular(Optional bytFlg As Byte = 1)     '' Adjust Irregular TextBoxes
If bytFlg = 1 Then     '' True
    Call ChangeState(txtT2, 1)          '' 2nd
    Call ChangeState(txtT3, 1)          '' 3rd
    Call ChangeState(txtT4, 1)          '' 4th
    Call ChangeState(txtT5, 1)          '' 5th
    Call ChangeState(txtT6, 1)          '' 6th
    Call ChangeState(txtT7, 1)          '' 7th
Else
    Call ChangeState(txtT2, 2)          '' 2nd
    Call ChangeState(txtT3, 2)          '' 3rd
    Call ChangeState(txtT4, 2)          '' 4th
    Call ChangeState(txtT5, 2)          '' 5th
    Call ChangeState(txtT6, 2)          '' 6th
    Call ChangeState(txtT7, 2)          '' 7th
End If
End Sub

Private Sub LoadSpecifics()                 '' Procedure Called when the Form is Loaded
On Error GoTo ERR_P
blnFileFound = False                        '' Initialize File not Found to False
bytMode = 0                                 '' Set Mode to 0 for Month and Year Selection
TB1.TabEnabled(1) = False                   '' Set Tab 1 to Disbled
strFileName = ""                            '' Set File Name to Blank
    cboMonth.Text = MonthName(Month(Date))      '' Set the Month to Current
    cboYear.Text = CStr(Year(Date))             '' Set the Year to Current
bytMode = 1                                 '' Set bytMode to View / Normal
Call ChangeMode                             '' Adjust According to the Mode
Call OpenMasters                            '' Open Masters
Call FillInstalltypes                       '' Fill Install types
Exit Sub
ERR_P:
    ShowError ("Load Specifics :: " & Me.Caption)
End Sub

Private Sub ValidMonthYear()    '' Procedure to Check if Valid Monthly transaction File
On Error GoTo ERR_P             '' for the Selected Month and Year is Available or not
strFileName = ""
strFileName = Left(cboMonth.Text, 3) & Right(cboYear.Text, 2) & "Trn"
If Not FindTable(Trim(strFileName)) Then
    If bytMode <> 0 Then
        MsgBox NewCaptionTxt("13035", adrsC) & cboMonth.Text, vbExclamation
    End If
        blnFileFound = False    '' File not Found
        Exit Sub
End If
blnFileFound = True             '' File Found
Exit Sub
ERR_P:
    ShowError ("ValidEmpMonthYear :: " & Me.Caption)
    blnFileFound = False
End Sub

Private Sub FillGrid()           '' Fills the Grid
On Error GoTo ERR_P
If adrsLeave.State = 1 Then adrsLeave.Close
adrsLeave.Open "Select * from " & strFileName & " Where Empcode='" & cboEmp.List(cboEmp.ListIndex, 1) & "' " & _
" order by " & strKDate & " ", ConMain, adOpenStatic
MSF1.Rows = 1
Do While Not adrsLeave.EOF
    adrsLeave.Find ("Empcode='" & cboEmp.List(cboEmp.ListIndex, 1) & "'")   '' Searches for the Specified
    If Not adrsLeave.EOF Then
        MSF1.Rows = MSF1.Rows + 1
    Else
        Exit Do
    End If
    With MSF1
        '' Code for Grid Display
        
        .TextMatrix(MSF1.Rows - 1, 0) = DateDisp(adrsLeave("Date"))               '' Date
        .TextMatrix(MSF1.Rows - 1, 1) = IIf(IsNull(adrsLeave("Shift")), "", _
                                        adrsLeave("Shift"))                       '' Shift
        .TextMatrix(MSF1.Rows - 1, 2) = adrsLeave("Entry")                        '' Entry
        .TextMatrix(MSF1.Rows - 1, 3) = adrsLeave("Presabs")                      '' Status
        .TextMatrix(MSF1.Rows - 1, 4) = IIf(IsNull(adrsLeave("ArrTim")), "0.00", _
                                        Format(adrsLeave("ArrTim"), "0.00"))      '' Arrival
        .TextMatrix(MSF1.Rows - 1, 5) = IIf(IsNull(adrsLeave("LateHrs")), "0.00", _
                                        Format(adrsLeave("LateHrs"), "0.00"))     '' Late
        .TextMatrix(MSF1.Rows - 1, 6) = IIf(IsNull(adrsLeave("DepTim")), "0.00", _
                                        Format(adrsLeave("DepTim"), "0.00"))      '' Departure
        .TextMatrix(MSF1.Rows - 1, 7) = IIf(IsNull(adrsLeave("EarlHrs")), "0.00", _
                                        Format(adrsLeave("EarlHrs"), "0.00"))     '' Early
        .TextMatrix(MSF1.Rows - 1, 8) = IIf(IsNull(adrsLeave("WrkHrs")), "0.00", _
                                        Format(adrsLeave("WrkHrs"), "0.00"))      '' Work Hours
        .TextMatrix(MSF1.Rows - 1, 9) = IIf(IsNull(adrsLeave("OvTim")), "0.00", _
                                        Format(adrsLeave("OvTim"), "0.00"))       '' Overtime
        .TextMatrix(MSF1.Rows - 1, 10) = IIf(IsNull(adrsLeave("Present")), "0.00", _
                                         Format(adrsLeave("Present"), "0.00"))    '' Present
        .RowHeight(.Rows - 1) = 205
    End With
    adrsLeave.MoveNext
Loop
If MSF1.Rows = 1 Then
    MsgBox NewCaptionTxt("13036", adrsC) & cboEmp.Text, vbExclamation
    TB1.TabEnabled(1) = False
Else
    TB1.TabEnabled(1) = True
    Call SettypEmp                      '' Set Employee Type
    Call FillEmptype(cboEmp.List(cboEmp.ListIndex, 1))       '' Fill Employee Details
    Call SettypCat                      '' Set the Category Type
    Call FillCattype(typEmp.strECat)    '' Fill Category Details
    Call FillOTType(typOTVars.bytOTCode)
    Call FillCOType(typCOVars.bytCOCode)
End If
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub Display(ByVal strDateTrn As String)        '' Display Records
On Error GoTo ERR_P
adrsLeave.MoveFirst
adrsLeave.Find "EmpCode='" & cboEmp.List(cboEmp.ListIndex, 1) & "'"
If adrsLeave.EOF Then                   '' If Employee is not Found
    MsgBox NewCaptionTxt("13037", adrsC) & strDateTrn
    TB1.Tab = 0
    Exit Sub
End If
''adrsLeave.Find "" & strKDate & " =" & strDTEnc & DateCompDate(strDateTrn) & strDTEnc
adrsLeave.Find "Date=" & strDTEnc & DateCompDate(strDateTrn) & strDTEnc
If adrsLeave.EOF Then                   '' If Date is not Found
    MsgBox NewCaptionTxt("13037", adrsC) & strDateTrn
    TB1.Tab = 0
    Exit Sub
Else
    If cboEmp.List(cboEmp.ListIndex, 1) <> adrsLeave("EmpCode") Then '' if the Employee is Different
        MsgBox NewCaptionTxt("13037", adrsC) & strDateTrn
        TB1.Tab = 0
        Exit Sub
    Else
        Call RetFields              '' Display Details
        Call SetMiscVars            '' Set the Misceellaneous Variables
        typVar.strStatus = strCapSND       '' Retreive the Status
        typVar.strShiftTmp = strRotPass    '' Retrieve the WO/HL if any
        Call CalcNumShiftPunches    '' Get Other Punches
        Call SetCOOTCaps                 '' Set CO/OT Flags
        Call SetStatus
    End If
End If
Exit Sub
ERR_P:
    ShowError ("Display :: " & Me.Caption)
    ''Resume Next
End Sub

Private Sub SetStatus()
    Select Case adrsLeave("Presabs")
    Case typVar.strPrsCode & typVar.strPrsCode
        cboStatus.Text = typVar.strPrsCode
    Case typVar.strAbsCode & typVar.strAbsCode
        cboStatus.Text = typVar.strAbsCode
    Case typVar.strWosCode & typVar.strWosCode
        cboStatus.Text = typVar.strWosCode
    Case typVar.strHlsCode & typVar.strHlsCode
        cboStatus.Text = typVar.strHlsCode
    Case typVar.strAbsCode & Right(typVar.strStatus, 2)
        cboStatus.Text = "AF"
    Case Left(typVar.strStatus, 2) & typVar.strAbsCode
        cboStatus.Text = "AS"
    Case typVar.strPrsCode & Right(typVar.strStatus, 2)
        cboStatus.Text = "PF"
    Case Left(typVar.strStatus, 2) & typVar.strPrsCode
        cboStatus.Text = "PS"
    Case Else
        cboStatus.Text = ""
    End Select
End Sub

Private Sub RetFields()         '' Retreives from the RecordSets
On Error GoTo ERR_P
'' Date,Shift,Entry Details
lblDate.Caption = DateDisp(adrsLeave("Date"))                   '' Date
Call DisOthers                                                  '' Shift
cboEntry.Value = adrsLeave("Entry")                             '' Entry
'' Miscellaneous Details
Call DisOthers(2)                                               '' Status
cboPresent.Value = IIf(IsNull(adrsLeave("Present")), "0.0", _
                   Format(adrsLeave("Present"), "0.0"))         '' Present
txtRest.Text = IIf(IsNull(adrsLeave("ActBreak")), "0.00", _
               Format(adrsLeave("ActBreak"), "0.00"))           '' Break Hours
cboCO.Value = IIf(IsNull(adrsLeave("Cof")), "0.0", _
              Format(adrsLeave("Cof"), "0.0"))                  '' CO Days
'' Time Details (ADWLEO)
txtArr.Text = IIf(IsNull(adrsLeave("ArrTim")), "0.00", _
              Format(adrsLeave("ArrTim"), "0.00"))              '' Arrival Time
txtDept.Text = IIf(IsNull(adrsLeave("DepTim")), "0.00", _
               Format(adrsLeave("DepTim"), "0.00"))             '' Departure Time
txtWork.Text = IIf(IsNull(adrsLeave("WrkHrs")), "0.00", _
               Format(adrsLeave("WrkHrs"), "0.00"))             '' Work Hours
txtLate.Text = IIf(IsNull(adrsLeave("LateHrs")), "0.00", _
               Format(adrsLeave("LateHrs"), "0.00"))            '' Late Hours

txtEarly.Text = IIf(IsNull(adrsLeave("EarlHrs")), "0.00", _
                Format(adrsLeave("EarlHrs"), "0.00"))           '' Early Hours
txtOT.Text = IIf(IsNull(adrsLeave("OvTim")), "0.00", _
             Format(adrsLeave("OvTim"), "0.00"))                '' Overtime
'' Irregular Details
chkIrr.Value = IIf(adrsLeave("Chq") = "*", 1, 0)                '' Irregular Check Box
txtT2.Text = IIf(IsNull(adrsLeave("actrt_o")), "0.00", _
             Format(adrsLeave("actrt_o"), "0.00"))              '' Break In
txtT3.Text = IIf(IsNull(adrsLeave("actrt_i")), "0.00", _
             Format(adrsLeave("actrt_i"), "0.00"))              '' Break Out
txtT4.Text = IIf(IsNull(adrsLeave("Time5")), "0.00", _
             Format(adrsLeave("Time5"), "0.00"))                '' Time5
txtT5.Text = IIf(IsNull(adrsLeave("Time6")), "0.00", _
             Format(adrsLeave("Time6"), "0.00"))                '' Time6
txtT6.Text = IIf(IsNull(adrsLeave("Time7")), "0.00", _
             Format(adrsLeave("Time7"), "0.00"))                '' Time7
txtT7.Text = IIf(IsNull(adrsLeave("Time8")), "0.00", _
             Format(adrsLeave("Time8"), "0.00"))                '' Time8
'' On Duty
txtOnFrom.Text = IIf(IsNull(adrsLeave("Od_From")), "0.00", _
                 Format(adrsLeave("Od_From"), "0.00"))          '' On Duty From
txtOnTo.Text = IIf(IsNull(adrsLeave("Od_To")), "0.00", _
               Format(adrsLeave("Od_To"), "0.00"))              '' On Duty To
'' Off Duty
txtOffFrom.Text = IIf(IsNull(adrsLeave("Ofd_From")), "0.00", _
                  Format(adrsLeave("Ofd_From"), "0.00"))        '' Off Duty From
txtOffTo.Text = IIf(IsNull(adrsLeave("Ofd_To")), "0.00", _
                Format(adrsLeave("Ofd_To"), "0.00"))            '' Off Duty To
'' Permission Cards
txtPLate.Text = IIf(IsNull(adrsLeave("Aflg")), "", _
               adrsLeave("Aflg"))                               '' Aflg
txtPEarly.Text = IIf(IsNull(adrsLeave("Dflg")), "", _
               adrsLeave("Dflg"))                               '' Dflg
'' OT Remarks
typOTVars.strOTRem = IIf(IsNull(adrsLeave("OTRem")), "", _
               adrsLeave("OTRem"))                               '' OT Remarks
Exit Sub
ERR_P:
    Select Case Err.Number
    Case 380        '' If invalid text Property is Set to the ComboBox
        MsgBox NewCaptionTxt("13038", adrsC), vbCritical
        'Resume Next
    Case Else
        ShowError ("RetFields :: " & Me.Caption)
    End Select
    Resume Next
End Sub

Private Sub DisOthers(Optional bytFlg As Byte = 1)      '' Displays the Other Details
On Error GoTo ERR_P
Select Case bytFlg
    Case 1
        cboShift.Value = IIf(IsNull(adrsLeave("Shift")), "", _
                         Trim(adrsLeave("Shift")))            '' Shift
    Case 2
        cboStatus.Value = ""                            '' Status
        If InStr(MSF1.TextMatrix(bytRowVal, 3), typVar.strWosCode) > 0 Then
            typVar.strShiftTmp = typVar.strWosCode                '' Checks for WO
        Else
            typVar.strShiftTmp = ""
        End If
        If InStr(MSF1.TextMatrix(bytRowVal, 3), typVar.strHlsCode) > 0 Then
            typVar.strShiftTmp = typVar.strHlsCode                '' Checks for HL
        End If
        '' strCapSND = MSF1.TextMatrix(MSF1.Row, 3)        '' Gets the Status
        strCapSND = MSF1.TextMatrix(bytRowVal, 3)        '' Gets the Status
        strRotPass = ""
        strRotPass = typVar.strShiftTmp                        '' Gets the WO/HL
End Select
Exit Sub
ERR_P:
End Sub

Private Sub ChangeMode()        '' Procedure when the Mode Changes
On Error GoTo ERR_P
Select Case bytMode
    Case 1      '' Normal / View
        Call ViewAction
    Case 2      '' Shift
        Call ShiftAction
    Case 3      '' Record
        Call RecAction
    Case 4      '' Status
        Call StatusAction
    Case 5      '' On
        Call OnDutyAction
    Case 6      '' Off
        Call OffDutyAction
    Case 7      '' OT
        Call OTAction
    Case 8      '' Time
        Call TimeAction
End Select
Exit Sub
ERR_P:
    ShowError ("ChangeMode :: " & Me.Caption)
End Sub

Private Sub ToggleOthers(Optional bytFlg As Byte = 1)   '' Toggles the State of Buttons
On Error GoTo ERR_P
Select Case bytFlg
    Case 1          '' Disables
        cmdShift.Enabled = False
        cmdRec.Enabled = False
        cmdStatus.Enabled = False
        cmdOn.Enabled = False
        cmdOff.Enabled = False
        frEmp.Enabled = False
        txtlate1.Enabled = True          ' for IOC
        txtlate2.Enabled = True
    Case 2          '' Enables
        cmdShift.Enabled = True
        cmdRec.Enabled = True
        cmdStatus.Enabled = True
        cmdOn.Enabled = True
        cmdOff.Enabled = True
        frEmp.Enabled = True
        txtlate1.Enabled = False
        txtlate2.Enabled = False
End Select
Exit Sub
ERR_P:
    ShowError ("Toggle Others :: " & Me.Caption)
End Sub

Private Function ValidateModMaster() As Boolean
On Error GoTo ERR_P         '' Validates Records before Updating
ValidateModMaster = True
Select Case bytMode
    Case 2          '' Shift
        If Trim(cboShift.Text) = "" Then
            MsgBox NewCaptionTxt("13039", adrsC), vbExclamation
            cboShift.SetFocus
            ValidateModMaster = False
            Exit Function
        End If
        typVar.strStatus = ""
        Call SettypShift                        '' Set Shift Type
        typVar.strShiftOfDay = cboShift.Text           '' Put the Shift
        Call FillShifttype(typVar.strShiftOfDay)       '' Fill Shift TYpe
    Case 3          '' Record
        Call MakeFormats        '' Checks for Valid Times
        If Not CheckDecimals(1) Then ValidateModMaster = False: Exit Function
        If Not CheckDecimals(2) Then ValidateModMaster = False: Exit Function
        If Not CheckDecimals(3) Then ValidateModMaster = False: Exit Function
        If Not CheckDecimals(4) Then ValidateModMaster = False: Exit Function
        If Not CheckDecimals(5) Then ValidateModMaster = False: Exit Function
        If Not CheckDecimals(7) Then ValidateModMaster = False: Exit Function
        If Not CheckDecimals(8) Then ValidateModMaster = False: Exit Function
        Select Case Trim(txtPLate.Text)
            Case "", "1", "3"
            Case Else
                MsgBox NewCaptionTxt("13040", adrsC), vbExclamation
                txtPLate.SetFocus
                ValidateModMaster = False: Exit Function
        End Select
        Select Case Trim(txtPEarly.Text)
            Case "", "2"
            Case Else
                MsgBox NewCaptionTxt("13041", adrsC), vbExclamation
                txtPEarly.SetFocus
                ValidateModMaster = False: Exit Function
        End Select
    Case 4          '' Status
    Case 5          '' On
        Call MakeFormats        '' Checks for Valid Times
        If Not CheckDecimals(6) Then ValidateModMaster = False: Exit Function
        If Not CheckDecimals(4) Then ValidateModMaster = False: Exit Function
    Case 6          '' Off
        Call MakeFormats        '' Checks for Valid Times
        If Not CheckDecimals(6) Then ValidateModMaster = False: Exit Function
        If Not CheckDecimals(5) Then ValidateModMaster = False: Exit Function
    Case 7          '' OT
        Call MakeFormats        '' Checks for Valid Times
        If Not CheckDecimals(6) Then ValidateModMaster = False: Exit Function
    Case 8          '' Time
        Call MakeFormats        '' Checks for Valid Times
        If Not CheckDecimals(2) Then ValidateModMaster = False: Exit Function
        If Not CheckDecimals(3) Then ValidateModMaster = False: Exit Function
        If Not CheckDecimals(4) Then ValidateModMaster = False: Exit Function
        If Not CheckDecimals(5) Then ValidateModMaster = False: Exit Function
        If Not CheckDecimals(7) Then ValidateModMaster = False: Exit Function
        If Not CheckDecimals(8) Then ValidateModMaster = False: Exit Function
        typVar.strShiftOfDay = cboShift.Text           '' Put the Shift
        Call SettypShift                        '' Set Shift Type
        Call FillShifttype(typVar.strShiftOfDay)       '' Fill Shift TYpe
End Select
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
    ValidateModMaster = False
End Function

Private Function SaveShift() As Boolean '' Saves Record in the Shift Mode
On Error GoTo ERR_P
SaveShift = True
'' Save the Record
Call GetStatus(1)                  '' Get Status"
Call DoProcess      '' Do Process
Exit Function
ERR_P:
    ShowError ("SaveShift :: " & Me.Caption)
    SaveShift = False
End Function

Private Sub CalcNumShiftPunches()           '' Gets the Details Necessary for Process
Dim bytTmp As Byte
Call SettypDH
'' Set OT & CO Vars
bytTmp = 0
strArr = typTR.sngTimeIn
strdep = typTR.sngTimeOut
strarr1 = strArr
strdep1 = strdep
dte1 = DateCompDate(lblDate.Caption)
typTR.sngTimeIn = Val(txtArr.Text)
If Val(txtArr.Text) > 0 Then bytTmp = bytTmp + 1                '' In Time
typTR.sngTimeOut = Val(txtDept.Text)
If Val(txtDept.Text) > 0 Then bytTmp = bytTmp + 1               '' Out Time
typTR.sngBreakE = Val(txtT2.Text)
If Val(txtT2.Text) > 0 Then bytTmp = bytTmp + 1                 '' T2, Break In
typTR.sngBreakS = Val(txtT3.Text)
If Val(txtT3.Text) > 0 Then bytTmp = bytTmp + 1                 '' T3, Break Out
typTR.sngTime5 = Val(txtT4.Text)
If Val(txtT4.Text) > 0 Then bytTmp = bytTmp + 1                 '' T4
typTR.sngTime6 = Val(txtT5.Text)
If Val(txtT5.Text) > 0 Then bytTmp = bytTmp + 1                 '' T5
typTR.sngTime7 = Val(txtT6.Text)
If Val(txtT6.Text) > 0 Then bytTmp = bytTmp + 1                 '' T6
typTR.sngTime8 = Val(txtT7.Text)
If Val(txtT7.Text) > 0 Then bytTmp = bytTmp + 1                 '' T7
typTR.sngODFrom = Val(txtOnFrom.Text)                           '' On From
typTR.sngODTo = Val(txtOnTo.Text)                               '' On To
typTR.sngOFDFrom = Val(txtOffFrom.Text)                         '' Off From
typTR.sngOFDTo = Val(txtOffTo.Text)                             '' Off To
typVar.strAflg = txtPLate.Text                                         '' Late Card
typVar.strDflg = txtPEarly.Text                                        '' Early Card
typDT.dtFrom = DateCompDate(lblDate.Caption)                    '' Valid No. of Punches
typVar.bytTmpEnt = bytTmp                                              '' Total no. of Punches
End Sub

Private Function SaveRecord() As Boolean    '' Saves Record in the Record Mode
On Error GoTo ERR_P
SaveRecord = True
'' Update Statement
ConMain.Execute "Update " & strFileName & " Set Entry=" & Val(cboEntry.Text) & _
",EntReq=" & typEmp.bytEntry & ",Shift='" & cboShift.Text & "',ArrTim=" & txtArr.Text & _
",LateHrs=" & txtLate.Text & ",Actrt_O=" & txtT2.Text & ",Actrt_I=" & txtT3.Text & _
",ActBreak=" & txtRest.Text & ",Deptim=" & txtDept.Text & ",EarlHrs=" & txtEarly.Text & _
",WrkHrs=" & txtWork.Text & ",OvTim=" & txtOT.Text & ",Time5=" & txtT4.Text & ",Time6=" & _
txtT5.Text & ",Time7=" & txtT6.Text & ",Time8=" & txtT7.Text & ",Od_From=" & txtOnFrom.Text & _
",Od_To=" & txtOnTo.Text & ",Ofd_From=" & txtOffFrom.Text & ",Ofd_To=" & txtOffTo.Text & _
",Ofd_Hrs=" & TimDiff(Val(txtOffTo.Text), Val(txtOffFrom.Text)) & ",Present=" & cboPresent.Text & _
",Presabs='" & RetStatus & "',Aflg='" & txtPLate.Text & "',Dflg='" & txtPEarly.Text & _
"',Cof=" & cboCO.Text & ",Chq='" & IIf(chkIrr.Value = 1, "*", "") & _
"',Remarks=Remarks " & StrKConcat & " '" & IIf(InStr(adrsLeave("Remarks"), "-C") > 0, "", "-C") & _
"' Where EmpCode='" & cboEmp.List(cboEmp.ListIndex, 1) & "' and " & strKDate & " =" & strDTEnc & _
DateCompStr(lblDate.Caption) & strDTEnc

Exit Function
ERR_P:
    ShowError ("SaveRecord :: " & Me.Caption)
    SaveRecord = False
End Function

Private Sub RecAction()     '' Procedure Called When the Form Gets in Record Mode
Call ShiftAction
Call StatusAction
Call OnDutyAction
Call OffDutyAction
Call TimeAction
Call OTAction
Call ChangeState(txtRest)
Call ChangeState(txtLate)
Call ChangeState(txtEarly)
Call ChangeState(txtPLate)
Call ChangeState(txtPEarly)
chkIrr.Enabled = True
Call AdjustIrregular
cboShift.SetFocus
End Sub

Private Sub MakeFormats()   '' Formats All the Time type Data to 0.00 Format
On Error GoTo ERR_P
'' Frame Details
txtRest.Text = IIf(txtRest.Text = "", "0.00", Format(txtRest.Text, "0.00"))
'' Frame Time
txtlate1.Text = IIf(txtlate1.Text = "", "0.00", Format(txtlate1.Text, "0.00"))
txtlate2.Text = IIf(txtlate2.Text = "", "0.00", Format(txtlate2.Text, "0.00"))
txtArr.Text = IIf(txtArr.Text = "", "0.00", Format(txtArr.Text, "0.00"))
txtDept.Text = IIf(txtDept.Text = "", "0.00", Format(txtDept.Text, "0.00"))
txtWork.Text = IIf(txtWork.Text = "", "0.00", Format(txtWork.Text, "0.00"))
txtLate.Text = IIf(txtLate.Text = "", "0.00", Format(txtLate.Text, "0.00"))
txtEarly.Text = IIf(txtEarly.Text = "", "0.00", Format(txtEarly.Text, "0.00"))
txtOT.Text = IIf(txtOT.Text = "", "0.00", Format(txtOT.Text, "0.00"))
'' Frame Irregular
txtT2.Text = IIf(txtT2.Text = "", "0.00", Format(txtT2.Text, "0.00"))
txtT3.Text = IIf(txtT3.Text = "", "0.00", Format(txtT3.Text, "0.00"))
txtT4.Text = IIf(txtT4.Text = "", "0.00", Format(txtT4.Text, "0.00"))
txtT5.Text = IIf(txtT5.Text = "", "0.00", Format(txtT5.Text, "0.00"))
txtT6.Text = IIf(txtT6.Text = "", "0.00", Format(txtT6.Text, "0.00"))
txtT7.Text = IIf(txtT7.Text = "", "0.00", Format(txtT7.Text, "0.00"))
'' Frame On Duty
txtOnFrom.Text = IIf(txtOnFrom.Text = "", "0.00", Format(txtOnFrom.Text, "0.00"))
txtOnTo.Text = IIf(txtOnTo.Text = "", "0.00", Format(txtOnTo.Text, "0.00"))
'' Frame Off Duty
txtOffFrom.Text = IIf(txtOffFrom.Text = "", "0.00", Format(txtOffFrom.Text, "0.00"))
txtOffTo.Text = IIf(txtOffTo.Text = "", "0.00", Format(txtOffTo.Text, "0.00"))
Exit Sub
ERR_P:
    ShowError ("Make Formats :: " & Me.Caption)
End Sub

Private Function CheckDecimals(ByVal bytFlg As Byte) As Boolean
On Error GoTo ERR_P
CheckDecimals = True        '' Checks for Decimals and Other Valid Values
Select Case bytFlg
    Case 1  '' Rest Hrs
        If Val(Right(txtRest.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtRest.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtRest.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtRest.SetFocus
            CheckDecimals = False
            Exit Function
        End If
    Case 2  '' Frame Time
        If Val(Right(txtArr.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtArr.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtArr.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtArr.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(Right(txtDept.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtDept.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtDept.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtDept.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtArr.Text) > 0 Then
            If Val(txtDept.Text) <= 0 Then
                MsgBox NewCaptionTxt("13044", adrsC), vbExclamation
                txtDept.SetFocus
                CheckDecimals = False
                Exit Function
            End If
        End If
        If Val(txtDept.Text) > 0 Then
            If Val(txtArr.Text) > Val(txtDept.Text) Then
                MsgBox NewCaptionTxt("13045", adrsC), vbExclamation
                txtArr.SetFocus
                CheckDecimals = False
                Exit Function
            End If
            If Val(txtArr.Text) <= 0 Then
                MsgBox NewCaptionTxt("13046", adrsC), vbExclamation
                txtArr.SetFocus
                CheckDecimals = False
                Exit Function
            End If
        End If
        If Val(Right(txtWork.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtWork.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtWork.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtWork.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(Right(txtLate.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtLate.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtLate.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtLate.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(Right(txtEarly.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtEarly.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtEarly.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtEarly.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(Right(txtOT.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtOT.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtOT.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtOT.SetFocus
            CheckDecimals = False
            Exit Function
        End If
    Case 3  '' Frame Irregular
        If Val(Right(txtT2.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtT2.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtT2.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtT2.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtT2.Text) > 0 And (Val(txtT2.Text) < Val(txtArr.Text) Or _
        Val(txtT2.Text) > Val(txtDept.Text)) Then
            MsgBox NewCaptionTxt("13047", adrsC), vbExclamation
            txtT2.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(Right(txtT3.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtT3.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtT3.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtT3.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If (Val(txtT3.Text) > 0) And (Val(txtT3.Text) < Val(txtArr.Text) Or _
        Val(txtT3.Text) > Val(txtDept.Text)) Then
            MsgBox NewCaptionTxt("13047", adrsC), vbExclamation
            txtT3.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(Right(txtT4.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtT4.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtT4.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtT4.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtT4.Text) > 0 And (Val(txtT4.Text) < Val(txtArr.Text) Or _
        Val(txtT4.Text) > Val(txtDept.Text)) Then
            MsgBox NewCaptionTxt("13047", adrsC), vbExclamation
            txtT4.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(Right(txtT5.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtT5.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtT5.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtT5.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtT5.Text) > 0 And (Val(txtT5.Text) < Val(txtArr.Text) Or _
        Val(txtT5.Text) > Val(txtDept.Text)) Then
            MsgBox NewCaptionTxt("13047", adrsC), vbExclamation
            txtT5.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(Right(txtT6.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtT6.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtT6.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtT6.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtT6.Text) > 0 And (Val(txtT6.Text) < Val(txtArr.Text) Or _
        Val(txtT6.Text) > Val(txtDept.Text)) Then
            MsgBox NewCaptionTxt("13047", adrsC), vbExclamation
            txtT6.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(Right(txtT7.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtT7.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtT7.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtT7.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtT7.Text) > 0 And (Val(txtT7.Text) < Val(txtArr.Text) Or _
        Val(txtT7.Text) > Val(txtDept.Text)) Then
            MsgBox NewCaptionTxt("13047", adrsC), vbExclamation
            txtT7.SetFocus
            CheckDecimals = False
            Exit Function
        End If
    Case 4  '' Frame On Duty
        If Val(Right(txtOnFrom.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtOnFrom.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtOnFrom.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtOnFrom.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(Right(txtOnTo.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtOnTo.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtOnTo.Text) > 0 And Val(txtOnTo.Text) < Val(txtOnFrom.Text) Then
            MsgBox NewCaptionTxt("13048", adrsC), vbExclamation
            txtOnTo.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtOnTo.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtOnTo.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtOnTo.Text) > 0 And Val(txtOnFrom.Text) > Val(txtOnTo.Text) Then
            MsgBox NewCaptionTxt("13049", adrsC), vbExclamation
            txtOnFrom.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtOnTo.Text) > 0 And Val(txtOnFrom.Text) = 0 Then
            MsgBox NewCaptionTxt("13050", adrsC), vbExclamation
            txtOnFrom.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtOnFrom.Text) > 0 Then
            If Val(txtOnFrom.Text) < Val(txtArr.Text) Then
                MsgBox NewCaptionTxt("13051", adrsC), vbExclamation
                If bytMode = 8 Then txtArr.SetFocus
                If bytMode = 3 Or bytMode = 5 Then txtOnFrom.SetFocus
                CheckDecimals = False
                Exit Function
            End If
            If Val(txtOnFrom.Text) > Val(txtDept.Text) Then
                MsgBox NewCaptionTxt("13051", adrsC), vbExclamation
                If bytMode = 8 Then txtDept.SetFocus
                If bytMode = 3 Or bytMode = 5 Then txtOnFrom.SetFocus
                CheckDecimals = False
                Exit Function
            End If
        End If
        If Val(txtOnTo.Text) > 0 Then
            If Val(txtOnTo.Text) < Val(txtArr.Text) Then
                MsgBox NewCaptionTxt("13052", adrsC), vbExclamation
                If bytMode = 8 Then txtArr.SetFocus
                If bytMode = 3 Or bytMode = 5 Then txtOnTo.SetFocus
                CheckDecimals = False
                Exit Function
            End If
            If Val(txtOnTo.Text) > Val(txtDept.Text) Then
                MsgBox NewCaptionTxt("13052", adrsC), vbExclamation
                If bytMode = 8 Then txtDept.SetFocus
                If bytMode = 3 Or bytMode = 5 Then txtOnTo.SetFocus
                CheckDecimals = False
                Exit Function
            End If
        End If
    Case 5  '' Frame Off Duty
        If Val(Right(txtOffFrom.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtOffFrom.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtOffFrom.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtOffFrom.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(Right(txtOffTo.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtOffTo.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtOffTo.Text) > 0 And Val(txtOffTo.Text) < Val(txtOffFrom.Text) Then
            MsgBox NewCaptionTxt("13048", adrsC), vbExclamation
            txtOffTo.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtOffTo.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtOffTo.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtOffTo.Text) > 0 And Val(txtOffFrom.Text) > Val(txtOffTo.Text) Then
            MsgBox NewCaptionTxt("13048", adrsC), vbExclamation
            txtOffFrom.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtOffTo.Text) > 0 And Val(txtOffFrom.Text) = 0 Then
            MsgBox NewCaptionTxt("13053", adrsC), vbExclamation
            txtOffFrom.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtOffFrom.Text) > 0 Then
            If Val(txtOffFrom.Text) < Val(txtArr.Text) Then
                MsgBox NewCaptionTxt("13054", adrsC), vbExclamation
                If bytMode = 8 Then txtArr.SetFocus
                If bytMode = 3 Or bytMode = 6 Then txtOffFrom.SetFocus
                CheckDecimals = False
                Exit Function
            End If
            If Val(txtOffFrom.Text) > Val(txtDept.Text) Then
                MsgBox NewCaptionTxt("13054", adrsC), vbExclamation
                If bytMode = 8 Then txtDept.SetFocus
                If bytMode = 3 Or bytMode = 6 Then txtOffFrom.SetFocus
                CheckDecimals = False
                Exit Function
            End If
        End If
        If Val(txtOffTo.Text) > 0 Then
            If Val(txtOffTo.Text) < Val(txtArr.Text) Then
                MsgBox NewCaptionTxt("13055", adrsC), vbExclamation
                If bytMode = 8 Then txtArr.SetFocus
                If bytMode = 3 Or bytMode = 6 Then txtOffTo.SetFocus
                CheckDecimals = False
                Exit Function
            End If
            If Val(txtOffTo.Text) > Val(txtDept.Text) Then
                MsgBox NewCaptionTxt("13055", adrsC), vbExclamation
                If bytMode = 8 Then txtDept.SetFocus
                If bytMode = 3 Or bytMode = 6 Then txtOffTo.SetFocus
                CheckDecimals = False
                Exit Function
            End If
        End If
    Case 6      '' Overtime Check
        If Val(Right(txtOT.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            txtOT.SetFocus
            CheckDecimals = False
            Exit Function
        End If
        If Val(txtOT.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            txtOT.SetFocus
            CheckDecimals = False
            Exit Function
        End If
    Case 7      '' T2 is 0
        If Val(txtT3.Text) > 0 And Val(txtT2.Text) = 0 Then
            MsgBox NewCaptionTxt("13056", adrsC), vbExclamation
            txtT2.SetFocus
            CheckDecimals = False
            Exit Function
        End If  '' T3 is 0
        If Val(txtT4.Text) > 0 And Val(txtT3.Text) = 0 Then
            MsgBox NewCaptionTxt("13057", adrsC), vbExclamation
            txtT3.SetFocus
            CheckDecimals = False
            Exit Function
        End If  '' T4 is 0
        If Val(txtT5.Text) > 0 And Val(txtT4.Text) = 0 Then
            MsgBox NewCaptionTxt("13058", adrsC), vbExclamation
            txtT4.SetFocus
            CheckDecimals = False
            Exit Function
        End If  '' T5 is 0
        If Val(txtT6.Text) > 0 And Val(txtT5.Text) = 0 Then
            MsgBox NewCaptionTxt("13059", adrsC), vbExclamation
            txtT5.SetFocus
            CheckDecimals = False
            Exit Function
        End If  '' T6 is 0
        If Val(txtT7.Text) > 0 And Val(txtT6.Text) = 0 Then
            MsgBox NewCaptionTxt("13060", adrsC), vbExclamation
            txtT6.SetFocus
            CheckDecimals = False
            Exit Function
        End If
    Case 8      '' T2 > T3
        If Val(txtT3.Text) > 0 And Val(txtT2.Text) > Val(txtT3.Text) Then
            MsgBox NewCaptionTxt("13061", adrsC), vbExclamation
            txtT2.SetFocus
            CheckDecimals = False
            Exit Function
        End If  '' T3 > T4
        If Val(txtT4.Text) > 0 And Val(txtT3.Text) > Val(txtT4.Text) Then
            MsgBox NewCaptionTxt("13062", adrsC), vbExclamation
            txtT3.SetFocus
            CheckDecimals = False
            Exit Function
        End If  '' T4 > T5
        If Val(txtT5.Text) > 0 And Val(txtT4.Text) > Val(txtT5.Text) Then
            MsgBox NewCaptionTxt("13063", adrsC), vbExclamation
            txtT4.SetFocus
            CheckDecimals = False
            Exit Function
        End If  '' T5 > T6
        If Val(txtT6.Text) > 0 And Val(txtT5.Text) > Val(txtT6.Text) Then
            MsgBox NewCaptionTxt("13064", adrsC), vbExclamation
            txtT5.SetFocus
            CheckDecimals = False
            Exit Function
        End If  '' T6 > T7
        If Val(txtT7.Text) > 0 And Val(txtT6.Text) > Val(txtT7.Text) Then
            MsgBox NewCaptionTxt("13065", adrsC), vbExclamation
            txtT6.SetFocus
            CheckDecimals = False
            Exit Function
        End If
End Select
Exit Function
ERR_P:
    ShowError ("Check Decimals :: " & Me.Caption)
    CheckDecimals = False
End Function

Private Function RetStatus() As String      '' Returns Updated Status
On Error GoTo ERR_P
Select Case cboStatus.Text
    Case ""                 '' Empty String
        RetStatus = typVar.strStatus
    Case "AF"               '' Absent First Half
        RetStatus = typVar.strAbsCode & Right(typVar.strStatus, 2)
    Case "AS"               '' Absent Second Half
        RetStatus = Left(typVar.strStatus, 2) & typVar.strAbsCode
    Case "PF"               '' Present First Half
        RetStatus = typVar.strPrsCode & Right(typVar.strStatus, 2)
    Case "PS"               '' Present Second Half
        RetStatus = Left(typVar.strStatus, 2) & typVar.strPrsCode
    Case typVar.strAbsCode     '' Total Absent
        RetStatus = typVar.strAbsCode & typVar.strAbsCode
    Case typVar.strHlsCode     '' Total Holiday
        RetStatus = typVar.strHlsCode & typVar.strHlsCode
    Case typVar.strPrsCode     '' Total Present
        RetStatus = typVar.strPrsCode & typVar.strPrsCode
    Case typVar.strWosCode     '' Total Week Off
        RetStatus = typVar.strWosCode & typVar.strWosCode
    Case Else               '' Unknown Case
        RetStatus = typVar.strStatus
End Select
Exit Function
ERR_P:
    ShowError ("Return Status :: " & Me.Caption)
    RetStatus = typVar.strStatus
End Function

Private Function SaveOthers() As Boolean    '' Saves Record in Status,On Duty,Off Duty,
On Error GoTo ERR_P                         '' OT Mode
SaveOthers = True
Select Case bytMode
    Case 4          '' Status
        ConMain.Execute "update " & strFileName & " Set Presabs='" & RetStatus & _
        "',Present=" & cboPresent.Text & ",Remarks=Remarks " & StrKConcat & " '" & IIf(InStr(adrsLeave("Remarks"), "-C") > 0, "", "-C") & _
            "' Where EmpCode='" & cboEmp.List(cboEmp.ListIndex, 1) & "' and " & strKDate & " =" & _
        strDTEnc & DateCompStr(lblDate.Caption) & strDTEnc
    Case 5          '' On Duty
        ConMain.Execute "update " & strFileName & " Set Od_From =" & _
        txtOnFrom.Text & ",Od_To=" & txtOnTo.Text & ",WrkHrs=" & txtWork.Text & ",OvTim=" & _
        txtOT.Text & ",Remarks=Remarks " & StrKConcat & " '" & IIf(InStr(adrsLeave("Remarks"), "-C") > 0, "", "-C") & _
            "' Where EmpCode='" & cboEmp.List(cboEmp.ListIndex, 1) & "' and " & strKDate & " =" & strDTEnc & _
        DateCompStr(lblDate.Caption) & strDTEnc
    Case 6          '' Off Duty
        ConMain.Execute "update " & strFileName & " Set Ofd_From =" & _
        txtOffFrom.Text & ",Ofd_To=" & txtOffTo.Text & ",WrkHrs=" & txtWork.Text & ",OvTim=" & _
        txtOT.Text & ",Remarks=Remarks " & StrKConcat & " '" & IIf(InStr(adrsLeave("Remarks"), "-C") > 0, "", "-C") & _
            "' Where EmpCode='" & cboEmp.List(cboEmp.ListIndex, 1) & "' and " & strKDate & " =" & strDTEnc & _
        DateCompStr(lblDate.Caption) & strDTEnc
    Case 7          '' OT
        typDH.sngCOHrs = Val(cboCO.Text)
        If typDH.sngCOHrs > 0 Then
            Call PutCOHrs
        Else
            Call SetZeroCO      '' Set the CO to Zero
        End If
        ConMain.Execute "update " & strFileName & " Set cof=" & _
        typDH.sngCOHrs & ",Remarks=Remarks " & StrKConcat & " '" & IIf(InStr(adrsLeave("Remarks"), "-C") > 0, "", "-C") & _
        "' Where EmpCode='" & cboEmp.List(cboEmp.ListIndex, 1) & "' and " & strKDate & " =" & strDTEnc & _
        DateCompStr(lblDate.Caption) & strDTEnc
    Case 8
        Call CalcNumShiftPunches        '' Calculate Details Needed for Process
        typVar.bytTmpEnt = Val(cboEntry.Text)  '' Take the Entries
        Call SettypDH                   '' Set type DH
        Call SetBreakHours              '' Set Type BreakHours
        Call DoProcess      '' Do Process
End Select
       
Exit Function
ERR_P:
    ShowError ("SaveOthers :: " & Me.Caption)
    SaveOthers = False
End Function

Private Sub DoProcess()
On Error GoTo ERR_P
Select Case typEmp.bytEntry         '' Depending Upon the Entry Required
    Case 0      '' 0 Entry
        
        Call GetStatus(1)
        ''
        Call GetHoursZeroEnt(1)
        Exit Sub
    Case 1      '' 1 Entry
        Call GetHoursOneEnt
        Exit Sub
    Case Else
        '' Do Nothing
End Select
Select Case typVar.bytTmpEnt               '' Depending Upon the Entries Found
    Case 0:
        typVar.strRemarks = typVar.strRemarks & "-C"
        Call AddRecordsToTrn(1)
        Exit Sub
    Case 1, 3, 5, 7, 9              '' Odd Entries
        Call GetStatus(1)           '' Get Status for Odd Entries
        Call GetLateHours(1)           '' Get Late Hours
        Call GetPresent
        Call GetIrrMark
    Case Else
        Call ProcessHours
End Select
Call PutHours           '' Put Hours
typVar.strRemarks = typVar.strRemarks & "-C"
''For Mauritius 11-08-2003
''If bytMode = 2 Then
''    Call UpdateShift
''End If
Call AddRecordsToTrn(1)
Exit Sub
ERR_P:
    ShowError ("Do Process :: " & Me.Caption)
End Sub

Private Sub SetCOOTCaps()   '' Set OT/CO Captions
cmdOTSave.Caption = "CO"
End Sub

Private Sub SetZeroCO()     '' Set CO to Zero
If FindTable("LvBal" & Right(pVStar.YearSel, 2)) Then
    If adrsPaid.State = 1 Then adrsPaid.Close
    adrsPaid.Open "Select LvCode,Run_Wrk From LeavDesc Where LvCode='CO' and Cat='" & _
    typCat.strCat & "'", ConMain, adOpenStatic
    If Not (adrsPaid.EOF And adrsPaid.BOF) Then
        If FieldExists("LvBal" & Right(pVStar.YearSel, 2), "CO") Then
            If adrsDept1.State = 1 Then adrsDept1.Close
            adrsDept1.Open "Select * from LvInfo" & Right(pVStar.YearSel, 2) & _
            " Where EmpCode='" & typEmp.strEmp & "' and FromDate=" & strDTEnc & _
            DateCompStr(typDT.dtFrom) & strDTEnc & " and LCode='" & _
            adrsPaid("LvCode") & "' and Trcd=2", ConMain
            If Not (adrsDept1.EOF And adrsDept1.BOF) Then
                ConMain.Execute "Delete from LvInfo" & Right(pVStar.YearSel, 2) & _
                " Where EmpCode='" & typEmp.strEmp & "' and FromDate=" & strDTEnc & _
                DateCompStr(typDT.dtFrom) & strDTEnc & " and LCode='" & _
                adrsPaid("LvCode") & "' and Trcd=2", ConMain
                ConMain.Execute "update LvBal" & Right(pVStar.YearSel, 2) & _
                " Set CO=CO-" & IIf(IsNull(adrsDept1("days")), "0", adrsDept1("days")) & _
                " Where EmpCode='" & typEmp.strEmp & "'" & " and CO is not NULL"
                ConMain.Execute "update LvBal" & Right(pVStar.YearSel, 2) & _
                " Set CO=" & IIf(IsNull(adrsDept1("days")), "0", adrsDept1("days")) & _
                " Where EmpCode='" & typEmp.strEmp & "'" & " and CO is NULL"
                  
            End If
        Else
            MsgBox NewCaptionTxt("13066", adrsC), vbExclamation
        End If
    Else
        MsgBox NewCaptionTxt("13067", adrsC), vbExclamation
    End If
Else
    MsgBox NewCaptionTxt("13068", adrsC) & vbCrLf & _
    NewCaptionTxt("13069", adrsC), vbExclamation
End If
End Sub
''
