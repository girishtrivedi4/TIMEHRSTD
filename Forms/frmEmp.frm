VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmEmp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   DrawWidth       =   10
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   435
      Left            =   7200
      TabIndex        =   19
      Top             =   5850
      Width           =   2475
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command1"
      Height          =   435
      Left            =   4800
      TabIndex        =   18
      Top             =   5850
      Width           =   2415
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command1"
      Height          =   435
      Left            =   2400
      TabIndex        =   17
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command1"
      Height          =   435
      Left            =   0
      TabIndex        =   16
      Top             =   5850
      Width           =   2415
   End
   Begin MSMAPI.MAPIMessages EmpMessage 
      Left            =   930
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
   Begin MSMAPI.MAPISession EmpSession 
      Left            =   1755
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin TabDlg.SSTab TB1 
      Height          =   5205
      Left            =   0
      TabIndex        =   21
      Top             =   600
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   9181
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
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
      TabPicture(0)   =   "frmEmp.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command3"
      Tab(0).Control(1)=   "Command2"
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(3)=   "MSF1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Official Details"
      TabPicture(1)   =   "frmEmp.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fr1T"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Personal Details"
      TabPicture(2)   =   "frmEmp.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "PerDetailsFrame"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Other Details"
      TabPicture(3)   =   "frmEmp.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fr3Det"
      Tab(3).ControlCount=   1
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70800
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Click to sort on this Column"
         Top             =   430
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69600
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Click to sort on this Column"
         Top             =   430
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74140
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Click to sort on this Column"
         Top             =   430
         Width           =   255
      End
      Begin VB.Frame PerDetailsFrame 
         Height          =   4005
         Left            =   60
         TabIndex        =   66
         Top             =   360
         Width           =   8535
         Begin VB.Frame frAdd 
            Height          =   2070
            Left            =   60
            TabIndex        =   77
            Top             =   1890
            Width           =   8370
            Begin VB.TextBox txtmedical 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   4920
               MaxLength       =   15
               TabIndex        =   34
               Tag             =   "D"
               Text            =   " "
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox txtAdd1 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   1305
               MaxLength       =   50
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   30
               Top             =   180
               Width           =   2070
            End
            Begin VB.TextBox txtPin 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   1320
               MaxLength       =   6
               TabIndex        =   81
               Text            =   " "
               Top             =   1620
               Width           =   2055
            End
            Begin VB.TextBox txtAdd2 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   1320
               MaxLength       =   50
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   31
               Top             =   660
               Width           =   2070
            End
            Begin MSMask.MaskEdBox txtPhone 
               Height          =   345
               Left            =   4920
               TabIndex        =   33
               Top             =   210
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               AutoTab         =   -1  'True
               MaxLength       =   15
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
            Begin MSMask.MaskEdBox txtCity 
               Height          =   345
               Left            =   1320
               TabIndex        =   32
               Top             =   1140
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               AutoTab         =   -1  'True
               MaxLength       =   14
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
            Begin VB.Label lblmedical 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Medical Date"
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
               Left            =   3720
               TabIndex        =   98
               Top             =   720
               Width           =   1140
            End
            Begin VB.Label lblPhone 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone No"
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
               Left            =   3840
               TabIndex        =   82
               Top             =   240
               Width           =   855
            End
            Begin VB.Label lblPin 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pin Code"
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
               Left            =   420
               TabIndex        =   80
               Top             =   1650
               Width           =   795
            End
            Begin VB.Label lblCity 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "City"
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
               Left            =   450
               TabIndex        =   79
               Top             =   1230
               Width           =   345
            End
            Begin VB.Label lblAdd 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
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
               Left            =   420
               TabIndex        =   78
               Top             =   270
               Width           =   720
            End
         End
         Begin VB.Frame frSal 
            Height          =   645
            Left            =   60
            TabIndex        =   74
            Top             =   1245
            Width           =   8385
            Begin MSMask.MaskEdBox txtRef 
               Height          =   330
               Left            =   3810
               TabIndex        =   29
               Top             =   210
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               AutoTab         =   -1  'True
               MaxLength       =   20
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
            Begin MSMask.MaskEdBox txtSal 
               Height          =   315
               Left            =   1320
               TabIndex        =   28
               Top             =   240
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               AutoTab         =   -1  'True
               MaxLength       =   5
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
            Begin VB.Label lblRef 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reference"
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
               TabIndex        =   76
               Top             =   240
               Width           =   870
            End
            Begin VB.Label lblSal 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Basic Salary"
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
               TabIndex        =   75
               Top             =   240
               Width           =   1110
            End
         End
         Begin VB.Frame fr2Det 
            Height          =   1125
            Left            =   60
            TabIndex        =   67
            Top             =   120
            Width           =   8400
            Begin VB.TextBox txtEmail 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   5970
               MaxLength       =   45
               TabIndex        =   27
               Top             =   720
               Width           =   2355
            End
            Begin VB.TextBox txtDOB 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   1320
               MaxLength       =   15
               TabIndex        =   22
               Tag             =   "D"
               Text            =   " "
               Top             =   210
               Width           =   1215
            End
            Begin VB.TextBox txtDOJ 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   3840
               MaxLength       =   15
               TabIndex        =   23
               Tag             =   "D"
               Text            =   " "
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox txtConf 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   3840
               MaxLength       =   15
               TabIndex        =   26
               Tag             =   "D"
               Text            =   " "
               Top             =   720
               Width           =   1215
            End
            Begin MSMask.MaskEdBox txtBlood 
               Height          =   315
               Left            =   1320
               TabIndex        =   25
               Top             =   720
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
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
               PromptChar      =   "_"
            End
            Begin VB.Label lblSex 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sex"
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
               Left            =   5160
               TabIndex        =   72
               Top             =   300
               Width           =   345
            End
            Begin VB.Label lblConf 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Confirm Date"
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
               Left            =   2640
               TabIndex        =   71
               Top             =   750
               Width           =   1125
            End
            Begin VB.Label lblDOJ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date Of Join"
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
               Left            =   2640
               TabIndex        =   70
               Top             =   300
               Width           =   1065
            End
            Begin VB.Label lblBlood 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Blood group"
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
               TabIndex        =   69
               Top             =   735
               Width           =   1035
            End
            Begin VB.Label lblDOB 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date Of Birth"
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
               TabIndex        =   68
               Top             =   285
               Width           =   1125
            End
            Begin VB.Label lblEmail 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Email"
               Height          =   195
               Left            =   5160
               TabIndex        =   73
               Top             =   780
               Width           =   375
            End
            Begin MSForms.ComboBox cboSex 
               Height          =   375
               Left            =   5970
               TabIndex        =   24
               Top             =   240
               Width           =   645
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "1138;661"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
      End
      Begin VB.Frame fr1T 
         Height          =   4635
         Left            =   -74880
         TabIndex        =   49
         Top             =   360
         Width           =   14565
         Begin VB.CheckBox chkAuto 
            Caption         =   "Auto Shift Change"
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
            Left            =   120
            TabIndex        =   112
            Top             =   3960
            Width           =   1935
         End
         Begin VB.Frame frAuto 
            Height          =   735
            Left            =   2160
            TabIndex        =   109
            Top             =   3720
            Width           =   2775
            Begin VB.CommandButton cmdAuto 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2160
               TabIndex        =   111
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtAutoG 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   110
               Top             =   240
               Width           =   2055
            End
         End
         Begin VB.Frame fr1Det 
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
            Left            =   405
            TabIndex        =   50
            Top             =   135
            Width           =   8265
            Begin VB.TextBox txtName2 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   4305
               MaxLength       =   20
               TabIndex        =   4
               Text            =   " "
               Top             =   1000
               Width           =   3630
            End
            Begin VB.TextBox txtName 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4320
               MaxLength       =   50
               TabIndex        =   2
               Top             =   210
               Width           =   3630
            End
            Begin MSMask.MaskEdBox txtCard 
               Height          =   315
               Left            =   1200
               TabIndex        =   1
               Top             =   600
               Width           =   1110
               _ExtentX        =   1958
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
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtCode 
               Height          =   315
               Left            =   1200
               TabIndex        =   0
               Top             =   210
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               AllowPrompt     =   -1  'True
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
            Begin MSForms.ComboBox cboDesig 
               Height          =   315
               Left            =   4320
               TabIndex        =   3
               Top             =   600
               Width           =   3615
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "6376;556"
               ColumnCount     =   2
               MatchEntry      =   0
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label lblName2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fathers Name"
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
               Left            =   2610
               TabIndex        =   55
               Top             =   1035
               Width           =   1260
            End
            Begin VB.Label lblCode 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Code No"
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
               Top             =   240
               Width           =   750
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Card No"
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
               Top             =   630
               Width           =   705
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
               Left            =   2640
               TabIndex        =   53
               Top             =   255
               Width           =   510
            End
            Begin VB.Label lblDesc 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Designation"
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
               Left            =   2640
               TabIndex        =   54
               Top             =   630
               Width           =   1020
            End
         End
         Begin VB.Frame fr1Iden 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2205
            Left            =   0
            TabIndex        =   56
            Top             =   1440
            Width           =   9300
            Begin VB.TextBox txtFreeF 
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   1200
               MaxLength       =   8
               TabIndex        =   113
               Top             =   1560
               Width           =   1545
            End
            Begin VB.Frame fraSal 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   2880
               TabIndex        =   102
               Top             =   1395
               Visible         =   0   'False
               Width           =   6315
               Begin MSMask.MaskEdBox txtAvgDays 
                  Height          =   315
                  Left            =   5400
                  TabIndex        =   106
                  Top             =   210
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   556
                  _Version        =   393216
                  Appearance      =   0
                  AllowPrompt     =   -1  'True
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
               Begin MSMask.MaskEdBox txtNewSal 
                  Height          =   315
                  Left            =   3360
                  TabIndex        =   107
                  Top             =   210
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
                  _Version        =   393216
                  Appearance      =   0
                  AllowPrompt     =   -1  'True
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
               Begin MSMask.MaskEdBox txtOldSal 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   108
                  Top             =   210
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
                  _Version        =   393216
                  Appearance      =   0
                  AllowPrompt     =   -1  'True
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
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Avg Days"
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
                  Left            =   4560
                  TabIndex        =   105
                  Top             =   247
                  Width           =   945
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "New Salary"
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
                  TabIndex        =   104
                  Top             =   247
                  Width           =   990
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Old Salary"
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
                  Left            =   120
                  TabIndex        =   103
                  Top             =   187
                  Width           =   915
               End
            End
            Begin VB.Label lblSGroup 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SAP Code"
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
               Left            =   240
               TabIndex        =   114
               Top             =   1560
               Width           =   915
            End
            Begin MSForms.ComboBox cboCompany 
               Height          =   315
               Left            =   6960
               TabIndex        =   13
               Top             =   960
               Width           =   2265
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "3995;556"
               MatchEntry      =   0
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboCat 
               Height          =   315
               Left            =   3840
               TabIndex        =   8
               Top             =   240
               Width           =   2055
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "3625;556"
               ColumnCount     =   2
               MatchEntry      =   0
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboEnt 
               Height          =   285
               Left            =   1080
               TabIndex        =   5
               Top             =   255
               Width           =   1665
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "2937;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label lblEnt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Min. Entry"
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
               TabIndex        =   101
               Top             =   277
               Width           =   900
            End
            Begin VB.Label lblCat 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               Left            =   3000
               TabIndex        =   100
               Top             =   277
               Width           =   780
            End
            Begin MSForms.ComboBox cboLoca 
               Height          =   315
               Left            =   6960
               TabIndex        =   11
               Top             =   240
               Width           =   2265
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "3995;556"
               MatchEntry      =   0
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label lblLoca 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               Left            =   6120
               TabIndex        =   99
               Top             =   277
               Width           =   735
            End
            Begin MSForms.ComboBox cboGroup 
               Height          =   315
               Left            =   3840
               TabIndex        =   10
               Top             =   960
               Width           =   2055
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "3625;556"
               ColumnCount     =   2
               MatchEntry      =   0
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboDept 
               Height          =   315
               Left            =   3840
               TabIndex        =   9
               Top             =   600
               Width           =   2055
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "3625;556"
               ColumnCount     =   2
               MatchEntry      =   0
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   180
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboDiv 
               Height          =   315
               Left            =   6960
               TabIndex        =   12
               Top             =   600
               Width           =   2265
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "3995;556"
               MatchEntry      =   0
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboCORule 
               Height          =   285
               Left            =   1080
               TabIndex        =   7
               Top             =   975
               Width           =   1665
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "2937;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label lblCORule 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CO Rule"
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
               Left            =   240
               TabIndex        =   58
               Top             =   997
               Width           =   735
            End
            Begin MSForms.ComboBox cboOTRule 
               Height          =   285
               Left            =   1080
               TabIndex        =   6
               Top             =   615
               Width           =   1665
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "2937;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label lblOTRule 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "OT Rule"
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
               TabIndex        =   57
               Top             =   637
               Width           =   705
            End
            Begin VB.Label lblComp 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Company"
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
               Left            =   6120
               TabIndex        =   62
               Top             =   960
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.Label lblDiv 
               AutoSize        =   -1  'True
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
               Height          =   330
               Left            =   6240
               TabIndex        =   61
               Top             =   592
               Width           =   660
            End
            Begin VB.Label lblDept 
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
               Height          =   240
               Left            =   2760
               TabIndex        =   59
               Top             =   637
               Width           =   1005
            End
            Begin VB.Label lblGroup 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               Left            =   3240
               TabIndex        =   60
               Top             =   997
               Width           =   525
            End
         End
         Begin VB.Frame frSch 
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
            Left            =   5002
            TabIndex        =   63
            Top             =   3720
            Width           =   2010
            Begin VB.CommandButton cmdSch 
               Caption         =   "Define Shedule"
               Height          =   375
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   1785
            End
         End
         Begin VB.Frame frLeft 
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
            Left            =   7080
            TabIndex        =   64
            Top             =   3720
            Width           =   2235
            Begin VB.TextBox txtLeft 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1080
               MaxLength       =   15
               TabIndex        =   15
               Tag             =   "D"
               Text            =   " "
               Top             =   330
               Width           =   975
            End
            Begin VB.Label lblLeft 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Left Date"
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
               Left            =   240
               TabIndex        =   65
               Top             =   375
               Width           =   780
            End
         End
      End
      Begin VB.Frame fr3Det 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   -74610
         TabIndex        =   83
         Top             =   420
         Width           =   8055
         Begin MSMask.MaskEdBox txtNat 
            Height          =   375
            Left            =   5760
            TabIndex        =   42
            Top             =   1980
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox txtState 
            Height          =   375
            Left            =   5760
            TabIndex        =   40
            Top             =   1410
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox txtRoad 
            Height          =   375
            Left            =   5760
            TabIndex        =   38
            Top             =   870
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox txtArea 
            Height          =   360
            Left            =   5760
            TabIndex        =   36
            Top             =   330
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox txtTel 
            Height          =   360
            Left            =   2280
            TabIndex        =   41
            Top             =   1980
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   12
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
         Begin MSMask.MaskEdBox txtDist 
            Height          =   375
            Left            =   2280
            TabIndex        =   39
            Top             =   1410
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox txtVill 
            Height          =   375
            Left            =   2280
            TabIndex        =   37
            Top             =   840
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox txtSpl 
            Height          =   375
            Index           =   1
            Left            =   4200
            TabIndex        =   44
            Top             =   2940
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox txtSpl 
            Height          =   345
            Index           =   0
            Left            =   750
            TabIndex        =   43
            Top             =   2940
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox txtHouse 
            Height          =   345
            Left            =   2280
            TabIndex        =   35
            Top             =   330
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
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
         Begin VB.Label lblHouse 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "House No / Name"
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
            TabIndex        =   84
            Top             =   360
            Width           =   1545
         End
         Begin VB.Label lblVill 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City / Village"
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
            Left            =   615
            TabIndex        =   85
            Top             =   870
            Width           =   1110
         End
         Begin VB.Label lblDist 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "District"
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
            Left            =   1095
            TabIndex        =   86
            Top             =   1500
            Width           =   615
         End
         Begin VB.Label lblTel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tel No"
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
            Left            =   1080
            TabIndex        =   87
            Top             =   2040
            Width           =   555
         End
         Begin VB.Label lblArea 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Area"
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
            Left            =   4800
            TabIndex        =   88
            Top             =   390
            Width           =   405
         End
         Begin VB.Label lblRoad 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Road"
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
            Left            =   4800
            TabIndex        =   89
            Top             =   960
            Width           =   450
         End
         Begin VB.Label lblState 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "State"
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
            Left            =   4800
            TabIndex        =   90
            Top             =   1530
            Width           =   465
         End
         Begin VB.Label lblNat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nationality"
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
            Left            =   4725
            TabIndex        =   91
            Top             =   2070
            Width           =   915
         End
         Begin VB.Label lblSpl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Special Comments"
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
            Left            =   3480
            TabIndex        =   92
            Top             =   2550
            Width           =   1785
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   4575
         Left            =   -75000
         TabIndex        =   45
         Top             =   360
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   8070
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   12632256
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
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   3975
      Left            =   9960
      TabIndex        =   96
      Top             =   1200
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   12632256
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
   Begin MSForms.ComboBox cboName 
      Height          =   375
      Left            =   6120
      TabIndex        =   47
      Top             =   45
      Width           =   3285
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "5794;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cbocode 
      Height          =   315
      Left            =   3210
      TabIndex        =   46
      Top             =   45
      Width           =   1485
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2619;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   255
      Left            =   3210
      TabIndex        =   97
      Top             =   45
      Width           =   1485
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2619;450"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblCodeCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Left            =   30
      TabIndex        =   20
      Top             =   120
      Width           =   570
   End
   Begin VB.Label lblNameCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Left            =   4800
      TabIndex        =   48
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "frmEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private blnCodeMatch As Boolean
''
Dim adrsC As New ADODB.Recordset
Dim strSort As String

Private Sub cboCode_Click()
On Error GoTo ERR_P
If cboCode.Text = "" Then Exit Sub      '' if Nothing is Selected
If MSF1.Rows = 1 Then Exit Sub
adrsEmp.MoveFirst
adrsEmp.Find ("EmpCode Like '" & cboCode.Text & "%'")      '' Find the Employee on Code
If adrsEmp.EOF Then                                 '' If Record is not Found
    MsgBox NewCaptionTxt("23047", adrsC), vbInformation
    cboCode.Text = txtCode.Text
    Exit Sub
End If
blnCodeMatch = True
If TB1.Tab <> 1 Then
    TB1.Tab = 1
Else
    If adrsEmp("EmpCode") = txtCode.Text Then Exit Sub  '' If Record is already Displayed
    Call Display                                        '' Display the Record
End If
blnCodeMatch = False
Exit Sub
ERR_P:
    ShowError ("Code :: " & Me.Caption)
End Sub

Private Sub cboCode_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then Call cboCode_LostFocus
End Sub

Private Sub cboCode_LostFocus()
    Call cboCode_Click
End Sub

Private Sub cboName_Click()
On Error GoTo ERR_P
If cboName.Text = "" Then Exit Sub          '' If Nothing is Selected
If MSF1.Rows = 1 Then Exit Sub
adrsEmp.MoveFirst
adrsEmp.Find ("Name Like '" & cboName.Text & "%'")
If adrsEmp.EOF Then                         '' Find the Employee on Name
    MsgBox NewCaptionTxt("23047", adrsC), vbInformation
    cboName.Text = Trim(txtName.Text)
    Exit Sub
End If
blnCodeMatch = True
If TB1.Tab <> 1 Then
    TB1.Tab = 1
Else
    If adrsEmp("EmpCode") = txtCode.Text Then Exit Sub  '' If Record is already Displayed
    Call Display                                        '' Display the Record
End If
blnCodeMatch = False
Exit Sub
ERR_P:
    ShowError ("Name : " & Me.Caption)
End Sub

Private Sub cboname_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then Call cboName_LostFocus
End Sub

Private Sub cboName_LostFocus()
    Call cboName_Click
End Sub

Private Sub chkAuto_Click()

If chkAuto.Value = 1 Then
    frAuto.Visible = True
Else
    frAuto.Visible = False
End If
''
End Sub

Private Sub cmdAuto_Click()
strAutoG = strAutoG & "Em"
frmSelectShift.Show vbModal
If Right(strAutoG, 2) = "Em" Then strAutoG = Left(strAutoG, Len(strAutoG) - 2)
txtAutoG.Text = strAutoG
End Sub
''
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSch_Click()
On Error GoTo ERR_P
If txtCode.Text = "" Then
    MsgBox NewCaptionTxt("23048", adrsC), vbExclamation
    If bytMode = 2 Then txtCode.SetFocus
    Exit Sub
End If
bytShfMode = 1
Shft.Empcode = txtCode.Text
frmEmpShift.Show vbModal
Exit Sub
ERR_P:
    ShowError ("Schedule :: " & Me.Caption)
End Sub

Private Sub Command1_Click()
strSort = "Empcode"
Call FillGrid
End Sub

Private Sub Command2_Click()
strSort = "Card"
Call FillGrid
End Sub

Private Sub Command3_Click()
strSort = "[Name]"
Call FillGrid
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)                        '' Sets the Form Icon
Call RetCaptions                            '' Gets and Sets the Form Captions
Call GetRights                              '' Gets and Sets the Form Rights
Call FillCombos                             '' Fills the Necessary Combos                       '' Opens the Employee Master Table
strSort = "Empcode"
Call OpenMasterTable
Call FillGrid                               '' Fill the Master Grid
Call LoadSpecifics                          '' Other Actions to be Taken on Form Load
End Sub


Private Sub MSF1_DblClick()
    blnCodeMatch = False
    TB1.Tab = 1
End Sub

Private Sub TB1_Click(PreviousTab As Integer)
If bytMode = 2 Then Exit Sub        '' If Add Mode
If bytMode > 10 Then                '' If Temporary Mode to set the Tab Visible
    bytMode = bytMode - 10
    Exit Sub
End If
If TB1.Tab = 2 Or TB1.Tab = 3 Then  '' If Personal Or Others Tab is Clicked
    If MSF1.Text <> "" Then Exit Sub '' If Already Record is Present
End If
If TB1.Tab = 0 Then Exit Sub     '' If it is the List Tab
If TB1.Tab = 1 Then
If MSF1.Text = txtCode.Text Then Exit Sub
End If
If Not blnCodeMatch Then            '' If Selection is not From Code & Name Combos
    MSF1.Col = 0                    '' Set to the First Column
    If MSF1.Text = NewCaptionTxt("23007", adrsC) Then Exit Sub
    If MSF1.Text = txtCode.Text Then Exit Sub   '' If record is Already Displayed
    adrsEmp.MoveFirst                           '' Move to the First Record
    adrsEmp.Find "EmpCode='" & MSF1.Text & "'"  '' Find the Record Selected
End If
Call Display                                     '' Display the Record
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtAdd1_GotFocus()
    Call GF(txtAdd1)
End Sub

Private Sub txtAdd2_GotFocus()
    Call GF(txtAdd2)
End Sub

Private Sub txtAdd1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 8)
End If
End Sub

Private Sub txtAdd2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 8)
End If
End Sub

Private Sub txtArea_GotFocus()
    Call GF(txtArea)
End Sub

Private Sub txtArea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 8)
End If
End Sub

Private Sub txtAutoG_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 6)
End If
End Sub
''
Private Sub txtBlood_GotFocus()
    Call GF(txtBlood)
End Sub

Private Sub txtBlood_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 8)
End If
End Sub

Private Sub txtCard_GotFocus()
    Call GF(txtCard)
End Sub

Private Sub txtCard_KeyPress(KeyAscii As Integer)
''this AlphaNumeicEmpCode add by
'If Not GetFlagStatus("AlphaNumeicEmpCard") Then
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    Else
        KeyAscii = KeyPressCheck(KeyAscii, 2)
    End If
'End If
End Sub

Private Sub txtCity_GotFocus()
    Call GF(txtCity)
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 1)
End If
End Sub

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


Private Sub txtmedical_click()
varCalDt = ""
varCalDt = Trim(txtmedical.Text)
txtmedical.Text = ""
Call ShowCalendar
End Sub

Private Sub txtmedical_GotFocus()
If TB1.Tab <> 2 Then
        bytMode = bytMode + 10
        TB1.Tab = 2
    End If
    Call GF(txtmedical)
End Sub

Private Sub txtmedical_KeyPress(KeyAscii As Integer)
Call CDK(txtmedical, KeyAscii)
End Sub

Private Sub txtmedical_Validate(Cancel As Boolean)
If Not ValidDate(txtmedical) Then txtmedical.SetFocus: Cancel = True
End Sub

Private Sub txtName2_GotFocus()
    Call GF(txtName2)
End Sub

Private Sub txtName2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 3))))
End If
End Sub

Private Sub txtFreeF_GotFocus()
    Call GF(txtFreeF)
End Sub

Private Sub txtFreeF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 3))))
End If
End Sub

Private Sub txtDist_GotFocus()
    Call GF(txtDist)
End Sub

Private Sub txtDist_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 8)
End If
End Sub

Private Sub txtEmail_GotFocus()
    Call GF(txtEmail)
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 7)
End If
End Sub

Private Sub txtHouse_GotFocus()
    If TB1.Tab <> 3 Then
        bytMode = bytMode + 10
        TB1.Tab = 3
    End If
    Call GF(txtHouse)
End Sub

Private Sub txtHouse_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 8)
End If
End Sub

Private Sub txtLeft_Click()
varCalDt = ""
varCalDt = Trim(txtLeft.Text)
txtLeft.Text = ""
Call ShowCalendar
End Sub

Private Sub txtLeft_GotFocus()
    Call GF(txtLeft)
End Sub

Private Sub txtLeft_KeyPress(KeyAscii As Integer)
    Call CDK(txtLeft, KeyAscii)
End Sub

Private Sub txtLeft_Validate(Cancel As Boolean)
    If Not ValidDate(txtLeft) Then txtLeft.SetFocus: Cancel = True
End Sub

Private Sub txtDOJ_Click()
varCalDt = ""
varCalDt = Trim(txtDOJ.Text)
txtDOJ.Text = ""
Call ShowCalendar
End Sub

Private Sub txtDOJ_GotFocus()
    If TB1.Tab <> 2 Then
        bytMode = bytMode + 10
        TB1.Tab = 2
    End If
    Call GF(txtDOJ)
End Sub

Private Sub txtDOJ_KeyPress(KeyAscii As Integer)
    Call CDK(txtDOJ, KeyAscii)
End Sub

Private Sub txtDOJ_Validate(Cancel As Boolean)
    If Not ValidDate(txtDOJ) Then txtDOJ.SetFocus: Cancel = True
End Sub

Private Sub txtDOB_Click()
varCalDt = ""
varCalDt = Trim(txtDOB.Text)
txtDOB.Text = ""
Call ShowCalendar
End Sub

Private Sub txtDOB_GotFocus()
    If TB1.Tab <> 2 Then
        bytMode = bytMode + 10
        TB1.Tab = 2
    End If
    Call GF(txtDOB)
End Sub

Private Sub txtDOB_KeyPress(KeyAscii As Integer)
    Call CDK(txtDOB, KeyAscii)
End Sub

Private Sub txtDOB_Validate(Cancel As Boolean)
    If Not ValidDate(txtDOB) Then txtDOB.SetFocus: Cancel = True
End Sub

Private Sub txtConf_Click()
varCalDt = ""
varCalDt = Trim(txtConf.Text)
txtConf.Text = ""
Call ShowCalendar
End Sub

Private Sub txtConf_GotFocus()
    Call GF(txtConf)
End Sub

Private Sub txtConf_KeyPress(KeyAscii As Integer)
    Call CDK(txtConf, KeyAscii)
End Sub

Private Sub txtConf_Validate(Cancel As Boolean)
    If Not ValidDate(txtConf) Then txtConf.SetFocus: Cancel = True
End Sub

Private Sub txtName_GotFocus()
    Call GF(txtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 3)
End If
End Sub

Private Sub txtNat_GotFocus()
    Call GF(txtNat)
End Sub

Private Sub txtNat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 1)
End If
End Sub

Private Sub txtPhone_GotFocus()
    Call GF(txtPhone)
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 2)
End If
End Sub

Private Sub txtPin_GotFocus()
    Call GF(txtPin)
End Sub

Private Sub txtPin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 2)
End If
End Sub

Private Sub txtRef_GotFocus()
    Call GF(txtRef)
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 3)
End If
End Sub

Private Sub txtRoad_GotFocus()
    Call GF(txtRoad)
End Sub

Private Sub txtRoad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 8)
End If
End Sub

Private Sub txtSal_GotFocus()
    Call GF(txtSal)
End Sub

Private Sub txtSal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 4)
End If
End Sub

Private Sub txtSpl_GotFocus(Index As Integer)
    Call GF(txtSpl(Index))
End Sub

Private Sub txtSpl_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 8)
End If
End Sub

Private Sub txtState_GotFocus()
    Call GF(txtState)
End Sub

Private Sub txtState_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 1)
End If
End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 2)
End If
End Sub

Private Sub txtVill_GotFocus()
    Call GF(txtVill)
End Sub

Private Sub txtVill_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 8)
End If
End Sub

Private Sub RetCaptions()                   '' Gets and Sets the Form Captions
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '23%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("23001", adrsC)              '' Form Caption
'' Tab Captions
TB1.TabCaption(0) = NewCaptionTxt("00013", adrsMod)       '' List
TB1.TabCaption(1) = NewCaptionTxt("23004", adrsC)       '' Official Details
TB1.TabCaption(2) = NewCaptionTxt("23005", adrsC)       '' Personal Details
TB1.TabCaption(3) = NewCaptionTxt("23006", adrsC)       '' Other Details
Call OtherCaptions                          '' Sets the Captions for Other Controls
Call CapGrid
End Sub

Private Sub OtherCaptions()     '' Sets Captions to the Other Controls
'' General
lblCodeCap.Caption = NewCaptionTxt("23002", adrsC)          '' Find Employee With ...
lblNameCap.Caption = NewCaptionTxt("23003", adrsC)          '' or Having Name
'' Tab 1
'' Details Frame
fr1Det.Caption = NewCaptionTxt("00014", adrsMod)              '' Details
lblCode.Caption = NewCaptionTxt("23012", adrsC)             '' Code No.
lblCard.Caption = NewCaptionTxt("23013", adrsC)             '' Card No.
lblName.Caption = NewCaptionTxt("00048", adrsMod)             '' Name
lblDesc.Caption = NewCaptionTxt("23014", adrsC)             '' Designation
lblName2.Caption = NewCaptionTxt("00120", adrsMod)            '' Fathers Name
'' Identity Frame
'fr1Iden.Caption = NewCaptionTxt("23015", adrsC)             '' Identification
lblEnt.Caption = NewCaptionTxt("23016", adrsC)              '' Min. Entry
lblOTRule.Caption = NewCaptionTxt("00090", adrsMod)           '' Min. Entry
lblCORule.Caption = NewCaptionTxt("00091", adrsMod)           '' Min. Entry
chkAuto.Caption = NewCaptionTxt("23018", adrsC)             '' AutoShift Change
lblCat.Caption = NewCaptionTxt("00051", adrsMod)              '' Category
lblDiv.Caption = NewCaptionTxt("23020", adrsC)              '' Division
lblComp.Caption = NewCaptionTxt("00057", adrsMod)             '' Company Code
lblDept.Caption = NewCaptionTxt("00058", adrsMod)             '' Department
lblGroup.Caption = NewCaptionTxt("00059", adrsMod)            '' Group
lblLoca.Caption = NewCaptionTxt("00110", adrsMod)             '' Location
'' Schedule Frame
frSch.Caption = NewCaptionTxt("23021", adrsC)               '' Working Schedule
cmdSch.Caption = NewCaptionTxt("23022", adrsC)              '' Define Schedule
'' Left frame
frLeft.Caption = NewCaptionTxt("23023", adrsC)              '' Past Employee
lblLeft.Caption = NewCaptionTxt("23024", adrsC)             '' Left Date
'' Tab 2
'' Details Frame
lblDOB.Caption = NewCaptionTxt("23025", adrsC)              '' Date of Birth
lblBlood.Caption = NewCaptionTxt("23026", adrsC)            '' Blood Group
lblDOJ.Caption = NewCaptionTxt("23027", adrsC)              '' Date of Join
lblConf.Caption = NewCaptionTxt("23028", adrsC)             '' Conf. Date
lblSex.Caption = NewCaptionTxt("23029", adrsC)              '' Sex
lblEmail.Caption = NewCaptionTxt("23030", adrsC)            '' Email-ID
'' Salary Frame
lblSal.Caption = NewCaptionTxt("23031", adrsC)              '' Basic Salary
lblRef.Caption = NewCaptionTxt("23032", adrsC)              '' Reference
'' Address Frame
lblAdd.Caption = NewCaptionTxt("23033", adrsC)              '' Address
lblCity.Caption = NewCaptionTxt("23034", adrsC)             '' City
lblPin.Caption = NewCaptionTxt("23035", adrsC)              '' Pin Code
lblPhone.Caption = NewCaptionTxt("23036", adrsC)            '' Phone No.
'' Tab 3
'' Details Frame
fr3Det.Caption = NewCaptionTxt("23037", adrsC)              '' Permanent Address
lblHouse.Caption = NewCaptionTxt("23038", adrsC)            '' House No/ Name
lblVill.Caption = NewCaptionTxt("23039", adrsC)             '' City/Village
lblDist.Caption = NewCaptionTxt("23040", adrsC)             '' District
lblTel.Caption = NewCaptionTxt("23041", adrsC)              '' Tel No.
lblArea.Caption = NewCaptionTxt("23042", adrsC)             '' Area
lblRoad.Caption = NewCaptionTxt("23043", adrsC)             '' Road
lblState.Caption = NewCaptionTxt("23044", adrsC)            '' State
lblNat.Caption = NewCaptionTxt("23045", adrsC)              '' Nationality
lblSpl.Caption = NewCaptionTxt("23046", adrsC)              '' Special Comments
End Sub

Private Sub CapGrid()       '' Labels,Sizes and Aligns the Grid
'' Set the Grid Captions
With MSF1
    .TextMatrix(0, 0) = NewCaptionTxt("23007", adrsC)           '' Employee Code
    .TextMatrix(0, 1) = NewCaptionTxt("23008", adrsC)           '' Employee Name
    .TextMatrix(0, 2) = NewCaptionTxt("23009", adrsC)           '' Employee Card
    .TextMatrix(0, 3) = NewCaptionTxt("23010", adrsC)           '' Join Date
    .TextMatrix(0, 4) = NewCaptionTxt("23011", adrsC)           '' Confirm Date
    .TextMatrix(0, 5) = "Left date"
End With
'' Aligns
With MSF1
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    .ColAlignment(4) = flexAlignLeftCenter
    .ColAlignment(5) = flexAlignLeftCenter
End With
'' Sizing
With MSF1
    .ColWidth(1) = .ColWidth(1) * 3.45
    .ColWidth(2) = .ColWidth(2) * 1
    .ColWidth(3) = .ColWidth(3) * 1.2
    .ColWidth(4) = .ColWidth(4) * 1.2
    .ColWidth(5) = .ColWidth(5) * 1.2
End With
End Sub

Private Sub GetRights()     '' Gets and Sets the Rights for the Particular Form
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 11, 1)
If Mid(strTmp, 1, 1) = "1" Then AddRights = True
If Mid(strTmp, 2, 1) = "1" Then EditRights = True
If Mid(strTmp, 3, 1) = "1" Then DeleteRights = True
Exit Sub
ERR_P:
    ShowError ("GetRights :: " & Me.Caption)
    AddRights = False
    EditRights = False
    DeleteRights = False
End Sub

Private Sub FillCombos()        '' Fills All the Combos in the Form
On Error GoTo ERR_P
Dim bytTmp As Byte
'' Employee Code Combo
'Call FillEmpCodeName
'' Entry Combo
For bytTmp = 1 To 6
    cboEnt.AddItem Choose(bytTmp, "0", "1", "2", "4", "6", "8")
Next bytTmp
'' Category Combo,Department Combo,Group Combo
Call FillCatDepGroupCombo
'' Travel Combo

'' Sex Combo
cboSex.AddItem "M"
cboSex.AddItem "F"
FillEmpCodeName
Exit Sub
ERR_P:
    ShowError ("FillCombos :: " & Me.Caption)
End Sub

Private Sub FillEmpCodeName()       '' Fills Employee Code & Name Combos
On Error GoTo ERR_P
cboCode.clear       '' Clear Code Combo
cboName.clear       '' Clear Name Combo
Dim strTmp As String
If strCurrentUserType = HOD Then
    ''Original ->strTmp = " Where Dept=" & intCurrDept & " "
    strTmp = Replace(strCurrData, "'", "")
End If

If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select Empcode from Empmst " & strTmp & " Order by Empcode", _
ConMain, adOpenKeyset, adLockOptimistic       '' Fill Code Combo
If Not (adrsDept1.BOF And adrsDept1.EOF) Then
    Do While Not adrsDept1.EOF
        cboCode.AddItem adrsDept1(0)
        adrsDept1.MoveNext
    Loop
End If
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select Name from Empmst " & strTmp & " Order by Name", _
ConMain, adOpenKeyset, adLockOptimistic       '' Fill Name Combo
If Not (adrsDept1.BOF And adrsDept1.EOF) Then
    Do While Not adrsDept1.EOF
        cboName.AddItem adrsDept1(0)
        adrsDept1.MoveNext
    Loop
End If

Exit Sub
ERR_P:
    ShowError (" FillEmpCode :: " & Me.Caption)
End Sub

Private Sub FillCatDepGroupCombo()      '' Fills Category,Department and Group Combo
On Error GoTo ERR_P
'' Category Combo
Call ComboFill(cboCat, 3, 2)
'' Department combo
Call ComboFill(cboDept, 2, 2)
'' Group Combo
Call ComboFill(cboGroup, 8, 2)
'' Company Combo
Call ComboFill(cboCompany, 5, 2)

'' Division Combo
Call ComboFill(cboDiv, 13, 2)
'' OT Rule Combo
Call ComboFill(cboOTRule, 9, 2)
cboOTRule.AddItem "None"
'' CO Rule Combo
Call ComboFill(cboCORule, 10, 2)
'' Location Combo
Call ComboFill(cboLoca, 11, 2)

Call ComboFill(cboDesig, 20, 2)
cboCORule.AddItem "None"

Exit Sub
ERR_P:
End Sub

Private Sub LoadSpecifics()     '' Action to be Taken when the Form is Getting Loaded
blnCodeMatch = False            '' Default Grid Action
bytMode = 1                     '' Set Mode to View Mode
Call ChangeMode                 '' TakeAction Accordingly
Select Case InVar.bytCom
    Case "1", "", "0"
        lblComp.Visible = False
        cboCompany.Visible = False
    Case Else
        lblComp.Visible = True
        cboCompany.Visible = True
End Select

txtCode.MaxLength = pVStar.CodeSize
txtCard.MaxLength = pVStar.CardSize
fraSal.Visible = GetFlagStatus("pratham")
End Sub

Private Sub OpenMasterTable()       '' Open Employee Master
On Error GoTo ERR_P
If adrsEmp.State = 1 Then adrsEmp.Close
If 1 = 2 Then
If strCurrentUserType <> HOD Then strCurrData = " ,Deptdesc, Groupmst, Location, Division, Company WHERE deptdesc.dept = Empmst.dept  and  Division.Div = Empmst.div and GroupMst.Group = Empmst.group and  Location.Location = Empmst.Location and  company.Company = Empmst.company "


strSql = "SELECT Empmst.*, frmDesignation.DesigName, catdesc.Desc as CDesc, deptdesc.desc as DDesc, GroupMst.GrupDesc, Location.LocDesc, Division.DivDesc, company.CName, CORul.CODesc, OTRul.OTDesc"
strSql = strSql + " From Empmst ,CORul, OTRul, FrmDesignation, CatDesc  " + strCurrData + "  and Empmst.designatn = frmDesignation.DesigCode AND catdesc.cat = Empmst.cat and  Empmst.COCOde = CORul.COCode and Empmst.OTCode = OTRul.OTCode"
'strSql = strSql + " FROM company INNER JOIN (((Location INNER JOIN (GroupMst INNER JOIN (Division INNER JOIN (deptdesc INNER JOIN (catdesc INNER JOIN"
'strSql = strSql + "  (Empmst INNER JOIN frmDesignation ON Empmst.designatn = frmDesignation.DesigCode) ON catdesc.cat = Empmst.cat) ON deptdesc.dept = Empmst.dept) "
'strSql = strSql + " ON Division.Div = Empmst.div) ON GroupMst.Group = Empmst.group) ON Location.Location = Empmst.Location) INNER JOIN CORul ON"
'strSql = strSql + " Empmst.COCOde = CORul.COCode) INNER JOIN OTRul ON Empmst.OTCode = OTRul.OTCode) ON company.Company = Empmst.company"
strSql = strSql + " Order by Empmst." & strSort

Else
    If strCurrentUserType <> HOD Then strCurrData = " ,Deptdesc, Groupmst, Location, Division, Company WHERE deptdesc.dept = Empmst.dept  and  Division.Div = Empmst.div and GroupMst.[Group] = Empmst.[group] and  Location.Location = Empmst.Location and  company.Company = Empmst.company "
    strSql = "SELECT Empmst.*, frmDesignation.DesigName, catdesc.[Desc] as CDesc, deptdesc.[desc] as DDesc, GroupMst.GrupDesc, Location.LocDesc, Division.DivDesc, company.CName, CORul.CODesc, OTRul.OTDesc"
    strSql = strSql + " From Empmst ,CORul, OTRul, FrmDesignation, CatDesc  " + strCurrData + "  and Empmst.designatn = frmDesignation.DesigCode AND catdesc.cat = Empmst.cat and  Empmst.COCOde = CORul.COCode and Empmst.OTCode = OTRul.OTCode"
    strSql = strSql + " Order by Empmst." & strSort
End If

adrsEmp.Open Replace(strSql, "'", ""), ConMain, adOpenKeyset, adLockOptimistic
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Sub FillGrid()                          '' Fills the Master Grid
On Error GoTo ERR_P
Dim intCounter As Integer
OpenMasterTable             '' Requeries the Recordset for any Updated Values
'adrsEmp.Sort = strSort
'' Put Appropriate Rows in the Grid
If adrsEmp.EOF Then
    MSF1.Rows = 1
    TB1.TabEnabled(1) = False                   '' Disables Tab 1 If no Records are Found
    TB1.TabEnabled(2) = False                   '' Disables Tab 2 If no Records are Found
    TB1.TabEnabled(3) = False                   '' Disables Tab 3 If no Records are Found
    Exit Sub
End If
''SG05
MSF1.clear
With MSF1
    .TextMatrix(0, 0) = NewCaptionTxt("23007", adrsC)           '' Employee Code
    .TextMatrix(0, 1) = NewCaptionTxt("23008", adrsC)           '' Employee Name
    .TextMatrix(0, 2) = NewCaptionTxt("23009", adrsC)           '' Employee Card
    .TextMatrix(0, 3) = NewCaptionTxt("23010", adrsC)           '' Join Date
    .TextMatrix(0, 4) = NewCaptionTxt("23011", adrsC)           '' Confirm Date
    .TextMatrix(0, 5) = "Left Date"          '' Left Date
End With
''
MSF1.Rows = adrsEmp.RecordCount + 1             '' Sets Rows Appropriately
adrsEmp.MoveFirst
For intCounter = 1 To adrsEmp.RecordCount       '' Fills the Grid
    With MSF1
        .TextMatrix(intCounter, 0) = adrsEmp("EmpCode")     '' Employee Code
        '' Employee Name
        .TextMatrix(intCounter, 1) = IIf(IsNull(adrsEmp("Name")), "", adrsEmp("Name"))
        '' Employee Card
        .TextMatrix(intCounter, 2) = IIf(IsNull(adrsEmp("Card")), "", adrsEmp("Card"))
        If Not IsNull(adrsEmp("JoinDate")) Then             '' JoinDate
            .TextMatrix(intCounter, 3) = DateDisp(adrsEmp("JoinDate"))
        End If
        If Not IsNull(adrsEmp("ConfmDt")) Then              '' Confirm Date
            .TextMatrix(intCounter, 4) = DateDisp(adrsEmp("ConfmDt"))
        End If
        .TextMatrix(intCounter, 5) = DateDisp(IIf(IsNull(adrsEmp("leavdate")), "", adrsEmp("leavdate")))
    End With
    adrsEmp.MoveNext
Next
TB1.TabEnabled(1) = True                        '' If Records Found Enables the TAB 1
TB1.Tab = 0
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub Display()       '' Displays the Employee Record
On Error GoTo ERR_P
If Not (adrsEmp.EOF) Then
    '' Tab 1
    '' Details Frame
    txtCode.Text = adrsEmp("EmpCode")
    txtName.Text = IIf(IsNull(adrsEmp("Name")), "", adrsEmp("Name"))
    txtCard.Text = IIf(IsNull(adrsEmp("Card")), "", adrsEmp("Card"))
    cboDesig.Text = IIf(IsNull(adrsEmp("Designame")), "", adrsEmp("Designame"))
    txtName2.Text = IIf(IsNull(adrsEmp("Name2")), "", adrsEmp("Name2"))
    
    cboEnt.Text = IIf(IsNull(adrsEmp("Entry")), 0, adrsEmp("Entry"))
    chkAuto.Value = IIf(adrsEmp("Shf_Chg") = 0, 0, 1)
    
    If chkAuto.Value = 0 Then
        frAuto.Visible = False
    Else
        frAuto.Visible = True
    End If
    txtAutoG.Text = IIf(IsNull(adrsEmp("AutoG")), "", adrsEmp("AutoG"))
    strAutoG = txtAutoG.Text
    ''
    ''For Mauritius 20-08-2003
    txtFreeF.Text = IIf(IsNull(adrsEmp("Qualf")), "", adrsEmp("Qualf"))
    ''
    '' Call the Following Procedure to Put Values in  ComboBoxes to Trap Error in Case
    Call PutComboValues
    '' Define Schedule
    Shft.Empcode = adrsEmp("EmpCode")
    If Not IsNull(adrsEmp("Shf_Date")) Then             '' Shift Date
        Shft.startdate = DateDisp(adrsEmp("Shf_Date"))
    Else
        Shft.startdate = DateCompDate(CStr(Date))
    End If
    strRotPass = adrsEmp("Name")
    If adrsEmp("STyp") = "F" Then
        Shft.ShiftType = "F"
        Shft.ShiftCode = adrsEmp("F_Shf")
    Else
        Shft.ShiftType = "R"
        Shft.ShiftCode = adrsEmp("SCode")
    End If
    Shft.WO = IIf(IsNull(adrsEmp("Off")), "", adrsEmp("Off"))
    Shft.WO1 = IIf(IsNull(adrsEmp("Off2")), "", adrsEmp("Off2"))
    Shft.WO2 = IIf(IsNull(adrsEmp("WO_1_3")), "", adrsEmp("WO_1_3"))
    Shft.WO3 = IIf(IsNull(adrsEmp("WO_2_4")), "", adrsEmp("WO_2_4"))
    ''
    Shft.WOHLAction = IIf(IsNull(adrsEmp("WOHLAction")), 0, adrsEmp("WOHLAction"))
    Shft.Action3Shift = IIf(IsNull(adrsEmp("Action3Shift")), "", adrsEmp("Action3Shift"))
    Shft.AutoOnPunch = IIf(adrsEmp("AutoForPunch") = 1, True, False)
    Shft.ActionBlank = IIf(IsNull(adrsEmp("ActionBlank")), "", adrsEmp("ActionBlank"))
    '' Past Employee
    If Not IsNull(adrsEmp("LeavDate")) Then
        txtLeft.Text = DateDisp(adrsEmp("LeavDate"))
    Else
        txtLeft.Text = ""
    End If
    '' Tab 2
    '' Details Frame
    If Not IsNull(adrsEmp("Birth_Dt")) Then             '' Date of Birth
        txtDOB.Text = DateDisp(adrsEmp("Birth_Dt"))
    Else
        txtDOB.Text = ""
    End If
    If Not IsNull(adrsEmp("JoinDate")) Then             '' Date of Join
        txtDOJ.Text = DateDisp(adrsEmp("JoinDate"))
    Else
        txtDOJ.Text = ""
    End If
    If Not IsNull(adrsEmp("ConfmDt")) Then              '' Date of Confirmation
        txtConf.Text = DateDisp(adrsEmp("ConfmDt"))
    Else
        txtConf.Text = ""
    End If
    txtBlood.Text = IIf(IsNull(adrsEmp("BG")), "", adrsEmp("BG"))   '' Blood Group
    cboSex.Text = IIf(IsNull(adrsEmp("Sex")), "M", adrsEmp("Sex"))
    txtEmail.Text = IIf(IsNull(adrsEmp("Email_Id")), "", adrsEmp("Email_Id"))    '' Email Id
    '' Salary Frame
    txtSal.Text = IIf(IsNull(adrsEmp("Salary")), "0", adrsEmp("Salary"))    '' Salary
    txtRef.Text = IIf(IsNull(adrsEmp("Reference")), "", adrsEmp("Reference"))
    '' AddRess Frame
    txtAdd1.Text = IIf(IsNull(adrsEmp("ResAdd1")), "", adrsEmp("ResAdd1"))
    txtAdd2.Text = IIf(IsNull(adrsEmp("ResAdd2")), "", adrsEmp("ResAdd2"))
    txtCity.Text = IIf(IsNull(adrsEmp("City")), "", adrsEmp("City"))
    txtPin.Text = IIf(IsNull(adrsEmp("Pin")), "", adrsEmp("Pin"))
    txtPhone.Text = IIf(IsNull(adrsEmp("Phone")), "", adrsEmp("Phone"))

    '' Details Frame
    txtHouse.Text = IIf(IsNull(adrsEmp("udf1")), "", adrsEmp!udf1)
    txtArea.Text = IIf(IsNull(adrsEmp("udf2")), "", adrsEmp!udf2)
    txtVill.Text = IIf(IsNull(adrsEmp("udf3")), "", adrsEmp!udf3)
    txtRoad.Text = IIf(IsNull(adrsEmp("udf4")), "", adrsEmp!udf4)
    txtDist.Text = IIf(IsNull(adrsEmp("udf5")), "", adrsEmp!udf5)
    txtState.Text = IIf(IsNull(adrsEmp("udf6")), "", adrsEmp!udf6)
    txtTel.Text = IIf(IsNull(adrsEmp("udf7")), "", adrsEmp("udf7"))
    txtNat.Text = IIf(IsNull(adrsEmp("udf8")), "", adrsEmp!udf8)
    txtSpl(0).Text = IIf(IsNull(adrsEmp("udf9")), "", adrsEmp!udf9)
    txtSpl(1).Text = IIf(IsNull(adrsEmp("udf10")), "", adrsEmp!udf10)
    If GetFlagStatus("pratham") = True Then
        txtOldSal.Text = IIf(IsNull(adrsEmp("OldSal")), "", adrsEmp!OldSal)
        txtNewSal.Text = IIf(IsNull(adrsEmp("NewSal")), "", adrsEmp!NewSal)
        txtAvgDays.Text = IIf(IsNull(adrsEmp("AvgDays")), "", adrsEmp!AvgDays)
    End If
 End If
Exit Sub
ERR_P:
    If Err.Number = 380 Then
        MsgBox "Given The Reportess Code " & adrsEmp!ReporteesName & "  Is Not Available", vbExclamation, "Invalied Reportees"
        'Resume Next
    End If
    ShowError ("Display :: " & Me.Caption)
    'Resume Next
End Sub

Private Sub PutComboValues()            '' Puts Values of the Selected Employee into
On Error GoTo ERR_P                     '' their ComboBoxes
Dim strTmp As String
strTmp = "Categoty"
cboCat.Text = adrsEmp("Cdesc")            '' Category
strTmp = "Department"
cboDept.Text = adrsEmp("Ddesc")          '' Department
strTmp = "Group"

cboGroup.Text = IIf(IsNull(adrsEmp("GrupDesc")), "0", adrsEmp("GrupDesc"))        '' Group
strTmp = "Location"
cboLoca.Text = IIf(IsNull(adrsEmp("LocDesc")), "0", adrsEmp("LocDesc"))  '' Location
strTmp = "Company"
cboCompany.Value = IIf(IsNull(adrsEmp("CName")), "", adrsEmp("CName"))  '' Company
strTmp = "Division"
cboDiv.Value = IIf(IsNull(adrsEmp("DivDesc")), "", adrsEmp("DivDesc"))  '' Division
strTmp = "OT Rule"
cboOTRule.Value = IIf(IsNull(adrsEmp("OTDesc")) Or adrsEmp("OTDesc") = 100, "None", _
adrsEmp("OTDEsc")) '' OT Rule
strTmp = "CO Rule"
cboCORule.Value = IIf(IsNull(adrsEmp("CODesc")) Or adrsEmp("CODesc") = 100, "None", _
adrsEmp("CODesc")) '' CO Rule
Exit Sub
ERR_P:
    ShowError ("Display :: PutComboValues :: " & strTmp & " :: " & Me.Caption)
    Resume Next
End Sub

Private Sub ChangeMode()        '' Clled when the Mode of Operation on the Form Changes
Select Case bytMode
    Case 1
        Call ViewAction
    Case 2
        Call AddAction
    Case 3
        adrsEmp.MoveFirst                           '' Move to the First Record
        adrsEmp.Find "EmpCode='" & MSF1.Text & "'"  '' Find the Record Selected
        If txtCode.Text <> "" Then
        MSF1.Text = txtCode.Text
        End If
        Call EditAction
End Select
End Sub

Private Sub ViewAction()        '' Called when the Form Comes in the Edit Mode
'' Enable Needed Controls
cmdAddSave.Enabled = True       '' Enable ADD/SAVE Button
cmdEditCan.Enabled = True       '' Enable Edit/Cancel Button
cmdDel.Enabled = True           '' Enable Delete Button
cboCode.Enabled = True          '' Enable Code ComboBox
cboName.Enabled = True          '' Enable Name ComboBox
'' Disable Needed Controls
'' Tab 1
fr1Det.Enabled = False
fr1Iden.Enabled = False
frSch.Enabled = False

frAuto.Enabled = False
''
frLeft.Enabled = False
'' Tab 2
fr2Det.Enabled = False
frSal.Enabled = False
frAdd.Enabled = False
'' Tab 3
fr3Det.Enabled = False
'' Give Captions to the Needed Controls
Call SetGButtonCap(Me)
TB1.Tab = 0     '' Set the Tab to the First Tab
End Sub

Private Sub AddAction()         '' Action Taken when the Form is put into Add Mode
'' Enable Necessary Controls
If TB1.TabEnabled(1) = False Then TB1.TabEnabled(1) = True
If TB1.TabEnabled(2) = False Then TB1.TabEnabled(2) = True
If TB1.TabEnabled(3) = False Then TB1.TabEnabled(3) = True
TB1.Tab = 1
'' Tab 1
fr1Det.Enabled = True
fr1Iden.Enabled = True
frSch.Enabled = True

frAuto.Enabled = True
''
frLeft.Enabled = True
'' Tab 2
fr2Det.Enabled = True
frSal.Enabled = True
frAdd.Enabled = True
fraSal.Enabled = True
'' Tab 3
fr3Det.Enabled = True
txtCode.Enabled = True      '' Enable Code TextBox
'' Disable Necessary Controls
cmdDel.Enabled = False      '' Disable Delete Button
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Clear & Set Necessary Controls
Call ClearControls
txtCode.SetFocus        '' Set Focus to the Code TextBox
End Sub

Private Sub EditAction()        '' Action Taken When the form is put into Edit Mode
'' Enable Necessary Controls
If TB1.TabEnabled(1) = False Then TB1.TabEnabled(1) = True
If TB1.TabEnabled(2) = False Then TB1.TabEnabled(2) = True
If TB1.TabEnabled(3) = False Then TB1.TabEnabled(3) = True
TB1.Tab = 1
'' Tab 1
fr1Det.Enabled = True
fr1Iden.Enabled = True
frSch.Enabled = True

frAuto.Enabled = True
''
frLeft.Enabled = True
'' Tab 2
fr2Det.Enabled = True
frSal.Enabled = True
fraSal.Enabled = True
frAdd.Enabled = True
'' Tab 3
fr3Det.Enabled = True
txtCode.Enabled = False         '' Enable Code TextBox
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
txtCard.SetFocus                '' Set Focus to the Card
End Sub
    
Private Sub ClearControls()     '' Clears Controls before Adding
'' Tab 1
'' Details Frame
txtCode.Text = ""
txtName.Text = ""
txtCard.Text = ""
txtName2.Text = ""
''For Mauritius 20-08-2003
txtFreeF.Text = ""
''
'' Identification Frame
cboEnt.Text = "0"
chkAuto.Value = 0
'' Call the Following Procedure to Put Values in  ComboBoxes to Trap Error in Case
cboOTRule.Text = "None"
cboCORule.Text = "None"
cboCat.Value = ""
cboDept.Value = ""
cboDesig.Value = ""
'cboCompany.Text = ""
If cboLoca.ListCount > 0 Then cboLoca.ListIndex = 0
If cboGroup.ListCount > 0 Then cboGroup.ListIndex = 0
If cboCompany.ListCount > 0 Then cboCompany.ListIndex = 0
If cboDiv.ListCount > 0 Then cboDiv.ListIndex = 0
cboLoca.Value = ""
cboGroup.Value = ""
cboCompany.Value = ""
cboDiv.Value = ""
'' Define Schedule
Shft.Empcode = txtCode.Text
Shft.startdate = Date
strRotPass = ""
Shft.ShiftType = "F"
Shft.ShiftCode = ""
Shft.WO = ""
Shft.WO1 = ""
Shft.WO2 = ""
Shft.WO3 = ""
Shft.WOHLAction = 0
Shft.Action3Shift = ""
Shft.AutoOnPunch = False
Shft.ActionBlank = ""
'' Past Employee
txtLeft.Text = ""
'' Tab 2
'' Details Frame
txtDOB.Text = ""
txtDOJ.Text = ""
txtConf.Text = ""
txtBlood.Text = ""
cboSex.Text = "M"
txtEmail.Text = ""
'' Salary Frame
txtSal.Text = "0.00"
txtRef.Text = ""
'' AddRess Frame
txtAdd1.Text = ""
txtAdd2.Text = ""
txtCity.Text = ""
txtPin.Text = ""
txtPhone.Text = ""
'txtmedical.Text = ""
'' Tab 3
'' Details Frame
txtHouse.Text = ""
txtArea.Text = ""
txtVill.Text = ""
txtRoad.Text = ""
txtDist.Text = ""
txtState.Text = ""
txtTel.Text = ""
txtNat.Text = ""
txtSpl(0).Text = ""
txtSpl(1).Text = ""
txtAutoG.Text = ""
txtOldSal.Text = ""
txtNewSal.Text = ""
txtAvgDays.Text = ""
End Sub

Private Sub cmdAddSave_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        '' Check for Rights
        ''
        Call GetRights
        If Not AddRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        bytMode = 2
        Call ChangeMode
    Case 2          '' Add Mode
        If Not ValidateAddmaster Then Exit Sub  '' Validate For Add
        If Not SaveModMaster Then Exit Sub      '' Save for Add
        Call SaveAddLog                         '' Save the Add Log
        '' Update Employee Leaves
        Call OpenLeaveMaster
        Call UpdateNewEmpLeave(txtCode.Text, DateCompDate(txtDOJ.Text), _
        cboCat.Text, Val(pVStar.YearSel))
        '' Update Employee Shifts
        Call FillEmpShift
        Call FillEmpCodeName                    '' Reflect the ComboBoxes
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

Private Function ValidateAddmaster() As Boolean     '' Validates New record before Saving
On Error GoTo ERR_P
ValidateAddmaster = True
'' Check for Leave Balance File
If Not FindTable("LvBal" & Right(pVStar.YearSel, 2)) Then
    MsgBox NewCaptionTxt("23049", adrsC), vbExclamation
    ValidateAddmaster = False
    Exit Function
End If
'' Check for Demo Version
If InVar.blnVerType = "1" Then
    If CInt(InVar.lngEmp) <= adrsEmp.RecordCount Then
        MsgBox NewCaptionTxt("23050", adrsC) & InVar.lngEmp, vbExclamation
        txtCode.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If

        
If Len(txtCode.Text) = 0 Then
    MsgBox NewCaptionTxt("23048", adrsC), vbExclamation
    txtCode.SetFocus
    ValidateAddmaster = False
    Exit Function
End If

'' Check for Existing Employee Code
If MSF1.Rows > 1 Then
    adrsEmp.MoveFirst
    adrsEmp.Find "EmpCode='" & txtCode.Text & "'"
    If Not adrsEmp.EOF Then
        MsgBox NewCaptionTxt("23051", adrsC), vbExclamation
        txtCode.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If

    If MSF1.Rows > 1 Then
        adrsEmp.MoveFirst
        adrsEmp.Find "Card='" & txtCard.Text & "'"
        If Not adrsEmp.EOF Then
            MsgBox "This card is already assigned to " & vbCrLf & vbCrLf & adrsEmp("empcode") & " : " & adrsEmp("Name"), vbCritical
            txtCard.SetFocus
            ValidateAddmaster = False
            Exit Function
        End If
    End If

If IsDate(txtDOB.Text) Then
    If Year(Now()) - Year(CDate(txtDOB.Text)) < 18 Then
        MsgBox "Employee Birth Date Must Be Between 18 to 65 Years ", vbExclamation
        txtDOB.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If


If Not IsDate(txtDOJ.Text) Then
    MsgBox "Please Enter Proper Joinig Date ", vbExclamation
    txtDOJ.SetFocus
    ValidateAddmaster = False
    Exit Function
End If

If Year(Now()) - Year(CDate(txtDOJ.Text)) > 65 Then
    MsgBox "Joining Date Should Be Less Than 65 Years ", vbExclamation
    txtDOJ.SetFocus
    ValidateAddmaster = False
    Exit Function
End If

ValidateAddmaster = CheckValidation

Exit Function
ERR_P:
    ShowError ("ValidateAddmaster :: " & Me.Caption)
    ValidateAddmaster = False
    Resume Next
End Function

'Private Function PermCardCheck() As Boolean
'On Error GoTo ERR_P                         '' Checks if Card No is Not Existing in
'PermCardCheck = True                        '' List of Permissions
'If adRsInstall.State = 1 Then adRsInstall.Close
'adRsInstall.Open "select * from Install", conmain
'If Val(txtCard.Text) >= adRsInstall("pstart") And Val(txtCard.Text) <= adRsInstall("pend") Then
'    MsgBox txtCard.Text & NewCaptionTxt("23067", adrsC), vbExclamation
'    PermCardCheck = False
'End If
'Exit Function
'ERR_P:
'    ShowError ("Validations :: PermCardCheck :: " & Me.Caption)
'    PermCardCheck = False
'End Function

Private Sub FillEmpShift()          '' Fills Employee Shift For the Newly Added Employee
On Error GoTo ERR_P                 '' for the Current Month
'' Start Date Checks
Dim dttmp As Date
'' If the Employee has Already Left
If Trim(txtLeft.Text) <> "" Then
    dttmp = FdtLdt(Month(Date), CStr(Year(Date)), "F")
    If DateCompDate(txtLeft.Text) <= dttmp Then Exit Sub
End If
'' Get Current Months Last Process Date
dttmp = FdtLdt(Month(Date), CStr(Year(Date)), "L")
'' Check on JoinDate
If DateCompDate(DateCompDate(txtDOJ.Text)) > dttmp Then Exit Sub
'' Check on Shift Date
If DateCompDate(Shft.startdate) > dttmp Then Exit Sub
'' End Date Checks
If Not FindTable(Left(MonthName(Month(Date)), 3) & Right(CStr(Year(Date)), 2) & "Shf") Then
    'conmain.Execute "Select * into " & Left(MonthName(Month(Date)), 3) & _
    Right(CStr(Year(Date)), 2) & "Shf" & " from shfinfo where " & "1=2"
    Call CreateTableIntoAs("*", "shfinfo", Left(MonthName(Month(Date)), 3) & _
    Right(CStr(Year(Date)), 2) & "Shf", " Where 1=2")
    Call CreateTableIndexAs("MONYYSHF", Left(MonthName(Month(Date)), 3), Right(CStr(Year(Date)), 2))
    Call GetSENums(MonthName(Month(Date)), CStr(Year(Date)))
End If
Call GetSENums(MonthName(Month(Date)), CStr(Year(Date)))
adrsEmp.Requery
Call FillEmployeeDetails(txtCode.Text)
If Month(Date) = Month(Shft.startdate) And Year(Shft.startdate) = Year(Date) Then Call AdjustSENums(DateCompDate(Shft.startdate))
If typEmpRot.strShifttype = "F" Then
        '' If Fixed Shifts
        Call FixedShifts(txtCode.Text, MonthName(Month(Date)), CStr(Year(Date)))
    Else
        '' if Rotation Shifts
        '' Fill Other Skip Pattern and Shift Pattern Array
        Call FillArrays
        Select Case strCapSND
            Case "O"        '' After Specific Number of Days
                Call SpecificDaysShifts(txtCode.Text, MonthName(Month(Date)), CStr(Year(Date)))
            Case "D"        '' Only on Fixed Days
                Call FixedDaysShifts(txtCode.Text, MonthName(Month(Date)), CStr(Year(Date)))
            Case "W"        '' Only On Fixed Week days
                Call WeekDaysShifts(txtCode.Text, MonthName(Month(Date)), CStr(Year(Date)))
        End Select
    End If
    '' Add that Record to the Shift File
    Call AddRecordsToShift(MonthName(Month(Date)), CStr(Year(Date)), txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("FillEmpShift :: " & Me.Caption)
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
    If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion) _
    = vbYes Then                                '' Delete the Record
        ConMain.Execute "delete from EmpMst where EmpCode='" & txtCode.Text _
        & "'"
        Call AddActivityLog(lgDelete_Action, 1, 14)     '' Delete Log
        Call AuditInfo("DELETE", Me.Caption, "Deleted Employee: " & txtCode.Text)
        
        ConMain.Execute "delete from lvbal" & Right(pVStar.YearSel, 2) & " where EmpCode='" & txtCode.Text & "'"
        ConMain.Execute "delete from EmpMst where EmpCode='" & txtCode.Text & "'"
   
    End If
    Call FillEmpCodeName                        '' Reflect the ComboBoxes
    Call FillGrid                               '' Reflect the Grid
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
ERR_P:
    If Err.Number = -2147467259 Or Err.Number = -2147217900 Or Err.Number = -2147217873 Then
            MsgBox "Employee Cannot be deleted because Leaves are assigned to this employee.", vbCritical, Me.Caption
            If (MsgBox("Do you want to Delete the Employee?", vbYesNo)) = vbYes Then
                ConMain.Execute "delete from lvbal" & Right(pVStar.YearSel, 2) & " where EmpCode='" & txtCode.Text & "'"
                ConMain.Execute "delete from EmpMst where EmpCode='" & txtCode.Text & "'"
            End If
            Call FillEmpCodeName                        '' Reflect the ComboBoxes
            Call FillGrid                               '' Reflect the Grid
            bytMode = 1
            Call ChangeMode
            Exit Sub
    End If
    ShowError ("Delete Record :: " & Me.Caption)
End Sub

Private Function ValidateModMaster() As Boolean     '' Validates Record Before Saving
On Error GoTo ERR_P                                 '' Called for Edit
ValidateModMaster = True
'' Check for Empty Card No
If Len(txtCard.Text) < pVStar.CardSize Then
    MsgBox NewCaptionTxt("23052", adrsC) & pVStar.CardSize & NewCaptionTxt("23053", adrsC), _
    vbExclamation
    txtCard.SetFocus
    ValidateModMaster = False
    Exit Function
End If

    If MSF1.Rows > 1 Then
        adrsEmp.MoveFirst
        adrsEmp.Find "Card='" & txtCard.Text & "' "
        If Not adrsEmp.EOF Then
            If adrsEmp("EmpCode") <> txtCode.Text Then
                MsgBox "This card is already assigned to " & vbCrLf & vbCrLf & adrsEmp("empcode") & " : " & adrsEmp("Name"), vbCritical
                txtCard.SetFocus
                ValidateModMaster = False
                Exit Function
            End If
        End If
    End If
If IsDate(txtDOB.Text) Then
    If Year(Now()) - Year(CDate(txtDOB.Text)) < 18 Then
        MsgBox "Enter Proper Date Of Birth", vbExclamation
        txtDOB.SetFocus
        ValidateModMaster = False
        Exit Function
    End If

    If Year(CDate(txtDOJ.Text)) - Year(CDate(txtDOB.Text)) > 65 Then
        MsgBox "Enter Proper Joining Date", vbExclamation
        txtDOJ.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
End If
ValidateModMaster = CheckValidation

Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
    ValidateModMaster = False
End Function

Private Function CheckValidation() As Boolean
    CheckValidation = True

   
If Trim(txtName.Text) = "" Then
    MsgBox NewCaptionTxt("23055", adrsC), vbExclamation
    txtName.SetFocus
    CheckValidation = False
    Exit Function
End If

If cboDesig.Text = "" Then
    MsgBox "Please Select Employee designation"
    cboDesig.SetFocus
    CheckValidation = False
    Exit Function
End If


'' Check for Empty OT Rule
If Trim(cboOTRule.Text) = "" Then
    MsgBox NewCaptionTxt("23068", adrsC), vbExclamation
    cboOTRule.SetFocus
    CheckValidation = False
    Exit Function
End If
'' Check for Empty CO Rule
If Trim(cboCORule.Text) = "" Then
    MsgBox NewCaptionTxt("23069", adrsC), vbExclamation
    cboCORule.SetFocus
    CheckValidation = False
    Exit Function
End If

If Trim(cboCat.Text) = "" Then
    MsgBox NewCaptionTxt("23056", adrsC), vbExclamation
    cboCat.SetFocus
    CheckValidation = False
    Exit Function
End If
'' Check for Empty Department
If Trim(cboDept.Text) = "" Then
    MsgBox NewCaptionTxt("23057", adrsC), vbExclamation
    cboDept.SetFocus
    CheckValidation = False
    Exit Function
End If
'' Check for Empty Group
If Trim(cboGroup.Text) = "" Then
    MsgBox NewCaptionTxt("23058", adrsC), vbExclamation
    cboGroup.SetFocus
    CheckValidation = False
    Exit Function
End If
'' Check for Empty Location
If Trim(cboLoca.Text) = "" Then
    MsgBox NewCaptionTxt("23070", adrsC), vbExclamation
    cboLoca.SetFocus
    CheckValidation = False
    Exit Function
End If
If Trim(cboCompany.Text) = "" Then
    MsgBox NewCaptionTxt("23059", adrsC), vbExclamation
    cboCompany.SetFocus
    CheckValidation = False
    Exit Function
End If

'If Not PermCardCheck Then
'    txtCard.SetFocus
'    CheckValidation = False
'    Exit Function
'End If

If Len(txtCard.Text) < pVStar.CardSize Then
    MsgBox NewCaptionTxt("23052", adrsC) & pVStar.CardSize & NewCaptionTxt("23053", adrsC), _
    vbExclamation
    txtCard.SetFocus
    CheckValidation = False
    Exit Function
End If

'' Check for Empty JoinDate
If Trim(txtDOJ.Text) = "" Then
    MsgBox NewCaptionTxt("23061", adrsC), vbExclamation
    txtDOJ.SetFocus
    CheckValidation = False
    Exit Function
End If
'' Check for JoinDate Greater then Shift Date
If DateCompDate(txtDOJ.Text) > DateCompDate(Shft.startdate) Then
    MsgBox NewCaptionTxt("23062", adrsC), vbExclamation
    txtDOJ.SetFocus
    CheckValidation = False
    Exit Function
End If
'' Check for JoinDate Greater then Leave Date
If Trim(txtLeft.Text) <> "" Then
    If DateCompDate(txtDOJ.Text) > DateCompDate(txtLeft.Text) Then
        MsgBox NewCaptionTxt("23063", adrsC), vbExclamation
        txtLeft.SetFocus
        CheckValidation = False
        Exit Function
    End If
End If

If Trim(txtDOB.Text) <> "" Then
    If DateCompDate(txtDOB.Text) > DateCompDate(txtDOJ.Text) Then
        MsgBox NewCaptionTxt("23065", adrsC), vbExclamation
        txtDOB.SetFocus
        CheckValidation = False
        Exit Function
    End If
End If
'' Check for Confirm Date Less than Join date
If Trim(txtConf.Text) <> "" Then
    If DateCompDate(txtConf.Text) < DateCompDate(txtDOJ.Text) Then
        MsgBox NewCaptionTxt("23066", adrsC), vbExclamation
        txtConf.SetFocus
        CheckValidation = False
        Exit Function
    End If
End If

If Trim(cboDiv.Value) = "" Then
        MsgBox "Division should not blank", vbExclamation
        cboDiv.SetFocus
        CheckValidation = False
        Exit Function
End If
    
If chkAuto.Value = 1 Then
    If txtAutoG.Text = "" Then
        MsgBox "Autoshift should not blank", vbExclamation
        txtAutoG.SetFocus
        CheckValidation = False
    End If
End If

If Trim(txtLeft.Text) <> "" Then
    If DateCompDate(Shft.startdate) > DateCompDate(txtLeft.Text) Then
        MsgBox NewCaptionTxt("23064", adrsC), vbExclamation
        txtLeft.SetFocus
        CheckValidation = False
        Exit Function
    End If
End If

If Shft.ShiftCode = "" Then
    MsgBox NewCaptionTxt("23060", adrsC), vbExclamation
    cmdSch.SetFocus
    CheckValidation = False
    Exit Function
End If

Exit Function
ERR_P:
    ShowError ("CheckValidation :: " & Me.Caption)
    CheckValidation = False
End Function

Private Function SaveModMaster() As Boolean     '' Procedure Called to Save Edited Record
On Error GoTo ERR_P
Dim strTmp(4) As String
With adrsEmp
    If bytMode = 2 Then
       .AddNew
       .Fields("EmpCode") = txtCode.Text
    Else
        .MoveFirst
        .Find "EmpCode='" & txtCode.Text & "'"
    End If
    .Fields("Styp") = IIf(Shft.ShiftType = "F", "F", "R")
    .Fields("F_Shf") = IIf(Shft.ShiftType = "F", Shft.ShiftCode, 100)
    .Fields("Scode") = IIf(Shft.ShiftType = "F", 100, Shft.ShiftCode)
    .Fields("Shf_Date") = Format(Shft.startdate, "dd/mmm/yy")
    .Fields("Off") = Shft.WO
    .Fields("Off2") = Shft.WO1
    .Fields("WO_1_3") = Shft.WO2
    .Fields("WO_2_4") = Shft.WO3
    
    .Fields("Card") = txtCard.Text
    .Fields("Name") = txtName.Text
    .Fields("Entry") = cboEnt.Text
    .Fields("LeavDate") = IIf(Trim(txtLeft.Text) = "", Null, Format(txtLeft.Text, "dd/mmm/yy"))
    .Fields("Birth_Dt") = IIf(Trim(txtDOB.Text) = "", Null, Format(txtDOB.Text, "dd/mmm/yy"))
    .Fields("Confmdt") = IIf(Trim(txtConf.Text) = "", Null, Format(txtConf.Text, "dd/mmm/yy"))
    .Fields("Joindate") = Format(txtDOJ.Text, "dd/mmm/yy")
    .Fields("OTCode") = IIf(cboOTRule.Text = "None", 100, cboOTRule.List(cboOTRule.ListIndex, 1))
    .Fields("COCode") = IIf(cboCORule.Text = "None", 100, cboCORule.List(cboCORule.ListIndex, 1))
    .Fields("Designatn") = cboDesig.List(cboDesig.ListIndex, 1)
    .Fields("Shf_Chg") = IIf(chkAuto.Value = 0, 0, 1)
    .Fields("Cat") = cboCat.List(cboCat.ListIndex, 1)
    .Fields("Dept") = cboDept.List(cboDept.ListIndex, 1)
    .Fields("Group") = cboGroup.List(cboGroup.ListIndex, 1)
    .Fields("Location") = cboLoca.List(cboLoca.ListIndex, 1)
    .Fields("Company") = cboCompany.List(cboCompany.ListIndex, 1)
    .Fields("Div") = cboDiv.List(cboDiv.ListIndex, 1)
    .Fields("BG") = txtBlood.Text
    .Fields("Sex") = cboSex.Text
    .Fields("Qualf") = txtFreeF.Text
    .Fields("Salary") = Val(txtSal.Text)
    .Fields("Reference") = txtRef.Text
    .Fields("Email_Id") = txtEmail.Text
    .Fields("ResAdd1") = txtAdd1.Text
    .Fields("ResAdd2") = txtAdd2.Text
    .Fields("City") = txtCity.Text
    .Fields("Pin") = txtPin.Text
    .Fields("Phone") = txtPhone.Text
    .Fields("Udf1") = txtHouse.Text
    .Fields("Udf2") = txtArea.Text
    .Fields("UDF3") = txtVill.Text
    .Fields("UDF4") = txtRoad.Text
    .Fields("UDF5") = txtDist.Text
    .Fields("UDF6") = txtState.Text
    .Fields("UDF7") = txtTel.Text
    .Fields("UDF8") = txtNat.Text
    .Fields("UDF9") = txtSpl(0).Text
    .Fields("UDF10") = txtSpl(1).Text
    .Fields("Name2") = Trim(txtName2.Text)
    .Fields("WOHLAction") = Shft.WOHLAction
    .Fields("Action3Shift") = Shft.Action3Shift
    .Fields("AutoForPunch") = IIf(Shft.AutoOnPunch, 1, 0)
    .Fields("ActionBlank") = Shft.ActionBlank
    .Fields("AutoG") = IIf(chkAuto.Value = 1, txtAutoG.Text, "")
    .Update
End With

SaveModMaster = True
Exit Function
ERR_P:
    ShowError ("SaveModMaster :: " & Me.Caption)
    Set adrsEmp = Nothing
    SaveModMaster = False
    Resume Next
End Function

Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo ERR_P
Call AddActivityLog(lgADD_MODE, 1, 14)     '' Add Activity
Call AuditInfo("ADD", Me.Caption, "Added Employee: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 14)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edited Employee: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub
