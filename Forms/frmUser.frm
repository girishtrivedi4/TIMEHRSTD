VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Accounts"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frMain 
      Caption         =   "List of Current Users"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1065
      Index           =   9
      Left            =   7800
      TabIndex        =   14
      Top             =   1080
      Width           =   3135
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Label for Description"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   8
         Left            =   100
         TabIndex        =   15
         Top             =   300
         Width           =   8235
      End
   End
   Begin VB.Frame frMain 
      Caption         =   "Master Tables Rights"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   945
      Index           =   4
      Left            =   7410
      TabIndex        =   1
      Top             =   30
      Width           =   2415
      Begin VB.CommandButton cmdSelUN 
         Caption         =   "Select/Unselect"
         Height          =   585
         Index           =   0
         Left            =   6690
         TabIndex        =   25
         Top             =   1980
         Width           =   1545
      End
      Begin MSFlexGridLib.MSFlexGrid MSFMaster 
         Height          =   3945
         Left            =   1410
         TabIndex        =   24
         Top             =   1980
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   6959
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
      End
      Begin VB.Label lblMast 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Master Rights"
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
         Left            =   120
         TabIndex        =   23
         Top             =   1980
         Width           =   1185
      End
      Begin VB.Label lblUserNameCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:-"
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
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[USER NAME]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1500
         TabIndex        =   21
         Top             =   1560
         Width           =   1050
      End
      Begin VB.Label lblTypeCap 
         AutoSize        =   -1  'True
         Caption         =   "User Type:-"
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
         Index           =   0
         Left            =   2910
         TabIndex        =   20
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label lblUType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[USER TYPE]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4140
         TabIndex        =   19
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmUser.frx":0000
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Index           =   4
         Left            =   150
         TabIndex        =   10
         Top             =   300
         Width           =   8235
      End
   End
   Begin VB.Frame frMain 
      Caption         =   "User Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1425
      Index           =   2
      Left            =   240
      TabIndex        =   111
      Top             =   3840
      Width           =   2535
      Begin VB.TextBox txtUserName 
         Height          =   375
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   119
         Top             =   1680
         Width           =   3435
      End
      Begin VB.Frame frType 
         Caption         =   "Type of Users"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   60
         TabIndex        =   112
         Top             =   2160
         Width           =   8865
         Begin VB.OptionButton optType 
            Caption         =   "Administartor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   115
            Top             =   450
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton optType 
            Caption         =   "HOD (Deparrmeantal Head)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   210
            TabIndex        =   114
            Top             =   1290
            Width           =   3045
         End
         Begin VB.OptionButton optType 
            Caption         =   "General User"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   113
            Top             =   2340
            Width           =   1935
         End
         Begin VB.Label lblType 
            Caption         =   "If This user is assigned as ADMINISTRATOR, he will automatically be assigned all the RIGHTS and PRIVILEDGES of the application."
            Height          =   555
            Index           =   0
            Left            =   570
            TabIndex        =   118
            Top             =   780
            Width           =   7185
         End
         Begin VB.Label lblType 
            Caption         =   $"frmUser.frx":0162
            Height          =   645
            Index           =   1
            Left            =   600
            TabIndex        =   117
            Top             =   1620
            Width           =   7545
         End
         Begin VB.Label lblType 
            Caption         =   $"frmUser.frx":0232
            Height          =   525
            Index           =   2
            Left            =   600
            TabIndex        =   116
            Top             =   2730
            Width           =   7515
         End
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmUser.frx":02E7
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   360
         TabIndex        =   121
         Top             =   360
         Width           =   8235
      End
      Begin VB.Label lblUserNameCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name :-"
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
         Index           =   0
         Left            =   180
         TabIndex        =   120
         Top             =   1740
         Width           =   1125
      End
   End
   Begin VB.Frame frMain 
      Caption         =   "Other Rights"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   825
      Index           =   6
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
      Begin VB.CheckBox chkOther 
         Caption         =   "Edit Parameter"
         Height          =   255
         Index           =   1
         Left            =   390
         TabIndex        =   123
         Top             =   2700
         Width           =   1605
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Restore"
         Height          =   255
         Index           =   24
         Left            =   6300
         TabIndex        =   80
         Top             =   5100
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdSelUN 
         Caption         =   "Select/Unselect"
         Height          =   585
         Index           =   2
         Left            =   3840
         TabIndex        =   81
         Top             =   5490
         Width           =   1545
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Backup"
         Height          =   255
         Index           =   23
         Left            =   6300
         TabIndex        =   79
         Top             =   4830
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Compact Database"
         Height          =   255
         Index           =   22
         Left            =   6300
         TabIndex        =   78
         Top             =   4560
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Edit Paid Days"
         Height          =   255
         Index           =   21
         Left            =   6300
         TabIndex        =   77
         Top             =   4080
         Width           =   2055
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Delete Old Daily Data"
         Height          =   255
         Index           =   20
         Left            =   6240
         TabIndex        =   76
         Top             =   4320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Export Data"
         Height          =   255
         Index           =   19
         Left            =   6300
         TabIndex        =   75
         Top             =   3750
         Width           =   2055
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Admin Form"
         Height          =   255
         Index           =   18
         Left            =   6300
         TabIndex        =   74
         Top             =   3480
         Width           =   2055
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "View Daily Data"
         Height          =   255
         Index           =   17
         Left            =   6330
         TabIndex        =   72
         Top             =   2970
         Width           =   2055
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "General Reports"
         Height          =   255
         Index           =   16
         Left            =   6330
         TabIndex        =   71
         Top             =   2700
         Width           =   2055
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Time"
         Height          =   255
         Index           =   15
         Left            =   4260
         TabIndex        =   69
         Top             =   4380
         Width           =   1065
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Edit CO"
         Height          =   255
         Index           =   14
         Left            =   3120
         TabIndex        =   68
         Top             =   4380
         Width           =   945
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "OT Authorization"
         Height          =   255
         Index           =   13
         Left            =   4260
         TabIndex        =   67
         Top             =   4110
         Width           =   1605
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Off Duty"
         Height          =   255
         Index           =   12
         Left            =   3120
         TabIndex        =   66
         Top             =   4110
         Width           =   945
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "On Duty"
         Height          =   255
         Index           =   11
         Left            =   4260
         TabIndex        =   65
         Top             =   3840
         Width           =   1065
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Status"
         Height          =   255
         Index           =   10
         Left            =   3120
         TabIndex        =   64
         Top             =   3840
         Width           =   945
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Record"
         Height          =   255
         Index           =   9
         Left            =   4260
         TabIndex        =   63
         Top             =   3570
         Width           =   1065
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Shift"
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   62
         Top             =   3570
         Width           =   945
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Monthly Process"
         Height          =   255
         Index           =   7
         Left            =   3120
         TabIndex        =   60
         Top             =   3030
         Width           =   2055
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Daily Process"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   59
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Edit Shift Shedule"
         Height          =   255
         Index           =   5
         Left            =   390
         TabIndex        =   57
         Top             =   4380
         Width           =   2055
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Create Shift Schedule"
         Height          =   255
         Index           =   4
         Left            =   390
         TabIndex        =   56
         Top             =   4110
         Width           =   2055
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Update Leave Balances"
         Height          =   255
         Index           =   3
         Left            =   390
         TabIndex        =   54
         Top             =   3570
         Width           =   2055
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Create Files"
         Height          =   255
         Index           =   2
         Left            =   390
         TabIndex        =   53
         Top             =   3300
         Width           =   2055
      End
      Begin VB.Line Line2 
         Index           =   3
         X1              =   9150
         X2              =   9150
         Y1              =   2280
         Y2              =   5430
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   5910
         X2              =   5910
         Y1              =   2280
         Y2              =   5430
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   2550
         X2              =   2550
         Y1              =   2280
         Y2              =   5430
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   60
         X2              =   9150
         Y1              =   5430
         Y2              =   5430
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   60
         X2              =   60
         Y1              =   2280
         Y2              =   5430
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   60
         X2              =   9150
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label lblOtherHeads 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Install"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   122
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label lblOtherHeads 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Miscellaneous Rights"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   6030
         TabIndex        =   73
         Top             =   3210
         Width           =   1920
      End
      Begin VB.Label lblOtherHeads 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Rights"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   6090
         TabIndex        =   70
         Top             =   2430
         Width           =   1230
      End
      Begin VB.Label lblOtherHeads 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Data Correction Rights"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   2850
         TabIndex        =   61
         Top             =   3300
         Width           =   2520
      End
      Begin VB.Label lblOtherHeads 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Process Rights"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2880
         TabIndex        =   58
         Top             =   2460
         Width           =   1365
      End
      Begin VB.Label lblOtherHeads 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Schedule Rights"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   55
         Top             =   3840
         Width           =   1890
      End
      Begin VB.Label lblOtherHeads 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yearly Leave Files Rights"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   52
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label lblOtherCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Rights"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lblUserNameCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:-"
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
         Index           =   4
         Left            =   90
         TabIndex        =   50
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[USER NAME]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1470
         TabIndex        =   49
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label lblTypeCap 
         AutoSize        =   -1  'True
         Caption         =   "User Type:-"
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
         Index           =   2
         Left            =   2880
         TabIndex        =   48
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label lblUType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[USER TYPE]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   4110
         TabIndex        =   47
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmUser.frx":0451
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   7215
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   500
      Left            =   6360
      TabIndex        =   8
      Top             =   1290
      Width           =   1335
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Cancel"
      Height          =   500
      Left            =   4680
      TabIndex        =   7
      Top             =   1290
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   500
      Left            =   1530
      TabIndex        =   6
      Top             =   1260
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   500
      Left            =   90
      TabIndex        =   5
      Top             =   1260
      Width           =   1335
   End
   Begin VB.Frame frMain 
      Caption         =   "Leave transaction Rights"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1185
      Index           =   5
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdSelUN 
         Caption         =   "Select/Unselect"
         Height          =   585
         Index           =   1
         Left            =   5610
         TabIndex        =   46
         Top             =   2580
         Width           =   1545
      End
      Begin VB.CheckBox chkLR 
         Caption         =   "Check1"
         Height          =   315
         Index           =   7
         Left            =   4380
         TabIndex        =   45
         Top             =   4050
         Width           =   255
      End
      Begin VB.CheckBox chkLR 
         Caption         =   "Check1"
         Height          =   315
         Index           =   6
         Left            =   3510
         TabIndex        =   44
         Top             =   4050
         Width           =   255
      End
      Begin VB.CheckBox chkLR 
         Caption         =   "Check1"
         Height          =   315
         Index           =   5
         Left            =   4380
         TabIndex        =   42
         Top             =   3660
         Width           =   255
      End
      Begin VB.CheckBox chkLR 
         Caption         =   "Check1"
         Height          =   315
         Index           =   4
         Left            =   3510
         TabIndex        =   41
         Top             =   3660
         Width           =   255
      End
      Begin VB.CheckBox chkLR 
         Caption         =   "Check1"
         Height          =   315
         Index           =   3
         Left            =   4380
         TabIndex        =   39
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox chkLR 
         Caption         =   "Check1"
         Height          =   315
         Index           =   2
         Left            =   3510
         TabIndex        =   38
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox chkLR 
         Caption         =   "Check1"
         Height          =   315
         Index           =   1
         Left            =   4380
         TabIndex        =   36
         Top             =   2850
         Width           =   255
      End
      Begin VB.CheckBox chkLR 
         Caption         =   "Check1"
         Height          =   315
         Index           =   0
         Left            =   3510
         TabIndex        =   35
         Top             =   2850
         Width           =   255
      End
      Begin VB.Label lblTrans 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avail"
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
         Left            =   1680
         TabIndex        =   43
         Top             =   4080
         Width           =   450
      End
      Begin VB.Label lblTrans 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encash"
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
         Left            =   1680
         TabIndex        =   40
         Top             =   3690
         Width           =   675
      End
      Begin VB.Label lblTrans 
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
         Index           =   1
         Left            =   1680
         TabIndex        =   37
         Top             =   3270
         Width           =   525
      End
      Begin VB.Label lblTrans 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening"
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
         Left            =   1680
         TabIndex        =   34
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label lblDel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete"
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
         Left            =   4260
         TabIndex        =   33
         Top             =   2520
         Width           =   570
      End
      Begin VB.Label lblAdd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add"
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
         Left            =   3420
         TabIndex        =   32
         Top             =   2520
         Width           =   345
      End
      Begin VB.Label lblTransCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transactions"
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
         Left            =   1650
         TabIndex        =   31
         Top             =   2520
         Width           =   1110
      End
      Begin VB.Label lblLRights 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Rights"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   30
         Top             =   2190
         Width           =   1140
      End
      Begin VB.Label lblUType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[USER TYPE]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   4260
         TabIndex        =   29
         Top             =   1710
         Width           =   1005
      End
      Begin VB.Label lblTypeCap 
         AutoSize        =   -1  'True
         Caption         =   "User Type:-"
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
         Index           =   1
         Left            =   3030
         TabIndex        =   28
         Top             =   1710
         Width           =   1005
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[USER NAME]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1620
         TabIndex        =   27
         Top             =   1710
         Width           =   1050
      End
      Begin VB.Label lblUserNameCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:-"
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
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   1710
         Width           =   1065
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmUser.frx":053B
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   150
         TabIndex        =   11
         Top             =   300
         Width           =   7665
      End
   End
   Begin VB.Frame frMain 
      Caption         =   "List of Users"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   945
      Index           =   1
      Left            =   2880
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         Height          =   555
         Left            =   6990
         TabIndex        =   18
         Top             =   2220
         Width           =   1365
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   555
         Left            =   6990
         TabIndex        =   17
         Top             =   1560
         Width           =   1365
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   3615
         Left            =   120
         TabIndex        =   16
         Top             =   1530
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmUser.frx":068E
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3525
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   7995
      End
   End
   Begin VB.Frame frMain 
      Caption         =   "Data Frame"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Index           =   8
      Left            =   120
      TabIndex        =   124
      Top             =   2760
      Width           =   2775
      Begin VB.ListBox lst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   4
         Left            =   1440
         TabIndex        =   143
         Top             =   6720
         Width           =   1695
      End
      Begin VB.ListBox lst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   3
         Left            =   1080
         TabIndex        =   142
         Top             =   6720
         Width           =   1695
      End
      Begin VB.ListBox lst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   2
         Left            =   720
         TabIndex        =   141
         Top             =   6720
         Width           =   1695
      End
      Begin VB.ListBox lst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   1
         Left            =   360
         TabIndex        =   140
         Top             =   6720
         Width           =   1695
      End
      Begin VB.ListBox lst2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   4
         Left            =   9840
         TabIndex        =   139
         Top             =   600
         Width           =   1695
      End
      Begin VB.ListBox lst2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   3
         Left            =   9480
         TabIndex        =   138
         Top             =   600
         Width           =   1695
      End
      Begin VB.ListBox lst2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   2
         Left            =   9120
         TabIndex        =   137
         Top             =   600
         Width           =   1695
      End
      Begin VB.ListBox lst2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   1
         Left            =   8760
         TabIndex        =   136
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddAll 
         Caption         =   "Add All"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   135
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "Remove All"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   134
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   133
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddD 
         Caption         =   " Add"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   132
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ListBox lst2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Index           =   0
         Left            =   5760
         TabIndex        =   131
         Top             =   1560
         Width           =   2300
      End
      Begin VB.ListBox lst 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2280
         Index           =   0
         Left            =   1560
         TabIndex        =   130
         Top             =   1560
         Width           =   2300
      End
      Begin VB.OptionButton opt 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   129
         Top             =   2760
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "Division"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   128
         Top             =   3120
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   127
         Top             =   2400
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   126
         Top             =   2040
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "Dept"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   125
         Top             =   1680
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmUser.frx":0800
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   240
         TabIndex        =   144
         Top             =   600
         Width           =   7635
      End
   End
   Begin VB.Frame frMain 
      Caption         =   "Passwords"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1305
      Index           =   7
      Left            =   3120
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
      Begin VB.Frame frLocation 
         Caption         =   "Select Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   200
         TabIndex        =   145
         Top             =   300
         Visible         =   0   'False
         Width           =   1935
         Begin VB.ListBox lstLocation 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2280
            Left            =   500
            MultiSelect     =   1  'Simple
            TabIndex        =   146
            Top             =   400
            Width           =   1000
         End
      End
      Begin VB.CheckBox chkChange 
         Caption         =   "Edit Password"
         Height          =   255
         Left            =   150
         TabIndex        =   86
         Top             =   2130
         Width           =   2415
      End
      Begin VB.Frame frSecond 
         Caption         =   "Leave Request and Correction Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   4200
         TabIndex        =   94
         Top             =   2550
         Width           =   4125
         Begin VB.TextBox txtSecond 
            Appearance      =   0  'Flat
            Height          =   345
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1590
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   100
            Top             =   1140
            Width           =   2475
         End
         Begin VB.TextBox txtSecond 
            Appearance      =   0  'Flat
            Height          =   345
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1590
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   98
            Top             =   690
            Width           =   2475
         End
         Begin VB.TextBox txtSecond 
            Appearance      =   0  'Flat
            Height          =   345
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1590
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   96
            Top             =   240
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.Label lblSecond 
            AutoSize        =   -1  'True
            Caption         =   "Confirm Password"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   99
            Top             =   1200
            Width           =   1260
         End
         Begin VB.Label lblSecond 
            AutoSize        =   -1  'True
            Caption         =   "New Password"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   97
            Top             =   750
            Width           =   1065
         End
         Begin VB.Label lblSecond 
            AutoSize        =   -1  'True
            Caption         =   "Old Password"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   95
            Top             =   330
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.Frame frLogin 
         Caption         =   "Login Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   60
         TabIndex        =   87
         Top             =   2550
         Width           =   4125
         Begin VB.TextBox txtLogin 
            Appearance      =   0  'Flat
            Height          =   345
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1590
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   93
            Top             =   1140
            Width           =   2475
         End
         Begin VB.TextBox txtLogin 
            Appearance      =   0  'Flat
            Height          =   345
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1590
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   91
            Top             =   690
            Width           =   2475
         End
         Begin VB.TextBox txtLogin 
            Appearance      =   0  'Flat
            Height          =   345
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1590
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   89
            Top             =   240
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.Label lblLogin 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   92
            Top             =   1200
            Width           =   1260
         End
         Begin VB.Label lblLogin 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New Password"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   90
            Top             =   750
            Width           =   1065
         End
         Begin VB.Label lblLogin 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Old Password"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   88
            Top             =   330
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.Label lblUType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[USER TYPE]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   4140
         TabIndex        =   85
         Top             =   1620
         Width           =   1005
      End
      Begin VB.Label lblTypeCap 
         AutoSize        =   -1  'True
         Caption         =   "User Type:-"
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
         Index           =   4
         Left            =   2910
         TabIndex        =   84
         Top             =   1620
         Width           =   1005
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[USER NAME]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1500
         TabIndex        =   83
         Top             =   1620
         Width           =   1050
      End
      Begin VB.Label lblUserNameCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:-"
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
         Index           =   5
         Left            =   120
         TabIndex        =   82
         Top             =   1620
         Width           =   1065
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmUser.frx":0896
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   8085
      End
   End
   Begin VB.Frame frMain 
      Caption         =   "HOD Details and Rights"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   825
      Index           =   3
      Left            =   5400
      TabIndex        =   101
      Top             =   120
      Width           =   4095
      Begin VB.ListBox lstRights 
         Height          =   2595
         Left            =   1470
         MultiSelect     =   1  'Simple
         TabIndex        =   102
         Top             =   2460
         Width           =   4905
      End
      Begin VB.Label lblUserNameCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:-"
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
         Index           =   1
         Left            =   150
         TabIndex        =   109
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[USER NAME]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1530
         TabIndex        =   108
         Top             =   1500
         Width           =   1050
      End
      Begin VB.Label lblTypeCap 
         AutoSize        =   -1  'True
         Caption         =   "User Type:-"
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
         Index           =   3
         Left            =   2940
         TabIndex        =   107
         Top             =   1500
         Width           =   1005
      End
      Begin VB.Label lblUType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[USER TYPE]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   4170
         TabIndex        =   106
         Top             =   1500
         Width           =   1005
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department :-"
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
         Left            =   150
         TabIndex        =   105
         Top             =   1980
         Visible         =   0   'False
         Width           =   1170
      End
      Begin MSForms.ComboBox cboDept 
         Height          =   405
         Left            =   1500
         TabIndex        =   104
         Top             =   1890
         Visible         =   0   'False
         Width           =   1545
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2725;714"
         ColumnCount     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblSelRights 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Rights"
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
         Left            =   150
         TabIndex        =   103
         Top             =   2580
         Width           =   1155
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmUser.frx":09ED
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   150
         TabIndex        =   110
         Top             =   330
         Width           =   8235
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset
Dim adrsForm As New ADODB.Recordset
'' Constants and Variables to Keep Track of Frames
Private Const TOTAL_FRAMES = 7
Dim bytCurrentFrame As Byte
'' Constants and Variables to Keep track of Control Positions, Heights and Widths
'' Form Height=7250, Width=8500
Private Const FORM_HEIGHT = 7300, FORM_WIDTH = 8500
'' Frame Height=FormHeight- (ButtonHeight+10), Width=FormWidth-20,Top=30,Left=0
Private Const FRAME_HEIGHT = FORM_HEIGHT - 975
Private Const FRAME_WIDTH = FORM_WIDTH - 100
Private Const FRAME_TOP = 30
Private Const FRAME_LEFT = 0
'' Description Label Height=1200,Width=7500, Top=300, Left=150
Private Const LABEL_TOP = 300
Private Const LABEL_LEFT = 150
Private Const LABEL_HEIGHT = 1200
Private Const LABEL_WIDTH = 7500
'' Button Height=500, Width=2000
Private Const BUTTON_HEIGHT = 500
Private Const BUTTON_WIDTH = 2000
Private Const BUTTON_TOP = FRAME_HEIGHT + 25
Private Const BUTTON_LEFT = 600
'' Array for Master Arrays
Dim strMasterTables(1 To TOTAL_MASTER_TABLES) As String
'' Other Needed Variables
'' General
Dim strCurrUser As String, strCurrType As String
''For Mauritius 07-03-2003
Dim bytCnt As Byte, bytRec As Byte
Dim bytIndex As Byte, strMaster(6, 1) As String
Dim strData() As String
''
Dim strCurrMaster As String, strCurrDeptRights As String
Dim strCurrLvRights As String, strCurrOther As String
'' For Edit Only
''For Mauritius 15-08-2003
''-> Original Dim strCurrDept As String
Dim strTmpDept As String
''
Dim strLoginPass As String, strSecondPass As String

Private Sub chkChange_Click()
If chkChange.Value = 1 Then
    frLogin.Visible = True
    frSecond.Visible = True
    Select Case bytMode
        Case 1      '' View / Edit Mode
            '' txtLogin(0).Enabled = True
            txtSecond(1).Enabled = True
            txtLogin(1).SetFocus
        Case 2      '' Add Mode
            '' txtLogin(0).Enabled = False
            txtSecond(0).Enabled = False
            txtLogin(1).SetFocus
    End Select
Else
    frLogin.Visible = False
    frSecond.Visible = False
    cmdNext.SetFocus
End If
End Sub

Private Sub cmdAdd_Click()
If (MSF1.Rows - 1) >= CByte(InVar.bytUse) Then
    MsgBox NewCaptionTxt("65077", adrsC), vbInformation
    Exit Sub
End If
bytMode = 2
Call cmdNext_Click
End Sub

Private Sub cmdBack_Click()
Select Case bytCurrentFrame
    Case 4
        bytCurrentFrame = 2
    Case 1
        cmdBack.Enabled = False
        cmdCan.Enabled = False
        Call ToggleCaption(False)
    ''For Mauritius 07-08-2003
    Case 7
        Select Case strCurrType
            Case ADMIN
                bytCurrentFrame = 2
            Case HOD
                bytCurrentFrame = 3
            Case GENERAL
                bytCurrentFrame = bytCurrentFrame - 1
        End Select
        Call ToggleCaption(False)
    Case 8
        bytCurrentFrame = 7
        Call ToggleCaption(False)
    ''
    Case Else
        bytCurrentFrame = bytCurrentFrame - 1
End Select
If bytCurrentFrame = 1 Then
    bytMode = 1
    bytCurrentFrame = 1
    cmdBack.Enabled = False
    cmdCan.Enabled = False
    Call ToggleCaption(False)
    Call MakeCancel(False)
End If
Call LoadFrame(bytCurrentFrame)
cmdNext.Enabled = True
End Sub

Private Sub cmdCan_Click()
bytMode = 1
bytCurrentFrame = 1
Call LoadFrame(bytCurrentFrame)
cmdNext.Enabled = True
cmdBack.Enabled = False
cmdCan.Enabled = False
MakeCancel (False)
Call ToggleCaption(False)
End Sub

Private Sub cmdDel_Click()
With MSF1
    If .row = 0 Then Exit Sub
    If UCase(.TextMatrix(.row, 0)) = UCase(UserName) Then
        MsgBox NewCaptionTxt("65078", adrsC), vbCritical
        Exit Sub
    End If
    If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion) = vbYes Then
        ConMain.Execute "Delete from UserAccs where UserName='" & _
        .TextMatrix(.row, 0) & "'"
        Call FillUserGrid
    End If
End With
End Sub

Private Sub cmdNext_Click()
'' If in View/Edit Mode
If bytMode = 1 Then
    '' Check for CurrentFrame Validations
    If Not ValidateEditFrame(bytCurrentFrame) Then
        '' Invalid Data
        Exit Sub
    Else
        '' Proper Data
        TakeEditAction (bytCurrentFrame)
    End If
End If
'' If in Add Mode
If bytMode = 2 Then
    '' Check for CurrentFrame Validations
    If Not ValidateAddFrame(bytCurrentFrame) Then
        '' Invalid Data
        Exit Sub
    Else
        '' Proper Data
        TakeAddAction (bytCurrentFrame)
    End If
End If
Call LoadFrame(bytCurrentFrame)
Call TakePostActions
End Sub

Private Sub TakePostActions()
'' On Error Resume Next
Select Case bytMode
    Case 1      '' View / Edit
        Select Case bytCurrentFrame
            Case 1  '' List of Users
                cmdBack.Enabled = False
                cmdCan.Enabled = False
                Call MakeCancel(False)
                Call ToggleCaption(False)
            Case 2  '' User Details
                cmdBack.Enabled = True
                cmdCan.Enabled = True
                MakeCancel (True)
                cmdNext.SetFocus
            Case 3  '' HOD Details
                ''For Mauritius 11-08-2003
                'cboDept.SetFocus
                cmdBack.Enabled = True
                cmdCan.Enabled = True
            Case 4  '' Master Rights
                cmdSelUN(0).SetFocus
                cmdBack.Enabled = True
                cmdCan.Enabled = True
            Case 5  '' Leave Rights
                cmdSelUN(1).SetFocus
                cmdBack.Enabled = True
                cmdCan.Enabled = True
            Case 6  '' Other Rights
                cmdSelUN(2).SetFocus
                cmdBack.Enabled = True
                cmdCan.Enabled = True
            Case 7  '' Password
                Select Case strCurrType
                    Case HOD
                        chkChange.Visible = True
                        chkChange.Value = 0
                        Call chkChange_Click
                        cmdBack.Enabled = True
                        cmdCan.Enabled = True
                    Case Else
                        chkChange.Visible = True
                        chkChange.Value = 0
                        Call chkChange_Click
                        cmdBack.Enabled = True
                        cmdCan.Enabled = True
                        Call ToggleCaption(True)
                End Select
                If GetFlagStatus("LocationRights") Then
                    Call SetLocationFrame
                End If
                
            Case 8
                chkChange.Visible = True
                'chkChange.Value = 0
                'Call chkChange_Click
                cmdBack.Enabled = True
                cmdCan.Enabled = True
                Call ToggleCaption(True)
        End Select
    Case 2      '' Add
        cmdCan.Enabled = True
        MakeCancel (True)
        Select Case bytCurrentFrame
            Case 2  '' User Details
                txtUserName.SetFocus
                cmdBack.Enabled = True
                Call MakeCancel
            Case 3  '' HOD Details
                ''For Mauritius 11-08-2003
                ''cboDept.SetFocus
                cmdBack.Enabled = True
            Case 4  '' Master Rights
                cmdSelUN(0).SetFocus
                cmdBack.Enabled = True
            Case 5  '' Leave Rights
                cmdSelUN(1).SetFocus
                cmdBack.Enabled = True
            Case 6  '' Other Rights
                cmdSelUN(2).SetFocus
                cmdBack.Enabled = True
'' ORIGINAL as per For Mauritius 18-08-2003
''            Case 7  '' Passwords
''                chkChange.Value = 1
''                chkChange.Visible = False
''                Call chkChange_Click
''                cmdBack.Enabled = True
''                Call ToggleCaption(True)
'' END
            Case 7  '' Password
                Select Case strCurrType
                    Case HOD
                        chkChange.Value = 1
                        chkChange.Visible = False
                        Call chkChange_Click
                        cmdBack.Enabled = True
                        cmdCan.Enabled = True
                    Case Else
                        chkChange.Value = 1
                        chkChange.Visible = False
                        Call chkChange_Click
                        cmdBack.Enabled = True
                        cmdCan.Enabled = True
                        If GetFlagStatus("LocationRights") Then 'Girish 28-01-10
                            Call SetLocationFrame
                        End If
                        Call ToggleCaption(True)
                End Select
            Case 8
                chkChange.Visible = True
                chkChange.Value = 0
                Call chkChange_Click
                cmdBack.Enabled = True
                cmdCan.Enabled = True
                Call ToggleCaption(True)
        End Select
End Select
End Sub

Private Sub TakeEditAction(ByVal bytValFrame)
On Error GoTo ERR_P
Dim bytTmp As Byte, strTmp As String, bytCnt As Byte
Select Case bytValFrame
    Case 1      '' User Details
        txtUserName.Enabled = False
        txtUserName.Text = strCurrUser
        Select Case strCurrType
            Case ADMIN
                optType(0).Value = True
            Case HOD
                optType(1).Value = True
            Case Else
                optType(2).Value = True
        End Select
        Call SetUserDetails
        bytCurrentFrame = bytCurrentFrame + 1
    Case 2      '' HOD Details
        If optType(0).Value = True Then         '' Admin
            strCurrType = ADMIN
            bytCurrentFrame = 7
        ElseIf optType(1).Value = True Then     '' HOD
            strCurrType = HOD
            bytCurrentFrame = 3
        Else
            strCurrType = GENERAL
            bytCurrentFrame = 4
        End If
        strCurrUser = Trim(txtUserName.Text)
        For bytTmp = lblUserName.LBound To lblUserName.UBound
            lblUserName(bytTmp).Caption = strCurrUser
            lblUType(bytTmp).Caption = strCurrType
        Next
        If strCurrType = ADMIN Then
            '' Make Administrative String
            '' Make Master Rights String
            strCurrMaster = ""
            With MSFMaster
                For bytTmp = 1 To .Rows - 1
                    strCurrMaster = strCurrMaster & _
                    IIf(.TextMatrix(bytTmp, 1) = CON_YES, "1", "0")
                    strCurrMaster = strCurrMaster & _
                    IIf(.TextMatrix(bytTmp, 2) = CON_YES, "1", "0")
                    strCurrMaster = strCurrMaster & _
                    IIf(.TextMatrix(bytTmp, 3) = CON_YES, "1", "0")
                Next
            End With
            '' Make Leave Rights String
            strCurrLvRights = ""
            For bytTmp = chkLR.LBound To chkLR.UBound
                strCurrLvRights = strCurrLvRights & chkLR(bytTmp).Value
            Next
            '' Make Other Rights String
            strCurrOther = ""
            For bytTmp = chkOther.LBound To chkOther.UBound
                strCurrOther = strCurrOther & chkOther(bytTmp).Value
            Next
        End If
        If strCurrType = HOD Then
            '' Set HOD Rights
            With lstRights
                For bytTmp = 0 To .ListCount - 1
                    If strCurrDeptRights = "" Then
                        .Selected(bytTmp) = False
                    Else
                        .Selected(bytTmp) = IIf(Mid(strCurrDeptRights, bytTmp + 1, 1) = "1", _
                        True, False)
                    End If
                Next
            End With
            ''cboDept.Value = strCurrData
        End If
        If strCurrType = GENERAL Then
            With MSFMaster
                For bytTmp = 5 To ((.Rows - 1) * 4) - 1
                    If (bytTmp Mod 4) <> 0 Then .TextArray(bytTmp) = CON_NO
                Next
                If strCurrMaster <> "" Then
                    bytCnt = 1
                    For bytTmp = 1 To Len(strCurrMaster)
                        strTmp = Mid(strCurrMaster, 1, bytCnt * 3)
                        strTmp = Right(strTmp, 3)
                        .TextMatrix(bytCnt, 1) = IIf(Mid(strTmp, 1, 1) = "1", _
                        CON_YES, CON_NO)
                        .TextMatrix(bytCnt, 2) = IIf(Mid(strTmp, 2, 1) = "1", _
                        CON_YES, CON_NO)
                        .TextMatrix(bytCnt, 3) = IIf(Mid(strTmp, 3, 1) = "1", _
                        CON_YES, CON_NO)
                        bytCnt = bytCnt + 1
                        If bytCnt = .Rows Then Exit For
                    Next
                End If
            End With
        End If
    Case 3
        strCurrDeptRights = ""
        For bytTmp = 0 To lstRights.ListCount - 1
            strCurrDeptRights = strCurrDeptRights & _
            IIf(lstRights.Selected(bytTmp) = True, "1", "0")
        Next
        ''strCurrDept = cboDept.Text
        bytCurrentFrame = 7
    Case 4
        '' Make Master Rights String
        strCurrMaster = ""
        With MSFMaster
            For bytTmp = 1 To .Rows - 1
                strCurrMaster = strCurrMaster & _
                IIf(.TextMatrix(bytTmp, 1) = CON_YES, "1", "0")
                strCurrMaster = strCurrMaster & _
                IIf(.TextMatrix(bytTmp, 2) = CON_YES, "1", "0")
                strCurrMaster = strCurrMaster & _
                IIf(.TextMatrix(bytTmp, 3) = CON_YES, "1", "0")
            Next
        End With
        bytCurrentFrame = bytCurrentFrame + 1
        If strCurrLvRights <> "" Then
            With chkLR
                For bytTmp = .LBound To .UBound
                    .Item(bytTmp).Value = 0
                    strTmp = ""
                    strTmp = Mid(strCurrLvRights, bytTmp + 1, 1)
                    .Item(bytTmp).Value = IIf(strTmp = "1", 1, 0)
                Next
            End With
        Else
            With chkLR
                For bytTmp = .LBound To .UBound
                    .Item(bytTmp).Value = 0
                Next
            End With
        End If
    Case 5
        '' Make Leave Rights String
        strCurrLvRights = ""
        For bytTmp = chkLR.LBound To chkLR.UBound
            strCurrLvRights = strCurrLvRights & chkLR(bytTmp).Value
        Next
        bytCurrentFrame = bytCurrentFrame + 1
        If strCurrOther <> "" Then
            With chkOther
                For bytTmp = .LBound To .UBound
                    .Item(bytTmp).Value = 0
                    strTmp = ""
                    strTmp = Mid(strCurrOther, bytTmp, 1)
                    .Item(bytTmp).Value = IIf(strTmp = "1", 1, 0)
                Next
            End With
        Else
            With chkOther
                For bytTmp = .LBound To .UBound
                    .Item(bytTmp).Value = 0
                Next
            End With
        End If
    Case 6
        '' Make Other Rights String
        strCurrOther = ""
        For bytTmp = chkOther.LBound To chkOther.UBound
            strCurrOther = strCurrOther & chkOther(bytTmp).Value
        Next
        bytCurrentFrame = bytCurrentFrame + 1
    ''For Mauritius 07-08-2003
    Case 7
        Select Case strCurrType
            Case HOD
                bytCurrentFrame = bytCurrentFrame + 1
                opt(0).Value = True
            Case Else
                If Not SaveUser Then Exit Sub
                Call FillUserGrid
                bytMode = 1
                bytCurrentFrame = 1
        End Select
    Case 8
        If Not EnoughData Then Exit Sub
        If Not SaveUser Then Exit Sub
        Call FillUserGrid
        bytMode = 1
        bytCurrentFrame = 1
    ''
End Select
Exit Sub
ERR_P:
    ShowError ("TakeEditAction::" & Me.Caption)
    ''Resume Next
End Sub

Private Sub TakeAddAction(ByVal bytValFrame)
On Error GoTo ERR_P
Dim bytTmp As Byte
Select Case bytValFrame
    Case 1      '' Ready to Add New user
        bytMode = 2
        Call ClearControls
        ''For Mauritius 11-08-2003
        strTmpDept = "@@@@@"
        strData = Split(strTmpDept, "@")
        For bytCnt = 0 To lst.Count - 1
            Call FillList(bytCnt)
        Next
        bytCurrentFrame = bytCurrentFrame + 1
    Case 2      '' New user Name and type Specified
        If optType(1).Value = True Then
            bytCurrentFrame = bytCurrentFrame + 1
            strCurrType = HOD
        ElseIf optType(2).Value = True Then
            bytCurrentFrame = 4
            strCurrType = GENERAL
        Else
            bytCurrentFrame = 7
            strCurrType = ADMIN
        End If
        strCurrUser = Trim(txtUserName.Text)
        For bytTmp = lblUserName.LBound To lblUserName.UBound
            lblUserName(bytTmp).Caption = strCurrUser
            lblUType(bytTmp).Caption = strCurrType
        Next
        '' Make Administrative String
        '' Make Master Rights String
        strCurrMaster = ""
        With MSFMaster
            For bytTmp = 1 To .Rows - 1
                strCurrMaster = strCurrMaster & _
                IIf(.TextMatrix(bytTmp, 1) = CON_YES, "1", "0")
                strCurrMaster = strCurrMaster & _
                IIf(.TextMatrix(bytTmp, 2) = CON_YES, "1", "0")
                strCurrMaster = strCurrMaster & _
                IIf(.TextMatrix(bytTmp, 3) = CON_YES, "1", "0")
            Next
        End With
        '' Make Leave Rights String
        strCurrLvRights = ""
        For bytTmp = chkLR.LBound To chkLR.UBound
            strCurrLvRights = strCurrLvRights & chkLR(bytTmp).Value
        Next
        '' Make Other Rights String
        strCurrOther = ""
        For bytTmp = chkOther.LBound To chkOther.UBound
            strCurrOther = strCurrOther & chkOther(bytTmp).Value
        Next
    Case 3      '' Save HOD Details
        strCurrDeptRights = ""
        For bytTmp = 0 To lstRights.ListCount - 1
            strCurrDeptRights = strCurrDeptRights & _
            IIf(lstRights.Selected(bytTmp) = True, "1", "0")
        Next
        bytCurrentFrame = 7
    Case 4      '' Master Rights
        '' Make Master Rights String
        strCurrMaster = ""
        With MSFMaster
            For bytTmp = 1 To .Rows - 1
                strCurrMaster = strCurrMaster & _
                IIf(.TextMatrix(bytTmp, 1) = CON_YES, "1", "0")
                strCurrMaster = strCurrMaster & _
                IIf(.TextMatrix(bytTmp, 2) = CON_YES, "1", "0")
                strCurrMaster = strCurrMaster & _
                IIf(.TextMatrix(bytTmp, 3) = CON_YES, "1", "0")
            Next
        End With
        bytCurrentFrame = bytCurrentFrame + 1
    Case 5      '' Leave Rights
        '' Make Leave Rights String
        strCurrLvRights = ""
        For bytTmp = chkLR.LBound To chkLR.UBound
            strCurrLvRights = strCurrLvRights & chkLR(bytTmp).Value
        Next
        bytCurrentFrame = bytCurrentFrame + 1
    Case 6      '' Other Rights
        '' Make Other Rights String
        strCurrOther = ""
        For bytTmp = chkOther.LBound To chkOther.UBound
            strCurrOther = strCurrOther & chkOther(bytTmp).Value
        Next
        bytCurrentFrame = bytCurrentFrame + 1
    ''For Mauritius 07-08-2003
    Case 7
        Select Case strCurrType
            Case HOD
                bytCurrentFrame = bytCurrentFrame + 1
                opt(0).Value = True
            Case Else
                If Not SaveUser Then Exit Sub
                Call FillUserGrid
                bytMode = 1
                bytCurrentFrame = 1
        End Select
    Case 8
        If Not EnoughData Then Exit Sub
        If Not SaveUser Then Exit Sub
        Call FillUserGrid
        bytMode = 1
        bytCurrentFrame = 1
    ''
End Select
Exit Sub
ERR_P:
    ShowError ("TakeAddAction::" & Me.Caption)
End Sub

Private Sub SetLocationFrame()
If strCurrType = "GENERAL" Then
    frLocation.Visible = True
    frLogin.Top = 2000
    frSecond.Top = 2000
    frLocation.Top = 3650
    frLocation.Left = 1700
    frLocation.Width = 5000
    frLocation.Height = 2500
    lstLocation.Top = 300
    lstLocation.Left = 200
    lstLocation.Width = 4500
    lstLocation.Height = 2200
    
    If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "select * From Location Order By Location", ConMain, adOpenStatic, adLockOptimistic
        If Not (adrsTemp.EOF And adrsTemp.BOF) Then
            lstLocation.clear
            Do While Not adrsTemp.EOF
                lstLocation.AddItem adrsTemp(0) & ":: " & adrsTemp(1)
                adrsTemp.MoveNext
            Loop
        End If
Else
    frLocation.Visible = False
End If
End Sub

''For Mauritius 09-08-2003
Private Function EnoughData() As Boolean
On Error GoTo ERR_P
If lst2(0).ListCount <= 0 Then
    MsgBox NewCaptionTxt("65091", adrsC), vbInformation
    opt(0).Value = True
    Exit Function
End If
If lst2(1).ListCount <= 0 Then
    MsgBox NewCaptionTxt("65092", adrsC), vbInformation
    opt(1).Value = True
    Exit Function
End If
If lst2(2).ListCount <= 0 Then
    MsgBox NewCaptionTxt("65093", adrsC), vbInformation
    opt(2).Value = True
    Exit Function
End If
If lst2(3).ListCount <= 0 Then
    MsgBox NewCaptionTxt("65094", adrsC), vbInformation
    opt(3).Value = True
    Exit Function
End If
If lst2(4).ListCount <= 0 Then
    MsgBox NewCaptionTxt("65095", adrsC), vbInformation
    opt(4).Value = True
    Exit Function
End If
EnoughData = True
Exit Function
ERR_P:
    ShowError ("EnoughData :: " & Me.Caption)
End Function
''
Private Function ValidateEditFrame(ByVal bytValFrame) As Boolean
On Error GoTo ERR_P
Select Case bytValFrame
    Case 1              '' List of Users
        With MSF1
            If .row = 0 Then Exit Function
            If Trim(.TextMatrix(.row, 0)) = "" Then Exit Function
            strCurrUser = UCase(Trim(.TextMatrix(.row, 0)))
            strCurrType = UCase(Trim(.TextMatrix(.row, 1)))
            strTmpDept = UCase(Trim(.TextMatrix(.row, 2)))
            '' bytCurrentFrame = bytCurrentFrame + 1
        End With
    Case 2, 4, 5, 6     '' Non-Validable Frames
        '' No Validations
    Case 3              '' HOD Details
        ''For Mauritius 11-08-2003
        ''If Not ValidDepartment Then Exit Function
    Case 7
        If chkChange.Value = 0 Then
            ValidateEditFrame = True
            Exit Function
        End If

''
        If Not ValidPasswords Then Exit Function
End Select
ValidateEditFrame = True
Exit Function
ERR_P:
    ShowError ("ValidateEditFrame::" & Me.Caption)
End Function

Private Function ValidateAddFrame(ByVal bytValFrame) As Boolean
On Error GoTo ERR_P
Dim bytTmp As Byte
Select Case bytValFrame
    Case 1      '' List of Users
    Case 2      '' User Details
        '' Check for Print User Name
        If UCase(Trim(txtUserName.Text)) = UCase(strPrintUser) Then
            txtUserName.Text = ""
            txtUserName.SetFocus
            Exit Function
        End If
        '' Check if Any User Name is Entered
        If Trim(txtUserName.Text) = "" Then
            MsgBox NewCaptionTxt("65079", adrsC)
            txtUserName.SetFocus
            Exit Function
        End If
        '' Check if the User Name Already Exists
        With MSF1
            For bytTmp = 1 To .Rows - 1
                If UCase(.TextMatrix(bytTmp, 0)) = UCase(Trim(txtUserName.Text)) Then
                    MsgBox NewCaptionTxt("65080", adrsC), vbInformation
                    txtUserName.SetFocus
                    Exit Function
                End If
            Next
        End With
    Case 3      '' HOD Details
        ''For Mauritius 11-08-2003
        ''If Not ValidDepartment Then Exit Function
    Case 4      '' Master Rights
        '' No Special Validations
    Case 5      '' Leave Rights
        '' No Special Validations
    Case 6      '' Other Rights
        '' No Special Validations
    Case 7      '' Passwords
        '' Check for Blank Login Password
        If Not ValidPasswords Then Exit Function
End Select
ValidateAddFrame = True
Exit Function
ERR_P:
    ShowError ("ValidateAddFrame::" & Me.Caption)
End Function

Private Function ValidPasswords() As Boolean
On Error GoTo ERR_P
'' Login Passwords
If Trim(txtLogin(1).Text) = "" Then
    MsgBox NewCaptionTxt("65081", adrsC), vbInformation
    txtLogin(1).SetFocus
    Exit Function
End If
If Trim(txtLogin(1).Text) <> Trim(txtLogin(2).Text) Then
    MsgBox NewCaptionTxt("65082", adrsC), vbInformation
    txtLogin(1).SetFocus
    Exit Function
End If
If InStr(DEncryptDat(UCase(Trim(txtLogin(1).Text)), 1), "'") Then
    MsgBox NewCaptionTxt("65083", adrsC), vbInformation
    txtLogin(1).SetFocus
    Exit Function
End If
If InStr(DEncryptDat(UCase(Trim(txtLogin(1).Text)), 1), Chr(34)) Then
    MsgBox NewCaptionTxt("65083", adrsC), vbInformation
    txtLogin(1).SetFocus
    Exit Function
End If
'' Second Level Passwords
If Trim(txtSecond(1).Text) = "" Then
    MsgBox NewCaptionTxt("65084", adrsC), vbInformation
    txtSecond(1).SetFocus
    Exit Function
End If
If Trim(txtSecond(1).Text) <> Trim(txtSecond(2).Text) Then
    MsgBox NewCaptionTxt("65085", adrsC), vbInformation
    txtSecond(1).SetFocus
    Exit Function
End If
If InStr(DEncryptDat(UCase(Trim(txtSecond(1).Text)), 1), "'") Then
    MsgBox NewCaptionTxt("65086", adrsC), vbInformation
    txtSecond(1).SetFocus
    Exit Function
End If
If InStr(DEncryptDat(UCase(Trim(txtSecond(1).Text)), 1), Chr(34)) Then
    MsgBox NewCaptionTxt("65086", adrsC), vbInformation
    txtSecond(1).SetFocus
    Exit Function
End If
ValidPasswords = True
Exit Function
ERR_P:
    ShowError ("ValidPasswords::" & Me.Caption)
End Function

Private Function ValidDepartment() As Boolean
On Error GoTo ERR_P
If Trim(cboDept.Text) = "" Then
    MsgBox NewCaptionTxt("65087", adrsC), vbInformation
    cboDept.SetFocus
    Exit Function
End If
ValidDepartment = True
Exit Function
ERR_P:
    ShowError ("ValidDepartment::" & Me.Caption)
End Function

Private Sub SetUserDetails()
On Error Resume Next
Dim bytTmp As Byte, strTmp As String, strMidSet As String
adrsForm.MoveFirst
adrsForm.Find "UserName='" & strCurrUser & "'"
'' Login Password
If IsNull(adrsForm("Password")) Then
    strLoginPass = ""
Else
    strLoginPass = DEncryptDat(adrsForm("Password"), 1)
End If
'' HOD Rights
strCurrDeptRights = IIf(IsNull(adrsForm("HODRights")), "", adrsForm("HODRights"))
'' Master Rights
strCurrMaster = IIf(IsNull(adrsForm("MasterRights")), "", adrsForm("MasterRights"))
'' Leave Rights
strCurrLvRights = IIf(IsNull(adrsForm("LeaveRights")), "", adrsForm("LeaveRights"))
'' Other Rights
strCurrOther = IIf(IsNull(adrsForm("OtherRights1")), "", adrsForm("OtherRights1"))
'' Second Password
If IsNull(adrsForm("OtherPass1")) Then
    strSecondPass = ""
Else
    strSecondPass = DEncryptDat(adrsForm("OtherPass1"), 1)
End If
''For Mauritius 09-08-2003
strTmpDept = IIf(IsNull(adrsForm("Dept")) Or Trim(adrsForm("Dept")) = "", "@@@@@", adrsForm("Dept"))
strData = Split(strTmpDept, "@")
For bytCnt = 0 To lst.Count - 1
    Call FillList(bytCnt)
Next
''

End Sub

Private Function SaveUser() As Boolean
On Error GoTo ERR_P
''For Mauritius 07-08-2003
Dim strVals As String
If strCurrType = HOD Then
    For bytCnt = 0 To lst2.Count - 1
        For bytRec = 0 To lst2(bytCnt).ListCount - 1
            strVals = strVals & Left(lst2(bytCnt).List(bytRec), InStr(1, lst2(bytCnt).List(bytRec), ":") - 1) & ","
        Next
        strVals = strVals & "@"
    Next
End If
''
Select Case bytMode
    Case 1
        ConMain.Execute "Update UserAccs Set UserType='" & _
        strCurrType & "',Dept='" & strVals & "',HODRights='" & _
        strCurrDeptRights & "',MasterRights='" & strCurrMaster & "',LeaveRights='" & _
        strCurrLvRights & "',OtherRights1='" & strCurrOther & "',UserModDate=" & _
        strDTEnc & DateCompStr(Date) & strDTEnc & ",UserModUser='" & _
        IIf(UCase(UserName) = strPrintUser, "******", UserName) & _
        "' Where UserName='" & strCurrUser & "'"
        If chkChange.Value = 1 Then
            ConMain.Execute "Update UserAccs Set Password='" & _
            DEncryptDat(UCase(Trim(txtLogin(1).Text)), 1) & "'," & _
            "OtherPass1='" & DEncryptDat(UCase(Trim(txtSecond(1).Text)), 1) & _
            "' Where UserName='" & strCurrUser & "'"
        End If
    Case 2
        ConMain.Execute "insert into UserAccs Values('" & _
        strCurrUser & "','" & DEncryptDat(UCase(Trim(txtLogin(1).Text)), 1) & _
        "','" & strCurrType & "','" & strVals & "','" & _
        strCurrDeptRights & "','" & strCurrMaster & _
        "','" & strCurrLvRights & "','" & strCurrOther & "','','" & _
        DEncryptDat(UCase(Trim(txtSecond(1).Text)), 1) & "',''," & strDTEnc & _
        DateCompStr(Date) & strDTEnc & ",'" & IIf(UCase(UserName) = strPrintUser, _
        "******", UserName) & "'," & strDTEnc & DateCompStr(Date) & strDTEnc & _
        ",'" & IIf(UCase(UserName) = strPrintUser, "******", UserName) & "')"
End Select
If strCurrType = "GENERAL" And GetFlagStatus("LocationRights") Then 'Girish
    Call LocationRights
End If
SaveUser = True
Exit Function
ERR_P:
    ShowError ("SaveUser::" & Me.Caption)
    Resume Next
End Function

Private Sub LocationRights()    'Girish
    Dim Loc As String
    For i = 0 To lstLocation.ListCount - 1
        If lstLocation.Selected(i) Then
            Loc = Loc & Left(lstLocation.List(i), InStr(1, lstLocation.List(i), ":") - 1) & ","
        End If
    Next
    If Len(Trim(Loc)) > 0 Then Loc = Left(Loc, Len(Loc) - 1)
    ConMain.Execute "Update UserAccs Set Dept = '" & Loc & "'Where UserName='" & strCurrUser & "'"
End Sub

Private Sub ClearControls()
'' On Error Resume Next
Dim bytTmp As Byte
'' Frame 2 - User Type and Name Selection
txtUserName.Enabled = True
txtUserName.Text = ""
'' Frame 3 - HOD Department and Rights Selection
cboDept.Value = ""
For bytTmp = 0 To lstRights.ListCount - 1
    lstRights.Selected(bytTmp) = False
Next
'' Frame 4 - Master Rights
With MSFMaster
    For bytTmp = 1 To .Rows - 1
        .TextMatrix(bytTmp, 1) = CON_NO
        .TextMatrix(bytTmp, 2) = CON_NO
        .TextMatrix(bytTmp, 3) = CON_NO
    Next
End With
'' Frame 5 - Leave Rights
With chkLR
    For bytTmp = .LBound To .UBound
        .Item(bytTmp).Value = 0
    Next
End With
'' Frame 6 - Other Rights
With chkOther
    For bytTmp = .LBound To .UBound
        .Item(bytTmp).Value = 0
    Next
End With
'' Frame 7 - Passwords
With txtLogin
    For bytTmp = .LBound To .UBound
        .Item(bytTmp).Text = ""
    Next
    .Item(0).Enabled = False
End With
With txtSecond
    For bytTmp = .LBound To .UBound
        .Item(bytTmp).Text = ""
    Next
    .Item(0).Enabled = False
End With
'' Clear Variables
strCurrUser = "": strCurrType = ""
strCurrDeptRights = ""
strCurrMaster = "": strCurrLvRights = "": strCurrOther = ""
End Sub

Private Sub ToggleCaption(Optional blnSave As Boolean = False)
Select Case blnSave
    Case True
        cmdNext.Caption = "Save"
    Case False
        cmdNext.Caption = "Next"
End Select
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSelUN_Click(Index As Integer)
Dim bytTmp As Byte
Static blnSelUnSel(0 To 2) As Boolean
Select Case Index
    Case 0      '' Master Rights
        With MSFMaster
            For bytTmp = 1 To .Rows - 1
                .TextMatrix(bytTmp, 1) = IIf(blnSelUnSel(0), CON_NO, CON_YES)
                .TextMatrix(bytTmp, 2) = IIf(blnSelUnSel(0), CON_NO, CON_YES)
                .TextMatrix(bytTmp, 3) = IIf(blnSelUnSel(0), CON_NO, CON_YES)
            Next
        End With
        blnSelUnSel(0) = Not blnSelUnSel(0)
    Case 1      '' Leave Transaction Rights
        With chkLR
            For bytTmp = .LBound To .UBound
                .Item(bytTmp).Value = IIf(blnSelUnSel(1), 0, 1)
            Next
        End With
        blnSelUnSel(1) = Not blnSelUnSel(1)
    Case 2      '' Other Rights
        With chkOther
            For bytTmp = .LBound To .UBound
                .Item(bytTmp).Value = IIf(blnSelUnSel(2), 0, 1)
            Next
        End With
        blnSelUnSel(2) = Not blnSelUnSel(2)
End Select
End Sub

Private Sub Form_Activate()
If strCurrentUserType <> ADMIN Then
    MsgBox NewCaptionTxt("00001", adrsMod)
    Unload Me
End If
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Call SetFormIcon(Me)
Call RetCaptions
Call LoadSpecifics
Exit Sub
ERR_P:
    ShowError ("Load::" & Me.Caption)
End Sub

Private Sub RetCaptions()                   '' Gets and Sets the Form Captions
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '65%'", ConMain, adOpenStatic

cmdRemove.Caption = "Remove"           ''Remove
cmdAddAll.Caption = "Add All"          ''Add All
cmdRemoveAll.Caption = "Remove All"    ''Remove All
'lblDesc(9).Caption = NewCaptionTxt("65096", adrsC)
''
End Sub

Private Sub LoadSpecifics()
'' On Error Resume Next
Dim bytTmp As Byte
With Me
    .Height = FORM_HEIGHT
    .Width = FORM_WIDTH
End With
bytCurrentFrame = 1
Call LoadFrame(bytCurrentFrame)
Call SetButtonPositions
Call CapGrid
Call FillControls
Call FillArrayOfMasterTables
Call LoadMasterGrid
'' Button Adjustments
cmdBack.Enabled = False
cmdCan.Enabled = False
Call MakeCancel(False)
'' Make Empty Labels
For bytTmp = lblUserName.LBound To lblUserName.UBound
    lblUserName(bytTmp).Caption = ""
    lblUType(bytTmp).Caption = ""
Next
'' Set Default Mode as View Mode
bytMode = 1
'' Open the Users Table Recordset
Call OpenMasterTable
Call FillUserGrid
Call FillArray
End Sub

''For Mauritius 07-08-2003
''Start
Private Sub FillArray()
strMaster(0, 0) = "Deptdesc": strMaster(0, 1) = "Dept," & strKDesc
strMaster(1, 0) = "Company": strMaster(1, 1) = "Company,cname"
strMaster(2, 0) = "Groupmst": strMaster(2, 1) = strKGroup & ",grupdesc"
strMaster(3, 0) = "Division": strMaster(3, 1) = "Div,divdesc"
strMaster(4, 0) = "Location": strMaster(4, 1) = "Location,LocDesc"
End Sub

Private Sub FillList(bytTmp As Byte)
On Error GoTo ERR_P
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "select " & strMaster(bytTmp, 1) & " from " & strMaster(bytTmp, 0) & " order by " & strMaster(bytTmp, 1), ConMain, adOpenStatic, adLockOptimistic
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    lst(bytTmp).clear
    lst2(bytTmp).clear
    Do While Not adrsTemp.EOF
        If UBound(strData) > 0 Then
            If InStr(1, strData(bytTmp), adrsTemp(0)) > 0 Then
                lst2(bytTmp).AddItem adrsTemp(0) & ":: " & adrsTemp(1)
            Else
                lst(bytTmp).AddItem adrsTemp(0) & ":: " & adrsTemp(1)
            End If
        Else
            lst(bytTmp).AddItem adrsTemp(0) & ":: " & adrsTemp(1)
        End If
        adrsTemp.MoveNext
    Loop
End If
Exit Sub
ERR_P:
    ShowError ("FillList :: " & Me.Caption)
    ''Resume Next
End Sub
''End
Private Sub FillUserGrid()
On Error GoTo ERR_P
Dim bytTmp As Byte
adrsForm.Requery
With MSF1
    If adrsForm.RecordCount = 0 Then
        .Rows = 1
        Exit Sub
    End If
    adrsForm.MoveFirst
    Do While Not adrsForm.EOF
        bytTmp = bytTmp + 1
        .Rows = bytTmp + 1
        .TextMatrix(bytTmp, 0) = adrsForm("UserName")
        .TextMatrix(bytTmp, 1) = adrsForm("UserType")
        If adrsForm("UserType") = HOD Then
            '.TextMatrix(bytTmp, 2) = IIf(IsNull(adrsForm("Dept")), "", _
            adrsForm("Dept"))
            strTmpDept = IIf(IsNull(adrsForm("Dept")), "", adrsForm("Dept"))
        Else
            .TextMatrix(bytTmp, 2) = ""
        End If
        adrsForm.MoveNext
    Loop
End With
Exit Sub
ERR_P:
    ShowError ("FillUserGrid::" & Me.Caption)
End Sub

Private Sub OpenMasterTable()
On Error GoTo ERR_P
adrsForm.ActiveConnection = ConMain
adrsForm.CursorType = adOpenStatic
adrsForm.LockType = adLockReadOnly
If adrsForm.State = 1 Then adrsForm.Close
adrsForm.Open "Select * from UserAccs order by UserType"
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable::" & Me.Caption)
End Sub

Private Sub MakeCancel(Optional blnTmp As Boolean = True)
cmdCan.Cancel = blnTmp
cmdExit.Cancel = Not blnTmp
End Sub

Private Sub LoadMasterGrid()
On Error GoTo ERR_P
Dim bytTmp As Byte
For bytTmp = 1 To TOTAL_MASTER_TABLES
    With MSFMaster
        .TextMatrix(bytTmp, 0) = strMasterTables(bytTmp)
    End With
Next
Exit Sub
ERR_P:
    ShowError ("LoadMasterControls:" & Me.Caption)
End Sub

Private Sub FillControls()
On Error GoTo ERR_P
'' Fill Department Combo
Call ComboFill(cboDept, 2, 2)
'' Fill Rights Combo
With lstRights
        .clear
        .AddItem "Manipulate Employee Details"
        .AddItem "Manipulate Shifts"
        .AddItem "Manipulate Daily Data"
        .AddItem "Manipulate Monthly Data"
        .AddItem "Manipulate Leave Transactions"
        .AddItem "View-Generate Reports"
        .AddItem "Export Data"
End With
Exit Sub
ERR_P:
    ShowError ("FillControls::" & Me.Caption)
End Sub

Private Sub CapGrid()
'' On Error Resume Next
Dim bytTmp As Byte
With MSF1
    .Rows = 10
    '' Set Captions
    .TextMatrix(0, 0) = NewCaptionTxt("65037", adrsC)       ''User Name
    .TextMatrix(0, 1) = NewCaptionTxt("65035", adrsC)       ''User Type
    .TextMatrix(0, 2) = NewCaptionTxt("00058", adrsMod)     ''Department
    '' Set Sizes
    ''For Mauritius 19-08-2003
    .ColWidth(0) = .ColWidth(0) * 3.5 '2.15
    .ColWidth(1) = .ColWidth(1) * 3.5 '2.15
    .ColWidth(2) = .ColWidth(2) * 0
    '' Set Alignment
    .ColAlignment(0) = flexAlignCenterCenter
    .ColAlignment(1) = flexAlignCenterCenter
    .ColAlignment(2) = flexAlignCenterCenter
End With
With MSFMaster
    .Rows = TOTAL_MASTER_TABLES + 1
    .TextMatrix(0, 0) = "MASTER" 'NewCaptionTxt("65075", adrsC)   ''Table Name
    .TextMatrix(0, 1) = NewCaptionTxt("65003", adrsC)   ''Add
    .TextMatrix(0, 2) = NewCaptionTxt("65076", adrsC)   ''Edit
    .TextMatrix(0, 3) = NewCaptionTxt("65004", adrsC)   ''Delete
    .ColWidth(0) = .ColWidth(0) * 2.25
    .ColWidth(1) = .ColWidth(1) * 0.8
    .ColWidth(2) = .ColWidth(2) * 0.8
    .ColWidth(3) = .ColWidth(3) * 0.8
    For bytTmp = 0 To .Cols - 1
        .ColAlignment(bytTmp) = flexAlignCenterCenter
    Next
End With
End Sub

Private Sub FillArrayOfMasterTables()
'' On Error Resume Next
Dim bytTmp As Byte
For bytTmp = 1 To TOTAL_MASTER_TABLES
''For Mauritius 19-08-2003
    strMasterTables(bytTmp) = Choose(bytTmp, "CATEGORY", "DEPARTMENT", "GROUP", _
    "LOCATION", "SHIFT", "ROTATION", "OTRULES", "CORULES", _
    "EMPLOYEE", "HOLIDAY", "DECLARE", "LEAVES", "MANUAL ENTRY")
Next
End Sub

Private Sub SetButtonPositions()
'' On Error Resume Next
'' Manouverabilty Buttons
With cmdBack
    .Top = BUTTON_TOP
    .Left = BUTTON_LEFT
End With
With cmdNext
    .Top = BUTTON_TOP
    .Left = BUTTON_LEFT + cmdBack.Width + 30
End With
'' Action Buttons Buttons
With cmdCan
    .Top = BUTTON_TOP
    .Left = BUTTON_LEFT + cmdBack.Width + cmdNext.Width + 1000
End With
'' Form Buttons
With cmdExit
    .Top = BUTTON_TOP
    .Left = BUTTON_LEFT + cmdBack.Width + cmdNext.Width + cmdCan.Width + 2000
End With
End Sub

Private Sub LoadFrame(ByVal bytFrameToLoad As Byte)
'' On Error resume Next
Dim bytTmp As Byte
If bytFrameToLoad = 0 Then Exit Sub
frMain(bytFrameToLoad).Top = FRAME_TOP
frMain(bytFrameToLoad).Left = FRAME_LEFT
frMain(bytFrameToLoad).Height = FRAME_HEIGHT
frMain(bytFrameToLoad).Width = FRAME_WIDTH
lblDesc(bytFrameToLoad).Top = LABEL_TOP
lblDesc(bytFrameToLoad).Left = LABEL_LEFT
lblDesc(bytFrameToLoad).Height = LABEL_HEIGHT
lblDesc(bytFrameToLoad).Width = LABEL_WIDTH
For bytTmp = 1 To frMain.Count
    If bytTmp = bytFrameToLoad Then
        frMain(bytTmp).Visible = True
    Else
        frMain(bytTmp).Visible = False
    End If
Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bytCurrentFrame <> 1 Then
    If MsgBox(NewCaptionTxt("65088", adrsC), vbYesNo + vbQuestion) = vbNo Then Cancel = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call FreeRes
End Sub

Private Sub FreeRes()
On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
If adrsForm.State = 1 Then adrsForm.Close
Set adrsC = Nothing
Set adrsForm = Nothing
Erase strMasterTables
End Sub



Private Sub MSF1_DblClick()
Call cmdNext_Click
End Sub
''For Mauritius 07-08-2003
''Start
Private Sub opt_Click(Index As Integer)
bytIndex = Index
For bytCnt = 0 To lst.Count - 1
    lst(bytCnt).Visible = False
    lst2(bytCnt).Visible = False
Next
lst(Index).Left = lst(0).Left
lst(Index).Top = lst(0).Top
lst(Index).Width = lst(0).Width
lst(Index).Appearance = 0
lst(Index).Visible = True

lst2(Index).Left = lst2(0).Left
lst2(Index).Top = lst2(0).Top
lst2(Index).Width = lst2(0).Width
lst2(Index).Appearance = 0
lst2(Index).Visible = True
End Sub

Private Sub cmdAddD_Click()
If Trim(lst(bytIndex).Text) <> "" Then
    lst2(bytIndex).AddItem lst(bytIndex).Text
End If
If lst(bytIndex).ListIndex >= 0 Then
    lst(bytIndex).RemoveItem lst(bytIndex).ListIndex
End If
End Sub

Private Sub cmdAddAll_Click()
If lst(bytIndex).ListCount > 0 Then
    For bytCnt = 0 To lst(bytIndex).ListCount - 1
        lst2(bytIndex).AddItem lst(bytIndex).List(bytCnt)
    Next
    lst(bytIndex).clear
End If
End Sub

Private Sub cmdRemove_Click()
If Trim(lst2(bytIndex).Text) <> "" Then
    lst(bytIndex).AddItem lst2(bytIndex).Text
End If
If lst2(bytIndex).ListIndex >= 0 Then
    lst2(bytIndex).RemoveItem lst2(bytIndex).ListIndex
End If
End Sub

Private Sub cmdRemoveAll_Click()
If lst2(bytIndex).ListCount > 0 Then
    For bytCnt = 0 To lst2(bytIndex).ListCount - 1
        lst(bytIndex).AddItem lst2(bytIndex).List(bytCnt)
    Next
    lst2(bytIndex).clear
End If
End Sub
''End
Private Sub MSFMaster_DblClick()
'' On Error Resume Next
With MSFMaster
    If .row = 0 Then Exit Sub
    If .Col = 0 Then Exit Sub
    If .Text = CON_YES Then
        .Text = CON_NO
    Else
        .Text = CON_YES
    End If
End With
End Sub

Private Sub txtLogin_GotFocus(Index As Integer)
    Call GF(txtLogin(Index))
End Sub

Private Sub txtLogin_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9)
End Sub

Private Sub txtSecond_GotFocus(Index As Integer)
    Call GF(txtSecond(Index))
End Sub

Private Sub txtSecond_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9)
End Sub

Private Sub txtUserName_GotFocus()
    Call GF(txtUserName)
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF10 Then Call ShowF10("65")
If KeyCode = vbKeyF10 Then KeyCode = 0
End Sub

