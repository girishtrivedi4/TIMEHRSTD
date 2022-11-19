VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12270
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
   ScaleHeight     =   6480
   ScaleWidth      =   12270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5670
      TabIndex        =   151
      Top             =   4830
      Width           =   2115
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8820
      Top             =   3450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   285
      Left            =   0
      TabIndex        =   137
      TabStop         =   0   'False
      Top             =   6150
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   503
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      ForeColor       =   16744576
      FocusRect       =   0
      HighLight       =   0
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
   Begin VB.Frame frPeri 
      Caption         =   "Periodic &Reports"
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
      Left            =   9000
      TabIndex        =   84
      Top             =   2640
      Width           =   8745
      Begin VB.OptionButton optPer 
         Caption         =   "Unprocessed Leaves"
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
         Index           =   10
         Left            =   4890
         TabIndex        =   148
         Top             =   1245
         Width           =   2115
      End
      Begin VB.OptionButton optPer 
         Caption         =   "Permission cards"
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
         Index           =   9
         Left            =   4890
         TabIndex        =   147
         Top             =   1725
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.OptionButton optPer 
         Caption         =   "Summary"
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
         Left            =   4890
         TabIndex        =   146
         Top             =   900
         Width           =   1995
      End
      Begin VB.OptionButton optPer 
         Caption         =   "Continuous Absent"
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
         Left            =   2580
         TabIndex        =   94
         Top             =   1600
         Width           =   1995
      End
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
         Left            =   3930
         MaxLength       =   10
         TabIndex        =   86
         Tag             =   "D"
         Top             =   240
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
         Left            =   5700
         MaxLength       =   10
         TabIndex        =   88
         Tag             =   "D"
         Top             =   240
         Width           =   1155
      End
      Begin VB.OptionButton optPer 
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
         Height          =   240
         Index           =   0
         Left            =   210
         TabIndex        =   89
         Top             =   900
         Width           =   2355
      End
      Begin VB.OptionButton optPer 
         Caption         =   "Muster Report"
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
         Left            =   210
         TabIndex        =   90
         Top             =   1250
         Width           =   2355
      End
      Begin VB.OptionButton optPer 
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
         Index           =   2
         Left            =   210
         TabIndex        =   91
         Top             =   1600
         Width           =   2115
      End
      Begin VB.OptionButton optPer 
         Caption         =   "Late Arrival "
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
         Left            =   2610
         TabIndex        =   92
         Top             =   900
         Width           =   1995
      End
      Begin VB.OptionButton optPer 
         Caption         =   "Early Departure"
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
         Left            =   2580
         TabIndex        =   93
         Top             =   1250
         Width           =   2115
      End
      Begin VB.Label lblNotePeri 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reports Available for Maximum period of 30/31 days"
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
         Left            =   1110
         TabIndex        =   136
         Top             =   660
         Width           =   5415
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
         Left            =   1110
         TabIndex        =   85
         Top             =   270
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
         Left            =   5220
         TabIndex        =   87
         Top             =   270
         Width           =   300
      End
   End
   Begin VB.Frame frMast 
      Caption         =   "Master &Reports"
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
      Left            =   9000
      TabIndex        =   70
      Top             =   540
      Width           =   8745
      Begin VB.OptionButton optmas 
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
         Index           =   12
         Left            =   5790
         TabIndex        =   145
         Top             =   660
         Width           =   1305
      End
      Begin VB.OptionButton optmas 
         Caption         =   "Cost Centre"
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
         Index           =   11
         Left            =   4350
         TabIndex        =   140
         Top             =   1770
         Width           =   2145
      End
      Begin VB.OptionButton optmas 
         Caption         =   "Cost Centre"
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
         Index           =   10
         Left            =   4350
         TabIndex        =   139
         Top             =   1410
         Width           =   2115
      End
      Begin VB.OptionButton optmas 
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
         Height          =   270
         Index           =   9
         Left            =   4350
         TabIndex        =   138
         Top             =   1050
         Width           =   1155
      End
      Begin VB.OptionButton optmas 
         Caption         =   "Employee List"
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
         Left            =   300
         TabIndex        =   75
         Top             =   660
         Width           =   1635
      End
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
         Left            =   4620
         MaxLength       =   10
         TabIndex        =   74
         Tag             =   "D"
         Top             =   210
         Visible         =   0   'False
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
         Left            =   2910
         MaxLength       =   10
         TabIndex        =   72
         Tag             =   "D"
         Top             =   210
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optmas 
         Caption         =   "Employee Details"
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
         Left            =   300
         TabIndex        =   76
         Top             =   1050
         Width           =   1935
      End
      Begin VB.OptionButton optmas 
         Caption         =   "Left Employee"
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
         Left            =   300
         TabIndex        =   77
         Top             =   1410
         Width           =   1965
      End
      Begin VB.OptionButton optmas 
         Caption         =   "Leave"
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
         Left            =   300
         TabIndex        =   78
         Top             =   1770
         Width           =   885
      End
      Begin VB.OptionButton optmas 
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
         Height          =   285
         Index           =   4
         Left            =   2460
         TabIndex        =   79
         Top             =   660
         Width           =   735
      End
      Begin VB.OptionButton optmas 
         Caption         =   "Rotational Shift"
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
         Index           =   5
         Left            =   2460
         TabIndex        =   80
         Top             =   1050
         Width           =   1635
      End
      Begin VB.OptionButton optmas 
         Caption         =   "Holiday"
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
         Left            =   2460
         TabIndex        =   81
         Top             =   1410
         Width           =   1005
      End
      Begin VB.OptionButton optmas 
         Caption         =   "B. Unit"
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
         Index           =   7
         Left            =   2460
         TabIndex        =   82
         Top             =   1770
         Width           =   1425
      End
      Begin VB.OptionButton optmas 
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
         Height          =   270
         Index           =   8
         Left            =   4350
         TabIndex        =   83
         Top             =   660
         Width           =   1155
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
         Left            =   4110
         TabIndex        =   73
         Top             =   270
         Visible         =   0   'False
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
         Left            =   1770
         TabIndex        =   71
         Top             =   270
         Visible         =   0   'False
         Width           =   990
      End
   End
   Begin VB.Frame frYear 
      Caption         =   "Yearly &Reports"
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
      Left            =   9330
      TabIndex        =   62
      Top             =   4590
      Width           =   8745
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
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton optYea 
         Caption         =   "Absent"
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
         Index           =   0
         Left            =   960
         TabIndex        =   65
         Top             =   750
         Width           =   2445
      End
      Begin VB.OptionButton optYea 
         Caption         =   "Mandays"
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
         Index           =   1
         Left            =   960
         TabIndex        =   66
         Top             =   1140
         Width           =   2445
      End
      Begin VB.OptionButton optYea 
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
         Height          =   405
         Index           =   2
         Left            =   960
         TabIndex        =   67
         Top             =   1560
         Width           =   2445
      End
      Begin VB.OptionButton optYea 
         Caption         =   "Present"
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
         Index           =   3
         Left            =   4170
         TabIndex        =   68
         Top             =   750
         Width           =   2445
      End
      Begin VB.OptionButton optYea 
         Caption         =   "Leave Information"
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
         Index           =   4
         Left            =   4170
         TabIndex        =   69
         Top             =   1170
         Width           =   2445
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
         Left            =   1140
         TabIndex        =   63
         Top             =   330
         Width           =   1980
      End
   End
   Begin VB.Frame frWeek 
      Caption         =   "Weekly &Reports"
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
      Left            =   1800
      TabIndex        =   27
      Top             =   480
      Width           =   8745
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
         Left            =   4380
         MaxLength       =   10
         TabIndex        =   29
         Tag             =   "D"
         Top             =   270
         Width           =   1125
      End
      Begin VB.OptionButton optWee 
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
         Height          =   375
         Index           =   0
         Left            =   270
         TabIndex        =   30
         Top             =   750
         Width           =   2355
      End
      Begin VB.OptionButton optWee 
         Caption         =   "Absent"
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
         Index           =   1
         Left            =   270
         TabIndex        =   31
         Top             =   1170
         Width           =   2355
      End
      Begin VB.OptionButton optWee 
         Caption         =   "Attendace"
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
         Index           =   2
         Left            =   270
         TabIndex        =   32
         Top             =   1590
         Width           =   2355
      End
      Begin VB.OptionButton optWee 
         Caption         =   "Late Arrival"
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
         Index           =   3
         Left            =   2970
         TabIndex        =   33
         Top             =   750
         Width           =   2355
      End
      Begin VB.OptionButton optWee 
         Caption         =   "Early Departure"
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
         Index           =   4
         Left            =   2970
         TabIndex        =   34
         Top             =   1170
         Width           =   2355
      End
      Begin VB.OptionButton optWee 
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
         Height          =   375
         Index           =   5
         Left            =   2970
         TabIndex        =   35
         Top             =   1560
         Width           =   2355
      End
      Begin VB.OptionButton optWee 
         Caption         =   "Shift Schedule"
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
         Index           =   6
         Left            =   5850
         TabIndex        =   36
         Top             =   750
         Width           =   2355
      End
      Begin VB.OptionButton optWee 
         Caption         =   "Irregular"
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
         Index           =   7
         Left            =   5850
         TabIndex        =   37
         Top             =   1170
         Width           =   2355
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
         Left            =   1110
         TabIndex        =   28
         Top             =   330
         Width           =   3480
      End
   End
   Begin VB.Frame frDly 
      Caption         =   "Daily &Reports"
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
      Left            =   9000
      TabIndex        =   6
      Top             =   2190
      Width           =   8745
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
         Left            =   3630
         TabIndex        =   21
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtDlyCAbs 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Left            =   7890
         MaxLength       =   2
         TabIndex        =   14
         Top             =   1650
         Width           =   435
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
         Left            =   3450
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "D"
         Top             =   210
         Width           =   1125
      End
      Begin VB.OptionButton optDly 
         Caption         =   "Summary"
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
         Index           =   12
         Left            =   6420
         TabIndex        =   26
         Top             =   1200
         Width           =   1635
      End
      Begin VB.OptionButton optDly 
         Caption         =   "Outdoor Duty"
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
         Index           =   11
         Left            =   6420
         TabIndex        =   25
         Top             =   900
         Width           =   2115
      End
      Begin VB.OptionButton optDly 
         Caption         =   "Manpower"
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
         Index           =   10
         Left            =   6420
         TabIndex        =   24
         Top             =   600
         Width           =   2115
      End
      Begin VB.OptionButton optDly 
         Caption         =   "Shift Arrangement"
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
         Index           =   9
         Left            =   3630
         TabIndex        =   23
         Top             =   1800
         Width           =   2115
      End
      Begin VB.OptionButton optDly 
         Caption         =   "Entries"
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
         Index           =   8
         Left            =   3630
         TabIndex        =   22
         Top             =   1500
         Width           =   2115
      End
      Begin VB.OptionButton optDly 
         Caption         =   "Authorized OT"
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
         Index           =   7
         Left            =   3630
         TabIndex        =   20
         Top             =   900
         Width           =   2115
      End
      Begin VB.OptionButton optDly 
         Caption         =   "Irregular"
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
         Index           =   6
         Left            =   3630
         TabIndex        =   19
         Top             =   600
         Width           =   2115
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
         TabIndex        =   18
         Top             =   1800
         Width           =   2115
      End
      Begin VB.OptionButton optDly 
         Caption         =   "Early Departure"
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
         Index           =   4
         Left            =   480
         TabIndex        =   17
         Top             =   1500
         Width           =   2115
      End
      Begin VB.OptionButton optDly 
         Caption         =   "Late Arrival"
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
         Left            =   480
         TabIndex        =   16
         Top             =   1200
         Width           =   2115
      End
      Begin VB.OptionButton optDly 
         Caption         =   "Continuous Absent"
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
         Left            =   6420
         TabIndex        =   15
         Top             =   1800
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.OptionButton optDly 
         Caption         =   "Absent"
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
         Left            =   480
         TabIndex        =   12
         Top             =   900
         Width           =   2115
      End
      Begin VB.OptionButton optDly 
         Caption         =   "Physical Arrival"
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
         Index           =   0
         Left            =   480
         TabIndex        =   11
         Top             =   510
         Width           =   1755
      End
      Begin MSForms.ComboBox cboShift 
         Height          =   345
         Left            =   5700
         TabIndex        =   10
         Top             =   210
         Width           =   1515
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "2672;609"
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
         Left            =   4740
         TabIndex        =   9
         Top             =   270
         Width           =   375
      End
      Begin VB.Label lblDlyCAbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For how many days ?"
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
         Left            =   5820
         TabIndex        =   13
         Top             =   1710
         Width           =   1890
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
         Left            =   1110
         TabIndex        =   7
         Top             =   270
         Width           =   1980
      End
   End
   Begin VB.Frame frMonChk 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   125
      Top             =   5100
      Width           =   8745
      Begin VB.CheckBox chkPromp 
         Caption         =   "Prompt before Printing"
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
         Left            =   6180
         TabIndex        =   128
         Top             =   150
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox chkDotMa 
         Caption         =   "Use 132 Column Dot Matrix Printer"
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
         Left            =   2490
         TabIndex        =   127
         Top             =   150
         Width           =   3345
      End
      Begin VB.CheckBox chkDateT 
         Caption         =   "Print Date and Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   126
         Top             =   150
         Width           =   2115
      End
   End
   Begin VB.Frame frCmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   129
      Top             =   5520
      Width           =   8745
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         TabIndex        =   135
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "&File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6180
         TabIndex        =   134
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdSelPri 
         Caption         =   "Selec&t Printer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   130
         Top             =   180
         Width           =   1905
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   133
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Pre&view"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3660
         TabIndex        =   132
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   131
         Top             =   180
         Width           =   1275
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
      Height          =   2420
      Left            =   0
      TabIndex        =   95
      Top             =   2700
      Width           =   8745
      Begin VB.OptionButton optGrpDiv 
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
         Left            =   6990
         TabIndex        =   144
         Top             =   1100
         Width           =   1275
      End
      Begin VB.OptionButton optGrpLoc 
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
         Left            =   5670
         TabIndex        =   123
         Top             =   1100
         Width           =   1275
      End
      Begin VB.CheckBox chkNewP 
         Caption         =   "Start New Page When Group Changes"
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
         Left            =   5670
         TabIndex        =   124
         Top             =   1725
         Width           =   2925
      End
      Begin VB.OptionButton optGrpDC 
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
         Left            =   5670
         TabIndex        =   122
         Top             =   1380
         Width           =   2385
      End
      Begin VB.OptionButton optGrpGrp 
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
         Left            =   6990
         TabIndex        =   121
         Top             =   760
         Width           =   915
      End
      Begin VB.OptionButton optGrpCat 
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
         Left            =   5670
         TabIndex        =   120
         Top             =   760
         Width           =   1185
      End
      Begin VB.OptionButton optGrpDep 
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
         Left            =   6990
         TabIndex        =   119
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optGrpEmp 
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
         Left            =   5670
         TabIndex        =   118
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "S. Group"
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
         TabIndex        =   150
         Top             =   1290
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSForms.ComboBox cboMain 
         Height          =   315
         Left            =   4410
         TabIndex        =   149
         Top             =   1560
         Visible         =   0   'False
         Width           =   945
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1667;556"
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
      Begin MSForms.ComboBox cmbFrDivSel 
         Height          =   315
         Left            =   3270
         TabIndex        =   143
         Top             =   1560
         Width           =   945
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1667;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbToDivSel 
         Height          =   315
         Left            =   3270
         TabIndex        =   142
         Top             =   1920
         Width           =   945
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1667;556"
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
      Begin VB.Label lblDivSel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   3300
         TabIndex        =   141
         Top             =   1290
         Width           =   660
      End
      Begin VB.Label lblLocSel 
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
         Left            =   2100
         TabIndex        =   108
         Top             =   1260
         Width           =   735
      End
      Begin MSForms.ComboBox cmbFrLocSel 
         Height          =   315
         Left            =   2070
         TabIndex        =   114
         Top             =   1560
         Width           =   945
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1667;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbToLocSel 
         Height          =   315
         Left            =   2070
         TabIndex        =   115
         Top             =   1920
         Width           =   945
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1667;556"
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
         X1              =   5460
         X2              =   5460
         Y1              =   150
         Y2              =   2350
      End
      Begin MSForms.ComboBox cmbFrComSel 
         Height          =   315
         Left            =   4410
         TabIndex        =   116
         Top             =   570
         Width           =   945
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1667;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbToGrpSel 
         Height          =   315
         Left            =   660
         TabIndex        =   113
         Top             =   1890
         Width           =   1245
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2196;556"
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
      Begin MSForms.ComboBox cmbFrGrpSel 
         Height          =   315
         Left            =   660
         TabIndex        =   111
         Top             =   1530
         Width           =   1245
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2196;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbToCatSel 
         Height          =   315
         Left            =   3270
         TabIndex        =   106
         Top             =   930
         Width           =   945
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1667;556"
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
      Begin MSForms.ComboBox cmbFrCatSel 
         Height          =   315
         Left            =   3270
         TabIndex        =   105
         Top             =   570
         Width           =   945
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1667;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbToDepSel 
         Height          =   315
         Left            =   2070
         TabIndex        =   104
         Top             =   930
         Width           =   945
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1667;556"
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
      Begin MSForms.ComboBox cmbFrDepSel 
         Height          =   315
         Left            =   2070
         TabIndex        =   103
         Top             =   570
         Width           =   945
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1667;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbToEmpSel 
         Height          =   315
         Left            =   660
         TabIndex        =   102
         Top             =   930
         Width           =   1245
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2196;556"
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
      Begin MSForms.ComboBox cmbFrEmpSel 
         Height          =   315
         Left            =   660
         TabIndex        =   100
         Top             =   570
         Width           =   1245
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2196;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   178
         FontPitchAndFamily=   2
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
         Left            =   5670
         TabIndex        =   117
         Top             =   180
         Width           =   885
      End
      Begin VB.Label lblFr2Sel 
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
         Left            =   90
         TabIndex        =   110
         Top             =   1590
         Width           =   465
      End
      Begin VB.Label lblTo2Sel 
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
         Left            =   90
         TabIndex        =   112
         Top             =   1920
         Width           =   225
      End
      Begin VB.Label lblComSel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Business Unit"
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
         Left            =   4200
         TabIndex        =   109
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblGrpSel 
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
         Left            =   660
         TabIndex        =   107
         Top             =   1260
         Width           =   525
      End
      Begin VB.Label lblCatSel 
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
         Left            =   3270
         TabIndex        =   98
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lblDepSel 
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
         Left            =   2100
         TabIndex        =   97
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblEmpSel 
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
         Left            =   660
         TabIndex        =   96
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lblTo1Sel 
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
         Left            =   90
         TabIndex        =   101
         Top             =   960
         Width           =   225
      End
      Begin VB.Label lblFr1Sel 
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
         Left            =   90
         TabIndex        =   99
         Top             =   570
         Width           =   465
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   5490
         X2              =   5490
         Y1              =   150
         Y2              =   2350
      End
   End
   Begin VB.OptionButton cmdPeriodic 
      Caption         =   "P&eriodic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1425
   End
   Begin VB.OptionButton cmdMaster 
      Caption         =   "M&asters"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1605
   End
   Begin VB.OptionButton cmdYearly 
      Caption         =   "&Yearly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin VB.OptionButton cmdMonthly 
      Caption         =   "&Monthly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.OptionButton cmdWeekly 
      Caption         =   "&Weekly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.OptionButton cmdDaily 
      Caption         =   "&Daily"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   30
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1425
   End
   Begin MSMAPI.MAPISession ReportSession 
      Left            =   8940
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages ReportMessage 
      Left            =   8940
      Top             =   930
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
      Caption         =   "Monthly &Reports"
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
      TabIndex        =   38
      Top             =   480
      Width           =   8745
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
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Leave Consumption"
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
         Index           =   14
         Left            =   4440
         TabIndex        =   57
         Top             =   1860
         Width           =   2355
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Leave Balance"
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
         Index           =   9
         Left            =   2100
         TabIndex        =   52
         Top             =   1860
         Width           =   2355
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
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   180
         Width           =   1425
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Monthly Absent"
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
         Index           =   4
         Left            =   90
         TabIndex        =   47
         Top             =   1860
         Width           =   1965
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Shift schedule"
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
         Index           =   17
         Left            =   6780
         TabIndex        =   60
         Top             =   1230
         Width           =   1815
      End
      Begin VB.OptionButton optMon 
         Caption         =   "WO on Holiday"
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
         Index           =   18
         Left            =   6780
         TabIndex        =   61
         Top             =   1530
         Width           =   1815
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Early Departure Memo"
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
         Index           =   13
         Left            =   4440
         TabIndex        =   56
         Top             =   1530
         Width           =   2355
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Attendance"
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
         Index           =   1
         Left            =   90
         TabIndex        =   44
         Top             =   900
         Width           =   1965
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Overtime Paid"
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
         Index           =   6
         Left            =   2100
         TabIndex        =   49
         Top             =   900
         Width           =   2355
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Muster Report"
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
         Index           =   2
         Left            =   90
         TabIndex        =   45
         Top             =   1230
         Width           =   1965
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Absent Memo"
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
         Index           =   7
         Left            =   2100
         TabIndex        =   50
         Top             =   1230
         Width           =   2355
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Late Arrival Memo"
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
         Index           =   12
         Left            =   4440
         TabIndex        =   55
         Top             =   1230
         Width           =   2355
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Monthly Present"
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
         Index           =   3
         Left            =   90
         TabIndex        =   46
         Top             =   1530
         Width           =   1965
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Absent /Late /Early"
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
         Index           =   8
         Left            =   2100
         TabIndex        =   51
         Top             =   1530
         Width           =   2355
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Total Lates"
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
         Index           =   15
         Left            =   6780
         TabIndex        =   58
         Top             =   570
         Width           =   1695
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Late Arrival"
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
         Index           =   10
         Left            =   4440
         TabIndex        =   53
         Top             =   570
         Width           =   2355
      End
      Begin VB.OptionButton optMon 
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
         Height          =   285
         Index           =   5
         Left            =   2100
         TabIndex        =   48
         Top             =   570
         Width           =   2355
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Total Earlys"
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
         Index           =   16
         Left            =   6780
         TabIndex        =   59
         Top             =   900
         Width           =   1755
      End
      Begin VB.OptionButton optMon 
         Caption         =   "Early Departure"
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
         Index           =   11
         Left            =   4440
         TabIndex        =   54
         Top             =   900
         Width           =   2355
      End
      Begin VB.OptionButton optMon 
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
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   43
         Top             =   570
         Width           =   1965
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
         Left            =   5940
         TabIndex        =   41
         Top             =   240
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
         Left            =   1770
         TabIndex        =   39
         Top             =   240
         Width           =   2205
      End
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strLastDateM As String, strFirstDateM As String, strAttachPath As String
Dim rpGroup As String
''
Dim adrsC As New ADODB.Recordset

Private Sub chkCat_Click()
Call FrSelFill
End Sub

Private Sub chkDateT_Click(Index As Integer)    '' Decides, to show or not to show date & time
SaveSetting "Vstar", "PrjSettings", "Show Date and time", chkDateT(Index).Value
End Sub

Private Sub chkDotMa_Click(Index As Integer)    '' Decides, to use dotmatrix printer or not
If chkDotMa(1).Value = 1 Then
    If MsgBox(NewCaptionTxt("40064", adrsC) & vbCrLf & _
    NewCaptionTxt("00009", adrsMod), vbYesNo + vbQuestion) = vbNo Then
        chkDotMa(1).Value = 0
    End If
End If
End Sub

Private Sub chkNewP_Click()     '' Decides , to perform  page break if group changes
SaveSetting "Vstar", "PrjSettings", "Print on Next Page", chkNewP.Value
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
         MsgBox NewCaptionTxt("40065", adrsC) & cmbMonth.Text & Space(1) & _
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

Private Sub cmdSelPri_Click()   '' Sets Default Printer.
On Error GoTo Err_P
CommonDialog1.PrinterDefault = True
CommonDialog1.Flags = cdlSetNotSupported
CommonDialog1.ShowPrinter
Printer.TrackDefault = True
Exit Sub
Err_P:
    ShowError ("Select Printer:: " & Me.Caption)
End Sub

Private Sub cmdSend_Click()     '' SEND's report through email available in empmst.
On Error GoTo Err_P
If adRsInstall.State = 1 Then adRsInstall.Close
adRsInstall.Open "Select Email From Install", VstarDataEnv.cnDJConn, _
adOpenStatic, adLockOptimistic
If adRsInstall("Email") = True Then
    bytAction = 1
    Call ReportsMod
    Call SetVarEmpty
Else
    MsgBox NewCaptionTxt("40066", adrsC), vbExclamation
End If
Exit Sub
Err_P:
    ShowError ("Send Reports :: " & Me.Caption)
End Sub

Private Sub cmdPreview_Click()  '' Gives Preview of the selected report
bytAction = 2
Call ReportsMod
Call SetVarEmpty
End Sub

Private Sub cmdPrint_Click()    '' Prints the selecdted report
If chkPromp.Value = 1 Then
    If MsgBox(NewCaptionTxt("40067", adrsC), vbYesNo + vbQuestion) = vbNo Then Exit Sub
End If
bytAction = 3
Call ReportsMod
Call SetVarEmpty
End Sub

Private Sub cmdFile_Click()     '' Gives output of a report to a TXT file
bytAction = 4
Call ReportsMod
Call SetVarEmpty
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
'' Enable Date TextBoxes
txtDaily.Enabled = True
txtWeek.Enabled = True
txtFrPeri.Enabled = True
txtToPeri.Enabled = True
End Sub

Private Sub Form_Load()
On Error GoTo Err_P
CatFlag = True
Call SetFormIcon(Me)            '' Set the Forms Icon
bytRepMode = 1                  '' Load Mode
'' Disable Date TextBoxes
txtDaily.Enabled = False
txtWeek.Enabled = False
txtFrPeri.Enabled = False
txtToPeri.Enabled = False
bytAction = 0 ''No Action defined
MSF1.ColWidth(0) = MSF1.Width - 10
Call SetRepVars                 '' SETS DEFAULT VALUES OF TEXTBOXES AND RADIO BUTTONS
Call RetCaptions                '' RETRIEVES CAPTIONS FOR ALL LABELS
Call LoadSpecifics              '' procedure to Perform Other Actions on Load
If adrsDSR.State = 1 Then adrsDSR.Close
adrsDSR.Open "Select * from NewCaptions Where ID Like 'D%' or ID Like '00%'", VstarDataEnv.cnDJConn, adOpenStatic

'rpGroup = "groupmst.grupdesc,catdesc.cat,deptdesc.dept" & _
'" ,CatDesc." & strKDesc & ",deptdesc." & strKDesc & _
'",Location.Location,Location.LocDesc,Division.Div,Division.DivDesc,Company.Company,Company.CName"

rpGroup = "groupmst.grupdesc as groupmst,catdesc.cat as catdesccat,deptdesc.dept as " & _
" deptdescdept,CatDesc." & strKDesc & " as catdescdesc,deptdesc." & strKDesc & _
" as deptdescdesc,Location.Location,Location.LocDesc,Division.Div,Division.DivDesc,Company.Company,Company.CName"
Exit Sub
Err_P:
    ShowError ("Load :: " & Me.Caption)
End Sub

Private Sub RetCaptions()
On Error GoTo Err_P
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '40%'", VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
'form
Me.Caption = NewCaptionTxt("40001", adrsC)
'daily pane
cmdDaily.Caption = NewCaptionTxt("40002", adrsC)
frDly.Caption = NewCaptionTxt("40003", adrsC)
lblDaily.Caption = NewCaptionTxt("40004", adrsC)
lblShf.Caption = NewCaptionTxt("40005", adrsC)
optDly(0).Caption = NewCaptionTxt("40006", adrsC)
optDly(1).Caption = NewCaptionTxt("40007", adrsC)
optDly(2).Caption = NewCaptionTxt("40008", adrsC)
optDly(3).Caption = NewCaptionTxt("40009", adrsC)
optDly(4).Caption = NewCaptionTxt("40010", adrsC)
optDly(5).Caption = NewCaptionTxt("40011", adrsC)
optDly(6).Caption = NewCaptionTxt("40012", adrsC)
optDly(7).Caption = NewCaptionTxt("00119", adrsMod)
optDly(8).Caption = NewCaptionTxt("40013", adrsC)
optDly(9).Caption = NewCaptionTxt("40014", adrsC)
optDly(10).Caption = NewCaptionTxt("40015", adrsC)
optDly(11).Caption = NewCaptionTxt("40016", adrsC)
optDly(12).Caption = NewCaptionTxt("40017", adrsC)
optDly(13).Caption = NewCaptionTxt("00116", adrsMod)
'weekly pane
cmdWeekly.Caption = NewCaptionTxt("40018", adrsC)
frWeek.Caption = NewCaptionTxt("40019", adrsC)
lblWeek.Caption = NewCaptionTxt("40020", adrsC)
optWee(0).Caption = NewCaptionTxt("40011", adrsC)
optWee(1).Caption = NewCaptionTxt("40007", adrsC)
optWee(2).Caption = NewCaptionTxt("40021", adrsC)
optWee(3).Caption = NewCaptionTxt("40009", adrsC)
optWee(4).Caption = NewCaptionTxt("40010", adrsC)
optWee(5).Caption = NewCaptionTxt("00038", adrsMod)
optWee(6).Caption = NewCaptionTxt("40022", adrsC)
optWee(7).Caption = NewCaptionTxt("40012", adrsC)

'monthly pane
cmdMonthly.Caption = NewCaptionTxt("40023", adrsC)
frMonth.Caption = NewCaptionTxt("40024", adrsC)
lblMonth.Caption = NewCaptionTxt("40025", adrsC)
lblMonYear.Caption = NewCaptionTxt("00027", adrsMod)
optMon(0).Caption = NewCaptionTxt("40011", adrsC)
optMon(1).Caption = NewCaptionTxt("40021", adrsC)
optMon(2).Caption = NewCaptionTxt("40026", adrsC)
optMon(3).Caption = NewCaptionTxt("40027", adrsC)
optMon(4).Caption = NewCaptionTxt("40028", adrsC)
optMon(5).Caption = NewCaptionTxt("00038", adrsMod)
optMon(6).Caption = NewCaptionTxt("40029", adrsC)
optMon(7).Caption = NewCaptionTxt("40030", adrsC)
optMon(8).Caption = NewCaptionTxt("40031", adrsC)
optMon(9).Caption = NewCaptionTxt("40032", adrsC)
optMon(10).Caption = NewCaptionTxt("40009", adrsC)
optMon(11).Caption = NewCaptionTxt("40010", adrsC)
optMon(12).Caption = NewCaptionTxt("40033", adrsC)
optMon(13).Caption = NewCaptionTxt("40034", adrsC)
optMon(14).Caption = NewCaptionTxt("40035", adrsC)
optMon(15).Caption = NewCaptionTxt("40036", adrsC)
optMon(16).Caption = NewCaptionTxt("40037", adrsC)
optMon(17).Caption = NewCaptionTxt("40022", adrsC)
optMon(18).Caption = NewCaptionTxt("40038", adrsC)

'yearly pane
cmdYearly.Caption = NewCaptionTxt("40039", adrsC)
frYear.Caption = NewCaptionTxt("40040", adrsC)
lblYear.Caption = NewCaptionTxt("40041", adrsC)
optYea(0).Caption = NewCaptionTxt("40007", adrsC)
optYea(1).Caption = NewCaptionTxt("40042", adrsC)
optYea(2).Caption = NewCaptionTxt("40011", adrsC)
optYea(3).Caption = NewCaptionTxt("40043", adrsC)
optYea(4).Caption = NewCaptionTxt("40044", adrsC)

'master pane
cmdMaster.Caption = NewCaptionTxt("40045", adrsC)
frMast.Caption = NewCaptionTxt("40046", adrsC)
lblMastFr.Caption = NewCaptionTxt("00019", adrsMod)
lblMastTo.Caption = NewCaptionTxt("00011", adrsMod)
optmas(0).Caption = NewCaptionTxt("40047", adrsC)
optmas(1).Caption = NewCaptionTxt("40048", adrsC)
optmas(2).Caption = NewCaptionTxt("40049", adrsC)
optmas(3).Caption = NewCaptionTxt("00063", adrsMod)
optmas(4).Caption = NewCaptionTxt("00031", adrsMod)
optmas(5).Caption = NewCaptionTxt("40050", adrsC)
optmas(6).Caption = NewCaptionTxt("40051", adrsC)
optmas(7).Caption = NewCaptionTxt("00058", adrsMod)
optmas(8).Caption = NewCaptionTxt("00051", adrsMod)
optmas(9).Caption = NewCaptionTxt("00059", adrsMod)
optmas(10).Caption = NewCaptionTxt("00110", adrsMod)
optmas(11).Caption = NewCaptionTxt("00057", adrsMod)
optmas(12).Caption = NewCaptionTxt("00126", adrsMod)
'periodic pane
cmdPeriodic.Caption = NewCaptionTxt("40052", adrsC)
frPeri.Caption = NewCaptionTxt("40053", adrsC)
lblFrPeri.Caption = NewCaptionTxt("40054", adrsC)
lblToPeri.Caption = NewCaptionTxt("00011", adrsMod)
lblNotePeri.Caption = NewCaptionTxt("40055", adrsC)
optPer(0).Caption = NewCaptionTxt("40011", adrsC)
optPer(1).Caption = NewCaptionTxt("40026", adrsC)
optPer(2).Caption = NewCaptionTxt("00038", adrsMod)
optPer(3).Caption = NewCaptionTxt("40009", adrsC)
optPer(4).Caption = NewCaptionTxt("40010", adrsC)
optPer(5).Caption = NewCaptionTxt("40008", adrsC)
optPer(6).Caption = NewCaptionTxt("40017", adrsC)
optPer(9).Caption = NewCaptionTxt("40086", adrsC)
optPer(10).Caption = "Unprocessed Leaves" ''NewCaptionTxt("40035", adrsC)

'selection frame
frSel.Caption = NewCaptionTxt("40056", adrsC)
lblEmpSel.Caption = NewCaptionTxt("40057", adrsC)
lblDepSel.Caption = NewCaptionTxt("00058", adrsMod)
lblCatSel.Caption = NewCaptionTxt("00051", adrsMod)
lblGrpSel.Caption = NewCaptionTxt("00059", adrsMod)
lblLocSel.Caption = NewCaptionTxt("00110", adrsMod)
lblComSel.Caption = NewCaptionTxt("00057", adrsMod)
lblFr1Sel.Caption = NewCaptionTxt("00010", adrsMod)
lblTo1Sel.Caption = NewCaptionTxt("00011", adrsMod)
lblFr2Sel.Caption = lblFr1Sel.Caption
lblTo2Sel.Caption = lblTo1Sel.Caption
lblDivSel.Caption = NewCaptionTxt("00126", adrsMod)

'group by
lblGroupBy.Caption = NewCaptionTxt("40058", adrsC)
optGrpEmp(0).Caption = NewCaptionTxt("40057", adrsC)
optGrpDep(1).Caption = NewCaptionTxt("00058", adrsMod)
optGrpCat(2).Caption = NewCaptionTxt("00051", adrsMod)
optGrpGrp(3).Caption = NewCaptionTxt("00059", adrsMod)
optGrpDC(4).Caption = NewCaptionTxt("00057", adrsMod)
optGrpLoc(5).Caption = NewCaptionTxt("00110", adrsMod)
optGrpDiv(6).Caption = NewCaptionTxt("00126", adrsMod)
chkNewP.Caption = NewCaptionTxt("40060", adrsC)

'check boxes
chkDateT(0).Caption = NewCaptionTxt("40061", adrsC)
chkDotMa(1).Caption = NewCaptionTxt("40062", adrsC)
chkPromp.Caption = NewCaptionTxt("40063", adrsC)

'command buttons
cmdSelPri.Caption = NewCaptionTxt("00074", adrsMod)
cmdSend.Caption = NewCaptionTxt("00075", adrsMod)
cmdPreview.Caption = NewCaptionTxt("00076", adrsMod)
cmdPrint.Caption = NewCaptionTxt("00077", adrsMod)
cmdFile.Caption = NewCaptionTxt("00078", adrsMod)
cmdExit.Caption = NewCaptionTxt("00008", adrsMod)

Exit Sub
Err_P:
    ShowError ("RetCaptions :: " & Me.Caption)
    ''Resume Next
End Sub

Private Function FrVisible(Optional ByVal bytFrVal As Byte = 1)
frDly.Visible = False                       '' DEPENDING UPON THE CURRENT TAB SELECTED
frWeek.Visible = False                      '' VISIBLES APPROPRIATE FRAME
frMonth.Visible = False
frYear.Visible = False
frMast.Visible = False
frPeri.Visible = False

frSel.Enabled = True

lblDlyCAbs.Visible = False
txtDlyCAbs.Visible = False

Select Case bytFrVal
    Case 1: frDly.Visible = True            '' Daily
            frDly.Top = frMonth.Top
            frDly.Left = frMonth.Left
    Case 2: frWeek.Visible = True           '' Weekly
            frWeek.Top = frMonth.Top
            frWeek.Left = frMonth.Left
    Case 3: frMonth.Visible = True          '' Monthly
    Case 4: frYear.Visible = True           '' Yearly
            frYear.Top = frMonth.Top
            frYear.Left = frMonth.Left
    Case 5: frMast.Visible = True           '' Masters
            frMast.Top = frMonth.Top
            frMast.Left = frMonth.Left
    Case 6: frPeri.Visible = True           '' Periodic
            frPeri.Top = frMonth.Top
            frPeri.Left = frMonth.Left
End Select

End Function

Private Function MonthYearFill()
On Error GoTo Err_P
Dim i As Byte                       '' FILLS REQUIRED MONTH AND YEAR COMBOS
For i = 1 To 12
    cmbMonth.AddItem Choose(i, "January", "February", "March", "April", "May", _
    "June", "July", "August", "september", "October", "November", "December")
Next i

For i = 0 To 99
    cmbYear.AddItem (1997 + i)
    cmbMonYear.AddItem (1997 + i)
Next i

cmbYear.Text = Year(Date)
cmbMonth.Text = MonthName(Month(Date))
cmbMonYear.Text = Year(Date)
Exit Function
Err_P:
    ShowError ("MonthYearFill :: " & Me.Caption)
End Function

Private Sub LoadSpecifics()
On Error GoTo Err_P
Call SetToolTipText(Me)         '' Sets the ToolTipText for Date Text Boxes
Call MonthYearFill              '' FILLS MONTH AND YEAR COMBOS
Call FrSelFill                  '' FILLS SELECTION FRAME COMBOS
Call FrchkFill                  '' FILLS CHECKBOX FRAME
Call PutZeros                   '' SELECTS FIRST OPTION BUTTONS FOR ALL TABS
Call GetRights                  '' Check for Rights
Exit Sub
Err_P:
    ShowError ("Load Specifics :: " & Me.Caption)
End Sub

Private Sub GetRights()
On Error GoTo Err_P
Dim strTmp As String
strTmp = RetRights(4, 16, 6, 1)
If strTmp = "1" Then
    cmdPreview.Enabled = True
    cmdPrint.Enabled = True
    cmdSend.Enabled = True
    cmdFile.Enabled = True
Else
    cmdPreview.Enabled = False
    cmdPrint.Enabled = False
    cmdSend.Enabled = False
    cmdFile.Enabled = False
End If
Exit Sub
Err_P:
    ShowError ("GetRights::" & Me.Caption)
    cmdPreview.Enabled = False
    cmdPrint.Enabled = False
    cmdSend.Enabled = False
    cmdFile.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
CatFlag = False
End Sub

Private Sub optDly_Click(Index As Integer)
lblShf.Visible = True
cboShift.Visible = True
If Index = 12 Then  ''If summary report is selected disable grouping options
    Call ShowGroup(False)
Else
    Call ShowGroup
End If
If Index = 8 Then                       'FOR ENTRIES REPORT
    lblShf.Visible = False
    cboShift.Visible = False
    If Trim(txtDaily.Text) = "" Then
        MsgBox NewCaptionTxt("40068", adrsC), vbExclamation
        optDly(0).Value = True
        txtDaily.SetFocus
        Exit Sub
    Else
        Call SetRepVars(1)
    End If
    If MsgBox(NewCaptionTxt("40069", adrsC), vbOKCancel) = vbCancel Then
        optDly(0).Value = True
        Exit Sub
    Else
       If Not dlySeachEntries Then
            optDly(0).Value = True
            Exit Sub
       Else                                             ''if cancel button is clicked on
            If optDly(0).Value = True Then Index = 0    ''frmentries make the first opt bttn
       End If                                           ''enable instead of entries bttn
    End If
End If
typOptIdx.bytDly = Index
End Sub

Private Sub optWee_Click(Index As Integer)
typOptIdx.bytWek = Index
End Sub

Private Sub optMon_Click(Index As Integer)
typOptIdx.bytMon = Index
Select Case typOptIdx.bytMon
    Case 1, 6, 15, 16       'FOR ATTENDANCE, OT PAID, TOTAL LATES ,TOTAL EARLYS REPORTS
        MsgBox NewCaptionTxt("40070", adrsC) & vbCrLf & vbTab & _
        NewCaptionTxt("40071", adrsC), vbInformation
    Case 7, 12, 13          'FOR ABSENT,LATE,EARLY MEMO REPORTS
        bytMode = Index
        frmMemo.Show vbModal
End Select
End Sub

Private Sub optMon_DblClick(Index As Integer)
    Call optMon_Click(Index)
End Sub

Private Sub optYea_Click(Index As Integer)
    typOptIdx.bytYer = Index
End Sub

Private Sub optMas_Click(Index As Integer)
typOptIdx.bytMst = Index
Select Case Index
    Case 0, 1                                 'FOR EMPLOYEE REPORT
        lblMastFr.Visible = False
        lblMastTo.Visible = False
        txtMastFr.Visible = False
        txtMastTo.Visible = False
        Call ShowGroupEx
    Case 2                                  'FOR LEFT EMPLOYEE REPORT
        lblMastFr.Visible = True
        lblMastTo.Visible = True
        txtMastFr.Visible = True
        txtMastTo.Visible = True
        Call ShowGroupEx
        txtMastFr.SetFocus
    Case Else                               'FOR OTHER REPORTS
        lblMastFr.Visible = False
        lblMastTo.Visible = False
        txtMastFr.Visible = False
        txtMastTo.Visible = False
        Call ShowGroupEx(False)
End Select
End Sub

Private Sub optPer_Click(Index As Integer)
typOptIdx.bytPer = Index
Select Case Index
    Case 6 ''If summary report is selected disable grouping options
        Call ShowGroup(False)
    Case 7 ''Meal Allowance
        Call ShowGroup
    Case Else
        Call ShowGroup
End Select
End Sub

Private Sub optGrpCat_Click(Index As Integer)
chkNewP.Visible = True
End Sub

Private Sub optGrpDC_Click(Index As Integer)
chkNewP.Visible = True
End Sub

Private Sub optGrpDep_Click(Index As Integer)
chkNewP.Visible = True
End Sub

Private Sub optGrpEmp_Click(Index As Integer)
chkNewP.Visible = False
End Sub

Private Sub optGrpGrp_Click(Index As Integer)
chkNewP.Visible = True
End Sub

Private Sub optGrpLoc_Click(Index As Integer)
chkNewP.Visible = True
End Sub

Private Sub optGrpDiv_Click(Index As Integer)
chkNewP.Visible = True
End Sub

Private Sub cmdDaily_Click()    'DAILY TAB
bytRepMode = 1
Call SetMSF1Cap(1)
Call FrVisible(1)
Call ShowGroupEx
optDly(typOptIdx.bytDly).Value = False      '' These two lines are written just to get the
optDly(typOptIdx.bytDly).Value = True       '' call the option button's click.never remove this
txtDaily.SetFocus
End Sub

Private Sub cmdWeekly_Click()   'WEEKLY TAB
bytRepMode = 2
Call ShowGroup
Call SetMSF1Cap(2)
Call FrVisible(2)
Call ShowGroupEx
optWee(typOptIdx.bytWek).Value = True
txtWeek.SetFocus
End Sub

Private Sub cmdMonthly_Click()  'MONTHLY TAB
bytRepMode = 3
Call ShowGroup
Call SetMSF1Cap(3)
Call FrVisible(3)
Call ShowGroupEx
optMon(typOptIdx.bytMon).Value = True
cmbMonth.SetFocus
End Sub

Private Sub cmdYearly_Click()   'YEARLY TAB
bytRepMode = 4
Call ShowGroup
Call SetMSF1Cap(4)
Call FrVisible(4)
Call ShowGroupEx
optYea(typOptIdx.bytYer).Value = True
cmbYear.SetFocus
End Sub

Private Sub cmdMaster_Click()   'MASTER TAB
bytRepMode = 5
Call ShowGroup
Call SetMSF1Cap(5)
Call FrVisible(5)
optmas(typOptIdx.bytMst).Value = True
If typOptIdx.bytMst > 3 Then Call ShowGroupEx(False)
End Sub

Private Sub cmdPeriodic_Click() 'PERIODIC TAB
bytRepMode = 6
Call ShowGroup
Call SetMSF1Cap(6)
Call FrVisible(6)
Call ShowGroupEx
optPer(typOptIdx.bytPer).Value = True
txtFrPeri.SetFocus
End Sub

Private Sub ReportsMod()
On Error GoTo RepErr
Call RetValues 'Retrieves values selected or entered by user in selection frame
Select Case bytRepMode
    Case 1                                      'DAILY TAB
        If Not dlyValid Then Call SetMSF1Cap(1): Exit Sub
        Call SetRepVars(1)
        If Not dlyCreateFiles Then Call SetMSF1Cap(1): Exit Sub 'CREATES TEMPORARY FILE
        If Not dlyReportsMod Then Call SetMSF1Cap(1): Exit Sub  'DUMPS VALUES IN TEMP FILE
        If Not dlySetEmpstr3 Then Call SetMSF1Cap(1): Exit Sub  'SETS QUERY FOR DSR
        If Not dlyTotalCalc Then Call SetMSF1Cap(1): Exit Sub   'CALCULATES VALUES REQ. BY DSR
        If Not SetRepName Then Call SetMSF1Cap(1): Exit Sub     'SETS NAME OF THE REPORT,COMMAND ETC.
        If Not RecordsFound Then                'CHECKS FOR AVAILABILITY OF REQ.ED RECORDS
         '   Call ChkRepFile                     'DELETES TEMPORARY FILE
            Call SetMSF1Cap(1)
            Exit Sub
        End If
        If Not ChkPrinter(repname, bytPoLa) Then Exit Sub 'CHECKS FOR REQUIRED PAPERSIZE
    Case 2                                      'WEEKLY TAB
        If Not WkValid Then Call SetMSF1Cap(2): Exit Sub
        If Not WkCreateFiles Then Call SetMSF1Cap(2): Exit Sub
        If Not wkReportsMod Then Call SetMSF1Cap(2): Exit Sub
        If Not WkSetEmpstr3 Then Call SetMSF1Cap(2): Exit Sub
        If Not SetRepName Then Call SetMSF1Cap(2): Exit Sub
        If Not RecordsFound Then
            Call ChkRepFile                     'DELETES TEMPORARY FILE
            Call SetMSF1Cap(2)
            Exit Sub
        End If
        If Not ChkPrinter(repname, bytPoLa) Then Call SetMSF1Cap(2): Exit Sub
    Case 3                                      'MONTHLY TAB
        If Not monValid Then Call SetMSF1Cap(3): Exit Sub
        If Not monCreateFiles Then Call SetMSF1Cap(3): Exit Sub
        If Not monReportsMod Then Call SetMSF1Cap(3): Exit Sub
        If Not monSetEmpstr3 Then Call SetMSF1Cap(3): Exit Sub
        If Not SetRepName Then Call SetMSF1Cap(3): Exit Sub
        If Not RecordsFound Then
            Call ChkRepFile                     'DELETES TEMPORARY FILE
            Call SetMSF1Cap(3)
            Exit Sub
        End If
        If Not ChkPrinter(repname, bytPoLa) Then Call SetMSF1Cap(3): Exit Sub
    Case 4                                      'YEARLY TAB
        If Not yrValid Then Call SetMSF1Cap(4): Exit Sub
        If Not yrCreateFiles Then Call SetMSF1Cap(4): Exit Sub
        If Not yrReportsMod Then Call SetMSF1Cap(4): Exit Sub
        If Not yrSetEmpstr3 Then Call SetMSF1Cap(4): Exit Sub
        If Not SetRepName Then Call SetMSF1Cap(4): Exit Sub
        If Not RecordsFound Then
             Call ChkRepFile                    'DELETES TEMPORARY FILE
             Call SetMSF1Cap(4)
            Exit Sub
        End If
        If Not ChkPrinter(repname, bytPoLa) Then Call SetMSF1Cap(4): Exit Sub
    Case 5                                      'MASTER TAB
        If Not maValid Then Call SetMSF1Cap(5): Exit Sub
        If Not maReportsMod Then Call SetMSF1Cap(5): Exit Sub
        If Not maSetEmpstr3 Then Call SetMSF1Cap(5): Exit Sub
        If Not SetRepName Then Call SetMSF1Cap(5): Exit Sub
        If Not RecordsFound Then
            Call SetMSF1Cap(5)
            Exit Sub
        End If
       ' If Not ChkPrinter(repname, bytPoLa) Then Call SetMSF1Cap(5): Exit Sub
    Case 6                                      'PERIODIC TAB
        If Not PeValid Then Call SetMSF1Cap(6): Exit Sub
        If Not peCreateFiles Then Call SetMSF1Cap(6): Exit Sub
        If Not peReportsMod Then Call SetMSF1Cap(6): Exit Sub
        If Not peSetEmpstr3 Then Call SetMSF1Cap(6): Exit Sub
        If Not SetRepName Then Call SetMSF1Cap(6): Exit Sub
        If Not RecordsFound Then
            Call ChkRepFile                     'DELETES TEMPORARY FILE
            Call SetMSF1Cap(6)
            Exit Sub
        End If
        If Not ChkPrinter(repname, bytPoLa) Then Call SetMSF1Cap(6): Exit Sub
End Select
Call SetMSF1Cap(11)
'repname.Refresh
'Select Case bytAction ' DEPENDING UPON ACTION SELECTED PEROFRMING THE ACTION
'    Case 1 'SEND THROUGH E-MAIL (WITH ADDRESSES AVAILABLE IN EMPMST)
'        repname.ExportFormat rptKeyHTML, , , "C:\Program Files\Microsoft Visual Studio\VIntDev98\Templates\READMEDN.HTM"
'        repname.ExportReport rptKeyHTML, App.Path & "\" & strRepName
'        strAttachPath = App.Path & "\" & strRepName & ".htm"
'        EmailStaff.Show vbModal
'        Call SendThruMail
'    Case 2 'PREVIEW
'        cmdExit.Cancel = False
'        repname.Show vbModal
'        cmdExit.Cancel = True
'    Case 3 'PRINT
'        repname.PrintReport
'        While repname.AsyncCount > 0
'            DoEvents
'        Wend
'    Case 4 'COPY TO A FILE
'        If TextFileN(strRepName) Then
'            repname.ExportReport rptKeyText, CommonDialog1.FileName
'            While repname.AsyncCount > 0
'                DoEvents
'            Wend
'        End If
'End Select
Call SetMSF1Cap(bytRepMode)
Exit Sub
RepErr:
    Select Case Err.Number
        Case -2147217865
            Call SetMSF1Cap(10)
            MsgBox NewCaptionTxt("40072", adrsC), vbInformation
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
        End Select
        If strRepFile <> "" Then Call ChkRepFile
        Call SetVarEmpty
End Sub

Private Sub RetValues()             '' RETRIEVES SELECTIONS MADE BY USER
On Error GoTo Err_P
Dim strTmp As String
Dim catArray As String
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


If chkCat.Value = 1 Then
strSql = " and empmst.empcode between '" & Trim(cmbFrEmpSel.Text) & "' and '" & _
         Trim(cmbToEmpSel.Text) & "' AND deptdesc.dept between " & Trim(cmbFrDepSel.Text) & _
         " and " & Trim(cmbToDepSel.Text) & " AND catdesc.cat  between '" & _
         Trim(cmbFrCatSel.Text) & "' and '" & Trim(cmbToCatSel.Text) & "' And " & _
         "groupmst." & strKGroup & " between " & Trim(cmbFrGrpSel.Text) & " and " & _
         Trim(cmbToGrpSel.Text) & strTmp & " and Location.Location between " & _
         Trim(cmbFrLocSel.Text) & " and " & Trim(cmbToLocSel.Text) & " and " & _
         "Division.Div between " & Trim(cmbFrDivSel.Text) & " and " & _
         Trim(cmbToDivSel.Text) & " and empmst.dept = deptdesc.dept and " & _
         " empmst.cat = catdesc.cat and empmst." & strKGroup & " = groupmst." & strKGroup & " and " & _
         "empmst.company = company.company and empmst.Location = Location.Location " & _
         "and empmst.Div = Division.Div"
Else
catArray = ""
'' if not a single category is in visible mode and chkcat is not checked
If cmbFrCatSel.ListCount > 0 Then
For i = 0 To cmbFrCatSel.ListCount - 1
catArray = catArray & "'" & cmbFrCatSel.List(i) & "',"
Next i
catArray = Left(catArray, Len(catArray) - 1)
strSql = " and empmst.empcode between '" & Trim(cmbFrEmpSel.Text) & "' and '" & _
         Trim(cmbToEmpSel.Text) & "' AND deptdesc.dept between " & Trim(cmbFrDepSel.Text) & _
         " and " & Trim(cmbToDepSel.Text) & " AND catdesc.cat  in " & _
          "(" & catArray & ")And " & _
         "groupmst." & strKGroup & " between " & Trim(cmbFrGrpSel.Text) & " and " & _
         Trim(cmbToGrpSel.Text) & strTmp & " and Location.Location between " & _
         Trim(cmbFrLocSel.Text) & " and " & Trim(cmbToLocSel.Text) & " and " & _
         "Division.Div between " & Trim(cmbFrDivSel.Text) & " and " & _
         Trim(cmbToDivSel.Text) & " and empmst.dept = deptdesc.dept and " & _
         " empmst.cat = catdesc.cat and empmst." & strKGroup & " = groupmst." & strKGroup & " and " & _
         "empmst.company = company.company and empmst.Location = Location.Location " & _
         "and empmst.Div = Division.Div"

Else
MsgBox "There is no category selected"
End If
End If
''For Mauritius 20-08-2003
If UCase(Trim(cboMain.Text)) <> "ALL" Then
    strSql = strSql & " And Empmst.qualf = '" & cboMain.Text & "'"
End If
''

''Following is the code for grouping
If optGrpEmp(0).Value = True Then                       ''Empcodewise
    sqlStr = "empcode": headGrp = "catdescdesc"
    StrGroup1 = NewCaptionTxt("D0012", adrsDSR)
    StrGroup2 = StrGroup1
ElseIf optGrpDep(1).Value = True Then                   ''Department
    sqlStr = "deptdescdept": headGrp = "deptdescdesc"
    StrGroup1 = NewCaptionTxt("D0011", adrsDSR)
    StrGroup2 = StrGroup1
ElseIf optGrpCat(2).Value = True Then
    sqlStr = "catdesccat": headGrp = "catdescdesc"      ''Category
    StrGroup1 = NewCaptionTxt("D0012", adrsDSR)
    StrGroup2 = StrGroup1
ElseIf optGrpGrp(3).Value = True Then
    sqlStr = "groupmst": headGrp = "groupmst"           ''Groupwise
    StrGroup1 = NewCaptionTxt("D0013", adrsDSR)
    StrGroup2 = StrGroup1
''For Mauritius 01-08-2003
ElseIf optGrpDC(4).Value = True Then                    ''Department/Category
    sqlStr = "Company"
    headGrp = "CName"
    StrGroup1 = NewCaptionTxt("00057", adrsMod)
    StrGroup2 = StrGroup1
''
ElseIf optGrpLoc(5).Value = True Then
    sqlStr = "Location": headGrp = "LocDesc"           ''Locationwise
    StrGroup1 = NewCaptionTxt("00110", adrsDSR)
    StrGroup2 = StrGroup1
ElseIf optGrpDiv(6).Value = True Then
    sqlStr = "Div": headGrp = "DivDesc"           ''Divisionwise
    StrGroup1 = NewCaptionTxt("00126", adrsMod)
    StrGroup2 = StrGroup1

End If
Exit Sub
Err_P:
    ShowError ("RetVALUES :: Reports Form")
End Sub

Private Sub FrSelFill()
On Error GoTo Err_P
'filling all the combos in selection frame
Call ComboFill(cmbFrEmpSel, 1, 2)      'emp
cmbToEmpSel.List = cmbFrEmpSel.List
Call ComboFill(cmbFrDepSel, 2, 2)      'dept
cmbToDepSel.List = cmbFrDepSel.List
Call ComboFill(cmbFrCatSel, 3, 2)      'cat
cmbToCatSel.List = cmbFrCatSel.List
Call ComboFill(cmbFrGrpSel, 8, 2)      'group
cmbToGrpSel.List = cmbFrGrpSel.List
Call ComboFill(cmbFrLocSel, 11, 2)     'Location
cmbToLocSel.List = cmbFrLocSel.List
Call ComboFill(cmbFrDivSel, 13, 2)     'Division
cmbToDivSel.List = cmbFrDivSel.List

Call ComboFill(cmbFrComSel, 5, 2)      'company
Call FillShiftCombo                    'shift
cmbFrComSel.AddItem "All"
''For Mauritius 20-08-2003
Call FillMainCombo                      ''Maintain-nonMaintain


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

optGrpEmp(0).Value = True
Exit Sub
Err_P:
    ShowError ("Combos Selects :: Reports Form")
End Sub

Public Sub SetRepVars(Optional bytSetRep As Byte = 7)
On Error GoTo Err_P
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
        If optmas(2).Value = True Then ''( Left Employee)
            typRep.strLeftFr = txtMastFr.Text
            typRep.strLeftTo = txtMastTo.Text
        End If
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
Err_P:
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
    cmdDaily.Value = True
    txtDaily.SetFocus
    Cancel = True
End If
End Sub

Private Sub txtDlyCAbs_GotFocus()
    Call GF(txtDlyCAbs)
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
    cmdWeekly.Value = True
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
    cmdMaster.Value = True
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
    cmdMaster.Value = True
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
    cmdPeriodic.Value = True
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
    cmdPeriodic.Value = True
    txtToPeri.SetFocus
    Cancel = True
End If
End Sub

Private Sub FrchkFill()
'' Gets the setting from the registry
'' or the date & time display in the report
chkDateT(0).Value = GetSetting("Vstar", "PrjSettings", "Show Date and Time", 0)
'' For the prompt before printing
'' chkPromp.Value = GetSetting("Vstar", "PrjSettings", "Prompt Before Printing", 0)
'' For breaking page when group changes
chkNewP.Value = GetSetting("Vstar", "PrjSettings", "Print on Next Page", 0)
End Sub

Private Function RecordsFound() As Boolean  '' CHECKS IF THE REQUIRED RECORDS ARE
On Error GoTo Err_P                         '' AVAILABLE OR NOT
RecordsFound = True
Call SetMSF1Cap(9)
'If bytBackEnd = 2 And bytRepMode <> 5 Then Sleep (2000)
'If RsName.State Then RsName.Close
'RsName.Open empstr3
If blnIntz = True Then
   CRV.ViewReport
   frmCRV.Show vbModal
  Else
    Call SetMSF1Cap(10)
   RecordsFound = False
   Exit Function
End If
Exit Function
Err_P:
    ShowError ("Records Found :: " & Me.Caption)
    RecordsFound = False
End Function

Private Function dlyReportsMod() As Boolean
On Error GoTo Err_P
dlyReportsMod = False                   '' FUNCTION FOR DAILY REPORTS
'Call SetRepVars(1)
'' Adjust Shift Inclusion Statements based on the Report & Shift Selected
If cboShift.ListIndex <> -1 Then
    If cboShift.ListIndex <> cboShift.ListCount - 1 Then
        Select Case typOptIdx.bytDly
            Case 10, 12  ''manpower,summary
                strSql = strSql & " and " & strMon_Trn & ".Shift='" & _
                cboShift.Text & "' "
            Case 8  ''entries report

            Case Else
                If typOptIdx.bytDly <> 9 Then
                    strSql = strSql & " and " & strMon_Trn & ".Shift='" & _
                    cboShift.Text & "' and " & strRepFile & ".shift = " & strMon_Trn & _
                    ".shift  and " & strMon_Trn & ".empcode = " & strRepFile & _
                    ".empcode and " & strMon_Trn & "." & strKDate & " = " & strRepFile & _
                    "." & strKDate & ""
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
    Case 8 'Entries
        If Not DlyEntries Then Exit Function
        dlyReportsMod = True
    Case 10 'Manpower
        If Not DlyManpower Then Exit Function
        dlyReportsMod = True
    Case 0, 1, 3, 4, 5, 6, 7, 9, 11, 13
        dlyReportsMod = True
    Case 12 ''Summary
        If Not DlySummar Then Exit Function
        dlyReportsMod = True
    ''For Mauritius 16-08-2003
    Case 14 ''Punch Variation
        If Not dlyPunchVari Then Exit Function
        dlyReportsMod = True
    ''
    Case Else
        dlyReportsMod = False
End Select
Exit Function
Err_P:
    ShowError ("DailyReportsMod :: " & Me.Caption)
End Function

Private Function dlySeachEntries() As Boolean
On Error GoTo RepErr                            '' FUNCTION FOR DAILY ENTRIES REPORT
dlySeachEntries = False
Call RetValues
frmEntries.Show vbModal
dlySeachEntries = True
Call SetRepVars(1)
Exit Function
RepErr:
dlySeachEntries = False
    ShowError ("Error in dlySeachEntries " & Me.Caption)
End Function

Private Function dlyValid() As Boolean
On Error GoTo Err_P
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
    bytNoDay = IIf(IsEmpty(txtDlyCAbs.Text), 0, txtDlyCAbs.Text)
    If bytNoDay > 31 Then
        Call SetMSF1Cap(10)
        MsgBox NewCaptionTxt("40055", adrsC), vbExclamation
        txtDlyCAbs.SetFocus
        dlyValid = False
        Exit Function
    ElseIf bytNoDay <= 0 Then
        Call SetMSF1Cap(10)
        MsgBox NewCaptionTxt("00072", adrsMod), vbExclamation
        txtDlyCAbs.SetFocus
        dlyValid = False
        Exit Function
    End If
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
End If  ''End of Cont abs reports check
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
Err_P:
    ShowError ("dlyValid :: Reportsfrm")
    dlyValid = False
End Function

Private Function WkValid() As Boolean
On Error GoTo Err_P
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
    MsgBox NewCaptionTxt("40065", adrsC) & _
        MonthName(Month(CDate(txtWeek.Text) + 6)), vbExclamation
    WkValid = False
    txtWeek.SetFocus
    Exit Function
End If
Exit Function
Err_P:
    ShowError ("wkValid :: Reportsfrm")
    WkValid = False
End Function

Private Function monValid() As Boolean
On Error GoTo Err_P
monValid = True                             '' FUNCTION FOR MONTHLY REPORT VALIDATIONS
Call SetMSF1Cap(7)
''No check required for
''Leave consumption report,Leave Balance ,OT paid hrs.
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
            MsgBox NewCaptionTxt("40065", adrsC) & " " & cmbMonth.Text & _
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
Err_P:
    ShowError ("monValid :: Reportsfrm")
    monValid = False
End Function

Private Function yrValid() As Boolean
On Error GoTo Err_P
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
Exit Function
Err_P:
    ShowError ("yrValid :: Reportsfrm")
    yrValid = False
End Function

Private Function maValid() As Boolean
On Error GoTo Err_P
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
Exit Function
Err_P:
    ShowError ("maValid :: Reportsfrm")
    maValid = False
    'Resume Next
End Function

Private Function PeValid() As Boolean
On Error GoTo Err_P
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
If (CDate(txtFrPeri.Text) + 31) < CDate(txtToPeri.Text) Then
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("40079", adrsC), vbInformation
    txtFrPeri.SetFocus
    PeValid = False
    Exit Function
End If
If Month(CDate(txtToPeri.Text)) - Month(CDate(txtFrPeri.Text)) > 1 Then
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("40080", adrsC), vbInformation
    txtFrPeri.SetFocus
    PeValid = False
    Exit Function
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
    Exit Function '' No need to check transaction file for future leave availment report.
End If
''
If Not FindTable(Left(MonthName(Month(CDate(txtFrPeri.Text))), 3) & _
Right(Year(CDate(txtFrPeri.Text)), 2) & "trn") Then
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("40065", adrsC) & _
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
Err_P:
    ShowError ("peValid :: Reportsfrm")
    PeValid = False
End Function

Private Function wkReportsMod() As Boolean
On Error GoTo Err_P
wkReportsMod = False                    '' FUNCTION FOR WEEKLY REPORT
Call SetRepVars(2)
Call SetMSF1Cap(8)
Select Case typOptIdx.bytWek
Case 7  'Performance,Irregular
    If Not WkPerfo(typRep.strWkDate, CStr(DateCompDate(typRep.strWkDate) + 6)) Then Exit Function
     wkReportsMod = True
    ''For Mauriitus 11-08-2003
    Case 0
    If Not WKPerfOvt Then Exit Function
    wkReportsMod = True
Case 1, 2, 3, 4, 5
    If Not WkOtherRep Then Exit Function
    wkReportsMod = True
Case 6 'Shift schedule
    If Not WkShiftRep() Then Exit Function
    wkReportsMod = True
Case Else
    wkReportsMod = False
End Select
Exit Function
Err_P:
    ShowError ("wkReportsMod :: Reportsfrm")
End Function

Private Function dlySetEmpstr3() As Boolean
On Error GoTo RepErr                    '' FUNCTION TO CREATE COMMANDTEXT QUERIES FOR DAILY
dlySetEmpstr3 = True
empstr3 = ""
Select Case typOptIdx.bytDly
Case 0 'Physical Arrival
       
    empstr3 = "SHAPE{SELECT " & strRepFile & ".empcode," & strRepFile & ".shift," & _
    rpGroup & ",empmst.Name," & strRepFile & ".arrtim, " & strRepFile & ".latehrs, " & _
    strRepFile & ".presabs, " & strRepFile & ".remarks FROM " & strRepFile & "," & rpTables & _
    " WHERE " & strRepFile & ".arrtim >0 and  empmst.empcode = " & strRepFile & ".empcode AND " & _
    strRepFile & "." & strKDate & "=" & strDTEnc & DateCompStr(typRep.strDlyDate) & strDTEnc & _
    " " & strSql & " order by " & strRepFile & ".empcode} as dailyperfrpt compute " & _
    "dailyperfrpt by '" & sqlStr & "','" & headGrp & "'"
Case 1 'Absent
    empstr3 = "SHAPE{SELECT " & strRepFile & ".empcode, " & strRepFile & ".Shift," & rpGroup & _
    ",empmst.Name," & strRepFile & ".presabs FROM " & strRepFile & "," & rpTables & _
    " WHERE " & strRepFile & ".empcode = empmst.empcode " & strSql & " order by " & _
    strRepFile & ".empcode} as dailyperfrpt compute dailyperfrpt by '" & sqlStr & _
    "','" & headGrp & "'"
Case 2 'Cont Absent
    empstr3 = "SHAPE {SELECT " & strRepFile & ".PresAbsStr," & strRepFile & ".empcode," & _
    "empmst.Name," & rpGroup & " FROM " & strRepFile & "," & rpTables & " WHERE " & _
     strRepFile & ".empcode=empmst.empcode " & strSql & " order by " & strRepFile & _
     ".empcode} as weekreport compuTE weekreport BY '" & sqlStr & "','" & headGrp & "'"
Case 3 'Late Arrival
    empstr3 = "SHAPE{SELECT " & strRepFile & ".empcode, " & strRepFile & ".shift," & _
    rpGroup & ",empmst.Name," & strRepFile & ".arrtim, " & strRepFile & ".latehrs," & _
    strRepFile & ".presabs, " & strRepFile & ".remarks FROM " & strRepFile & "," & _
    rpTables & " WHERE " & strRepFile & ".empcode " & " = empmst.empcode AND " & strRepFile & _
    "." & strKDate & "=" & strDTEnc & DateCompStr(typRep.strDlyDate) & strDTEnc & " and " & _
    strRepFile & ".latehrs>0" & " " & strSql & " order by " & strRepFile & ".empcode} as " & _
    "dailyperfrpt compute dailyperfrpt by '" & sqlStr & "','" & headGrp & "'"
Case 4 'Early Dep
    empstr3 = "SHAPE{SELECT " & strRepFile & ".empcode," & strRepFile & ".shift," & _
    rpGroup & ",empmst.Name," & strRepFile & ".deptim, " & strRepFile & ".earlhrs," & _
    strRepFile & ".presabs, " & strRepFile & ".remarks FROM " & strRepFile & "," & _
    rpTables & " WHERE " & strRepFile & ".empcode = empmst.empcode AND " & strRepFile & _
    "." & strKDate & "=" & strDTEnc & DateCompStr(typRep.strDlyDate) & strDTEnc & " and " & _
    strRepFile & ".earlhrs>0" & " " & strSql & " order by " & strRepFile & ".empcode}" & _
    " as dailyperfrpt compute dailyperfrpt by '" & sqlStr & "','" & headGrp & "'"
Case 5 'Perf
    empstr3 = "SHAPE{SELECT " & strRepFile & ".empcode," & strRepFile & ".shift," & _
    rpGroup & ",empmst.Name," & strRepFile & ".arrtim," & strRepFile & ".latehrs," & _
    strRepFile & ".actrt_o," & strRepFile & ".actrt_i, " & strRepFile & ".time5, " & _
    strRepFile & ".time6, " & strRepFile & ".deptim, " & strRepFile & ".earlhrs, " & _
    strRepFile & ".wrkhrs, " & strRepFile & ".presabs, " & strRepFile & ".remarks," & _
    strRepFile & ".ovtim," & strRepFile & ".OTConf FROM " & strRepFile & "," & rpTables & " WHERE " & strRepFile & _
    ".empcode = empmst.empcode AND " & strRepFile & "." & strKDate & "=" & strDTEnc & _
    DateCompStr(typRep.strDlyDate) & strDTEnc & strSql & " order by " & strRepFile & _
    ".empcode} as dailyperfrpt compute dailyperfrpt by '" & sqlStr & "','" & headGrp & "'"
Case 6 'Irreg
    empstr3 = "SHAPE{SELECT " & strRepFile & ".empcode," & rpGroup & ",empmst.Name," & _
    strRepFile & ".arrtim," & strRepFile & ".latehrs," & strRepFile & ".actrt_o," & _
    strRepFile & ".actrt_i," & strRepFile & ".od_from," & strRepFile & ".od_to," & _
    strRepFile & ".deptim," & strRepFile & ".wrkhrs FROM " & strRepFile & "," & rpTables & _
    " WHERE " & strRepFile & ".empcode = empmst.empcode AND " & strRepFile & "." & strKDate & " = " & _
    strDTEnc & DateCompStr(typRep.strDlyDate) & strDTEnc & " AND " & strRepFile & ".chq ='*' " & _
    strSql & " order by " & strRepFile & ".empcode} as dailyperfrpt compute dailyperfrpt " & _
    "by '" & sqlStr & "','" & headGrp & "'"
Case 7, 13 'authorized / unauthorized OT
    empstr3 = "SHAPE{SELECT " & strRepFile & ".empcode," & strRepFile & ".shift," & _
    rpGroup & ",empmst.Name," & strRepFile & ".arrtim," & strRepFile & ".latehrs," & _
    strRepFile & ".actrt_o, " & strRepFile & ".actrt_i, " & strRepFile & ".deptim, " & _
    strRepFile & ".earlhrs, " & strRepFile & ".wrkhrs, " & strRepFile & ".ovtim," & _
    strRepFile & ".OTRem as remarks FROM " & strRepFile & "," & rpTables & " WHERE " & _
    strRepFile & ".ovtim>0 AND " & strRepFile & ".empcode = empmst.empcode AND " & _
    strRepFile & "." & strKDate & "=" & strDTEnc & DateCompStr(typRep.strDlyDate) & strDTEnc & _
    strSql & " order by " & strRepFile & ".empcode } as dailyperfrpt compute " & _
    "dailyperfrpt by '" & sqlStr & "','" & headGrp & "'"
Case 8 'Entries
    empstr3 = "SHAPE{SELECT " & strRepFile & ".empcode," & rpGroup & ",empmst.Name," & _
    strRepFile & ".punches from " & strRepFile & "," & rpTables & " where " & strRepFile & _
    ".empcode =  empmst.empcode " & strSql & " order by " & strRepFile & ".empcode} " & _
    "AS entries  COMPUTE entries BY '" & sqlStr & "','" & headGrp & "'"
Case 9 'Shift Arrangement
    empstr3 = "SHAPE{SELECT " & strMon_Trn & ".empcode," & strMon_Trn & ".d" & _
    Day(DateCompDate(typRep.strDlyDate)) & " as shift, " & rpGroup & ",empmst.Name FROM " & _
    strMon_Trn & "," & rpTables & " WHERE (" & strMon_Trn & ".d" & Day(DateCompDate(typRep.strDlyDate)) & _
    " <> '' OR " & strMon_Trn & ".d" & Day(DateCompDate(typRep.strDlyDate)) & " IS NOT NULL ) AND " & _
    strMon_Trn & ".empcode = empmst.empcode " & strSql & " order by " & strMon_Trn & _
    ".empcode } as dailyperfrpt compute dailyperfrpt by '" & sqlStr & "','" & headGrp & "'"
    strMon_Trn = ""
Case 10 'Manpower
    empstr3 = "SHAPE { SELECT " & strRepFile & ".srno," & strRepFile & ".empcode ," & _
    rpGroup & ",empmst.Name," & strRepFile & ".present," & strRepFile & ".absent," & _
    strRepFile & ".offs," & strMon_Trn & ".ovtim FROM " & strRepFile & "," & strMon_Trn & _
    "," & rpTables & " WHERE " & strRepFile & ".empcode = empmst.empcode AND " & strRepFile & _
    ".empcode = " & strMon_Trn & ".empcode AND empmst.empcode = " & strMon_Trn & ".empcode AND " & strMon_Trn & "." & strKDate & " = " & strDTEnc & _
    DateCompStr(typRep.strDlyDate) & strDTEnc & strSql & " order by " & strRepFile & _
    ".empcode} AS Manpower COMPUTE Manpower BY '" & sqlStr & "','" & headGrp & "'"
Case 11 'OutDoor
 'Atul   Adjusment done  for daily report
    'Apoorva for Fairfield on 21/02/2005 to display outdoor in daily report
    strLvloc = ""
    strLvloc = Left(MonthName(Month(typRep.strDlyDate)), 3) & Right(Year(typRep.strDlyDate), 2) & "trn"

    empstr3 = "SHAPE{SELECT " & strLvloc & ".empcode," & strLvloc & ".shift," & _
    rpGroup & ",empmst.Name," & strLvloc & "." & strKDate & "," & strLvloc & ".presabs FROM " & _
    strLvloc & "," & strRepFile & "," & rpTables & " WHERE " & strLvloc & ".empcode = empmst.empcode " & _
    "AND " & strLvloc & "." & strKDate & "=" & strDTEnc & DateCompStr(typRep.strDlyDate) & strDTEnc & _
    " and " & strLvloc & ".presabs = 'ODOD'" & strSql & " order by " & _
    strLvloc & ".empcode} as dailyperfrpt compute dailyperfrpt by '" & sqlStr & "'," & _
    "'" & headGrp & "'"
Case 12 'Summary
    empstr3 = "select * from " & strRepFile & " order by serial"
''For Mauritius 16-08-2003
Case 14 ''Punch Variationn
    empstr3 = "SHAPE{Select " & strRepFile & ".* ,Name," & rpGroup & " from " & strRepFile & "," & _
    rpTables & " where " & strRepFile & ".empcode = empmst.empcode " & strSql & " ORDER BY " & _
    strRepFile & ".Empcode} as dailyperfrpt compute dailyperfrpt by '" & sqlStr & "'," & _
    "'" & headGrp & "'"
''
End Select
Call CadSetting
Exit Function
RepErr:
    dlySetEmpstr3 = False
    ShowError ("Dlysetempstr3 :: " & Me.Caption)
End Function

Private Function WkSetEmpstr3() As Boolean
On Error GoTo RepErr                '' FUNCTION TO CREATE COMMANDTEXT QUERIES FOR WEEKLY
WkSetEmpstr3 = True
empstr3 = ""
Select Case typOptIdx.bytWek
    Case 0, 7 'Performance,Irregular
        empstr3 = "SHAPE {SELECT " & strRepFile & ".empcode,empmst.Name," & strRepFile & _
        "." & strKDate & "," & strRepFile & ".ArrStr," & strRepFile & ".DepStr," & strRepFile & _
        ".EarlStr, " & strRepFile & ".LateStr, " & strRepFile & ".OTStr," & strRepFile & _
        ".PresAbsStr, " & strRepFile & ".ShfStr, " & strRepFile & ".WorkStr," & strRepFile & _
        ".punches," & strRepFile & ".sumlate," & strRepFile & ".sumearly," & strRepFile & _
        ".sumwork," & strRepFile & ".sumextra," & rpGroup & " FROM " & strRepFile & "," & _
        rpTables & " WHERE  " & strRepFile & ".empcode = empmst.empcode " & strSql & _
        " ORDER BY " & strRepFile & ".empcode}  AS weekreport COMPUTE  " & _
        " weekreport BY '" & sqlStr & "','" & headGrp & "'"
    Case 1, 2, 3, 4, 5 'Absent,Attendance,Late Arrival,Early Departure,Overtime
         empstr3 = "SHAPE {select " & strRepFile & ".empcode," & rpGroup & ",empmst.Name," & _
         strRepFile & ".frw," & strRepFile & ".secw," & strRepFile & ".thw," & strRepFile & _
         ".fow," & strRepFile & ".fiw," & strRepFile & ".siw," & strRepFile & ".sevw from " & _
        strRepFile & "," & rpTables & " where empmst.empcode=" & strRepFile & ".empcode " & _
        strSql & " ORDER BY " & strRepFile & ".EMPCODE} AS WkOther COMPUTE WkOther BY " & _
        "'" & sqlStr & "','" & headGrp & "'"
    Case 6 'Shift Schedule
        empstr3 = "SHAPE {select " & strRepFile & ".empcode," & rpGroup & ",empmst.Name," & _
        strRepFile & ".frw," & strRepFile & ".secw," & strRepFile & ".thw," & strRepFile & _
        ".fow," & strRepFile & ".fiw," & strRepFile & ".siw," & strRepFile & ".sevw from " & _
        strRepFile & "," & rpTables & " where empmst.empcode=" & strRepFile & ".empcode " & _
        strSql & " } AS WkOther COMPUTE WkOther BY '" & sqlStr & "','" & headGrp & "'"
End Select
Call CadSetting
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
    Case 0, 5, 10, 11 '  Perofrmance,OverTime, Late Arrival,Early Departure
        empstr3 = "SHAPE {SELECT " & strRepFile & ".ArrStr," & strRepFile & ".DepStr," & _
        strRepFile & ".EarlStr, " & strRepFile & ".LateStr, " & strRepFile & ".OTStr," & _
        strRepFile & ".PresAbsStr, " & strRepFile & ".ShfStr, " & strRepFile & ".WorkStr," & _
        strRepFile & "." & strKDate & ", " & strRepFile & ".empcode, empmst.Name," & rpGroup & _
        ", " & strRepFile & ".sumlate," & strRepFile & ".sumearly," & strRepFile & _
        ".sumwork," & strRepFile & ".sumOT FROM " & strRepFile & "," & rpTables & " WHERE " & _
        strRepFile & ".empcode = empmst.empcode " & strSql & " ORDER BY " & strRepFile & _
        ".empcode} AS MnlReport COMPUTE MnlReport BY '" & sqlStr & "','" & headGrp & "'"
    Case 1, 2, 3, 4 '  Attendance, Muster Report,Monthly Present,Monthly Absent
        empstr3 = "SHAPE {SELECT " & strRepFile & ".empcode," & strRepFile & ".presabsstr," & _
        strRepFile & ".leavestr," & strRepFile & ".mndatestr," & strRepFile & ".pdaysstr," & _
        strRepFile & ".otstr," & strRepFile & ".wrkstr," & strRepFile & ".nightstr," & _
        strRepFile & ".lvval,empmst.Name," & rpGroup & " FROM " & strRepFile & "," & _
        rpTables & " WHERE " & strRepFile & ".Empcode = Empmst.empcode " & strSql & _
        " ORDER BY " & strRepFile & ".empcode} AS MonthAtt COMPUTE MonthAtt BY " & _
        "'" & sqlStr & "'" & "," & "'" & headGrp & "'"
    Case 6 '  OverTime Paid
        empstr3 = "SHAPE {SELECT " & strMon_Trn & ".empcode," & strMon_Trn & ".ot_hrs," & _
        strMon_Trn & ".otpd_hrs, empmst.Name, " & rpGroup & "  FROM " & strMon_Trn & _
        "," & rpTables & " WHERE " & strMon_Trn & ".LST_DATE = " & strDTEnc & _
        DateCompStr(strLastDateM) & strDTEnc & " AND " & strMon_Trn & ".otpd_hrs > 0 AND " & _
        "empmst.empcode = " & strMon_Trn & ".empcode " & strSql & " }  AS MonOtPd COMPUTE " & _
        "MonOtPd BY " & "'" & sqlStr & "'" & "," & "'" & headGrp & "'"
        strMon_Trn = ""
    Case 7, 12, 13 ' Absent Memo,Late Memo,Early Memo
        empstr3 = " shape{SELECT " & strRepFile & "." & strKDate & "," & strRepFile & ".empcode ," & _
        strRepFile & ".presabsstr," & strRepFile & ".latestr," & strRepFile & ".earlstr," & _
        "empmst.Name," & rpGroup & " FROM " & strRepFile & "," & rpTables & " WHERE " & _
        strRepFile & ".empcode=empmst.empcode " & strSql & "} AS Monmemo COMPUTE Monmemo BY " & _
        "'" & sqlStr & "'" & "," & "'" & headGrp & "'"
    Case 8 '  Absent/Late/Early
        empstr3 = "SHAPE {SELECT " & strRepFile & ".empcode, Empmst.Name," & strRepFile & _
        ".Latehrs AS sumlate," & strRepFile & ".LAteno AS noLate, " & strRepFile & _
        ".Earlyhrs AS sumearly," & strRepFile & ".earlyno AS noearl, " & strRepFile & _
        ".absent AS noAbsent," & rpGroup & " FROM " & strRepFile & "," & rpTables & " WHERE " & _
        "Empmst.empcode = " & strRepFile & ".empcode  " & strSql & " } AS MonALE COMPUTE " & _
        "MonALE BY " & "'" & sqlStr & "'" & "," & "'" & headGrp & "'"
    Case 9 '  Leave Balance
        empstr3 = "SHAPE {SELECT " & strRepFile & ".empcode,empmst.Name," & strRepFile & _
        ".leavestr," & strRepFile & ".lvval," & rpGroup & " FROM " & strRepFile & "," & _
        rpTables & " WHERE " & strRepFile & ".Empcode = Empmst.empcode " & strSql & _
        " ORDER BY " & strRepFile & ".empcode} " & "AS MonLvbal COMPUTE MonLvbal BY " & _
        "'" & sqlStr & "'" & "," & "'" & headGrp & "'"
    Case 14 ' Leave Consumption
        empstr3 = "SHAPE ( SHAPE {SELECT DISTINCT empmst.empcode, Name, " & rpGroup & _
        " From leavdesc," & strRepFile & "," & rpTables & " WHERE " & strRepFile & _
        ".empcode = Empmst.empcode AND " & strRepFile & ".lcode = Leavdesc.lvcode AND " & _
        "leavdesc.cat = catdesc.cat " & strSql & " order by empmst.empcode} AS leavetest " & _
        "APPEND ({select distinct " & strRepFile & ".empcode,lcode,fromdate,todate,days," & _
        "leave,trcd from " & strRepFile & ",leavdesc," & rpTables & " WHERE " & strRepFile & _
        ".empcode = empmst.empcode AND " & strRepFile & ".lcode = Leavdesc.lvcode AND " & _
        "leavdesc.cat = catdesc.cat " & strSql & " order by " & strRepFile & ".empcode} " & _
        "AS LeaveChild RELATE 'empcode' TO " & "'empcode') AS LeaveChild ) COMPUTE " & _
        "leavetest BY " & "'" & sqlStr & "'" & "," & "'" & headGrp & "'"
        strMon_Trn = ""
    Case 15, 16 ' Total Lates,Total Earlys
        empstr3 = " shape { SELECT Empmst.Name, " & strRepFile & ".empcode," & rpGroup & _
        "," & strRepFile & ".lvval as latecnt," & strRepFile & ".daysded From " & strRepFile & _
        "," & rpTables & " WHERE Empmst.empcode = " & strRepFile & ".empcode " & strSql & _
        "} AS monlateearlrpt COMPUTE monlateearlrpt  BY '" & sqlStr & "','" & headGrp & "'"
    Case 17 ' Shift Schedule
        empstr3 = "SHAPE {SELECT " & strRepFile & ".empcode,empmst.Name," & strRepFile & _
        ".shfstr," & rpGroup & " FROM " & strRepFile & "," & rpTables & " WHERE " & _
        strRepFile & ".empcode = Empmst.empcode " & strSql & "} AS MonShift COMPUTE " & _
        "MonShift BY " & "'" & sqlStr & "'" & "," & "'" & headGrp & "'"
    Case 18 ' WO On Holiday
        empstr3 = "SHAPE {SELECT " & strMon_Trn & ".empcode," & strMon_Trn & "." & strKDate & "," & _
        strMon_Trn & ".presabs, empmst.Name," & rpGroup & " From " & strMon_Trn & "," & _
        rpTables & " WHERE " & "(" & LeftStr(strLastDateM) & " = empmst." & strKOff & " OR " & _
        LeftStr(strLastDateM) & " = empmst.off2 OR " & LeftStr(strLastDateM) & " " & _
        "= empmst.wo_1_3 OR " & LeftStr(strLastDateM) & " = empmst.wo_2_4) AND " & _
        strMon_Trn & ".presabs = '" & pVStar.HlsCode & pVStar.HlsCode & "' AND " & _
        "empmst.empcode = " & strMon_Trn & ".empcode " & strSql & "} AS MonWo " & _
        "COMPUTE MonWo BY " & "'" & sqlStr & "'" & "," & "'" & headGrp & "'"
        strMon_Trn = ""
End Select
Call CadSetting
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
    Case 0, 3 'Absent,Present
        empstr3 = "SHAPE {select " & strRepFile & ".EmpCode," & strRepFile & ".yValStr," & _
        "empmst.Name, " & rpGroup & " from " & strRepFile & "," & rpTables & " where " & _
        strRepFile & ".empcode=empmst.empcode " & strSql & "}AS YrAbPr COMPUTE YrAbPr " & _
        "BY " & "'" & sqlStr & "'" & "," & "'" & headGrp & "'"
    Case 1, 2 'Mandays,Performance
        empstr3 = "SHAPE ( SHAPE {select distinct empmst.empcode,Name," & rpGroup & _
        " from " & rpTables & " ," & strRepFile & " WHERE empmst.empcode = " & strRepFile & _
        ".empcode " & strSql & " order by empmst.empcode} AS YrPerfo APPEND ({select " & _
        strRepFile & ".empcode," & strRepFile & ".ystr," & strRepFile & ".yvalstr," & _
        strRepFile & ".pddaysstr," & strRepFile & ".wrkstr," & strRepFile & ".nightstr," & _
        strRepFile & ".ltno," & strRepFile & ".latehrs," & strRepFile & ".erno," & _
        strRepFile & ".earlhrs," & strRepFile & ".counter," & rpGroup & " FROM " & _
        strRepFile & "," & rpTables & " WHERE Empmst.empcode = " & strRepFile & ".Empcode " & _
        strSql & " ORDER BY " & strRepFile & ".EMPCODE," & strRepFile & ".counter} AS " & _
        "yrperfochild RELATE 'empcode' TO 'empcode') as yrperfochild) COMPUTE yrperfo " & _
        " by '" & sqlStr & "','" & headGrp & "' "
    Case 4 'Leave Information
        empstr3 = "SHAPE ( SHAPE {select distinct empmst.empcode,Name," & rpGroup & _
        " from " & rpTables & "," & strRepFile & " WHERE empmst.empcode =" & strRepFile & _
        ".empcode " & strSql & " order by empmst.empcode} AS YrLeaveCon APPEND ({SELECT " & _
        strRepFile & ".empcode, " & strRepFile & ".lcode," & strRepFile & ".FromLv ," & _
        strRepFile & ".todate, " & strRepFile & ".AvailLv, " & strRepFile & ".CreditLv " & _
        "FROM " & strRepFile & "," & rpTables & " WHERE empmst.empcode = " & strRepFile & _
        ".empcode " & strSql & "} AS YrLvChild RELATE 'empcode' TO 'empcode') AS YrLvChild) " & _
        "COMPUTE YrLeaveCon BY '" & sqlStr & "','" & headGrp & "'"
End Select
Call CadSetting
Exit Function
RepErr:
    yrSetEmpstr3 = False
    ShowError ("yrSetEmpstr3 :: " & Me.Caption)
End Function

Private Function maSetEmpstr3() As Boolean
On Error GoTo RepErr                '' FUNCTION TO CREATE COMMANDTEXT QUERIES FOR MASTERS
maSetEmpstr3 = True
Dim Empstr1 As String, Empstr2 As String

'empstr3 = "shape {" & _
"SELECT Empmst.name  ,Empmst.Empcode,Empmst.card,Empmst.designatn,groupmst.grupdesc as groupmst,catdesc.cat as catdesccat," & _
"deptdesc.dept as deptdescdept,CatDesc.""Desc"" as catdescdesc,deptdesc.""Desc"" as deptdescdesc,Location.Location," & _
"Location.LocDesc,Division.Div,Division.DivDesc,Empmst.joindate,Empmst.""Group"",Empmst.styp,Empmst.entry,Empmst.birth_dt," & _
"Empmst.salary,Empmst.resadd1,Empmst.city,Empmst.phone,Empmst.pin,Empmst.sex,Empmst.bg,Empmst.udf1,Empmst.udf2,Empmst.udf3," & _
"Empmst.udf4,Empmst.udf5,Empmst.udf6,Empmst.udf7,Empmst.udf9,Empmst.udf10,Empmst.leavdate FROM empmst,catdesc,deptdesc,groupmst," & _
"company,Location,Division WHERE  empmst.leavdate is null  and empmst.empcode between '1001' and '1003' AND deptdesc.dept " & _
"between 1 and 1 AND catdesc.cat  between 'MGR' and 'MGR' And groupmst.""Group"" between 0 and 1 and Location.Location " & _
"between 1 and 1 and Division.Div between 1 and 1 and empmst.dept = deptdesc.dept and  empmst.cat = catdesc.cat and " & _
"empmst.""Group"" = groupmst.""Group"" and empmst.company = company.company and " & _
"empmst.Location = Location.Location and empmst.Div = Division.Div ORDER BY Empmst.Empcode,catdesc.cat,deptdesc.dept} as employee compute employee by '" & sqlStr & "','" & headGrp & "'"


Empstr1 = "SELECT Empmst.name,Empmst.Empcode,Empmst.card,Empmst.designatn," & _
            rpGroup & ",Empmst.joindate,Empmst." & strKGroup & ",Empmst.styp," & _
            "Empmst.entry,Empmst.birth_dt,Empmst.salary,Empmst.resadd1,Empmst.city," & _
            "Empmst.phone,Empmst.pin,Empmst.sex,Empmst.bg,Empmst.udf1,Empmst.udf2," & _
            "Empmst.udf3,Empmst.udf4,Empmst.udf5,Empmst.udf6,Empmst.udf7,Empmst.udf9," & _
            "Empmst.udf10,Empmst.leavdate FROM " & rpTables & " WHERE "
Empstr2 = " ORDER BY Empmst.Empcode,catdesc.cat,deptdesc.dept"
empstr3 = ""
Select Case typOptIdx.bytMst
    Case 0, 1 'Employee list,Employee details
        'empstr3 = "shape { " & Empstr1 & " empmst.leavdate is null " & strSql & _
        Empstr2 & " } as employee compute employee by '" & sqlStr & "','" & headGrp & "'"
        
        empstr3 = Empstr1 & " empmst.leavdate is null " & strSql & _
        Empstr2 & " "
        
        
        
    Case 2 'Left Employee
        empstr3 = "shape {" & Empstr1 & " empmst.leavdate is not NULL and " & _
        "empmst.leavdate between " & strDTEnc & DateCompStr(typRep.strLeftFr) & strDTEnc & _
        " and " & strDTEnc & DateCompStr(typRep.strLeftTo) & strDTEnc & " " & strSql & _
        " " & Empstr2 & "} as employee compute employee by '" & sqlStr & "','" & headGrp & "'"
    Case 3 'Leave Master
        empstr3 = "select * from leavdesc where isitleave = 'Y'"
    Case 4  ''Shift master
        ''SUPRIYA DATED 27/05/05
        If InVar.strSer = 1 Then
            empstr3 = "SELECT *,  shiftdd= " & _
                         "CASE " & _
                        "WHEN  isnumeric(shift) =0  and len(shift)=1 THEN ascii(substring(shift,1,1))" & _
                        "WHEN  isnumeric(shift) =0  and len(shift)=2 THEN ascii(substring(shift,1,1))" & _
                        " +  ascii(substring(shift,2,1))" & _
                        "WHEN  isnumeric(shift) =1   THEN Shift " & _
                        "End " & _
                        " From instshft where shift<>100" & _
                        "order by shiftdd"
        Else
            empstr3 = "Select hdend,hdstart,rst_in,rst_out,shf_hrs,shf_in,shf_out,shift,shiftname,rst_brk from instshft where shift <> '100' ORDER BY shift"
        End If
    Case 5  ''Rotation Shift Master
        If InVar.strSer = 1 Then
            empstr3 = "SELECT SCode,Name,Skp,Pattern,Mon_Oth,Tot_Shf,Tot_Skp,Day_Skp,  shiftdd= " & _
            "CASE " & _
            "WHEN  isnumeric(SCode) =0  and len(SCode)=1 THEN ascii(substring(SCode,1,1))" & _
            "WHEN  isnumeric(SCode) =0  and len(SCode)=2 THEN ascii(substring(SCode,1,1))" & _
            "           +  ascii(substring(SCode,2,1))" & _
            "WHEN  isnumeric(SCode) =1   THEN SCode " & _
            "End " & _
            " From Ro_Shift where scode<>100" & _
            "order by shiftdd"
        Else
            empstr3 = "select scode,Name,mon_oth,pattern,skp from ro_shift where Scode <> '100' order by scode"
        End If
    Case 6 'Holiday
        empstr3 = "select Catdesc." & strKDesc & " as cat," & strKDate & ",Holiday." & strKDesc & " from holiday,Catdesc where holiday.cat=catdesc.cat and " & _
        "holiday." & strKDate & " between " & _
        strDTEnc & DateCompStr(FdtLdt(CByte(pVStar.Yearstart), pVStar.YearSel, "f")) & _
        strDTEnc & " and " & strDTEnc & DateCompStr(FdtLdt(CByte(pVStar.Yearstart) - 1, _
        IIf(pVStar.Yearstart = "1", pVStar.YearSel, CStr(Val(pVStar.YearSel) + 1)), "l")) & _
        strDTEnc & " ORDER BY holiday." & strKDate & ",holiday.cat"
    Case 7  ''Department
        empstr3 = "Select dept," & strKDesc & " from deptdesc"
        empstr3 = "select deptdesc.dept ," & strKDesc & ", " & _
        " count(deptdesc.dept) as stre from deptdesc,empmst where " & _
        "empmst.dept = deptdesc.dept  group by deptdesc.dept," & strKDesc & ""
    Case 8  ''Category
        empstr3 = "Select * from Catdesc where cat <> '100'"
    Case 9  ''Group
        empstr3 = "select groupmst." & strKGroup & " as dept,grupdesc as " & strKDesc & ", " & _
        " count(groupmst." & strKGroup & ") as stre from groupmst,empmst where " & _
        "empmst." & strKGroup & " = groupmst." & strKGroup & "  group by groupmst." & strKGroup & ",grupdesc"
    Case 10 ''Location
        empstr3 = "select location.location as dept,locdesc as " & strKDesc & ", " & _
        " count(location.location) as stre from location,empmst where " & _
        "empmst.location = location.location  group by location.location,locdesc"
    Case 11 ''Company
        empstr3 = "select company.company as dept,cname as " & strKDesc & ", " & _
        " count(company.company) as stre from company,empmst where " & _
        "empmst.company = company.company  group by company.company,cname"
    Case 12 ''Division
        empstr3 = "select division.div as dept,divdesc as " & strKDesc & ",count(division.div) " & _
        " as stre from division,empmst where empmst.div = division.div  group by division.div,divdesc"
End Select
'Call CadSetting
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
    Case 0  'Performance
        empstr3 = "SELECT " & strRepFile & ".*,empmst.Name," & rpGroup & _
         " FROM " & strRepFile & "," & rpTables & " WHERE " & _
         strRepFile & ".empcode=empmst.empcode " & strSql & ""
    Case 2  'Overtime
         empstr3 = "SHAPE {SELECT " & strRepFile & ".empcode,empmst.Name," & strRepFile & _
         "." & strKDate & "," & strRepFile & ".ArrStr, " & strRepFile & ".DepStr," & strRepFile & _
         ".EarlStr, " & strRepFile & ".LateStr, " & strRepFile & ".OTStr," & strRepFile & _
         ".PresAbsStr, " & strRepFile & ".ShfStr, " & strRepFile & ".WorkStr," & strRepFile & _
         ".sumlate," & strRepFile & ".sumearly," & strRepFile & ".sumwork," & rpGroup & _
         "," & strRepFile & ".sumextra FROM " & strRepFile & "," & rpTables & " WHERE " & _
         strRepFile & ".empcode=empmst.empcode " & strSql & "} as weekreport compuTE " & _
         "weekreport BY " & "'" & sqlStr & "'" & "," & "'" & headGrp & "'"
    Case 1 'Muster Report
        empstr3 = "SHAPE {SELECT " & strRepFile & ".empcode," & strRepFile & ".presabsstr," & _
        "empmst.Name, " & rpGroup & " FROM " & strRepFile & "," & rpTables & " WHERE " & _
        strRepFile & ".Empcode = Empmst.empcode " & strSql & " ORDER BY " & strRepFile & _
        ".empcode} AS mnlreport  COMPUTE mnlreport  BY '" & sqlStr & "','" & headGrp & "'"
    Case 3, 4 'Late Arrival,Early Departure
        empstr3 = "SHAPE {SELECT " & strRepFile & ".empcode,empmst.Name," & strRepFile & _
        ".ArrStr, " & strRepFile & ".DepStr," & strRepFile & ".EarlStr, " & strRepFile & _
        ".LateStr, " & strRepFile & ".OTStr," & strRepFile & ".PresAbsStr, " & strRepFile & _
        ".ShfStr, " & strRepFile & ".WorkStr," & strRepFile & "." & strKDate & "," & strRepFile & _
        ".sumlate," & strRepFile & ".sumearly," & rpGroup & " FROM " & strRepFile & "," & _
        rpTables & " WHERE " & strRepFile & ".empcode=empmst.empcode " & strSql & "} as " & _
        "weekreport compuTE weekreport BY '" & sqlStr & "'," & "'" & headGrp & "'"
    Case 5 ''Continuous Absent
        empstr3 = "SHAPE {SELECT " & strRepFile & ".PresAbsStr,empmst.Name," & strRepFile & _
        ".empcode," & rpGroup & " FROM " & strRepFile & "," & rpTables & " WHERE " & _
        strRepFile & ".empcode=empmst.empcode " & strSql & " order by " & strRepFile & _
        ".empcode} as weekreport compuTE weekreport BY '" & sqlStr & "','" & headGrp & "'"
    Case 6  ''Summary
        empstr3 = "select * from " & strRepFile & " order by serial"
    ''For Mauritius 11-07-2003
    Case 7  ''Meal Allowance
         empstr3 = "SHAPE {SELECT " & strRepFile & ".empcode,empmst.Name," & strRepFile & _
         "." & strKDate & "," & strRepFile & ".ArrStr, " & strRepFile & ".DepStr," & strRepFile & _
         ".EarlStr, " & strRepFile & ".LateStr, " & strRepFile & ".OTStr," & strRepFile & _
         ".PresAbsStr, " & strRepFile & ".ShfStr, " & strRepFile & ".WorkStr," & strRepFile & _
         ".sumlate," & strRepFile & ".sumearly," & strRepFile & ".sumwork," & rpGroup & _
         "," & strRepFile & ".sumextra FROM " & strRepFile & "," & rpTables & " WHERE " & _
         strRepFile & ".empcode=empmst.empcode " & strSql & "} as weekreport compuTE " & _
         "weekreport BY " & "'" & sqlStr & "'" & "," & "'" & headGrp & "'"
    ''
    ''For Mauritius 12-07-2003
    Case 8  ''8 Punches
        empstr3 = "SHAPE {SELECT " & strRepFile & ".empcode,empmst.Name," & strRepFile & _
         "." & strKDate & "," & strRepFile & ".ArrStr, " & strRepFile & ".DepStr," & strRepFile & _
         ".EarlStr, " & strRepFile & ".LateStr, " & strRepFile & ".OTStr," & strRepFile & _
         ".PresAbsStr, " & strRepFile & ".ShfStr, " & strRepFile & ".WorkStr," & strRepFile & _
         ".sumlate," & strRepFile & ".sumearly," & strRepFile & ".sumwork,ACTRT_O,ACTRT_I ," & _
         "TIME5,TIME6,TIME7,TIME8, " & rpGroup & "," & strRepFile & ".sumextra FROM " & _
         strRepFile & "," & rpTables & " WHERE " & strRepFile & ".empcode=empmst.empcode " & _
         strSql & "} as weekreport compuTE weekreport BY '" & sqlStr & "','" & headGrp & "'"
    ''
    ''For Mauritius 14-07-2003
    Case 9  '' Permission cards
        empstr3 = "SHAPE {SELECT " & strRepFile & ".empcode,empmst.Name," & strRepFile & _
         "." & strKDate & "," & strRepFile & ".ArrStr, " & strRepFile & ".DepStr," & strRepFile & _
         ".EarlStr, " & strRepFile & ".LateStr, " & strRepFile & ".OTStr," & strRepFile & _
         ".PresAbsStr, " & strRepFile & ".ShfStr, " & strRepFile & ".WorkStr," & strRepFile & _
         ".sumlate," & strRepFile & ".sumearly," & strRepFile & ".sumwork," & rpGroup & _
         "," & strRepFile & ".sumextra FROM " & strRepFile & "," & rpTables & " WHERE " & _
         strRepFile & ".empcode=empmst.empcode " & strSql & "} as weekreport compuTE " & _
         "weekreport BY '" & sqlStr & "','" & headGrp & "'"
    ''
    ''For Mauritius 14-07-2003
    Case 10     ''peAvailment
        empstr3 = "SHAPE ( SHAPE {SELECT DISTINCT empmst.empcode, Name, " & rpGroup & _
        " From leavdesc," & strRepFile & "," & rpTables & " WHERE " & strRepFile & _
        ".empcode = Empmst.empcode AND " & strRepFile & ".lcode = Leavdesc.lvcode AND " & _
        "leavdesc.cat = catdesc.cat " & strSql & " order by empmst.empcode} AS leavetest " & _
        "APPEND ({select distinct " & strRepFile & ".empcode,lcode,fromdate,todate,days," & _
        "leave,trcd from " & strRepFile & ",leavdesc," & rpTables & " WHERE " & strRepFile & _
        ".empcode = empmst.empcode AND " & strRepFile & ".lcode = Leavdesc.lvcode AND " & _
        "leavdesc.cat = catdesc.cat " & strSql & " order by " & strRepFile & ".empcode} " & _
        "AS LeaveChild RELATE 'empcode' TO " & "'empcode') AS LeaveChild ) COMPUTE " & _
        "leavetest BY " & "'" & sqlStr & "'" & "," & "'" & headGrp & "'"
        strMon_Trn = ""
    ''
End Select
Call CadSetting
Exit Function
RepErr:
    peSetEmpstr3 = False
    ShowError ("peSetEmpstr3 :: " & Me.Caption)
End Function

Private Sub CadSetting()
If headGrp = "catdesccat','catdescdesc" And sqlStr = "deptdescdept','deptdescdesc" Then
    sqlStr = "cad"
End If
End Sub

Public Function ChkPrinter(ByRef Rptin As Object, Optional ByVal bytPorLan As Byte = 1) As Boolean
On Error GoTo Err_P
ChkPrinter = True                   '' SETS REPORT'S ORIENTATION
If chkDotMa(1).Value = 1 Then
    Rptin.Orientation = Printer.Orientation
Else
    If bytPorLan = 1 Then
        Rptin.Orientation = Printer.Orientation
    Else
        Rptin.Orientation = rptOrientLandscape
    End If
End If
Exit Function
Err_P:
    ShowError ("ChkPrinter :: " & Me.Caption)
    ChkPrinter = False
End Function

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
    Case 0
        If Not pePerfCryst Then Exit Function
       peReportsMod = True
        
    Case 2  'Performance,Overtime
        If Not pePerfOvt() Then Exit Function
        peReportsMod = True
    Case 1  'Muster Report
        If Not peMuster() Then Exit Function
        peReportsMod = True
    Case 3, 4 'Late Arrival,Early Departure
        If Not peLateEarl Then Exit Function
        peReportsMod = True
    Case 5 'continuous Absent
        If Not peContAbs Then Exit Function
        peReportsMod = True
    ''For Mauritius 11-07-2003
    Case 6 'Summary
        If Not peSummary Then Exit Function
        peReportsMod = True
    Case 7 'Meal Allowance
        If Not peMealAl Then Exit Function
        peReportsMod = True
    ''
    ''For Mauritius 12-07-2003
    Case 8  '' 8 Punches
        If Not pe8Punches Then Exit Function
        peReportsMod = True
    ''
    ''For Mauritius 14-07-2003
    Case 9  ''Permission cards
        If Not pePermission Then Exit Function
        peReportsMod = True
    ''
    ''For Mauritius 14-07-2003
    Case 10     ''Leave Availment Report
        If Not peLeaveAvail Then Exit Function
        peReportsMod = True
End Select
End Function

Private Function yrReportsMod() As Boolean
yrReportsMod = False            '' FUNCTION FOR YEARLY REPORTS
Call SetRepVars(4)
Call SetMSF1Cap(8)
Select Case typOptIdx.bytYer
    Case 0, 3 'Absent,Present
        If Not yrAbsPrs Then Exit Function
        yrReportsMod = True
    Case 1, 2 'Mandays,Perofrmance
        If Not yrManPerf Then Exit Function
        yrReportsMod = True
    Case 4 'Leave Info
        If Not yrLeaveInfo Then Exit Function
        yrReportsMod = True
End Select
End Function

Private Function monReportsMod() As Boolean
On Error GoTo Err_P
monReportsMod = False               '' FUNCTION FOR MONTHLY REPORTS
Call SetRepVars(3)
Call SetMSF1Cap(8)
Select Case typOptIdx.bytMon
    Case 0, 5, 10, 11 'Performance,Overtime,Late Arrival,Early Departure
        If Not monPerfOt Then Exit Function
        monReportsMod = True
    Case 1, 2, 3, 4 'Attendance,Muster Report,Monthly Present,Monthly Absent
        If Not monAtMuPA Then Exit Function
        monReportsMod = True
    Case 6 'Overtime Paid
        strMon_Trn = ""
        strLastDateM = FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "L")
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
    Case 7, 12, 13 'Absent Memo,Late Arrival Memo,Early Departure Memo
        If Not monALEMemo Then Exit Function
        monReportsMod = True
    Case 8 'Total Absent/Late/Early
        If Not monALERep Then Exit Function
        monReportsMod = True
    Case 9 'Leave Balance
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
            Exit Function
        End If
        strFirstDateM = FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "f")
        strLastDateM = FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "l")

        VstarDataEnv.cnDJConn.Execute " insert into " & strRepFile & "(empcode,lcode," & _
        "fromdate,todate,days,trcd ) select empcode,lcode,fromdate,todate,days,trcd  from " & _
        strMon_Trn & " where trcd in(4,6,7) and (fromdate between " & strDTEnc & DateCompStr(strFirstDateM) & _
        strDTEnc & " and" & strDTEnc & DateCompStr(strLastDateM) & strDTEnc & " or " & _
        "todate between " & strDTEnc & DateCompStr(strFirstDateM) & strDTEnc & " and " & _
        strDTEnc & DateCompStr(strLastDateM) & strDTEnc & ")"

        VstarDataEnv.cnDJConn.Execute "update " & strRepFile & " set trcd= ' ' where trcd ='4'"
        VstarDataEnv.cnDJConn.Execute "update " & strRepFile & " set trcd= 'Late Cut' where trcd ='6'"
        VstarDataEnv.cnDJConn.Execute "update " & strRepFile & " set trcd= 'Early Cut' where trcd ='7'"
        monReportsMod = True
    Case 15, 16 'Total Lates,Total Earlys
        If Not monTotLtEr Then Exit Function
        monReportsMod = True
    Case 17 'Shift schedule
        If Not monShiftSch Then Exit Function
        monReportsMod = True
    Case 18 'WO on Holiday
        strMon_Trn = ""
        strMon_Trn = Left(typRep.strMonMth, 3) & Right(typRep.strMonYear, 2) & "trn"
        Select Case bytBackEnd
            Case 1  ''SQL SERVER
                strLastDateM = "datename(dw," & strMon_Trn & "." & strKDate & ")"
            Case 2  ''MS-Access
                strLastDateM = "format(" & strMon_Trn & "." & strKDate & ",'dddd')"
            Case 3  ''Oracle
                strLastDateM = "TO_CHAR(" & strMon_Trn & "." & strKDate & ",'Day')"
        End Select
        monReportsMod = True
End Select
Exit Function
Err_P:
    ShowError ("monReportsMod :: " & Me.Caption)
End Function

Private Sub SetVarEmpty()
On Error GoTo Err_P
If strRepFile <> "" Then Call ChkRepFile        '' INITIALIZES REPORT'S ALL GLOBAL VARIABLES
empstr3 = "": sqlStr = "": headGrp = "": strSql = ""
Set Report = Nothing: Unload Report
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
Err_P:
    Select Case Err.Number
        Case 91
        Case Else
            ShowError ("SetVarEmpty :: Reports")
    End Select
End Sub

Private Sub SendThruMail()
On Error GoTo EmailErr                  '' PROCEDURE FOR SENDING REPORTS THROUGH E-MAIL
If pVStar.Use_Mail And EmailSend Then
    If EmailSendOpt = 0 Then 'send to selected employee
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "select email_id,name  from empmst where empcode= " & "'" & EmpId & _
        "'", VstarDataEnv.cnDJConn
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            'start the mapi session
            Call SetMSF1Cap(12)
            ReportSession.SignOn
            With ReportMessage
                .SessionID = ReportSession.SessionID
                .AddressCaption = "Visual Star Address Book"
                .Compose
                .RecipDisplayName = adrsTemp(0)
                .ResolveName
                .MsgSubject = EmailSub
                .MsgNoteText = " "
                .AttachmentPathName = strAttachPath
                .AttachmentPosition = 0
                .AttachmentType = 0
                .Send (0)
            End With
            ReportSession.SignOff
        Else
            MsgBox NewCaptionTxt("40083", adrsC) & adrsTemp(1), vbInformation
        End If
        If adrsTemp.State = 1 Then adrsTemp.Close
    End If
    If EmailSendOpt = 1 Then 'send to each employee
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "select email_id,name  from empmst order by empcode", _
        VstarDataEnv.cnDJConn
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            adrsTemp.MoveFirst
            ReportSession.SignOn
            Do
                If adrsTemp(0) <> "" Then
                    Call SetMSF1Cap(12)
                    With ReportMessage
                        .SessionID = ReportSession.SessionID
                        .AddressCaption = "Visual Star Address Book"
                        .Compose
                        .RecipDisplayName = adrsTemp(0)
                        .ResolveName
                        .MsgSubject = EmailSub
                        .MsgNoteText = " "
                        .AttachmentPathName = strAttachPath
                        .AttachmentPosition = 0
                        .AttachmentType = 0
                        .Send (0)
                    End With
                Else
                    MsgBox NewCaptionTxt("40083", adrsC) & adrsTemp(1), vbExclamation
                End If
                adrsTemp.MoveNext
            Loop Until adrsTemp.EOF
            ReportSession.SignOff
        End If
    End If
    If adrsTemp.State = 1 Then adrsTemp.Close
End If  'vstar.emailid
Exit Sub
EmailErr:
        ShowError ("Send Through Email :: Error" & Me.Caption)
End Sub

Private Sub PutZeros()                  '' SETS ALL TABS OPTION BUTTONS TO FIRST
optWee(0).Value = True
optMon(0).Value = True
optYea(0).Value = True
optmas(0).Value = True
optPer(0).Value = True
optDly(0).Value = True
End Sub

Private Sub FillShiftCombo()        '' Fills Shift ComboBox
On Error GoTo Err_P
Dim strArrTmp() As String, bytTmp As Byte
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select Shift,Shf_In,Shf_Out from Instshft where shift <> '100' Order by Shift", _
VstarDataEnv.cnDJConn, adOpenStatic
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
Err_P:
    ShowError ("FillShiftCombo :: " & Me.Caption)
End Sub

Private Sub FillMainCombo()
On Error GoTo Err_P
Dim strArrTmp() As String, bytTmp As Byte
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select distinct qualf from empmst where qualf is not null Order by qualf", VstarDataEnv.cnDJConn, adOpenStatic
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
Err_P:
    ShowError ("FillMainCombo :: " & Me.Caption)
End Sub

Private Sub ShowGroup(Optional ByVal blnShow As Boolean = True)
On Error GoTo Err_P

optGrpEmp(0).Enabled = blnShow
optGrpDep(1).Enabled = blnShow
optGrpCat(2).Enabled = blnShow
optGrpGrp(3).Enabled = blnShow
optGrpDC(4).Enabled = blnShow
optGrpLoc(5).Enabled = blnShow
optGrpDiv(6).Enabled = blnShow
Exit Sub
Err_P:
ShowError ("ShowGroup :: " & Me.Caption)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then Call ShowF10("40")
If KeyCode = vbKeyF10 Then KeyCode = 0
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

    optGrpEmp(0).Enabled = blnShow
    optGrpDep(1).Enabled = blnShow
    optGrpCat(2).Enabled = blnShow
    optGrpGrp(3).Enabled = blnShow
    optGrpDC(4).Enabled = blnShow
    optGrpLoc(5).Enabled = blnShow
    optGrpDiv(6).Enabled = blnShow

    chkNewP.Enabled = blnShow

End Sub
