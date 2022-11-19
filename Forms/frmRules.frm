VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRules 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLL 
      Caption         =   "Check1"
      Height          =   255
      Left            =   9240
      TabIndex        =   61
      Top             =   4200
      Width           =   2625
   End
   Begin VB.CheckBox chkE 
      Caption         =   "Check1"
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   600
      Width           =   2625
   End
   Begin VB.Frame frEarly 
      Caption         =   " "
      Height          =   3645
      Left            =   4560
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   4305
      Begin VB.Frame frE 
         Caption         =   "asdad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   90
         TabIndex        =   28
         Top             =   1200
         Width           =   4125
         Begin VB.OptionButton optPaidE 
            Caption         =   "Paid days"
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
            Left            =   120
            TabIndex        =   15
            Top             =   330
            Width           =   1380
         End
         Begin VB.OptionButton optLeavesE 
            Caption         =   "Leaves"
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
            Left            =   2025
            TabIndex        =   16
            Top             =   270
            Width           =   975
         End
      End
      Begin MSMask.MaskEdBox txtEarly 
         Height          =   330
         Left            =   3030
         TabIndex        =   14
         Top             =   795
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   582
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
      Begin MSMask.MaskEdBox txtDaysE 
         Height          =   315
         Left            =   810
         TabIndex        =   13
         Top             =   810
         Width           =   570
         _ExtentX        =   1005
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
      Begin MSMask.MaskEdBox txtTotalE 
         Height          =   345
         Left            =   3030
         TabIndex        =   12
         Top             =   300
         Width           =   675
         _ExtentX        =   1191
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
         PromptChar      =   " "
      End
      Begin VB.Label lblTotalE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Early Allowed in a Month"
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
         Top             =   360
         Width           =   2640
      End
      Begin VB.Label lblCutE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cut"
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
         Top             =   840
         Width           =   300
      End
      Begin VB.Label lblDayE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day for Every"
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
         Left            =   1470
         TabIndex        =   26
         Top             =   840
         Width           =   1155
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
         Left            =   3660
         TabIndex        =   27
         Top             =   840
         Width           =   450
      End
      Begin VB.Label lblCapLE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leaves"
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
         Left            =   2190
         TabIndex        =   29
         Top             =   2010
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbl1E 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1st Preference"
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
         Left            =   300
         TabIndex        =   30
         Top             =   2340
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lbl2E 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Preference"
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
         Left            =   300
         TabIndex        =   31
         Top             =   2820
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl3E 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3rd Preference"
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
         Left            =   270
         TabIndex        =   32
         Top             =   3300
         Visible         =   0   'False
         Width           =   1260
      End
      Begin MSForms.ComboBox cbo1E 
         Height          =   375
         Left            =   2190
         TabIndex        =   17
         Top             =   2310
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo2E 
         Height          =   375
         Left            =   2190
         TabIndex        =   18
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo3E 
         Height          =   375
         Left            =   2190
         TabIndex        =   19
         Top             =   3210
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   " "
      Height          =   435
      Left            =   6240
      TabIndex        =   22
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton cmdCanReset 
      Caption         =   " "
      Height          =   435
      Left            =   3360
      TabIndex        =   21
      Top             =   4440
      Width           =   2505
   End
   Begin VB.CommandButton cmdEditSave 
      Caption         =   " "
      Height          =   435
      Left            =   360
      TabIndex        =   20
      Top             =   4440
      Width           =   2505
   End
   Begin VB.CheckBox chkL 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2625
   End
   Begin VB.Frame frLate 
      Caption         =   " "
      Height          =   3645
      Left            =   120
      TabIndex        =   33
      Top             =   600
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Frame frL 
         Caption         =   "asdad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   90
         TabIndex        =   34
         Top             =   1200
         Width           =   4035
         Begin VB.OptionButton optPaidL 
            Caption         =   "Paid Days"
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
            Left            =   120
            TabIndex        =   6
            Top             =   330
            Width           =   1380
         End
         Begin VB.OptionButton optLeavesL 
            Caption         =   "Leaves"
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
            Left            =   1935
            TabIndex        =   7
            Top             =   285
            Width           =   975
         End
      End
      Begin MSMask.MaskEdBox txtLate 
         Height          =   330
         Left            =   3060
         TabIndex        =   5
         Top             =   765
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   582
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
      Begin MSMask.MaskEdBox txtDaysL 
         Height          =   315
         Left            =   810
         TabIndex        =   4
         Top             =   780
         Width           =   570
         _ExtentX        =   1005
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
      Begin MSMask.MaskEdBox txtTotalL 
         Height          =   345
         Left            =   3090
         TabIndex        =   3
         Top             =   300
         Width           =   675
         _ExtentX        =   1191
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
         PromptChar      =   " "
      End
      Begin VB.Label lblTotalL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Late Allowed in a Month"
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
         TabIndex        =   42
         Top             =   360
         Width           =   2565
      End
      Begin VB.Label lblCutL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cut"
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
         TabIndex        =   41
         Top             =   810
         Width           =   300
      End
      Begin VB.Label lblDayL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day for Every"
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
         Left            =   1500
         TabIndex        =   40
         Top             =   810
         Width           =   1155
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
         Left            =   3720
         TabIndex        =   39
         Top             =   810
         Width           =   375
      End
      Begin VB.Label lblCapLL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leaves"
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
         Left            =   2100
         TabIndex        =   38
         Top             =   2010
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbl1L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1st Preference"
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
         Left            =   300
         TabIndex        =   37
         Top             =   2340
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lbl2L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Preference"
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
         Left            =   300
         TabIndex        =   36
         Top             =   2820
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl3L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3rd Preference"
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
         Left            =   300
         TabIndex        =   35
         Top             =   3300
         Visible         =   0   'False
         Width           =   1260
      End
      Begin MSForms.ComboBox cbo1L 
         Height          =   375
         Left            =   2100
         TabIndex        =   8
         Top             =   2310
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo2L 
         Height          =   375
         Left            =   2100
         TabIndex        =   9
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo3L 
         Height          =   375
         Left            =   2100
         TabIndex        =   10
         Top             =   3210
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame frLateL 
      Caption         =   " "
      Height          =   3645
      Left            =   9120
      TabIndex        =   43
      Top             =   4200
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Frame frLL 
         Caption         =   "asdad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   44
         Top             =   1200
         Width           =   4035
         Begin VB.OptionButton optLeavesLL 
            Caption         =   "Leaves"
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
            Left            =   1935
            TabIndex        =   46
            Top             =   285
            Width           =   975
         End
         Begin VB.OptionButton optpaidLL 
            Caption         =   "Paid Days"
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
            Left            =   120
            TabIndex        =   45
            Top             =   330
            Width           =   1380
         End
      End
      Begin MSMask.MaskEdBox txtDL 
         Height          =   330
         Left            =   3060
         TabIndex        =   47
         Top             =   765
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   582
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
      Begin MSMask.MaskEdBox txtCL 
         Height          =   315
         Left            =   810
         TabIndex        =   48
         Top             =   780
         Width           =   570
         _ExtentX        =   1005
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
      Begin MSMask.MaskEdBox txtLL 
         Height          =   345
         Left            =   3090
         TabIndex        =   49
         Top             =   300
         Width           =   675
         _ExtentX        =   1191
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
         PromptChar      =   " "
      End
      Begin MSForms.ComboBox cbo3LL 
         Height          =   375
         Left            =   2100
         TabIndex        =   60
         Top             =   3210
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo2LL 
         Height          =   375
         Left            =   2100
         TabIndex        =   59
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo1LL 
         Height          =   375
         Left            =   2100
         TabIndex        =   58
         Top             =   2310
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lbl3LL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3rd Preference"
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
         Left            =   300
         TabIndex        =   57
         Top             =   3300
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lbl2LL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Preference"
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
         Left            =   300
         TabIndex        =   56
         Top             =   2820
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl1LL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1st Preference"
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
         Left            =   300
         TabIndex        =   55
         Top             =   2340
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leaves"
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
         Left            =   2100
         TabIndex        =   54
         Top             =   2010
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label4 
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
         Left            =   3720
         TabIndex        =   53
         Top             =   810
         Width           =   375
      End
      Begin VB.Label lblDL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day for Every"
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
         Left            =   1560
         TabIndex        =   52
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblCL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cut"
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
         TabIndex        =   51
         Top             =   810
         Width           =   300
      End
      Begin VB.Label lblLL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Late Allowed in a Month"
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
         TabIndex        =   50
         Top             =   360
         Width           =   2565
      End
   End
   Begin VB.Label lblCat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ewrwerwerw"
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
      Left            =   1665
      TabIndex        =   0
      Top             =   60
      Width           =   1245
   End
   Begin MSForms.ComboBox cboCat 
      Height          =   375
      Left            =   3641
      TabIndex        =   1
      Top             =   0
      Width           =   2055
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "3625;661"
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
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
''
Dim adrsC As New ADODB.Recordset

Private Sub RetCaption()
On Error GoTo ERR_P
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '44%'", ConMain, adOpenStatic, adLockReadOnly
Me.Caption = "Late / Early Rule"
'' Main
lblCat.Caption = NewCaptionTxt("44002", adrsC)
'' Late
'' Misc
chkL.Caption = NewCaptionTxt("44003", adrsC)
lblTotalL.Caption = NewCaptionTxt("44004", adrsC)
lblCutL.Caption = NewCaptionTxt("44005", adrsC)
lblDayL.Caption = NewCaptionTxt("44006", adrsC)
lblLate.Caption = NewCaptionTxt("00035", adrsMod)
'' Options
frL.Caption = NewCaptionTxt("44007", adrsC)
optPaidL.Caption = NewCaptionTxt("44008", adrsC)
optLeavesL.Caption = NewCaptionTxt("44009", adrsC)
'' Leaves
lblCapLL.Caption = NewCaptionTxt("44009", adrsC)
lbl1L.Caption = NewCaptionTxt("44010", adrsC)
lbl2L.Caption = NewCaptionTxt("44011", adrsC)
lbl3L.Caption = NewCaptionTxt("44012", adrsC)
'' Early
'' Misc
chkE.Caption = NewCaptionTxt("44013", adrsC)
lblTotalE.Caption = NewCaptionTxt("44014", adrsC)
lblCutE.Caption = NewCaptionTxt("44005", adrsC)
lblDayE.Caption = NewCaptionTxt("44006", adrsC)
lblEarly.Caption = NewCaptionTxt("00037", adrsMod)
'' Options
frE.Caption = NewCaptionTxt("44007", adrsC)
optPaidE.Caption = NewCaptionTxt("44008", adrsC)
optLeavesE.Caption = NewCaptionTxt("44009", adrsC)
'' Leaves
lblCapLE.Caption = NewCaptionTxt("44009", adrsC)
lbl1E.Caption = NewCaptionTxt("44010", adrsC)
lbl2E.Caption = NewCaptionTxt("44011", adrsC)
lbl3E.Caption = NewCaptionTxt("44012", adrsC)
''Lunchlate
chkLL.Caption = "LunchLate"
'Call SetButtonCap
Exit Sub
ERR_P:
    ShowError ("RetCaption :: " & Me.Caption)
End Sub

Private Sub SetButtonCap(Optional bytFlgCap As Byte = 1)    '' Sets Captions to the
Select Case bytFlgCap
    Case 1
        cmdEditSave.Caption = "Update" 'NewCaptionTxt("00005", adrsMod)
        cmdCanReset.Caption = "Reset" 'NewCaptionTxt("44015", adrsC)
        cmdExit.Caption = "Exit" 'NewCaptionTxt("00008", adrsMod)
    Case 2
        cmdEditSave.Caption = "Save" 'NewCaptionTxt("00007", adrsMod)
        cmdCanReset.Caption = "Cancel" 'NewCaptionTxt("00003", adrsMod)
End Select
End Sub

Private Sub cboCat_Change()

Call cboCat_Click
''
End Sub

Private Sub cboCat_Click()
If cboCat.Text = "" Then Exit Sub
Call FillLeaves             '' Fills the Leaves Combo
Call Display                '' Displays All the Details of the Selected Category
End Sub

Private Sub chkE_Click()
If chkE.Value = 1 Then
    frEarly.Visible = True
Else
    frEarly.Visible = False
End If
End Sub

Private Sub chkL_Click()
If chkL.Value = 1 Then
    frLate.Visible = True
Else
    frLate.Visible = False
End If
End Sub

Private Sub chkLL_Click()
If chkLL.Value = 1 Then
    frLateL.Visible = True
Else
    frLateL.Visible = False
End If
End Sub

Private Sub cmdCanReset_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 3
        bytMode = 1
        Call Display
        Call ChangeMode
    Case 1
        If Not DeleteRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        Else
            If cboCat.Text = "" Then Exit Sub
            If MsgBox(NewCaptionTxt("44016", adrsC), vbQuestion + vbYesNo) = vbYes Then
                '' Set All the Rules to Default Values
                '' Update
                ConMain.Execute "Update CatDesc Set LateRule='N'," & _
                "LtInMnth=0.00,LetCut=0.00,EverLet=0.00,DedLet='PD',FstLetPr='',SecLetPr=''," & _
                "TrdLetPr='',EarlRule='N',ErInMnth=0.00,ErlCut=0.00,EverErl=0.00,DedErl='PD'," & _
                "FstErlPr='',SecErlPr='',TrdErlPr='' Where Cat='" & cboCat.Text & "'"
                Call AddActivityLog(lgReset_Action, 1, 8)     '' Reset Activity
                Call AuditInfo("RESET", Me.Caption, "Reset Late/Early Rule Of Category: " & cboCat.Text)
                adrsDept1.Requery
                Call Display
            End If
        End If
End Select
Exit Sub
ERR_P:
    ShowError ("Cancel Reset :: " & Me.Caption)
End Sub

Private Sub cmdEditSave_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        If cboCat.Text = "" Then Exit Sub
        '' Check for Rights
        If Not EditRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        bytMode = 3
        Call ChangeMode
    Case 3      '' Edit Mode
        If Not ValidateModMaster Then Exit Sub  '' If not Valida for Edit
        If Not SaveModMaster Then Exit Sub      '' Save for Edit
        Call SaveModLog                         '' Save the Edit Log
        adrsDept1.Requery
        Call Display
        bytMode = 1
        Call ViewAction
End Select
Exit Sub
ERR_P:
    ShowError ("EditSave :: " & Me.Caption)
    'Resume Next
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()

chkLL.Visible = False

Call SetFormIcon(Me)        '' Sets the Form Icon
Call RetCaption             '' Sets the Captions
Call GetRights              '' Gets the Rights
Call OpenMasterTable        '' Opens the Master Table for Category
Call FillCombo              '' Fills the Category Combo
bytMode = 1                 '' Sets the Mode to the Default Mode
Call ChangeMode             '' Take Action Accordingly
End Sub

Private Sub ChangeMode()
Select Case bytMode
    Case 1          '' View Mode
        Call ViewAction
    Case 3          '' Edit Mde
        Call EditAction
End Select
End Sub

Private Sub ViewAction()
'' Disable Necessary Controls
frLate.Enabled = False          '' Disable Late Frame
frEarly.Enabled = False         '' Disable Early Frame
chkL.Enabled = False            '' Disable Late CheckBox
chkE.Enabled = False            '' Disable Early CheckBox
frLateL.Enabled = False
chkLL.Enabled = False
'' Set Necessary Captions
Call SetButtonCap
End Sub

Private Sub EditAction()
'' Enable Necessary Controls
frLate.Enabled = True       '' Enable Late Frame
frEarly.Enabled = True      '' Enable Early Frame
chkL.Enabled = True         '' Enable Late CheckBox
chkE.Enabled = True         '' Enable Early CheckBox
frLateL.Enabled = True
chkLL.Enabled = True
'' Change Necesary Captions
Call SetButtonCap(2)        '' SetCaptions of the Buttons
If chkL.Value = 1 Then
    txtTotalL.SetFocus          '' Sets the Focus to Total Late Allowed TextBox
Else
    chkL.SetFocus
End If
End Sub

Private Sub optLeavesE_Click()
cbo1E.Visible = True
cbo2E.Visible = True
cbo3E.Visible = True
lbl1E.Visible = True
lbl2E.Visible = True
lbl3E.Visible = True
lblCapLE.Visible = True
End Sub

Private Sub optLeavesL_Click()
cbo1L.Visible = True
cbo2L.Visible = True
cbo3L.Visible = True
lbl1L.Visible = True
lbl2L.Visible = True
lbl3L.Visible = True
lblCapLL.Visible = True
End Sub

Private Sub optPaidE_Click()
cbo1E.Visible = False
cbo2E.Visible = False
cbo3E.Visible = False
lbl1E.Visible = False
lbl2E.Visible = False
lbl3E.Visible = False
lblCapLE.Visible = False
End Sub

Private Sub optPaidL_Click()
cbo1L.Visible = False
cbo2L.Visible = False
cbo3L.Visible = False
lbl1L.Visible = False
lbl2L.Visible = False
lbl3L.Visible = False
lblCapLL.Visible = False
End Sub

Private Sub OpenMasterTable()       '' Opens the Category Master Table
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select Cat," & strKDesc & ",LateRule,LtInMnth,LetCut,EverLet,DedLet,FstLetPr,SecLetPr," & _
"TrdLetPr,EarlRule,ErInMnth,ErlCut,EverErl,DedErl,FstErlPr,SecErlPr,TrdErlPr From CatDesc" & _
" where cat <> '100' Order by Cat", ConMain, adOpenStatic
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Function ValidateModMaster() As Boolean
On Error GoTo ERR_P                 '' Validates the Record Befor Updating it to the
ValidateModMaster = True            '' Table
'' For Late
Call FormatAll
If chkL.Value = 1 Then
    If Right(txtDaysL.Text, 2) <> "00" And Right(txtDaysL.Text, 2) <> "50" Then
        MsgBox NewCaptionTxt("44017", adrsC)
        txtDaysL.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
    If Right(txtLate.Text, 2) <> "00" Then
        MsgBox NewCaptionTxt("44018", adrsC)
        txtLate.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
    If Right(txtTotalL.Text, 2) <> "00" Then
        MsgBox NewCaptionTxt("44018", adrsC)
        txtTotalL.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
    If optLeavesL.Value = True Then
        If cbo1L.Text = "" And cbo2L.Text = "" And cbo3L.Text = "" Then
            MsgBox NewCaptionTxt("44019", adrsC), vbExclamation
            cbo1L.SetFocus
            ValidateModMaster = False
            Exit Function
        End If
        If cbo3L.Text <> "" And cbo2L.Text = "" Then
            MsgBox NewCaptionTxt("44020", adrsC), vbExclamation
            cbo2L.SetFocus
            ValidateModMaster = False
            Exit Function
        End If
        If cbo2L.Text <> "" And cbo1L.Text = "" Then
            MsgBox NewCaptionTxt("44020", adrsC), vbExclamation
            cbo1L.SetFocus
            ValidateModMaster = False
            Exit Function
        End If
    End If
End If
If chkE.Value = 1 Then
     If Right(txtDaysE.Text, 2) <> "00" And Right(txtDaysE.Text, 2) <> "50" Then
        MsgBox NewCaptionTxt("44017", adrsC)
        txtDaysE.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
    If Right(txtEarly.Text, 2) <> "00" Then
        MsgBox NewCaptionTxt("44018", adrsC)
        txtEarly.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
    If Right(txtTotalE.Text, 2) <> "00" Then
        MsgBox NewCaptionTxt("44018", adrsC)
        txtTotalE.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
    If optLeavesE.Value = True Then
        If cbo1E.Text = "" And cbo2E.Text = "" And cbo3E.Text = "" Then
            MsgBox NewCaptionTxt("44019", adrsC), vbExclamation
            cbo1E.SetFocus
            ValidateModMaster = False
            Exit Function
        End If
        If cbo3E.Text <> "" And cbo2E.Text = "" Then
            MsgBox NewCaptionTxt("44020", adrsC), vbExclamation
            cbo2E.SetFocus
            ValidateModMaster = False
            Exit Function
        End If
        If cbo2E.Text <> "" And cbo1E.Text = "" Then
            MsgBox NewCaptionTxt("44021", adrsC), vbExclamation
            cbo1E.SetFocus
            ValidateModMaster = False
            Exit Function
        End If
    End If
End If
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
    ValidateModMaster = False
End Function

Private Sub FillCombo()     '' Fills the Category Combo
On Error GoTo ERR_P
Dim strCatArr() As String   '' Array for Categories
Dim bytCntTmp As Byte       '' Count of Records

If adrsDept1.EOF Then Exit Sub
cboCat.clear                '' Clear the ComboBox
adrsDept1.MoveFirst         '' Move to the First Record
ReDim strCatArr(adrsDept1.RecordCount - 1, 1)       '' Redimension the Array

For bytCntTmp = 0 To adrsDept1.RecordCount - 1
    strCatArr(bytCntTmp, 0) = adrsDept1("Cat")
    strCatArr(bytCntTmp, 1) = adrsDept1("Desc")
    adrsDept1.MoveNext
Next
adrsDept1.MoveFirst         '' Move to the First Record

cboCat.List = strCatArr     '' Assign the Array to the ComboBox
Exit Sub
ERR_P:
    ShowError ("FillCombo ::" & Me.Caption)
End Sub

Private Sub FillLeaves()            '' Fills the Combo with the Leaves Alloted to the
On Error GoTo ERR_P                 '' Employee
Dim strArrLeaves() As String        '' For the Leaves
Dim bytCntTmp As Byte               '' For the Count of Leaves
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open "Select Distinct(lvcode),Leave  from LeavDesc where Lvcode not in (" & "'" & _
pVStar.PrsCode & "'" & "," & "'" & pVStar.AbsCode & "'" & "," & "'" & pVStar.WosCode & _
"'" & "," & "'" & pVStar.HlsCode & "'" & ")" & " and type='Y'" & " and cat=" & "'" & _
cboCat.Text & "'", ConMain, adOpenStatic
If (adrsPaid.EOF And adrsPaid.BOF) Then
    Exit Sub
Else
    bytCntTmp = adrsPaid.RecordCount - 1
    ReDim strArrLeaves(bytCntTmp + 1, 1)
    For bytCntTmp = 0 To adrsPaid.RecordCount - 1
        strArrLeaves(bytCntTmp, 0) = adrsPaid("LvCode")
        strArrLeaves(bytCntTmp, 1) = adrsPaid("Leave")
        adrsPaid.MoveNext
    Next
    strArrLeaves(bytCntTmp, 0) = ""
    strArrLeaves(bytCntTmp, 1) = ""
End If
'' Clear All the Combos
cbo1L.clear
cbo2L.clear
cbo3L.clear
cbo1E.clear
cbo2E.clear
cbo3E.clear
cbo1LL.clear
cbo2LL.clear
cbo3LL.clear

'' Fill All the Combos
cbo1L.List = strArrLeaves
cbo2L.List = strArrLeaves
cbo3L.List = strArrLeaves
cbo1E.List = strArrLeaves
cbo2E.List = strArrLeaves
cbo3E.List = strArrLeaves
cbo1LL.List = strArrLeaves
cbo2LL.List = strArrLeaves
cbo3LL.List = strArrLeaves
Exit Sub
ERR_P:
    ShowError ("FillLeaves :: " & Me.Caption)
End Sub

Private Sub Display()
On Error GoTo ERR_P
adrsDept1.MoveFirst
adrsDept1.Find "Cat='" & cboCat.Text & "'"
If adrsDept1.EOF Then Exit Sub
'' Late
'' Others
chkL.Value = IIf(IsNull(adrsDept1("LateRule")) = True Or adrsDept1("LateRule") = "N", 0, 1)

txtTotalL.Text = IIf(IsNull(adrsDept1("LtInMnth")), "0.00", _
                Format(adrsDept1("LtInMnth"), "0.00"))
txtDaysL.Text = IIf(IsNull(adrsDept1("LetCut")), "0.00", _
                Format(adrsDept1("LetCut"), "0.00"))
txtLate.Text = IIf(IsNull(adrsDept1("EverLet")), "0.00", _
                Format(adrsDept1("EverLet"), "0.00"))
optPaidL.Value = IIf(IsNull(adrsDept1("DedLet")) Or adrsDept1("DedLet") = "PD", True, False)
optLeavesL.Value = IIf(adrsDept1("DedLet") = "LV", True, False)
'' Leaves
cbo1L.Value = GetValidLeave(IIf(IsNull(adrsDept1("FstLetPr")), "", adrsDept1("FstLetPr")))
cbo2L.Value = GetValidLeave(IIf(IsNull(adrsDept1("SecLetPr")), "", adrsDept1("SecLetPr")))
cbo3L.Value = GetValidLeave(IIf(IsNull(adrsDept1("TrdLetPr")), "", adrsDept1("TrdLetPr")))
'' Early
'' Others
chkE.Value = IIf(IsNull(adrsDept1("EarlRule")) = True Or adrsDept1("EarlRule") = "N", 0, 1)
txtTotalE.Text = IIf(IsNull(adrsDept1("ErInMnth")), "0.00", _
                Format(adrsDept1("ErInMnth"), "0.00"))
txtDaysE.Text = IIf(IsNull(adrsDept1("ErlCut")), "0.00", _
                Format(adrsDept1("ErlCut"), "0.00"))
txtEarly.Text = IIf(IsNull(adrsDept1("EverErl")), "0.00", _
                Format(adrsDept1("EverErl"), "0.00"))
optPaidE.Value = IIf(IsNull(adrsDept1("DedErl")) Or adrsDept1("DedErl") = "PD", True, False)
optLeavesE.Value = IIf(adrsDept1("DedErl") = "LV", True, False)
'' Leaves
cbo1E.Value = GetValidLeave(IIf(IsNull(adrsDept1("FstErlPr")), "", adrsDept1("FstErlPr")))
cbo2E.Value = GetValidLeave(IIf(IsNull(adrsDept1("SecErlPr")), "", adrsDept1("SecErlPr")))
cbo3E.Value = GetValidLeave(IIf(IsNull(adrsDept1("TrdErlPr")), "", adrsDept1("TrdErlPr")))
''Late Lunch
'chkLL.Value = IIf(IsNull(adrsDept1("LateLL")) = True Or adrsDept1("LateLL") = "N", 0, 1)
Exit Sub
ERR_P:
    ShowError ("Display :: " & Me.Caption)
End Sub

Private Function GetValidLeave(ByVal strLeave As String) As String
On Error GoTo ERR_P             '' Checks if the Leave is Valid to be Displayed inth ComboBox
If IsNull(strLeave) Or strLeave = "" Then
GetValidLeave = ""
Exit Function
End If
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open "Select LvCode from LeavDesc Where LvCode='" & strLeave & "'", _
ConMain, adOpenStatic
If (adrsPaid.EOF And adrsPaid.BOF) Then
    GetValidLeave = ""
Else
    GetValidLeave = adrsPaid("LvCode")
End If
Exit Function
ERR_P:
    ShowError ("GetValidLeave :: " & Me.Caption)
    GetValidLeave = ""
End Function

Private Function SaveModMaster() As Boolean     '' Saves the Data in the Edit Mode
On Error GoTo ERR_P
Dim strLate As String, strEarly As String
If optPaidL.Value = True Then
    strLate = "PD"
Else
    strLate = "LV"
End If
If optPaidE.Value = True Then
    strEarly = "PD"
Else
    strEarly = "LV"
End If
SaveModMaster = True
'' Update
ConMain.Execute "Update CatDesc Set LateRule='" & IIf(chkL.Value = 1, "Y", "N") & _
"',LtInMnth=" & txtTotalL.Text & ",LetCut=" & txtDaysL.Text & ",EverLet=" & txtLate.Text & _
",DedLet='" & strLate & "',FstLetPr='" & cbo1L.Value & "',SecLetPr='" & cbo2L.Value & "'," & _
"TrdLetPr='" & cbo3L.Value & "',EarlRule='" & IIf(chkE.Value = 1, "Y", "N") & "',ErInMnth=" & _
txtTotalE.Text & ",ErlCut=" & txtDaysE.Text & ",EverErl=" & txtEarly.Text & ",DedErl='" & _
strEarly & "',FstErlPr='" & cbo1E.Value & "',SecErlPr='" & cbo2E.Value & "',TrdErlPr='" & _
cbo3E.Value & "' Where Cat='" & cboCat.Text & "'"
Exit Function
ERR_P:
    ShowError ("SaveModMaster :: " & Me.Caption)
    SaveModMaster = False
End Function

Private Sub FormatAll()     '' Formats All the Numerical data to the 0.00 Format
txtTotalL.Text = IIf(txtTotalL.Text = "", "0.00", Format(txtTotalL.Text, "0.00"))
txtDaysL.Text = IIf(txtDaysL.Text = "", "0.00", Format(txtDaysL.Text, "0.00"))
txtLate.Text = IIf(txtLate.Text = "", "0.00", Format(txtLate.Text, "0.00"))
txtTotalE.Text = IIf(txtTotalE.Text = "", "0.00", Format(txtTotalE.Text, "0.00"))
txtDaysE.Text = IIf(txtDaysE.Text = "", "0.00", Format(txtDaysE.Text, "0.00"))
txtEarly.Text = IIf(txtEarly.Text = "", "0.00", Format(txtEarly.Text, "0.00"))
End Sub

Private Sub GetRights()         '' Gets and Sets the Particular Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 1)
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

Private Sub txtDaysE_GotFocus()
    Call GF(txtDaysE)
End Sub

Private Sub txtDaysE_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = keycheck(KeyAscii, txtDaysE)
End Select
End Sub

Private Sub txtDaysL_GotFocus()
    Call GF(txtDaysL)
End Sub

Private Sub txtDaysL_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = keycheck(KeyAscii, txtDaysL)
End Select
End Sub

Private Sub txtEarly_GotFocus()
    Call GF(txtEarly)
End Sub

Private Sub txtEarly_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = keycheck(KeyAscii, txtEarly)
End Select
End Sub

Private Sub txtLate_GotFocus()
    Call GF(txtLate)
End Sub

Private Sub txtLate_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = keycheck(KeyAscii, txtLate)
End Select
End Sub

Private Sub txtTotalE_GotFocus()
    Call GF(txtTotalE)
End Sub

Private Sub txtTotalE_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = keycheck(KeyAscii, txtTotalE)
End Select
End Sub

Private Sub txtTotalL_GotFocus()
    Call GF(txtTotalL)
End Sub

Private Sub txtTotalL_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = keycheck(KeyAscii, txtTotalL)
End Select
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 8)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edit Late/Early Rule Of Category: " & cboCat.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub
