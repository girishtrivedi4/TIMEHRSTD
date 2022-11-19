VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Schedule"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frDet 
      Height          =   2805
      Left            =   0
      TabIndex        =   13
      Top             =   930
      Width           =   6825
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   34
         Left            =   5700
         Locked          =   -1  'True
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   2490
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   34
         Left            =   6150
         TabIndex        =   95
         Top             =   2490
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   33
         Left            =   5700
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   2160
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   33
         Left            =   6150
         TabIndex        =   93
         Top             =   2160
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   32
         Left            =   5700
         Locked          =   -1  'True
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1830
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   32
         Left            =   6150
         TabIndex        =   91
         Top             =   1830
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   31
         Left            =   5700
         Locked          =   -1  'True
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   1500
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   31
         Left            =   6150
         TabIndex        =   89
         Top             =   1500
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   30
         Left            =   5700
         Locked          =   -1  'True
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1170
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   30
         Left            =   6150
         TabIndex        =   87
         Top             =   1170
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   29
         Left            =   5700
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   840
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   29
         Left            =   6150
         TabIndex        =   85
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Index           =   28
         Left            =   5700
         Locked          =   -1  'True
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   28
         Left            =   6150
         TabIndex        =   83
         Top             =   480
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   27
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   2490
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   27
         Left            =   5010
         TabIndex        =   81
         Top             =   2490
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   26
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   2160
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   26
         Left            =   5010
         TabIndex        =   79
         Top             =   2160
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   25
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   1830
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   25
         Left            =   5010
         TabIndex        =   77
         Top             =   1830
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   24
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1500
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   24
         Left            =   5010
         TabIndex        =   75
         Top             =   1500
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   23
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1170
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   23
         Left            =   5010
         TabIndex        =   73
         Top             =   1170
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   22
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   840
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   22
         Left            =   5010
         TabIndex        =   71
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Index           =   21
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   21
         Left            =   5010
         TabIndex        =   69
         Top             =   480
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   2490
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   20
         Left            =   3900
         TabIndex        =   67
         Top             =   2490
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   2160
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   19
         Left            =   3900
         TabIndex        =   65
         Top             =   2160
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1830
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   18
         Left            =   3900
         TabIndex        =   63
         Top             =   1830
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1500
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   17
         Left            =   3900
         TabIndex        =   61
         Top             =   1500
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1170
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   16
         Left            =   3900
         TabIndex        =   59
         Top             =   1170
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   840
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   15
         Left            =   3900
         TabIndex        =   57
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Index           =   14
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   14
         Left            =   3900
         TabIndex        =   55
         Top             =   480
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2490
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   13
         Left            =   2760
         TabIndex        =   53
         Top             =   2490
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2160
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   2760
         TabIndex        =   51
         Top             =   2160
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1830
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   2760
         TabIndex        =   49
         Top             =   1830
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1500
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   2760
         TabIndex        =   47
         Top             =   1500
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1170
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   2760
         TabIndex        =   45
         Top             =   1170
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   840
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   2760
         TabIndex        =   43
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   2760
         TabIndex        =   41
         Top             =   480
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2490
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1650
         TabIndex        =   39
         Top             =   2490
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2160
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1650
         TabIndex        =   37
         Top             =   2160
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1830
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1650
         TabIndex        =   35
         Top             =   1830
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1500
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1650
         TabIndex        =   33
         Top             =   1500
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1170
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1650
         TabIndex        =   31
         Top             =   1170
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   840
         Width           =   405
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1650
         TabIndex        =   29
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1650
         TabIndex        =   27
         Top             =   480
         Width           =   555
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   480
         Width           =   405
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   4
         X1              =   6750
         X2              =   6750
         Y1              =   150
         Y2              =   2850
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   1200
         X2              =   6720
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   3
         X1              =   5640
         X2              =   5640
         Y1              =   180
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   2
         X1              =   4500
         X2              =   4500
         Y1              =   180
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   3390
         X2              =   3390
         Y1              =   150
         Y2              =   2760
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   2250
         X2              =   2250
         Y1              =   150
         Y2              =   2850
      End
      Begin VB.Label lblWDay 
         BackColor       =   &H80000012&
         Caption         =   "SAT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   6
         Left            =   330
         TabIndex        =   25
         Top             =   2520
         Width           =   705
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   6
         Left            =   120
         Top             =   2490
         Width           =   975
      End
      Begin VB.Label lblWDay 
         BackColor       =   &H80000012&
         Caption         =   "FRI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   5
         Left            =   330
         TabIndex        =   24
         Top             =   2160
         Width           =   705
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   5
         Left            =   120
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblWDay 
         BackColor       =   &H80000012&
         Caption         =   "THU"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   4
         Left            =   330
         TabIndex        =   23
         Top             =   1830
         Width           =   705
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   4
         Left            =   120
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label lblWDay 
         BackColor       =   &H80000012&
         Caption         =   "WED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   3
         Left            =   330
         TabIndex        =   22
         Top             =   1530
         Width           =   705
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   3
         Left            =   120
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label lblWDay 
         BackColor       =   &H80000012&
         Caption         =   " TUE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   21
         Top             =   1170
         Width           =   705
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   2
         Left            =   120
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label lblWDay 
         BackColor       =   &H80000012&
         Caption         =   "MON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   20
         Top             =   840
         Width           =   705
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   1
         Left            =   120
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblWDay 
         BackColor       =   &H80000012&
         Caption         =   "SUN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   19
         Top             =   540
         Width           =   705
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   0
         Left            =   120
         Top             =   510
         Width           =   975
      End
      Begin VB.Label lblWeek1 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1st Week"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   14
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label lblWeek5 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5th Week"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5640
         TabIndex        =   18
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label lblWeek4 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4th Week"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4530
         TabIndex        =   17
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label lblWeek3 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3rd Week"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3420
         TabIndex        =   16
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label lblWeek2 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2nd Week"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2310
         TabIndex        =   15
         Top             =   150
         Width           =   1095
      End
   End
   Begin VB.Frame frEmp 
      Height          =   975
      Left            =   0
      TabIndex        =   8
      Top             =   -60
      Width           =   6825
      Begin MSForms.ComboBox cboYear 
         Height          =   300
         Left            =   5160
         TabIndex        =   3
         Top             =   607
         Width           =   975
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "1720;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboMonth 
         Height          =   300
         Left            =   5160
         TabIndex        =   2
         Top             =   187
         Width           =   1335
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2355;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   990
      End
      Begin MSForms.ComboBox cboDept 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   180
         Width           =   2835
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "5001;556"
         TextColumn      =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4680
         TabIndex        =   12
         Top             =   660
         Width           =   375
      End
      Begin VB.Label lblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4560
         TabIndex        =   11
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   10
         Top             =   652
         Width           =   1320
      End
      Begin MSForms.ComboBox cboCode 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   2835
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "5001;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CommandButton cmdExitCan 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   405
      Left            =   5070
      TabIndex        =   7
      Top             =   3780
      Width           =   1785
   End
   Begin VB.CommandButton cmdEditSave 
      Caption         =   "Edit"
      Height          =   405
      Left            =   3300
      TabIndex        =   6
      Top             =   3780
      Width           =   1785
   End
   Begin VB.CommandButton cmdMaster 
      Caption         =   "Master"
      Height          =   405
      Left            =   1650
      TabIndex        =   5
      Top             =   3780
      Width           =   1665
   End
   Begin VB.CommandButton cmdPeriod 
      Caption         =   "Period"
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   3780
      Width           =   1665
   End
End
Attribute VB_Name = "frmSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Private blnShiftFound As Boolean
'Master shift then with schmst unload create the shift file
Private MasterRights As Boolean
Dim adrsC As New ADODB.Recordset

Private Sub cboCode_Click()
On Error GoTo ERR_P
If cboCode.Text = "" Then Exit Sub
Call Display
Exit Sub
ERR_P:
    ShowError ("Employee : " & Me.Caption)
End Sub

Private Sub cboDept_Click()
On Error GoTo ERR_P
Dim bytTmp As Byte
If cboDept.ListIndex < 0 Then Exit Sub
If cboDept.Text = "ALL" Then
    Call ComboFill(cboCode, 12, 2)
Else
    Call ComboFill(cboCode, 12, 2, cboDept.List(cboDept.ListIndex, 1))
End If
For bytTmp = 0 To txtShift.UBound
    txtShift(bytTmp).Text = ""
Next
Exit Sub
ERR_P:
    ShowError ("Department::" & Me.Caption)
End Sub

Private Sub cboMonth_Change()

'   If cbodept.Text <> "" Then
'   Call ComboFill(cbocode, 12, 2, cbodept.List(cbodept.ListIndex, 0))
'   End If

End Sub

Private Sub cboMonth_Click()
'Call ComboFill(cboCode, 12, 2)
Call ShowDays
End Sub

Private Sub cboYear_Change()

    If cboDept.Text <> "" Then
'    Call ComboFill(cbocode, 12, 2, cbodept.List(cbodept.ListIndex, 1))
    End If

End Sub

Private Sub cboYear_Click()
'Call ComboFill(cboCode, 12, 2)
Call ShowDays
End Sub

Private Sub cmdEditSave_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1                  '' View Mode
        '' Check For Rights
        If blnShiftFound = True Then
            bytMode = 2
            Call ChangeMode
        End If
    Case 2                  '' Edit Mode
        If Not SaveModMaster Then Exit Sub
        Call AuditInfo("UPDATE", Me.Caption, "Edit Shift Schdule of employee " & cboCode.Text & " for " & cboMonth.Text & " " & cboYear.Text)
        Call SaveModLog                         '' Save the Edit Log
        bytMode = 1
        Call ChangeMode
End Select
Exit Sub
ERR_P:
    ShowError ("Edit Save :: " & Me.Caption)
End Sub

Private Sub cmdExitCan_Click()
Select Case bytMode
    Case 1
        Unload Me
    Case 2
        bytMode = 1
        Call ChangeMode
        Call Display
End Select
End Sub

Private Sub cmdMaster_Click()
On Error GoTo ERR_P
If blnShiftFound = False Then Exit Sub
bytShfMode = 3
If cboCode.ListIndex = -1 Then
    MsgBox "Select Proper Employee Code"
    cboCode.SetFocus
    Exit Sub
End If
If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "Select EmpCode,Name,Cat,STyp,F_Shf,SCode," & strKOff & ",Off2,WO_1_3,WO_2_4,Shf_Date," & _
"JoinDate,LeavDate,WOHLAction,Action3Shift,AutoForPunch,ActionBlank, Location From EmpMst Where EmpCode='" & cboCode.List(cboCode.ListIndex, 1) & "'", ConMain
Shft.Empcode = adrsEmp("EmpCode")
Shft.startdate = DateDisp(adrsEmp("Shf_Date"))
strRotPass = adrsEmp("Name")
If adrsEmp("STyp") = "F" Then
    Shft.ShiftType = "F"
    Shft.ShiftCode = IIf(IsNull(adrsEmp("F_Shf")), "", adrsEmp("F_Shf"))
Else
    Shft.ShiftType = "R"
    Shft.ShiftCode = IIf(IsNull(adrsEmp("SCode")), "", adrsEmp("SCode"))
End If
Shft.WO = IIf(IsNull(adrsEmp("Off")), "", adrsEmp("Off"))
Shft.WO1 = IIf(IsNull(adrsEmp("Off2")), "", adrsEmp("Off2"))
Shft.WO2 = IIf(IsNull(adrsEmp("WO_1_3")), "", adrsEmp("WO_1_3"))
Shft.WO3 = IIf(IsNull(adrsEmp("WO_2_4")), "", adrsEmp("WO_2_4"))
'' For Details regarding Daily Processing
Shft.WOHLAction = IIf(IsNull(adrsEmp("WOHLAction")), 0, adrsEmp("WOHLAction"))
Shft.Action3Shift = IIf(IsNull(adrsEmp("Action3Shift")), "", adrsEmp("Action3Shift"))
Shft.AutoOnPunch = IIf(adrsEmp("AutoForPunch") = 1, True, False)
Shft.ActionBlank = IIf(IsNull(adrsEmp("ActionBlank")), "", adrsEmp("ActionBlank"))
EmailSend = False
frmEmpShift.Show vbModal
If EmailSend = True Then Call FillSchShift
Exit Sub
ERR_P:
    ShowError ("Master :: " & Me.Caption)
End Sub

Private Sub cmdPeriod_Click()
On Error GoTo ERR_P
If cboDept.Text = "" Then Exit Sub
strDjFileN = cboCode.Text & ":" & cboCode.Text & ":" & cboMonth.Text & ":" & _
           cboYear.Text & ":" & typSENum.bytEnd & ":"
           
bytLstInd = cboDept.ListIndex
If cboCode.ListCount < 1 Then
    MsgBox "No Employee In This Department", vbExclamation
    Exit Sub
End If

frmCP.Show vbModal
If bytShfMode = 8 And cboMonth.Text <> "" And cboYear.Text <> "" Then Call Display
Exit Sub
ERR_P:
    ShowError ("Period :: " & Me.Caption)
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)            '' Sets the Form Icon
Call RetCaptions                '' Gets and Sets the Form Captions
Call GetRights                  '' Gets and Sets the Rights for the Form
Call FillCombos                 '' Fills the Form Combos
Call LoadSpecifics              '' Action to be Taken when the Form is Getting Loaded
End Sub

Private Sub RetCaptions()
On Error GoTo ERR_P
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '45%'", ConMain, adOpenStatic, adLockReadOnly
'Me.Caption = NewCaptionTxt("45001", adrsC)         '' Forms Captions
'' Employee Frame
Call SetCritLabel(lblDept)
lblCode.Caption = NewCaptionTxt("00061", adrsMod)    '' Employee Code
'lblMonth.Caption = NewCaptionTxt("00028", adrsMod)    '' Month
'lblYear.Caption = NewCaptionTxt("00029", adrsMod)      '' Year
lblWeek1.Caption = NewCaptionTxt("45002", adrsC)    '' 1st Week
lblWeek2.Caption = NewCaptionTxt("45003", adrsC)    '' 2nd Week
lblWeek3.Caption = NewCaptionTxt("45004", adrsC)    '' 3rd Week
lblWeek4.Caption = NewCaptionTxt("45005", adrsC)    '' 4th Week
lblWeek5.Caption = NewCaptionTxt("45006", adrsC)    '' 5th Week
Exit Sub
ERR_P:
End Sub

Private Sub SetButtonCap(Optional bytFlgCap As Byte = 1)    '' Sets Captions to the Main
If bytFlgCap = 1 Then                                       '' Buttons
    cmdPeriod.Caption = "Period"
    cmdMaster.Caption = "Master"
    cmdEditSave.Caption = "Update"
    cmdExitCan.Caption = "Exit"
    cmdExitCan.Cancel = True
Else
    cmdEditSave.Caption = "Save"
    cmdExitCan.Caption = "Cancel"
    cmdExitCan.Cancel = False
End If
End Sub

Private Sub LoadSpecifics()
bytShfMode = 0              '' Set the Global Mode to No Mode.
Call SetWeekCaps            '' Set the WeekOff Captions
Call ShowDays               '' Set the Days in the TextBox
bytMode = 1                 '' Normal Mode
Call ChangeMode             '' Take Action According to the Mode
blnShiftFound = False       '' Set the Shift Found to False
End Sub

Private Sub SetWeekCaps()
On Error GoTo ERR_P
Dim intStartID As Long, bytElm As Byte, bytRecStart As Byte
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "select weekfrom from install ", ConMain
'' Get the Start Week day Number
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    Select Case adrsTemp("WeekFrom")
        Case Is = Null
            bytRecStart = 0
        Case Is = Empty
            bytRecStart = 0
        Case ""
            bytRecStart = 0
        Case Else
            bytRecStart = adrsTemp("WeekFrom")
    End Select
Else
    bytRecStart = 0
End If
If bytRecStart = 1 Then bytRecStart = 8     '' If Sunday
intStartID = 45009 + bytRecStart - 2        '' Get the Start ID **
For bytElm = 0 To (6 - bytRecStart + 2)     '' Loop from Start Id Till Sunday
    lblWDay(bytElm).Caption = NewCaptionTxt(CStr(intStartID), adrsC)
    intStartID = intStartID + 1             '' Goto Next ID
Next
intStartID = 45009                          '' Startfrom the Start ID
For bytElm = bytElm To 6                    '' Loop Remaining of the Week Days
    lblWDay(bytElm).Caption = NewCaptionTxt(CStr(intStartID), adrsC)
    intStartID = intStartID + 1             '' Goto Next ID
Next
Exit Sub
ERR_P:
    ShowError ("SetWeekCaps :: " & Me.Caption)
    
End Sub

Private Sub ShowDays()
On Error GoTo ERR_P
If cboMonth.Text = "" Then Exit Sub
If cboYear.Text = "" Then Exit Sub
Dim bytTmp As Byte, strTmp As String
'' Clear All Elements
For bytTmp = 0 To txtNum.UBound
    txtNum(bytTmp).Text = ""
Next
bytTmp = 0
'' Get the Start and the End Day Numbers
Call GetSENums(cboMonth.Text, cboYear.Text)
'' Get Start Week
strTmp = WeekdayName(WeekDay(GetDateOfDay(1, cboMonth.Text, cboYear.Text), vbUseSystemDayOfWeek))
For bytTmp = 0 To 6
    If UCase(Left(lblWDay(bytTmp).Caption, 2)) = UCase(Left(strTmp, 2)) Then Exit For
Next
'' Put Value from the Start of the Month to End of the Month
For typSENum.bytStart = 1 To typSENum.bytEnd
    If bytTmp <= txtNum.UBound Then
        txtNum(bytTmp).Text = typSENum.bytStart
        bytTmp = bytTmp + 1
    Else
        bytTmp = bytTmp + 1
        Exit For
    End If
Next
'' If All the Days Still Dont Fit in the Limit Start from Element 0
If bytTmp = 36 Then
    If Val(txtNum(34).Text) < typSENum.bytEnd Then
        bytTmp = 0
        For typSENum.bytStart = Val(txtNum(34).Text) + 1 To typSENum.bytEnd
            If typSENum.bytStart <= typSENum.bytEnd Then
                txtNum(bytTmp).Text = typSENum.bytStart
                bytTmp = bytTmp + 1
            End If
        Next
    End If
End If
If cboCode.Text = "" Then Exit Sub
Call Display
Exit Sub
ERR_P:
    ShowError ("ShowDays :: " & Me.Caption)
End Sub

Private Function GetDateOfDay(ByVal bytDay As Byte, ByVal strMonth As String, _
strYear As String) As String        '' Function to make Date
Select Case bytDateF
    Case 1      '' American (MM/DD/YY)
        GetDateOfDay = Format(MonthNumber(strMonth), "00") & "/" & Format(bytDay, "00") & _
        "/" & strYear
    Case 2      '' British  (DD/MM/YY)
        GetDateOfDay = Format(bytDay, "00") & "/" & Format(MonthNumber(strMonth), "00") & _
        "/" & strYear
End Select
End Function

Private Sub GetRights()
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(4, 5, 2, 1)
If strTmp = "1" Then
    MasterRights = True
    cmdMaster.Enabled = True
    cmdPeriod.Enabled = True
    cmdEditSave.Enabled = True
Else
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    MasterRights = False
    cmdMaster.Enabled = False
    cmdPeriod.Enabled = False
    cmdEditSave.Enabled = False
End If
Exit Sub
ERR_P:
    ShowError ("GetRights :: " & Me.Caption)
        cmdMaster.Enabled = False
        cmdPeriod.Enabled = False
        cmdEditSave.Enabled = False
End Sub

Private Sub FillCombos()
On Error GoTo ERR_P
Dim intTmp As Integer
cboMonth.clear
For intTmp = 1 To 12
    cboMonth.AddItem Choose(intTmp, "January", "February", "March", "April", "May", "June", _
    "July", "August", "September", "October", "November", "December")
Next
cboYear.clear
For intTmp = 1996 To 2097
    cboYear.AddItem CStr(intTmp)
Next

cboMonth.Text = MonthName(Month(Date))
cboYear.Text = pVStar.YearSel
Call SetCritCombos(cboDept)
If strCurrentUserType <> HOD Then cboDept.Text = "ALL"
Exit Sub
ERR_P:
    ShowError ("FillCombos :: " & Me.Caption)
End Sub

Private Sub Display()
On Error GoTo ERR_P
blnShiftFound = False
Dim bytTmp As Byte, strTmp As String
'' Clear the Old Shifts
For bytTmp = 0 To txtShift.UBound
    txtShift(bytTmp).Text = ""
Next
'' Check if File Exists
If Not FindTable(Left(cboMonth.Text, 3) & Right(cboYear.Text, 2) & "Shf") Then
    MsgBox NewCaptionTxt("45016", adrsC) & " " & cboMonth.Text & " " & cboYear.Text, vbExclamation
    Exit Sub
End If

'' Check if Record Exists
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "Select * from " & Left(cboMonth.Text, 3) & Right(cboYear.Text, 2) & "Shf" & _
" Where EmpCode='" & cboCode.List(cboCode.ListIndex, 1) & "'", ConMain, adOpenStatic
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    '' Display Shift
    For bytTmp = 0 To txtNum.UBound
        If txtNum(bytTmp).Text <> "" Then
            strTmp = "D" & txtNum(bytTmp).Text
            txtShift(bytTmp).Text = IIf(IsNull(adrsTemp(strTmp)), "", adrsTemp(strTmp))
        End If
    Next
    blnShiftFound = True
Else
    MsgBox NewCaptionTxt("45017", adrsC) & cboCode.Text & NewCaptionTxt("45018", adrsC), vbExclamation
    Exit Sub
End If
Exit Sub
ERR_P:
    ShowError ("Display :: " & Me.Caption)
    blnShiftFound = False
End Sub

Private Sub ChangeMode()
Select Case bytMode
    Case 1              '' View / Normal Mode
        Call ViewAction
    Case 2              '' Edit Mode
        Call EditAction
End Select
End Sub

Private Sub ViewAction()
On Error GoTo ERR_P
Dim bytTmp As Byte
'' Enable Necessary Controls
frEmp.Enabled = True
If MasterRights = True Then cmdPeriod.Enabled = True
If MasterRights = True Then cmdMaster.Enabled = True
'' Disable Necessary Controls
For bytTmp = 0 To txtShift.UBound
    txtShift(bytTmp).Enabled = False
Next
'' Set Necessary Captions
Call SetButtonCap
Exit Sub
ERR_P:
    ShowError ("ViewAction :: " & Me.Caption)
End Sub

Private Sub EditAction()
On Error GoTo ERR_P
Dim bytTmp As Byte
'' Enable Necessary Controls
For bytTmp = 0 To txtNum.UBound
    If txtNum(bytTmp).Text <> "" Then txtShift(bytTmp).Enabled = True
Next
'' Disable Necessary Controls
frEmp.Enabled = False
cmdPeriod.Enabled = False
cmdMaster.Enabled = False
'' Set Necessary Captions
Call SetButtonCap(2)
Exit Sub
ERR_P:
    ShowError ("EditAction :: " & Me.Caption)
End Sub

Private Sub txtShift_Click(Index As Integer)
bytShfMode = 6
frmSingleS.Show vbModal
If bytShfMode = 9 Then
    txtShift(Index).Text = strDjFileN
End If
bytShfMode = 1
End Sub

Private Sub txtShift_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    bytShfMode = 6
    frmSingleS.Show vbModal
    If bytShfMode = 9 Then
        txtShift(Index).Text = strDjFileN
    End If
    bytShfMode = 1
Else
    KeyAscii = 0
End If
End Sub

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
Dim bytTmp As Byte, strTmp As String
SaveModMaster = True
strTmp = ""
For bytTmp = 0 To txtNum.UBound
    If txtNum(bytTmp).Text <> "" Then
        strTmp = strTmp & "D" & txtNum(bytTmp).Text & "='" & txtShift(bytTmp).Text & "',"
    End If
Next
If Right(strTmp, 1) = "," Then strTmp = Left(strTmp, Len(strTmp) - 1)
ConMain.Execute "Update " & Left(cboMonth.Text, 3) & Right(cboYear.Text, 2) & _
"Shf Set " & strTmp & " Where EmpCode='" & cboCode.List(cboCode.ListIndex, 1) & "'"
Exit Function
ERR_P:
    ShowError ("SaveModMaster :: " & Me.Caption)
    SaveModMaster = False
End Function

Private Sub FillSchShift()          '' Fills Employee Shift For the Newly Added Employee
On Error GoTo ERR_P                 '' for the Current Month
'' Start Date Checks
Dim dttmp As Date
'' If the Employee has Already Left
If Not IsNull(adrsEmp("LeavDate")) Then
    dttmp = FdtLdt(cboMonth.ListIndex + 1, cboYear.Text, "F")
    If DateCompDate(adrsEmp("LeavDate")) <= dttmp Then Exit Sub
End If
'' Get Current Months Last Process Date
dttmp = FdtLdt(cboMonth.ListIndex + 1, cboYear.Text, "L")
'' Check on JoinDate
If DateCompDate(adrsEmp("JoinDate")) > dttmp Then Exit Sub
'' Check on Shift Date
If DateCompDate(Shft.startdate) > dttmp Then Exit Sub
'' End Date Checks
Call GetSENums(cboMonth.Text, cboYear.Text)
adrsEmp.Requery                     '' Requery the Recordset
Call FillEmployeeDetails(cboCode.List(cboCode.ListIndex, 1))
If UCase(MonthName(Month(Shft.startdate))) = UCase(cboMonth.Text) And Year(Shft.startdate) = CInt(cboYear.Text) Then Call AdjustSENums(DateCompDate(Shft.startdate))
If typEmpRot.strShifttype = "F" Then
    '' If Fixed Shifts
    Call FixedShifts(cboCode.List(cboCode.ListIndex, 1), cboMonth.Text, cboYear.Text)
Else
    '' if Rotation Shifts
    '' Fill Other Skip Pattern and Shift Pattern Array
    Call FillArrays
    Select Case strCapSND
        Case "O"        '' After Specific Number of Days
            Call SpecificDaysShifts(cboCode.List(cboCode.ListIndex, 1), cboMonth.Text, cboYear.Text)
        Case "D"        '' Only on Fixed Days
            Call FixedDaysShifts(cboCode.List(cboCode.ListIndex, 1), cboMonth.Text, cboYear.Text)
        Case "W"        '' Only On Fixed Week days
            Call WeekDaysShifts(cboCode.List(cboCode.ListIndex, 1), cboMonth.Text, cboYear.Text)
    End Select
End If
''
'' Add that Record to the Shift File
Call UpdateAfterShiftDate(cboMonth.Text, cboYear.Text, cboCode.List(cboCode.ListIndex, 1))
Call Display
Exit Sub
ERR_P:
    ShowError ("FillSchShift :: " & Me.Caption)
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 3, 17)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edit Shift Schdule Of Employee " & cboCode.Text & "For " & cboMonth.Text & " " & cboYear.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub
