VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3405
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   8775
   ClipControls    =   0   'False
   FillColor       =   &H00400000&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.ivsofttech.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1260
      TabIndex        =   13
      Top             =   3120
      Width           =   1725
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "support@ivsofttech.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5910
      TabIndex        =   12
      Top             =   3120
      Width           =   2070
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web Site   :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   180
      TabIndex        =   11
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail   :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5070
      TabIndex        =   10
      Top             =   3120
      Width           =   765
   End
   Begin VB.Label lblAdd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1290
      TabIndex        =   9
      Top             =   2430
      Width           =   675
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address     :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   2430
      Width           =   1035
   End
   Begin VB.Label lblFax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax No. :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5070
      TabIndex        =   7
      Top             =   2910
      Width           =   780
   End
   Begin VB.Label lblPh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No. :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   2910
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2190
      Width           =   1665
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   60
      X2              =   8675
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Label lblLic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "License"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1890
      Width           =   660
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0FFFF&
      Height          =   1515
      Left            =   60
      Top             =   1850
      Width           =   8655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   1450
      Width           =   8655
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   8740
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   30
      X2              =   8770
      Y1              =   1420
      Y2              =   1420
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   3405
      Left            =   30
      Top             =   30
      Width           =   8745
   End
   Begin VB.Label lblPlease 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please wait while the Application initializes..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   345
      Left            =   3000
      TabIndex        =   2
      Top             =   1050
      Width           =   5700
   End
   Begin VB.Label lblVer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time Attendance Processing System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   345
      Left            =   3000
      TabIndex        =   1
      Top             =   660
      Width           =   5700
   End
   Begin VB.Image ImgWin 
      Height          =   1245
      Left            =   60
      Top             =   60
      Width           =   2720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IV Softtechub Private Limited"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3735
      TabIndex        =   0
      Top             =   150
      Width           =   4005
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   3000
      Top             =   60
      Width           =   5700
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error Resume Next
Call ReadIni
With ImgWin
    .Picture = LoadPicture(App.Path & "\Images\IMG-20130823-WA0001.gif")
End With
Label4.Caption = "Todays Date is " & Date & " and Time is " & Time
lblLic.Caption = "Licenced to " & InVar.strCOM & " - " & InVar.bytUse & " User(s)"
'lblVer.Caption = lblVer.Caption & " (Ver. " & App.Major & "." & App.Minor & "." & App.Revision & ")"
lblVer.Caption = lblVer.Caption & " (" & App.Major & "." & App.Minor & "." & App.Revision & ")"
lblPh.Caption = "" 'lblPh.Caption & " 91-0251-240 6365 "
lblFax.Caption = "" 'lblFax.Caption & " 91-22-2413 5883"
lblAdd.Caption = "1,Shree Ganpat CHS,Near Vitawa Octroi Naka" & _
                vbCrLf & "Thane Belapur Road,Vitawa, Thane. 400605. India."
End Sub

