VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Visual Star (VSTAR)"
   ClientHeight    =   4080
   ClientLeft      =   2340
   ClientTop       =   2040
   ClientWidth     =   5730
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2816.088
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   3720
      Width           =   1260
   End
   Begin VB.Image picIcon 
      Height          =   735
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label WebsiteAddr 
      AutoSize        =   -1  'True
      Caption         =   "Also visit us : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1020
      TabIndex        =   5
      Top             =   2760
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   1
      X1              =   -42.257
      X2              =   5380.766
      Y1              =   2153.479
      Y2              =   2153.479
   End
   Begin VB.Label lblDescription 
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Visual System for "
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   1470
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   690
      Width           =   525
   End
   Begin VB.Label lblDisclaimer 
      AutoSize        =   -1  'True
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   45
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset

Private Sub cmdOK_Click()
    Unload frmAbout
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Call SetFormIcon(Me)
picIcon.Picture = LoadPicture(App.Path & "\Images\IMWA0002.gif")
Call RetCaptions
Exit Sub
ERR_P:
    ShowError ("Load :: " & Me.Caption)
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions Where ID Like '05%' or Id Like '00%'", VstarDataEnv.cnDJConn, adOpenStatic
Me.Caption = NewCaptionTxt("05001", adrsC)
lblTitle.Caption = NewCaptionTxt("05002", adrsC)
lblVersion.Caption = NewCaptionTxt("05003", adrsC) & "       :     " & App.Major & "." & App.Minor & "." & App.Revision
lblDescription.Caption = NewCaptionTxt("05004", adrsC) & vbCrLf & NewCaptionTxt("05005", adrsC) & _
vbCrLf & NewCaptionTxt("05006", adrsC) & vbCrLf & NewCaptionTxt("05007", adrsC) & _
vbCrLf & NewCaptionTxt("05008", adrsC)
WebsiteAddr.Caption = NewCaptionTxt("05009", adrsC)
lblDisclaimer.Caption = NewCaptionTxt("05010", adrsC)
cmdOk.Caption = NewCaptionTxt("00002", adrsC)
End Sub

