VERSION 5.00
Begin VB.Form frmRotWD 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdWeekDays 
      Caption         =   "SUN"
      Height          =   345
      Index           =   6
      Left            =   5070
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdWeekDays 
      Caption         =   "SAT"
      Height          =   345
      Index           =   5
      Left            =   4230
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdWeekDays 
      Caption         =   "FRI"
      Height          =   345
      Index           =   4
      Left            =   3390
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdWeekDays 
      Caption         =   "THU"
      Height          =   345
      Index           =   3
      Left            =   2550
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdWeekDays 
      Caption         =   "WED"
      Height          =   345
      Index           =   2
      Left            =   1710
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdWeekDays 
      Caption         =   "TUE"
      Height          =   345
      Index           =   1
      Left            =   870
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdWeekDays 
      Caption         =   "MON"
      Height          =   345
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   " "
      Height          =   405
      Left            =   4410
      MaskColor       =   &H80000004&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   30
      UseMaskColor    =   -1  'True
      Width           =   1065
   End
   Begin VB.TextBox txtDates 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   4335
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   " "
      Height          =   375
      Left            =   4710
      TabIndex        =   11
      Top             =   870
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   " "
      Height          =   375
      Left            =   3510
      TabIndex        =   10
      Top             =   870
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   " "
      Height          =   375
      Left            =   2310
      TabIndex        =   9
      Top             =   870
      Width           =   1215
   End
End
Attribute VB_Name = "frmRotWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset
Private Sub cmdCan_Click()
    Unload Me
End Sub

Private Sub cmdUndo_Click()
On Error Resume Next
If txtDates.Text = "" Then Exit Sub
txtDates.Text = Left(txtDates.Text, Len(txtDates.Text) - 3)
End Sub

Private Sub cmdWeekDays_Click(Index As Integer)
On Error Resume Next
If InStr(txtDates.Text, Left(cmdWeekDays(Index).Caption, 2) & ",") <= 0 Then
    txtDates.Text = txtDates.Text & Left(cmdWeekDays(Index).Caption, 2) & ","
End If
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Call SetFormIcon(Me, True)
Call RetCaptions
txtDates = strRotPass
cmdUndo.Picture = LoadPicture(App.path & "\Images\Undo.Bmp")
Exit Sub
ERR_P:
    ShowError ("Load :: " & Me.Caption)
End Sub

Private Sub RetCaptions()
On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '43%'", ConMain, adOpenStatic, adLockReadOnly
Me.Caption = NewCaptionTxt("43001", adrsC)
With cmdWeekDays
    .Item(0).Caption = NewCaptionTxt("00065", adrsMod)
    .Item(1).Caption = NewCaptionTxt("00066", adrsMod)
    .Item(2).Caption = NewCaptionTxt("00067", adrsMod)
    .Item(3).Caption = NewCaptionTxt("00068", adrsMod)
    .Item(4).Caption = NewCaptionTxt("00069", adrsMod)
    .Item(5).Caption = NewCaptionTxt("00070", adrsMod)
    .Item(6).Caption = NewCaptionTxt("00071", adrsMod)
End With
cmdOk.Caption = NewCaptionTxt("00002", adrsMod)
cmdReset.Caption = NewCaptionTxt("00064", adrsMod)
cmdCan.Caption = NewCaptionTxt("00003", adrsMod)
End Sub

Private Sub cmdOK_Click()
strRotPass = txtDates.Text
Unload Me
End Sub

Private Sub cmdReset_Click()
txtDates.Text = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdOK_Click
End Sub
