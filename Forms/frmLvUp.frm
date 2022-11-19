VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLvUp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Leave Updation Form"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   30
      TabIndex        =   6
      Top             =   0
      Width           =   5010
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   225
         Left            =   90
         TabIndex        =   8
         Top             =   330
         Width           =   570
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
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
         Left            =   90
         TabIndex        =   7
         Top             =   660
         Width           =   570
      End
   End
   Begin VB.ComboBox cboYear 
      Height          =   315
      Left            =   2190
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1020
      Width           =   1065
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   345
      Left            =   405
      TabIndex        =   3
      Top             =   2010
      Visible         =   0   'False
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   609
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Finish"
      Height          =   495
      Left            =   2490
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   495
      Left            =   1050
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   585
      Left            =   375
      TabIndex        =   0
      Top             =   1410
      Visible         =   0   'False
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   1032
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   0
   End
   Begin VB.Label lblYear 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Year"
      Height          =   195
      Left            =   1710
      TabIndex        =   5
      Top             =   1080
      Width           =   330
   End
End
Attribute VB_Name = "frmLvUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo ERR_P
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open "Select lvupdtyear from install", ConMain
If CInt(cboYear.Text) < pVStar.YearSel Or cboYear.Text = adrsPaid("lvupdtyear") Then
    If (MsgBox(NewCaptionTxt("33009", adrsC), vbYesNo + vbQuestion) = vbYes) Then
UpdateYearlyLeave:
        Call AddActivityLog(lg_NoModeAction, 2, 7)      '' Leave Update Activity Log
        Call UpdateYearlyLeave(cboYear.Text, MSF1, Me)
        Call AuditInfo("UPDATE", Me.Caption, "Update Leave For Year " & cboYear.Text)
    End If
Else
    Call AddActivityLog(lg_NoModeAction, 2, 7)          '' Leave Update Activity Log
    Call UpdateYearlyLeave(cboYear.Text, MSF1, Me)
    Call AuditInfo("UPDATE", Me.Caption, "Update Leave For Year " & cboYear.Text)
End If
    MsgBox NewCaptionTxt("M6004", adrsMod), vbInformation
'cboYear.SetFocus
cmdExit.SetFocus
Exit Sub
ERR_P:
    ShowError ("Update" & Me.Caption)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdUpdate.Enabled = True Then Call cmdUpdate_Click
End Sub

Private Sub Form_Load()
Dim intTmp As Integer
Call SetFormIcon(Me)                '' Sets the Form Icon
Call RetCaptions                    '' Gets and Sets the Captions
Call GetRights                      '' Gets ans Sets the Right
'' Fill Year Combo Box
For intTmp = 1997 To 2096
    cboYear.AddItem CStr(intTmp)
Next
'' Set the Current Year
cboYear.Text = pVStar.YearSel
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '33%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("33001", adrsC)             '' Form Caption
Frame1.Caption = NewCaptionTxt("33002", adrsC)          '' Frame Caption
lbl1.Caption = NewCaptionTxt("33003", adrsC)            '' Label 1 Caption
lbl2.Caption = NewCaptionTxt("33004", adrsC)            '' Label 2 Caption
lblYear.Caption = NewCaptionTxt("00029", adrsMod)          '' &Year
cmdUpdate.Caption = "Update"
cmdExit.Caption = "Finish"
Call CapGrid                                '' Adjust Grids
End Sub

Private Sub CapGrid()
'' Sizing
MSF1.ColWidth(0) = MSF1.ColWidth(0) * 1.35
MSF1.ColWidth(1) = MSF1.ColWidth(1) * 1
MSF1.ColWidth(2) = MSF1.ColWidth(2) * 2.05
MSF2.ColWidth(0) = MSF2.ColWidth(0) * 4.34
'' Setting Captions
MSF1.TextMatrix(0, 0) = NewCaptionTxt("00061", adrsMod)    '' Employee Code
MSF1.TextMatrix(0, 1) = NewCaptionTxt("33005", adrsC)    '' Leave Code
MSF1.TextMatrix(0, 2) = NewCaptionTxt("33006", adrsC)    '' Leave Name
'' "Please Wait...Updating Yearly Leaves"
MSF2.TextMatrix(0, 0) = NewCaptionTxt("33007", adrsC)
End Sub

Private Sub GetRights()
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(3, 3, , 1)
If strTmp = "1" Then
    cmdUpdate.Enabled = True
Else
    cmdUpdate.Enabled = False
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
End If
Exit Sub
ERR_P:
End Sub
