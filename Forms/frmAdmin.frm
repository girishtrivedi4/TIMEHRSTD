VERSION 5.00
Begin VB.Form frmAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IVS "
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   525
      Left            =   0
      TabIndex        =   3
      Top             =   1530
      Width           =   2865
   End
   Begin VB.CommandButton cmdYearly 
      Caption         =   "Reset Yearly Lock"
      Height          =   525
      Left            =   0
      TabIndex        =   2
      Top             =   1020
      Width           =   2865
   End
   Begin VB.CommandButton cmdMonthly 
      Caption         =   "Reset Monthly Lock"
      Height          =   525
      Left            =   0
      TabIndex        =   1
      Top             =   510
      Width           =   2865
   End
   Begin VB.CommandButton cmdDaily 
      Caption         =   "Reset Daily Lock"
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2865
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset

Private Sub cmdDaily_Click()
On Error GoTo ERR_P
If MsgBox(NewCaptionTxt("06006", adrsC) & vbCrLf _
& NewCaptionTxt("00009", adrsMod), vbQuestion + vbYesNo, App.EXEName) = vbYes Then
    Call AddActivityLog(lgDaily_Action, 3, 28)       '' Daily Log
    Call AuditInfo("RESET DAILY LOCK", Me.Caption, "Reset Daily Lock")
    ConMain.Execute "Update Exc Set Daily=0"
End If
Exit Sub
ERR_P:
    ShowError ("Reset Daily :: " & Me.Caption)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdMonthly_Click()
On Error GoTo ERR_P
If MsgBox(NewCaptionTxt("06007", adrsC) & vbCrLf _
& NewCaptionTxt("00009", adrsMod), vbQuestion + vbYesNo, App.EXEName) = vbYes Then
    Call AddActivityLog(lgMonthly_Action, 3, 28)       '' Monthly Log
    Call AuditInfo("RESET MONTHLY LOCK", Me.Caption, "Reset Monthly Lock")
    ConMain.Execute "Update Exc Set Monthly=0"
End If
Exit Sub
ERR_P:
    ShowError ("Reset Monthly :: " & Me.Caption)
End Sub

Private Sub cmdYearly_Click()
On Error GoTo ERR_P
If MsgBox(NewCaptionTxt("06008", adrsC) & vbCrLf _
& NewCaptionTxt("00009", adrsMod), vbQuestion + vbYesNo, App.EXEName) = vbYes Then
    Call AddActivityLog(lgYearly_Action, 3, 28)       '' Yearly Log
    Call AuditInfo("RESET YEARLY LOCK", Me.Caption, "Reset Yearly Lock")
    ConMain.Execute "Update Exc Set Yearly=0"
End If
Exit Sub
ERR_P:
    ShowError ("Reset Yearly :: " & Me.Caption)
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Call SetFormIcon(Me)
Call RetCaptions
Call GetRights
Exit Sub
ERR_P:
    ShowError ("Load :: " & Me.Caption)
End Sub

Private Sub GetRights()
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(3, 18, , 1)
If strTmp = "1" Then
    cmdDaily.Enabled = True
    cmdMonthly.Enabled = True
    cmdYearly.Enabled = True
Else
    cmdDaily.Enabled = False
    cmdMonthly.Enabled = False
    cmdYearly.Enabled = False
End If
Exit Sub
ERR_P:
    ShowError ("GetRights::" & Me.Caption)
    cmdDaily.Enabled = False
    cmdMonthly.Enabled = False
    cmdYearly.Enabled = False
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions Where ID Like '06%'", ConMain, adOpenStatic
Me.Caption = "Admin Form"
End Sub
