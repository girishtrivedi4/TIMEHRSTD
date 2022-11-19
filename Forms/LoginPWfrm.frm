VERSION 5.00
Begin VB.Form LoginPWfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox UserTxt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   " "
      Top             =   330
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox UserNameTxt 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton CancelCmd 
      Cancel          =   -1  'True
      Caption         =   " "
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton OKCmd 
      Caption         =   " "
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox PassWdTxt 
      Height          =   375
      HideSelection   =   0   'False
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   " "
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label UserNameLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password   :"
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label LoginPwLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password   :"
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
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1155
   End
End
Attribute VB_Name = "LoginPWfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adrsC As New ADODB.Recordset

Private Sub CancelCmd_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
PassWdTxt = ""
UserNametxt.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ERR_P
If KeyAscii = 28 Then
    UserTxt.Visible = True
    UserNametxt.Visible = False
    PassWdTxt.Text = ""
    UserTxt.Text = ""
    UserTxt.SetFocus
End If
Exit Sub
ERR_P:
    ShowError ("Keypress :: " & Me.Caption)
End Sub

Private Sub Form_Load()
On Error GoTo ERR_Particular
Call SetFormIcon(Me)        '' Sets the Form Icon
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "userinfo", VstarDataEnv.cnDJConn, adOpenStatic
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    For i = 1 To adrsTemp.RecordCount
        UserNametxt.AddItem adrsTemp!UserName
        adrsTemp.MoveNext
    Next i
End If
adrsTemp.Close
Call RetCaptions
Exit Sub
ERR_Particular:
    ShowError ("Load :: " & Me.Caption)
End Sub

Private Sub Okcmd_Click()
On Error GoTo LoginErr
Dim strTmp As String, strPassTmp As String
'' Get Username in a temporary variable
If UserNametxt.Visible = True Then   'combo
    strTmp = UCase(Trim(UserNametxt))
ElseIf UserTxt.Visible = True Then
    strTmp = UCase(Trim(UserTxt))
End If
'' Check for empty username
If strTmp = "" Then
    MsgBox NewCaptionTxt("52005", adrsC), vbExclamation
    Exit Sub
End If
'' Get the Password
strPassTmp = PassWdTxt
'' Check for UserName and Password
If UCase(strTmp) <> strPrintUser Then
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "Select * from Userinfo where UserName='" & strTmp & "' and Password='" & _
    DEncryptDat(UCase(strPassTmp), 1) & "'", VstarDataEnv.cnDJConn, adOpenStatic
    If adrsTemp.EOF Then
        MsgBox NewCaptionTxt("52007", adrsC), vbExclamation
        PassWdTxt.Text = ""
        PassWdTxt.SetFocus
        strTmp = strTmp
        Exit Sub
    Else
        If adrsTemp("UserRights") = 1 Then
            'blnAdmin = True
        Else
            'blnAdmin = False
        End If
    End If
Else
    If Trim(strPassTmp) <> strPrintPass Then
        MsgBox NewCaptionTxt("52007", adrsC), vbExclamation
        PassWdTxt.Text = ""
        PassWdTxt.SetFocus
        strTmp = strTmp
        Exit Sub
    Else
        'blnAdmin = True
    End If
End If
UserName = strTmp
Unload Me
Call AddActivityLog(lg_NoModeAction, 2, 34)     '' Add Add Log
If Not blnDiff Then
    frmSplash.Show
    DoEvents
    Load MainForm
    Unload frmSplash
    MainForm.Show
End If
blnDiff = True
Exit Sub
LoginErr:
    ShowError ("OK :: " & Me.Caption)
    PassWdTxt = ""
    If UserNametxt.Visible Then
        UserNametxt.SetFocus
    ElseIf UserTxt.Visible Then
        UserTxt.SetFocus
    End If
End Sub

Private Sub PassWdTxt_GotFocus()
    Call GF(PassWdTxt)
End Sub

Private Sub PassWdTxt_KeyPress(KeyAscii As Integer)
KeyAscii = KeyPressCheck(KeyAscii, 6)
End Sub

Private Sub UserNameTxt_Click()
On Error Resume Next
If UserNametxt.Text <> "" Then PassWdTxt.SetFocus
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from Newcaptions where ID like '52%'", VstarDataEnv.cnDJConn
LoginPWfrm.Caption = NewCaptionTxt("52001", adrsC)      '' Login
LoginPwLbl.Caption = NewCaptionTxt("52002", adrsC)      '' Password
OKCmd.Caption = NewCaptionTxt("00002", adrsMod)           '' OK
CancelCmd.Caption = NewCaptionTxt("00003", adrsMod)       '' Cancel
UserNameLbl.Caption = NewCaptionTxt("52003", adrsC)     '' User Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then Call ShowF10("52")
If KeyCode = vbKeyF10 Then KeyCode = 0
End Sub
