VERSION 5.00
Begin VB.Form frmPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change User Password"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   405
      Left            =   3810
      TabIndex        =   7
      Top             =   660
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   3810
      TabIndex        =   6
      Top             =   120
      Width           =   1155
   End
   Begin VB.TextBox txtCPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1860
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   810
      Width           =   1575
   End
   Begin VB.TextBox txtNPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1860
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   450
      Width           =   1575
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1860
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   90
      Width           =   1575
   End
   Begin VB.Label lblCPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      Height          =   195
      Left            =   30
      TabIndex        =   4
      Top             =   840
      Width           =   1260
   End
   Begin VB.Label lblNPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      Height          =   195
      Left            =   270
      TabIndex        =   2
      Top             =   480
      Width           =   1065
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   270
      TabIndex        =   0
      Top             =   120
      Width           =   690
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adrsC As New ADODB.Recordset

Private Sub cmdCan_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
'' Checking Validations
If UCase(UserName) = strPrintUser Then Exit Sub
If Trim(txtPass.Text) = "" Then
    MsgBox NewCaptionTxt("00108", adrsMod), vbExclamation
    txtPass.SetFocus
    Exit Sub
End If
If Trim(txtNPass.Text) = "" Then
    MsgBox NewCaptionTxt("00108", adrsMod), vbExclamation
    txtNPass.SetFocus
    Exit Sub
End If
If UCase(Trim(txtNPass.Text)) <> UCase(Trim(txtCPass.Text)) Then
    MsgBox NewCaptionTxt("61005", adrsC), vbExclamation
    txtNPass.Text = ""
    txtCPass.Text = ""
    txtNPass.SetFocus
    Exit Sub
End If
If InStr(DEncryptDat(UCase(txtPass.Text), 1), "'") > 0 Then
    MsgBox NewCaptionTxt("00107", adrsMod), vbExclamation
    txtPass.Text = ""
    txtPass.SetFocus
    Exit Sub
End If
If InStr(DEncryptDat(UCase(txtNPass.Text), 1), "'") > 0 Then
    MsgBox NewCaptionTxt("00107", adrsMod), vbExclamation
    txtNPass.Text = ""
    txtNPass.SetFocus
    Exit Sub
End If
If InStr(DEncryptDat(UCase(txtCPass.Text), 1), Chr(34)) > 0 Then
    MsgBox NewCaptionTxt("00107", adrsMod), vbExclamation
    txtCPass.Text = ""
    txtCPass.SetFocus
    Exit Sub
End If
'' Check for Correct Password
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select * from Userinfo where UserName='" & UserName & "' and Password='" & _
DEncryptDat(UCase(txtPass.Text), 1) & "'", ConMain, adOpenStatic
If (adrsDept1.EOF And adrsDept1.BOF) Then
    MsgBox NewCaptionTxt("61007", adrsC), vbExclamation
    txtPass.Text = ""
    txtNPass.Text = ""
    txtCPass.Text = ""
    txtPass.SetFocus
    Exit Sub
Else
    ConMain.Execute "Update Userinfo set Password='" & _
    DEncryptDat(UCase(Trim(txtNPass.Text)), 1) & "' Where UserName='" & _
    UserName & "'"
    '' Password Changed Successfully.
    MsgBox NewCaptionTxt("00103", adrsMod), vbExclamation
    txtPass.Text = ""
    txtNPass.Text = ""
    txtCPass.Text = ""
    txtPass.SetFocus
End If
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)
Call RetCaptions
End Sub

Private Sub RetCaptions()
'' On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '61%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("61001", adrsC)
lblPass.Caption = NewCaptionTxt("61002", adrsC)
lblNPass.Caption = NewCaptionTxt("61003", adrsC)
lblCPass.Caption = NewCaptionTxt("61004", adrsC)
cmdOk.Caption = NewCaptionTxt("00002", adrsMod)
cmdCan.Caption = NewCaptionTxt("00003", adrsMod)
End Sub
