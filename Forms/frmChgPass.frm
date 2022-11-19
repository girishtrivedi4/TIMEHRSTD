VERSION 5.00
Begin VB.Form frmChgPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Passwords"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   585
      Left            =   4260
      TabIndex        =   8
      Top             =   3330
      Width           =   1185
   End
   Begin VB.Frame frLogin 
      Caption         =   "Login Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
         Height          =   495
         Index           =   0
         Left            =   4170
         TabIndex        =   3
         Top             =   570
         Width           =   1185
      End
      Begin VB.TextBox txtLogin 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1590
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   240
         Width           =   2475
      End
      Begin VB.TextBox txtLogin 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1590
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   690
         Width           =   2475
      End
      Begin VB.TextBox txtLogin 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1590
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1140
         Width           =   2475
      End
      Begin VB.Label lblLogin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   300
         Width           =   975
      End
      Begin VB.Label lblLogin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   11
         Top             =   750
         Width           =   1065
      End
      Begin VB.Label lblLogin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   12
         Top             =   1200
         Width           =   1260
      End
   End
   Begin VB.Frame frSecond 
      Caption         =   "Leave Request and Correction Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   30
      TabIndex        =   13
      Top             =   1680
      Width           =   5415
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
         Height          =   495
         Index           =   1
         Left            =   4170
         TabIndex        =   7
         Top             =   570
         Width           =   1185
      End
      Begin VB.TextBox txtSecond 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1590
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   240
         Width           =   2475
      End
      Begin VB.TextBox txtSecond 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1590
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   690
         Width           =   2475
      End
      Begin VB.TextBox txtSecond 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1590
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1140
         Width           =   2475
      End
      Begin VB.Label lblSecond 
         AutoSize        =   -1  'True
         Caption         =   "Old Password"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   330
         Width           =   975
      End
      Begin VB.Label lblSecond 
         AutoSize        =   -1  'True
         Caption         =   "New Password"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   750
         Width           =   1065
      End
      Begin VB.Label lblSecond 
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   16
         Top             =   1200
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmChgPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset

Private Sub cmdChange_Click(Index As Integer)
Select Case Index
    Case 0      '' Login Password
        If Not ValiDatePass(1, txtLogin) Then Exit Sub
        Call ChangePass(1, txtLogin)
    Case 1      '' Second Password
        If Not ValiDatePass(2, txtSecond) Then Exit Sub
        Call ChangePass(2, txtSecond)
End Select
End Sub

Private Sub ChangePass(ByVal bytPass As Byte, txt As Object)
On Error GoTo ERR_P
Err.clear
Dim strTmp As String
Select Case bytPass
    Case 1
        strTmp = "[Password]"
    Case 2
        strTmp = "OtherPass1"
End Select
ConMain.Execute "Update UserAccs Set " & strTmp & "='" & _
DEncryptDat(UCase(Trim(txt(2).Text)), 1) & "' where UserName='" & UserName & "'"
If Err.Number = 0 Then
    Select Case bytPass
        Case 1: strPassword = UCase(Trim(txt(2).Text))
        Case 2: strOtherPass1 = UCase(Trim(txt(2).Text))
    End Select
    MsgBox NewCaptionTxt("00103", adrsMod)      '' "Password Changed Successfully"
Else
    MsgBox NewCaptionTxt("68008", adrsC)
End If
Exit Sub
ERR_P:
    ShowError ("ChangePass::" & Me.Caption)
End Sub

Private Function ValiDatePass(bytPass As Byte, txt As Object) As Boolean
On Error GoTo ERR_P
Dim strTmp As String, adrsTmp As New ADODB.Recordset
'' Set the Field Caption
Select Case bytPass
    Case 1: strTmp = "Password"
    Case 2: strTmp = "OtherPass1"
End Select
If Trim(txt(0).Text) = "" Then
    MsgBox "Blank Old Password", vbExclamation
    txt(0).SetFocus
    Exit Function
End If
If adrsTmp.State = 1 Then adrsTmp.Close
With adrsTmp
    .ActiveConnection = ConMain
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open "Select " & strTmp & " from UserAccs where UserName='" & UserName & "'"
    If Not .EOF Then
        If Not IsNull(adrsTmp(strTmp)) Then
            If UCase(Trim(txt(0).Text)) <> UCase(DEncryptDat(Trim(adrsTmp(strTmp)), 1)) Then
                MsgBox NewCaptionTxt("00086", adrsMod)
                txt(0).SetFocus
                Exit Function
            End If
        Else
            MsgBox NewCaptionTxt("68010", adrsC)
            txt(0).SetFocus
            Exit Function
        End If
    Else
            MsgBox NewCaptionTxt("68011", adrsC)
            txt(0).SetFocus
            Exit Function
    End If
End With
If Trim(txt(1).Text) = "" Then
    MsgBox NewCaptionTxt("68012", adrsC)
    txt(1).SetFocus
    Exit Function
End If
If UCase(Trim(txt(1).Text)) <> UCase(Trim(txt(2).Text)) Then
    MsgBox NewCaptionTxt("68013", adrsC)
    txt(1).SetFocus
    Exit Function
End If
If InStr(DEncryptDat(UCase(Trim(txt(1).Text)), 1), Chr(34)) > 0 Then
    MsgBox NewCaptionTxt("68014", adrsC)
    txt(1).SetFocus
    Exit Function
End If
If InStr(DEncryptDat(UCase(Trim(txt(1).Text)), 1), "'") > 0 Then
    MsgBox NewCaptionTxt("68014", adrsC)
    txt(1).SetFocus
    Exit Function
End If
ValiDatePass = True
Exit Function
ERR_P:
    ShowError ("ValiDatePass::" & Me.Caption)
End Function

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 13 Then Exit Sub
Select Case UCase(Trim(Me.ActiveControl.name))
    Case "TXTLOGIN"
        Select Case Me.ActiveControl.Index
            Case 0  '' Old Password
                txtLogin(1).SetFocus
            Case 1  '' New Password
                txtLogin(2).SetFocus
            Case 2  '' Confirm Password
                Call cmdChange_Click(0)
        End Select
    Case "TXTSECOND"
        Select Case Me.ActiveControl.Index
            Case 0  '' Old Password
                txtSecond(1).SetFocus
            Case 1  '' New Password
                txtSecond(2).SetFocus
            Case 2  '' Confirm Password
                Call cmdChange_Click(1)
        End Select
End Select
End Sub

Private Sub Form_Load()
    Call SetFormIcon(Me, True)
End Sub

Private Sub txtLogin_GotFocus(Index As Integer)
    Call GF(txtLogin(Index))
End Sub

Private Sub txtSecond_GotFocus(Index As Integer)
    Call GF(txtLogin(Index))
End Sub
