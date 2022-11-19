VERSION 5.00
Begin VB.Form frmMemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Memo Text"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Top             =   330
      Width           =   1545
   End
   Begin VB.TextBox txtDays 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1620
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "0"
      Top             =   360
      Width           =   345
   End
   Begin VB.TextBox txtMemo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MaxLength       =   124
      TabIndex        =   0
      Top             =   30
      Width           =   7665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2010
      TabIndex        =   5
      Top             =   360
      Width           =   555
   End
   Begin VB.Label lblDays 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ignore for"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   30
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Memo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   30
      TabIndex        =   2
      Top             =   60
      Width           =   645
   End
End
Attribute VB_Name = "frmMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Memo Form
Option Explicit
''
Dim adrsC As New ADODB.Recordset

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me, True)        '' Sets the Forms Icon
Call RetCaptions            '' Sets Captions of controls
Call FillMemoText           '' Gets the Memo Text
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '35%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("35001", adrsC)      '' Insert Memo Text
cmdOk.Caption = NewCaptionTxt("00002", adrsMod)   '' &OK
lblMemo.Caption = NewCaptionTxt("35002", adrsC) '' Memo
lblDays.Caption = NewCaptionTxt("35003", adrsC) '' Ignore for
Label1.Caption = NewCaptionTxt("00012", adrsMod)  '' Days
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DoSave
End Sub

Private Sub txtDays_GotFocus()
    Call GF(txtDays)
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdOK_Click
Else
    KeyAscii = KeyPressCheck(KeyAscii, 2)
End If
End Sub

Private Sub txtMemo_GotFocus()
    Call GF(txtMemo)
End Sub

Private Sub txtMemo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtDays.SetFocus
Else
    KeyAscii = KeyPressCheck(KeyAscii, 8)
End If
End Sub

Private Sub DoSave()
On Error GoTo ERR_P
Select Case bytMode
    Case 7                  '' Absent
        If Trim(txtMemo.Text) = "" Then
            strCapSND = NewCaptionTxt("00080", adrsMod)
        Else
            strCapSND = Trim(txtMemo.Text)
        End If
        ConMain.Execute "Update MemoTable Set Memotext='" & strCapSND & "'" & _
        " Where MemoNum=1"
    Case 12                 '' Late
        If Trim(txtMemo.Text) = "" Then
            strCapSND = NewCaptionTxt("00081", adrsMod)
        Else
            strCapSND = Trim(txtMemo.Text)
        End If
        ConMain.Execute "Update MemoTable Set Memotext='" & strCapSND & "'" & _
        " Where MemoNum=2"
    Case 13                 '' Early
        If Trim(txtMemo.Text) = "" Then
            strCapSND = NewCaptionTxt("00082", adrsMod)
        Else
            strCapSND = Trim(txtMemo.Text)
        End If
        ConMain.Execute "Update MemoTable Set Memotext='" & strCapSND & "'" & _
        " Where MemoNum=3"
End Select
If Val(txtDays.Text) <= 0 Then
    bytMode = 0
Else
    bytMode = Val(txtDays.Text)
End If
Exit Sub
ERR_P:
    ShowError ("DoSave :: " & Me.Caption)
End Sub

Private Sub FillMemoText()
On Error GoTo ERR_P
Dim adrsEmpCnt As New ADODB.Recordset
If adrsEmpCnt.State = 1 Then adrsEmpCnt.Close
adrsEmpCnt.Open "Select * from MemoTable ", ConMain, adOpenStatic
Select Case bytMode
    Case 7          '' Absent
        txtMemo.Text = adrsEmpCnt("Memotext")
    Case 12         '' Late
        adrsEmpCnt.MoveNext
        txtMemo.Text = adrsEmpCnt("Memotext")
    Case 13         '' Early
        adrsEmpCnt.MoveLast
        txtMemo.Text = adrsEmpCnt("Memotext")
End Select
Exit Sub
ERR_P:
    ShowError ("FillMemoText :: " & Me.Caption)
End Sub
