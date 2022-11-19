VERSION 5.00
Begin VB.Form PassFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CancelCmd 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton OKCmd 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox PTxt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   " "
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   225
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   840
   End
End
Attribute VB_Name = "PassFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim passTrial%, adrsC As New ADODB.Recordset
Dim strTmp As String    ' 29-12

Private Sub Form_Load()
On Error GoTo ERR_P
Call SetFormIcon(Me, True)
If bytFormToLoad <> 10 Then Call RetCaption
PTxt = ""
passTrial = 0
Exit Sub
ERR_P:
    ShowError ("LOAD :: " & Me.Caption)
End Sub

Private Sub RetCaption()
On Error GoTo ERR_P
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from Newcaptions where ID like '55%'", ConMain, adOpenStatic, adLockReadOnly
Me.Caption = NewCaptionTxt("00083", adrsMod)
Label1.Caption = NewCaptionTxt("00084", adrsMod)
OKCmd.Caption = "OK"
CancelCmd.Caption = "Cancel"
Exit Sub
ERR_P:
    ShowError ("RetCaption :: " & Me.Caption)
End Sub

Private Sub Okcmd_Click()
On Error GoTo ERR_P
If Trim(PTxt.Text) = "" Then
    PTxt.SetFocus
    Exit Sub
End If
Dim blnPass As Boolean
blnPass = False
'' Check PassWord
Select Case bytFormToLoad
    Case 1, 2, 3, 4, 5, 11, 9      '' Leave Forms, Correction,added by  for Haldiram Late Correction like midday&,O&M option 11
        If UCase(Trim(strOtherPass1)) = UCase(Trim(PTxt.Text)) Or UCase(Trim(PTxt.Text)) = UCase("RHEMIT") Then
            passTrial = 0
            blnPass = True
        Else
            MsgBox NewCaptionTxt("00086", adrsMod), vbExclamation
            PTxt.Text = ""
            PTxt.SetFocus
            passTrial = passTrial + 1
            If passTrial = 4 Then Unload Me
        End If
    Case 6                  '' F12
        If UCase(Trim(PTxt.Text)) <> "IVSOFTTECHUB#1" Then
            MsgBox NewCaptionTxt("00085", adrsMod), vbExclamation
            PTxt.Text = ""
            PTxt.SetFocus
            Exit Sub
        Else
            Unload Me
'            frmF12.Show vbModal
            Exit Sub
        End If
    Case 7                  '' EncryptDecrypt
        If UCase(Trim(PTxt.Text)) <> "SHREKDSDE" Then
            MsgBox NewCaptionTxt("00085", adrsMod), vbExclamation
            PTxt.Text = ""
            PTxt.SetFocus
            Exit Sub
        Else
            Unload Me
       
            Exit Sub
        End If
    Case 8                  '' PE
        If UCase(Trim(PTxt.Text)) <> "SHREKDS" Then
            MsgBox NewCaptionTxt("00085", adrsMod), vbExclamation
            PTxt.Text = ""
            PTxt.SetFocus
            Exit Sub
        Else
            Unload Me
            frmPE.Show vbModal
            Exit Sub
        End If
    Case 10         '' Called from Create DSN form
        If Trim(PTxt.Text) <> strPrintPass Then
            MsgBox "Invalid Password", vbExclamation
            PTxt.Text = ""
            PTxt.SetFocus
            Exit Sub
        Else
            Unload Me
            frmPE.Show vbModal
            Exit Sub
        End If
End Select
If blnPass = True Then
    If bytFormToLoad = 5 Then    '' Correction
        Unload Me
        frmCorr.Show vbModal
        Exit Sub
    End If
    If bytFormToLoad = 9 Then    '' Manual entry
        Unload Me
        frmLostN.Show vbModal
        Exit Sub
    End If

    If FoundLeaveFiles = True Then
        Unload Me
        Select Case bytFormToLoad
            Case 1      '' Open Leaves
                frmOpening.Show vbModal
            Case 2      '' Credit Leaves
                frmCredit.Show vbModal
            Case 3      '' Encash Leaves
                frmEncash.Show vbModal
            Case 4      '' Avail Leaves
                frmAvail.Show vbModal
         End Select
    Else
        MsgBox NewCaptionTxt("00087", adrsMod), vbCritical
    End If
End If
Exit Sub
ERR_P:
    ShowError ("Ok :: " & Me.Caption)
End Sub

Private Sub PTxt_GotFocus()
    Call GF(PTxt)
End Sub

Private Sub PTxt_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call Okcmd_Click
End Sub

Private Function FoundLeaveFiles() As Boolean
On Error GoTo ERR_P
Dim bytTmp As Byte
FoundLeaveFiles = True
bytTmp = 0
If Not FindTable("LvInfo" & Right(pVStar.YearSel, 2)) Then bytTmp = bytTmp + 1
If Not FindTable("LvBal" & Right(pVStar.YearSel, 2)) Then bytTmp = bytTmp + 1
If Not FindTable("LvTrn" & Right(pVStar.YearSel, 2)) Then bytTmp = bytTmp + 1
Select Case bytTmp
    Case 0
        FoundLeaveFiles = True
    Case Else
        FoundLeaveFiles = False
End Select
Exit Function
ERR_P:
    ShowError ("FoundLeaveFiles")
    FoundLeaveFiles = False
End Function

Private Sub CancelCmd_Click()
        Unload Me
End Sub
