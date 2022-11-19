VERSION 5.00
Begin VB.Form frmFirst 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrative User"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frDate 
      Height          =   465
      Left            =   150
      TabIndex        =   12
      Top             =   2220
      Width           =   5775
      Begin VB.OptionButton optAmc 
         Caption         =   "Option2"
         Height          =   195
         Left            =   3270
         TabIndex        =   4
         Top             =   180
         Width           =   2355
      End
      Begin VB.OptionButton optBrit 
         Caption         =   "Option1"
         Height          =   195
         Left            =   210
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   2145
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3060
      TabIndex        =   6
      Top             =   2760
      Width           =   1245
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      Height          =   345
      Left            =   1830
      TabIndex        =   5
      Top             =   2760
      Width           =   1245
   End
   Begin VB.TextBox txtRPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1530
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   1665
   End
   Begin VB.TextBox txtPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1530
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   1665
   End
   Begin VB.TextBox txtUser 
      Height          =   315
      Left            =   1530
      TabIndex        =   0
      Top             =   360
      Width           =   1665
   End
   Begin VB.Label lbldate 
      BackStyle       =   0  'Transparent
      Caption         =   "lblDate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   120
      TabIndex        =   11
      Top             =   1620
      Width           =   5835
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This User Will Have All the Administrative Rights"
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
      Left            =   180
      TabIndex        =   7
      Top             =   0
      Width           =   5040
   End
   Begin VB.Label lblRPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblRPass"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1140
      Width           =   615
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblPass"
      Height          =   195
      Left            =   150
      TabIndex        =   9
      Top             =   780
      Width           =   495
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblName"
      Height          =   195
      Left            =   150
      TabIndex        =   8
      Top             =   420
      Width           =   570
   End
End
Attribute VB_Name = "frmFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTmpYear As String
Dim strComp As String
''
Dim adrsC As New ADODB.Recordset

Private Sub cmdCan_Click()
    End
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_particular
Dim strDFormat As String, strInstallDt As String
If UCase(Trim(txtUser.Text)) = UCase(strPrintUser) Then      '' Check for Print User Name
    txtUser.Text = ""
    txtUser.SetFocus
    Exit Sub
End If
strTmpYear = Year(Date)
If optBrit.Value = True Then        '' Get the Confirmation Message and DateFormat needed
    strDFormat = NewCaptionTxt("28010", adrsC) & vbCrLf & _
    vbTab & NewCaptionTxt("28011", adrsC)
    'strDateFO = "DD/MM/YY"
    bytDateF = 2
Else
    strDFormat = NewCaptionTxt("28012", adrsC) & vbCrLf & _
    vbTab & NewCaptionTxt("28011", adrsC)
    strDateFO = "M/D/YY"
    bytDateF = 1
End If                  '' Ask for Confirmation
If MsgBox(strDFormat, vbYesNo + vbQuestion) = vbNo Then Exit Sub
If Not DateSettings(bytDateF) Then       '' Check the Necessary Date Format
    ''Msgbox "Unable To Set the Application Date Settings:: Cannot Proceed", vbCritical
    Exit Sub
End If                      '' Check for USER NAME and PASSWORD
If Trim(txtUser.Text) = "" Then     '' Check for Empty User Name
    txtUser.SetFocus
    Exit Sub
End If
If Trim(txtPass.Text) = "" Then     '' Check for Empty Password
    txtPass.SetFocus
    Exit Sub
End If
If Trim(txtRPass.Text) = "" Then    '' Check for Empty Confirm Password
    txtRPass.SetFocus
    Exit Sub
End If                              '' Check for Print User
If UCase(Trim(txtUser.Text)) = UCase(strPrintUser) Then
    MsgBox NewCaptionTxt("28014", adrsC), vbExclamation
    Unload Me
    Exit Sub
End If                              '' Check for Password Matching
If UCase(Trim(txtPass.Text)) <> UCase(Trim(txtRPass.Text)) Then
        MsgBox NewCaptionTxt("28013", adrsC), vbExclamation
        txtRPass.SetFocus
        Exit Sub
End If                              '' Check for Encryption with Single Quote
If InStr(DEncryptDat(UCase(Trim(txtPass.Text)), 1), "'") > 0 Then
        MsgBox NewCaptionTxt("28014", adrsC), vbExclamation
        txtPass.SetFocus
        Exit Sub
End If                              '' Check for Encryption with Double Quote
If InStr(DEncryptDat(UCase(Trim(txtPass.Text)), 1), Chr(34)) > 0 Then
        MsgBox NewCaptionTxt("28014", adrsC), vbExclamation
        txtPass.SetFocus
        Exit Sub
End If
Call SetDateFormatSQL       '' Set the DMY Statement if SQL and Date Format 2
Call AddUser                '' Add User to the Database
Call AddRecords             '' Add Records to Other Tables
MsgBox txtUser.Text & NewCaptionTxt("28015", adrsC), vbInformation
Unload Me
MainForm.Show
blnDiff = True
Exit Sub
Err_particular:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)
Call RetCaptions    '' Sets the Captions for All the Controls
bytDateF = 2
End Sub

Private Sub optAmc_Click()
    bytDateF = 1
End Sub

Private Sub optBrit_Click()
    bytDateF = 2
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtRPass.SetFocus
Else
    KeyAscii = KeyPressCheck(KeyAscii, 6)
End If
End Sub

Private Sub txtRPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdOK_Click
Else
    KeyAscii = KeyPressCheck(KeyAscii, 6)
End If
End Sub

Private Sub txtRPass_GotFocus()
        Call GF(txtRPass)
End Sub

Private Sub txtPass_GotFocus()
        Call GF(txtPass)
End Sub

Private Sub txtUser_GotFocus()
        Call GF(txtUser)
End Sub

Private Sub AddUser()
On Error GoTo Err_particular
'' insert into UserAccs Table
ConMain.Execute "insert into UserAccs Values('" & _
txtUser.Text & "','" & DEncryptDat(UCase(Trim(txtPass.Text)), 1) & "','ADMIN',NULL,NULL," & _
"NULL,NULL,NULL,NULL,'" & DEncryptDat(UCase(Trim(txtPass.Text)), 1) & "',NULL," & _
strDTEnc & DateCompStr(Date) & strDTEnc & ",'" & _
txtUser.Text & "'," & strDTEnc & DateCompStr(Date) & strDTEnc & ",'" & txtUser.Text & "')"
UserName = txtUser.Text
strCurrentUserType = ADMIN
strPassword = Trim(txtPass.Text)
strOtherPass1 = strPassword
Exit Sub
Err_particular:
    ShowError ("User Cannot be Created " & vbCrLf & _
    " Contact Application Vendor for Support :: " & Me.Caption)
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPass.SetFocus
Else
    KeyAscii = KeyPressCheck(KeyAscii, 6)
End If
End Sub

Private Sub AddRecords()
On Error GoTo Err_particular
'' Install
If NoInstall Then _
ConMain.Execute "insert into install(dt_install,american_dt,upto,e_codesize" & _
",e_cardsize,prsm_cards,cur_year,hl_ot,wo_ot,ot_ot,prn_line,postLt,posterl,filt_time" & _
",yearfrom,weekfrom,deductlter,otround,email,definod) values(" & strDTEnc & DateCompStr(Date) & strDTEnc & _
"," & IIf(bytDateF = 2, 0, -1) & ",10,4,4,0," & strTmpYear & ",0,0,0,0" & _
",0.0,0.0,0.0,1,1,0,0,0,0)"
'' Leavdesc
If NoLeaves Then
    ConMain.Execute "insert into Leavdesc values('Absent Days','A ','N','N','',0,0,0,'','N','100','','',0,0,0,0,'N')"
    ConMain.Execute "insert into Leavdesc values('Holiday Days','HL','N','Y','',0,0,0,'','N','100','','',0,0,0,0,'N')"
    ConMain.Execute "insert into Leavdesc values('Present Days','P ','N','Y','',0,0,0,'','N','100','','',0,0,0,0,'N')"
    ConMain.Execute "insert into Leavdesc values('Weekly Off','WO','N','Y','',0,0,0,'','N','100','','',0,0,0,0,'N')"
End If

'' Exc
ConMain.Execute "insert into Exc values(0,0,0,'',0,0,'" & strCapField & "')"
'' Shift
If NoOShift Then _
ConMain.Execute "insert into instshft values('O',0.01,23.59,8.0,0.0," & _
"0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,1,13.0,13.3,'Open Shift','0',0)"
'' MemoTable
ConMain.Execute "insert into MemoTable Values(1,'" & NewCaptionTxt("00080", adrsMod) & "')"
ConMain.Execute "insert into MemoTable Values(2,'" & NewCaptionTxt("00081", adrsMod) & "')"
ConMain.Execute "insert into MemoTable Values(3,'" & NewCaptionTxt("00082", adrsMod) & "')"
'' Group
If NoCompany("GroupMst") Then _
ConMain.Execute "insert into GroupMst values(1,'General Group')"
'' Company
If NoCompany("Company") Then
strCName = InVar.strCOM
strCName = Replace(strCName, "'", "''")
ConMain.Execute "insert into Company Values(1,'" & strCName & "','')"
End If
'' Location
If NoCompany("Location") Then _
ConMain.Execute "insert into Location Values(1,'General Location')"
'' Division
If NoCompany("Division") Then _
ConMain.Execute "insert into Division Values(1,'General Division')"
'' Log
ConLog.Execute "insert into RECNUM Values(0)"
Exit Sub
Err_particular:
If Err.Number = -2147217900 Then
Else
    ShowError ("AddRecords :: " & Me.Caption)
End If
    'Resume Next
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '28%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("28001", adrsC)          '' Form Caption
lblUser.Caption = NewCaptionTxt("28002", adrsC)     '' This User will ....
lblName.Caption = NewCaptionTxt("28003", adrsC)     '' User Name
lblPass.Caption = NewCaptionTxt("28004", adrsC)     '' Password
lblRPass.Caption = NewCaptionTxt("28005", adrsC)    '' Re-Type Password
lblDate.Caption = NewCaptionTxt("28006", adrsC)     '' The Dat Format ....
optBrit.Caption = NewCaptionTxt("28007", adrsC)     '' British
optAmc.Caption = NewCaptionTxt("28008", adrsC)      '' American
cmdOk.Caption = NewCaptionTxt("00002", adrsMod)       '' Ok
cmdCan.Caption = NewCaptionTxt("00003", adrsMod)      '' Cancel
frDate.Caption = NewCaptionTxt("28009", adrsC)      '' Date Format
End Sub

Private Function NoOShift() As Boolean      '' Checks for Exixtence of O Shift
On Error GoTo ERR_P
NoOShift = True
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "Select Count(*) from INSTSHFT Where Shift='O'", ConMain
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    If IsNull(adrsTemp(0)) Or IsEmpty(adrsTemp(0)) Or adrsTemp(0) <= 0 Then
        NoOShift = True
    Else
        NoOShift = False
    End If
End If
Exit Function
ERR_P:
    NoOShift = True
End Function

Private Function NoInstall() As Boolean     '' Checks for the Exixtence of Parameter Details
On Error GoTo ERR_P
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "Select * from Install", ConMain, adOpenStatic
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    NoInstall = False
Else
    NoInstall = True
End If
Exit Function
ERR_P:
    ShowError ("NoInstall :: " & Me.Caption)
    NoInstall = True
End Function

Private Function NoLeaves() As Boolean      '' Checks for the Exixtence of Records in LeavDesc
On Error GoTo ERR_P
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "Select * from LeavDesc", ConMain, adOpenStatic
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    NoLeaves = False
Else
    NoLeaves = True
End If
Exit Function
ERR_P:
    ShowError ("NoLeaves :: " & Me.Caption)
    NoLeaves = True
End Function

Private Function NoCompany(strTmp) As Boolean
On Error GoTo ERR_P
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "Select * from " & strTmp, ConMain, adOpenStatic
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    NoCompany = False
Else
    NoCompany = True
End If
Exit Function
ERR_P:
    ShowError ("NoCompany :: " & Me.Caption)
    NoCompany = True
End Function
