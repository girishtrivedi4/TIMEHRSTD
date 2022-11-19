VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5295
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   4200
         Top             =   120
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1320
         Width           =   2205
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   435
         Left            =   4095
         TabIndex        =   5
         Top             =   530
         Width           =   1035
      End
      Begin VB.CommandButton cmdExit 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   435
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1260
         Width           =   1005
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   5280
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1650
         TabIndex        =   0
         Top             =   300
         Width           =   1485
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1650
         TabIndex        =   3
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Image ImgWin 
         Height          =   1455
         Left            =   60
         Picture         =   "frmLogin.frx":0E42
         Stretch         =   -1  'True
         Top             =   320
         Width           =   1395
      End
      Begin VB.Line Line1 
         X1              =   1500
         X2              =   1500
         Y1              =   270
         Y2              =   1800
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   1530
         X2              =   1530
         Y1              =   270
         Y2              =   1800
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   1500
         Y1              =   270
         Y2              =   270
      End
      Begin MSForms.ComboBox cboUser 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   2205
         VariousPropertyBits=   612390939
         DisplayStyle    =   3
         Size            =   "3889;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset
Dim adrsForm As New ADODB.Recordset

Public Sub cmdOK_Click()
On Error GoTo ERR_P
Dim strTmp As String, strPassTmp As String, intTmpDept As Integer
Dim strTmpHODR As String, strTmpMASR As String, strTmpLEAVER As String
Dim strtmpOTHERR1 As String
Dim strTmpOPass1 As String
intTmpDept = 0
'' Get Username in a temporary variable
If txtUser.Visible = True Then   'combo
    strTmp = Trim(txtUser.Text)
Else
    strTmp = Trim(cboUser.Text)
End If
'' Check for empty username
If strTmp = "" Then
    MsgBox NewCaptionTxt("69003", adrsC), vbExclamation
    If txtUser.Visible = True Then
        txtUser.SetFocus
    Else
        cboUser.SetFocus
    End If
    Exit Sub
End If
'' Get the Password
strPassTmp = Trim(txtPass.Text)
If UCase(strTmp) <> strPrintUser Then
    If adrsForm.State = 1 Then adrsForm.Close
       adrsForm.Open "Select UserName,Password,UserType,Dept,HODRights,MasterRights," & _
    "LeaveRights,OtherRights1,OtherPass1 from UserAccs where UserName='" _
    & strTmp & "' and Password='" & DEncryptDat(UCase(strPassTmp), 1) & "'"
AutoProcess:
    If adrsForm.EOF Then
        MsgBox NewCaptionTxt("69004", adrsC), vbExclamation
        Call AuditInfo("LOGIN FAILED", Me.Caption, "Login Failed", strTmp)
        txtPass.Text = ""
        Exit Sub
    Else
        Select Case UCase(adrsForm("UserType"))
            Case ADMIN
                strCurrentUserType = ADMIN
                strCurrData = ""
            Case HOD
                strCurrentUserType = HOD
                intTmpDept = Val(adrsForm("Dept"))
                Call GetDataStr
            Case GENERAL
                strCurrentUserType = GENERAL
            Case Else
                strCurrentUserType = GENERAL
        End Select
        '' Get Rights
        strTmpHODR = IIf(IsNull(adrsForm("HODRights")), "", adrsForm("HODRights"))
        strTmpMASR = IIf(IsNull(adrsForm("MasterRights")), "", adrsForm("MasterRights"))
        strTmpLEAVER = IIf(IsNull(adrsForm("LeaveRights")), "", adrsForm("LeaveRights"))
        strtmpOTHERR1 = IIf(IsNull(adrsForm("OtherRights1")), "", adrsForm("OtherRights1"))
        strTmpOPass1 = IIf(IsNull(adrsForm("OtherPass1")), "", adrsForm("OtherPass1"))
        strTmpOPass1 = UCase(DEncryptDat(Trim(strTmpOPass1), 1))
    End If
Else
    If Trim(strPassTmp) <> strPrintPass Then
        MsgBox NewCaptionTxt("00086", adrsMod), vbExclamation
        txtPass.Text = ""
        Exit Sub
    Else
        strTmpOPass1 = "PRINT"
        strCurrentUserType = ADMIN
    End If
    MainForm.INI.Visible = True
End If
'' Set Values to Global Variables
UserName = strTmp
'intCurrDept = intTmpDept

strHODRights = strTmpHODR
strMasterRights = strTmpMASR
strLeaveRights = strTmpLEAVER
strOtherRights1 = strtmpOTHERR1
strPassword = strPassTmp
strOtherPass1 = strTmpOPass1
''
'
    frmLogin.Height = 3750
Unload Me
Call AddActivityLog(lg_NoModeAction, 2, 34)     '' Add Add Log
Call AuditInfo("LOGIN", Me.Caption, "Login User Name: " & typLog.strUsername)
Call LoginLog
If Not blnDiff Then

    DoEvents
    
    MainForm.Enabled = True
    MainForm.SetFocus
    Call SetCaptionMainForm
'    Call AddCO
End If
blnDiff = True
Exit Sub
ERR_P:
    ShowError ("OK::" & Me.Caption)
    'Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Call BootStatus
Call SetFormIcon(Me, True)
Call RetCaptions
Call LoadSpecifics
'Lblchk.Visible = False
With ImgWin
    .Picture = LoadPicture(App.path & "\Images\IMWA0002.gif")
End With
Exit Sub
ERR_P:
    ShowError ("Load ::" & Me.Caption)
End Sub

Private Sub RetCaptions()
''On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '69%'", ConMain, adOpenStatic
'version add by
Me.Caption = "User Login"
''command captions
cmdOk.Caption = NewCaptionTxt("00002", adrsMod)             '' &OK
cmdExit.Caption = NewCaptionTxt("00003", adrsMod)           '' &Cancel
''Label captions
lblUser.Caption = NewCaptionTxt("69002", adrsC)             ''
lblPass.Caption = NewCaptionTxt("00084", adrsMod)           ''
End Sub

Private Sub LoadSpecifics()
On Error GoTo ERR_P
With adrsForm
    If .State = 1 Then .Close
    .ActiveConnection = ConMain
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open "Select UserName,Password from UserAccs Order by UserName"
    cboUser.clear
    Do While Not .EOF
        If Not IsNull(adrsForm("UserName")) Then
            cboUser.AddItem adrsForm("UserName")
        End If
        .MoveNext
    Loop
End With
Exit Sub
ERR_P:
    ShowError ("LoadSpecifics::" & Me.Caption)
End Sub

Private Sub FreeRes()
'' On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
If adrsForm.State = 1 Then adrsForm.Close
Set adrsC = Nothing
Set adrsForm = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call FreeRes
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ERR_P
If KeyAscii = 28 Then
    txtUser.Visible = True
    cboUser.Visible = False
    txtPass.Text = ""
    txtUser.Text = ""
    txtUser.SetFocus
End If
Exit Sub
ERR_P:
    ShowError ("Keypress :: " & Me.Caption)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF10 Then KeyCode = 0
End Sub

Private Sub GetDataStr()
On Error GoTo ERR_P
Dim strArr()  As String, bytCnt As Byte
Dim strFrom As String, strWhere As String

Dim bytRcd As Byte, strTmp As String
Dim strTAr() As String
''
strArr = Split(adrsForm("Dept"), "@")
For bytCnt = 0 To 4
    Select Case bytCnt
        Case 0  ''Dept
            If Trim(strArr(0)) <> "" Then
                strFrom = strFrom & "Deptdesc,"
                
                strTAr = Split(strArr(0), ",")
                For bytRcd = 0 To UBound(strTAr) - 1
                    strTmp = strTmp & "'" & strTAr(bytRcd) & "',"
                Next

                If strTmp <> "" Then strTmp = Left(strTmp, Len(strTmp) - 1)
                strWhere = strWhere & "Deptdesc.Dept in (" & strTmp & ") And Empmst.Dept = Deptdesc.Dept AND "
                strCurrDept = " And Deptdesc.Dept in (" & strTmp & ") "
                strTmp = ""
            End If
        Case 1  ''Company
            If Trim(strArr(1)) <> "" Then
                strFrom = strFrom & "Company,"
                strWhere = strWhere & "Company.Company in (" & Left(strArr(1), Len(strArr(1)) - 1) & ") And Empmst.Company = Company.Company AND "
            End If
        Case 2  ''Groupmst
            If Trim(strArr(2)) <> "" Then
                strTAr = Split(strArr(2), ",")
                For bytRcd = 0 To UBound(strTAr) - 1
                    strTmp = strTmp & "'" & strTAr(bytRcd) & "',"
                Next
                If strTmp <> "" Then strTmp = Left(strTmp, Len(strTmp) - 1)
                strFrom = strFrom & "Groupmst,"
                strWhere = strWhere & "Groupmst." & strKGroup & " in (" & strTmp & ") And Empmst." & strKGroup & " = Groupmst." & strKGroup & " AND "
            End If
        Case 3  ''Division
            If Trim(strArr(3)) <> "" Then
                strFrom = strFrom & "Division,"
                strWhere = strWhere & "Division.Div in (" & Left(strArr(3), Len(strArr(3)) - 1) & ") And Empmst.Div = Division.Div AND "
            End If
        Case 4  ''Location
            If Trim(strArr(4)) <> "" Then
                strFrom = strFrom & "Location,"
                strWhere = strWhere & "Location.Location in (" & Left(strArr(4), Len(strArr(4)) - 1) & ") And Empmst.Location = Location.Location AND "
            End If
    End Select
Next
If Trim(strFrom) <> "" Then
    strFrom = "," & Left(strFrom, Len(strFrom) - 1)
End If
If Trim(strWhere) <> "" Then
    strWhere = " where " & Left(strWhere, Len(strWhere) - 4)
End If
strCurrData = Replace(strFrom, "'", "") & " " & Replace(strWhere, "'", "")
Exit Sub
ERR_P:
    ShowError ("GetDataStr :: " & Me.Caption)
    ''Resume Next
End Sub
