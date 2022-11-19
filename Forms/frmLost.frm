VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLost 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lost Entry"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   3510
      TabIndex        =   14
      Top             =   3660
      Width           =   1185
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2340
      TabIndex        =   13
      Top             =   3660
      Width           =   1185
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command3"
      Height          =   375
      Left            =   1170
      TabIndex        =   12
      Top             =   3660
      Width           =   1185
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command4"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3660
      Width           =   1185
   End
   Begin TabDlg.SSTab TB1 
      Height          =   3645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6429
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmLost.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSF1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmLost.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frLost"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame frLost 
         Height          =   3240
         Left            =   60
         TabIndex        =   2
         Top             =   330
         Width           =   4590
         Begin VB.TextBox txtPunchDate 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2160
            TabIndex        =   8
            Tag             =   "D"
            Text            =   " "
            Top             =   1500
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtPunchTime 
            Height          =   375
            Left            =   2160
            TabIndex        =   10
            Top             =   2190
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin VB.Label lblEmpCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empoyee"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   270
            TabIndex        =   3
            Top             =   570
            Width           =   825
         End
         Begin VB.Label lblEmpNameCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee  "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   270
            TabIndex        =   5
            Top             =   1050
            Width           =   990
         End
         Begin VB.Label lblLostDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date  "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   270
            TabIndex        =   7
            Top             =   1620
            Width           =   525
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time  "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   255
            TabIndex        =   9
            Top             =   2220
            Width           =   540
         End
         Begin VB.Shape Sh1 
            BorderColor     =   &H80000005&
            BorderWidth     =   2
            Height          =   3015
            Left            =   60
            Top             =   180
            Width           =   4485
         End
         Begin MSForms.ComboBox cboEmpCode 
            Height          =   375
            Left            =   2160
            TabIndex        =   4
            Top             =   510
            Width           =   1335
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "2355;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblEmpName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "a"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2190
            TabIndex        =   6
            Top             =   1020
            Width           =   90
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   3255
         Left            =   -74940
         TabIndex        =   1
         Top             =   360
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   5741
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         AllowBigSelection=   0   'False
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmLost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strLostT_Punch As String, strLostDate As String
''
Dim adrsC As New ADODB.Recordset

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close

adrsC.Open "Select * From NewCaptions Where ID Like '32%'", VstarDataEnv.cnDJConn, adOpenStatic
Me.Caption = NewCaptionTxt("32001", adrsC)              '' Form caption
TB1.TabCaption(0) = NewCaptionTxt("00013", adrsMod)       '' List
TB1.TabCaption(1) = NewCaptionTxt("00014", adrsMod)       '' Details
lblTime.Caption = NewCaptionTxt("32003", adrsC)         '' Time of Punch
lblLostDate.Caption = NewCaptionTxt("32002", adrsC)     '' Date of Punch
lblEmpNameCap.Caption = NewCaptionTxt("32004", adrsC)   '' Employee Name
lblempcode.Caption = NewCaptionTxt("00061", adrsMod)      '' Employee Code
Call SetGButtonCap(Me)                           '' Sets Appropriate Caption for the Buttons
Call CapGrid                                '' Captions For Grid
End Sub

Private Sub FillGrid()          '' Fills the Lost Grid
On Error GoTo ERR_P
Dim intCounter As Integer
adrsDept1.Requery               '' Requeries the Recordset for any Updated Values
'' Put Appropriate Rows in the Grid
If adrsDept1.EOF Then
    MSF1.Rows = 1
    TB1.TabEnabled(1) = False       '' Disables Tab 1 If no Records are Found
    Exit Sub
End If
MSF1.Rows = adrsDept1.RecordCount + 1   '' Sets Rows Appropriately
adrsDept1.MoveFirst
For intCounter = 1 To adrsDept1.RecordCount     '' Fills the Grid
    With MSF1
        .TextMatrix(intCounter, 0) = adrsDept1("Empcode")
        .TextMatrix(intCounter, 1) = DateDisp(adrsDept1("date"))
        .TextMatrix(intCounter, 2) = IIf(IsNull(adrsDept1("t_punch")), "0.00", _
                                     Format(adrsDept1("t_punch"), "0.00"))
    End With
    adrsDept1.MoveNext
Next
TB1.TabEnabled(1) = True        '' If Records Found Enables the TAB 1
TB1.Tab = 0
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub CapGrid()           '' Gives the Captions to the Grid
With MSF1
    '' Sets the Column Widhts
    .ColWidth(0) = .ColWidth(0) * 1.4
    .ColWidth(1) = .ColWidth(1) * 1.32
    .ColWidth(2) = .ColWidth(2) * 1.2
    '' Sets the Column Alignment
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    '' Sets the Appropriate Captions
    .TextMatrix(0, 0) = NewCaptionTxt("00061", adrsMod)   '' Employee Code 29103
    .TextMatrix(0, 1) = NewCaptionTxt("32002", adrsC)   '' Date of Punch 29104
    .TextMatrix(0, 2) = NewCaptionTxt("32003", adrsC)   '' Time of Punch 29105
End With
End Sub

Private Sub OpenMasterTable()             '' Open the Recordset for the Display purposes
On Error GoTo ERR_P
Dim strTmp As String
If strCurrentUserType = HOD Then
    strTmp = "Select " & strKDate & ",t_punch,Lost.Empcode from Lost,Empmst " & strCurrData & " And Empmst.Empcode = Lost.Empcode order by Lost.Empcode"
Else
    strTmp = "Select " & strKDate & ",t_punch,Lost.Empcode from Lost order by Empcode"
End If
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open strTmp, VstarDataEnv.cnDJConn, adOpenStatic
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Sub Display()       '' Displays the Given Master Records
On Error GoTo ERR_P
cboEmpCode.Text = MSF1.TextMatrix(MSF1.row, 0)  '' Employee Code
txtPunchDate = MSF1.TextMatrix(MSF1.row, 1)     '' Date
txtPunchTime.Text = MSF1.TextMatrix(MSF1.row, 2)      '' Time
'' Get Values in the Temporary Variables
strLostT_Punch = Format(txtPunchTime.Text, "0.00")
strLostDate = txtPunchDate.Text
'' Code to Display Employee Name
lblEmpName.Caption = cboEmpCode.List(cboEmpCode.ListIndex, 1)   '' Employee Name
Exit Sub
ERR_P:
    ShowError ("Display  :: " & Me.Caption)
End Sub

Private Sub cboEmpCode_Click()
On Error GoTo ERR_P
'' Displays Employee Name
If cboEmpCode.Text = "" Then Exit Sub
lblEmpName.Caption = cboEmpCode.List(cboEmpCode.ListIndex, 1)
Exit Sub
ERR_P:
    ShowError ("Employee :: " & Me.Caption)
End Sub

Private Sub cmdAddSave_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        '' Check for Rights
        If Not AddRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        bytMode = 2
        Call ChangeMode
    Case 2          '' Add Mode
        If Not ValidateAddmaster Then Exit Sub  '' Validate For Add
        If Not SaveAddMaster Then Exit Sub      '' Save for Add
        Call SaveAddLog                         '' Save the Add Log
        Call FillGrid                           '' Reflect the Grid
        bytMode = 1
        Call ChangeMode
    Case 3          '' Edit Mode
        If Not ValidateModMaster Then Exit Sub  '' Validate for Edit
        If Not SaveModMaster Then Exit Sub      '' Save for Edit
        Call SaveModLog                         '' Save the Edit Log
        Call FillGrid                           '' Reflect the Grid
        bytMode = 1
        Call ChangeMode
End Select
Exit Sub
ERR_P:
    ShowError ("AddSave :: " & Me.Caption)
End Sub

Private Sub cmdDel_Click()
On Error GoTo ERR_P
'' Check for Rights
If Not DeleteRights Then
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    Exit Sub
Else
    If TB1.TabEnabled(1) = False Then Exit Sub
    If TB1.Tab = 0 Then                         '' Do not Display Record if
        If TB1.TabEnabled(1) Then TB1.Tab = 1   '' Already Displayed
    End If
    If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion) _
    = vbYes Then        '' Delete the Record
        Call AddToLostLog("DEL")        ''Adding to Log
        VstarDataEnv.cnDJConn.Execute "delete from lost where Empcode=" & "'" & _
        cboEmpCode.Text & "'" & " and " & strKDate & "=" & _
        strDTEnc & DateCompStr(strLostDate) & strDTEnc & " and t_punch=" & strLostT_Punch
        Call AddActivityLog(lgDelete_Action, 1, 22)     '' Delete Log
        Call AuditInfo("DELETE", Me.Caption, "Delete Lost Enty Of Employee " & cboEmpCode.Text)
    End If
    Call FillGrid       '' Reflect the Grid
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
ERR_P:
    ShowError ("Delete Record :: " & Me.Caption)
End Sub

Private Sub cmdEditCan_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        If TB1.TabEnabled(1) = False Then Exit Sub
        '' Check for Rights
        If Not EditRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        bytMode = 3
        Call ChangeMode
    Case 2       '' Add Mode
        If MSF1.Rows = 1 Then
            TB1.TabEnabled(1) = False
            TB1.Tab = 0
        End If
        bytMode = 1
        Call ChangeMode
    Case 3      '' Edit Mode
        bytMode = 1
        Call ChangeMode
End Select
Exit Sub
ERR_P:
    ShowError ("EditCancel :: " & Me.Caption)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)        '' Set the Form Icon
Call SetToolTipText(Me)     '' Set the ToolTipText
Call RetCaptions            '' Retreive Captions
Call OpenMasterTable        '' Open Master Table
Call FillGrid               '' Fill Lost Grid
Call FillComboEmp           '' Fill EmployeeCombo
TB1.Tab = 0                 '' Set the Tab to List
Call GetRights              '' Gets Rights for the Operations
bytMode = 1
Call ChangeMode             '' Take Action on the Appropriate Mode
End Sub

Private Sub FillComboEmp()
On Error GoTo ERR_P
Call ComboFill(cboEmpCode, 1, 2)    '' Fill Employee Code Combo
Exit Sub
ERR_P:
    ShowError ("Fill Employee Combo :: " & Me.Caption)
End Sub

Private Sub ChangeMode()
Select Case bytMode
    Case 1  '' View
        Call ViewAction
    Case 2  '' Add
        Call AddAction
    Case 3  '' Modify
        Call EditAction
End Select
End Sub

Private Sub MSF1_DblClick()
If TB1.TabEnabled(1) = False Then Exit Sub
TB1.Tab = 1
End Sub

Private Sub GetRights()     '' Gets and Sets the Appropriate Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 15, 1)
If Mid(strTmp, 1, 1) = "1" Then AddRights = True
If Mid(strTmp, 2, 1) = "1" Then EditRights = True
If Mid(strTmp, 3, 1) = "1" Then DeleteRights = True
Exit Sub
ERR_P:
    ShowError ("Rights ::" & Me.Caption)
    AddRights = False
    EditRights = False
    DeleteRights = False
End Sub

Private Sub ViewAction()    '' Procedure for Viewing Mode
'' Enable Needed Controls
cmdAddSave.Enabled = True       '' Enable ADD/SAVE Button
cmdEditCan.Enabled = True       '' Enable Edit/Cancel Button
cmdDel.Enabled = True           '' Enable Disable Button
'' Disable Needed Controls
txtPunchDate.Enabled = False    '' Disable Date TextBox
txtPunchTime.Enabled = False    '' Disable Time TextBox
cboEmpCode.Enabled = False      '' Disable Employee Code Combo
'' Give Captions to the Needed Controls
Call SetGButtonCap(Me)
TB1.Tab = 0     '' Set the Tab to the First Tab
End Sub

Private Sub AddAction()     '' Procedure for Addition Mode
If TB1.TabEnabled(1) = False Then TB1.TabEnabled(1) = True
TB1.Tab = 1
'' Enable Necessary Controls
txtPunchDate.Enabled = True     '' Enable Date TextBox
txtPunchTime.Enabled = True     '' Enable Time TextBox
cboEmpCode.Enabled = True       '' Enable Employee Code Combo
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Clear Necessary Controls
cboEmpCode.Value = ""
lblEmpName.Caption = ""             '' Clear Employee Name
txtPunchDate.Text = DateDisp(Date)  '' Clear Date Control
txtPunchTime.Text = ""              '' Clear Time Control
cboEmpCode.SetFocus                 '' Set Focus on the Employee ComboBox
End Sub

Private Sub EditAction()    '' Procedure for Edit Mode
'' Enable Necessary Controls
txtPunchDate.Enabled = True     '' Enable Date TextBox
txtPunchTime.Enabled = True     '' Enable Time TextBox
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
txtPunchDate.SetFocus       '' Set Focus on the Date TextBox
If TB1.Tab = 1 Then
    Exit Sub
Else
    If TB1.TabEnabled(1) Then TB1.Tab = 1
End If
End Sub

Private Function ValidateAddmaster() As Boolean     '' Validate If in Addition Mode
On Error GoTo ERR_P
ValidateAddmaster = True
If cboEmpCode.Text = "" Then
    MsgBox NewCaptionTxt("32005", adrsC), vbExclamation
    cboEmpCode.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If txtPunchDate.Text = "" Then
    MsgBox NewCaptionTxt("00072", adrsMod), vbExclamation
    txtPunchDate.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
Select Case Val(txtPunchTime.Text)
    Case Is <= 0
        MsgBox NewCaptionTxt("32006", adrsC), vbExclamation
        txtPunchTime.SetFocus
        ValidateAddmaster = False
        Exit Function
    Case Is > 23.59
        MsgBox NewCaptionTxt("00025", adrsMod), vbExclamation
        txtPunchTime.SetFocus
        ValidateAddmaster = False
        Exit Function
End Select
If Val(txtPunchTime.Text) - Int(Val(txtPunchTime.Text)) > 0.59 Then
    MsgBox NewCaptionTxt("00024", adrsMod), vbExclamation
    txtPunchTime.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster :: " & Me.Caption)
End Function

Private Function ValidateModMaster() As Boolean     '' Validate If in Edit Mode
On Error GoTo ERR_P
ValidateModMaster = True
If txtPunchDate.Text = "" Then
    MsgBox NewCaptionTxt("00072", adrsMod), vbExclamation
    txtPunchDate.SetFocus
    ValidateModMaster = False
    Exit Function
End If
Select Case Val(txtPunchTime.Text)
    Case Is <= 0
        MsgBox NewCaptionTxt("32006", adrsC), vbExclamation
        txtPunchTime.SetFocus
        ValidateModMaster = False
        Exit Function
    Case Is > 23.59
        MsgBox NewCaptionTxt("00025", adrsMod), vbExclamation
        txtPunchTime.SetFocus
        ValidateModMaster = False
        Exit Function
    Case Val(txtPunchTime.Text) - Int(Val(txtPunchTime.Text)) > 0.59
        MsgBox NewCaptionTxt("00024", adrsMod), vbExclamation
        txtPunchTime.SetFocus
        ValidateModMaster = False
        Exit Function
End Select
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
End Function

Private Function SaveAddMaster() As Boolean
On Error GoTo ERR_P
SaveAddMaster = True        '' Insert
VstarDataEnv.cnDJConn.Execute "insert into Lost (Empcode," & strKDate & ",t_punch) values (" & _
"'" & Trim(cboEmpCode.Text) & "'" & "," & strDTEnc & DateSaveIns(txtPunchDate.Text) & strDTEnc & "," & _
txtPunchTime.Text & ")"
''For Mauritius 10-07-2003
''Add to LostLog
strLostDate = txtPunchDate.Text
strLostT_Punch = txtPunchTime.Text
Call AddToLostLog("ADD")
Exit Function
ERR_P:
    SaveAddMaster = False
    ShowError ("SaveAddMaster :: " & Me.Caption)
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
SaveModMaster = True        '' Update
VstarDataEnv.cnDJConn.Execute "update lost set " & strKDate & "=" & strDTEnc & _
DateSaveIns(txtPunchDate.Text) & strDTEnc & "," & "t_punch=" & _
txtPunchTime.Text & " where Empcode=" & "'" & cboEmpCode.Text & "'" & " and " & strKDate & "=" & _
strDTEnc & DateCompStr(strLostDate) & strDTEnc & " and " & "t_punch=" & strLostT_Punch
''For Mauritius 10-07-2003
''Add to LostLog for Mod
Call AddToLostLog("MOD")
Exit Function
ERR_P:
    SaveModMaster = False
    ShowError ("SaveModMaster :: " & Me.Caption)
End Function

Private Sub TB1_Click(PreviousTab As Integer)
If bytMode = 2 Then Exit Sub
If PreviousTab = 1 Then Exit Sub
MSF1.Col = 0
If MSF1.Text = NewCaptionTxt("00061", adrsMod) Then Exit Sub
Call Display
End Sub

Private Sub txtPunchDate_Click()
varCalDt = ""
varCalDt = Trim(txtPunchDate.Text)
txtPunchDate.Text = ""
Call ShowCalendar
End Sub

Private Sub txtPunchDate_GotFocus()
    Call GF(txtPunchDate)
End Sub

Private Sub txtPunchDate_KeyPress(KeyAscii As Integer)
     Call CDK(txtPunchDate, KeyAscii)
End Sub

Private Sub txtPunchDate_Validate(Cancel As Boolean)
    If Not ValidDate(txtPunchDate) Then txtPunchDate.SetFocus: Cancel = True
End Sub

Private Sub txtPunchTime_GotFocus()
    Call GF(txtPunchTime)
End Sub

Private Sub txtPunchTime_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtPunchTime)
End If
End Sub

Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo ERR_P
Call AddActivityLog(lgADD_MODE, 1, 22)     '' Add Activity
Call AuditInfo("ADD", Me.Caption, "Add Lost Enty Of Employee " & cboEmpCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 22)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edit Lost Enty Of Employee " & cboEmpCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub AddToLostLog(ByVal strOperation As String)
On Error GoTo ERR_P
Dim sngTime As Single
sngTime = IIf(Hour(Time) = 0, 24, Hour(Time)) & "." & Format(Minute(Time), "00") '' Time
VstarDataEnv.cnDJConn.Execute "Delete from LostLog where LEmpcode='" & cboEmpCode.Text & "' and " & _
" LDATE =" & strDTEnc & DateCompStr(strLostDate) & strDTEnc & " and LT_punch=" & _
strLostT_Punch

VstarDataEnv.cnDJConn.Execute "insert into LostLog Values('" & strOperation & "','" & _
IIf(UCase(Trim(UserName)) = UCase(strPrintUser), "*****", UserName) & "'," & strDTEnc & _
DateSaveIns(CStr(Date)) & strDTEnc & "," & sngTime & ",'" & cboEmpCode.Text & "'," & _
strDTEnc & DateSaveIns(strLostDate) & strDTEnc & "," & strLostT_Punch & ")"

Exit Sub
ERR_P:
    ShowError (" :: " & Me.Caption)
End Sub
