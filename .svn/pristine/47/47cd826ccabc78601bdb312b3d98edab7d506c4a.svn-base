VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMachineMst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine Master"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5445
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command1"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3360
      Width           =   1395
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1380
      TabIndex        =   2
      Top             =   3360
      Width           =   1395
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2730
      TabIndex        =   1
      Top             =   3360
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   3360
      Width           =   1395
   End
   Begin TabDlg.SSTab TB1 
      Height          =   4215
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "List"
      TabPicture(0)   =   "frmMachineMst.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSF1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Details"
      TabPicture(1)   =   "frmMachineMst.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblDesc"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblCode"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtDesc"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtCode"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         MaxLength       =   5
         TabIndex        =   6
         Top             =   960
         Width           =   1005
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2130
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1290
         Width           =   3255
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2865
         Left            =   -74640
         TabIndex        =   7
         Top             =   360
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   5054
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Machine Code"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   990
         Width           =   1035
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Machinen Name"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmMachineMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adrsC As New ADODB.Recordset

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)
Call RetCaptions
Call OpenMasterTable
Call FillGrid
TB1.Tab = 0
Call GetRights
bytMode = 1
Call ChangeMode
End Sub

Private Sub RetCaptions()
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '12%'", VstarDataEnv.cnDJConn, adOpenStatic
TB1.TabCaption(0) = NewCaptionTxt("00013", adrsMod)
TB1.TabCaption(1) = NewCaptionTxt("00014", adrsMod)
Call SetGButtonCap(Me)
Call CapGrid
Exit Sub
ERR_P:
    ShowError ("RetCaption :: " & Me.Caption)
End Sub

Private Sub CapGrid()
With MSF1
    .ColWidth(0) = .ColWidth(0) * 1.25
    .ColWidth(1) = .ColWidth(1) * 3.4
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .TextMatrix(0, 0) = "Machine Code"
    .TextMatrix(0, 1) = "Machine Name"
End With
End Sub

Private Sub GetRights()
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 6)
If Mid(strTmp, 1, 1) = "1" Then AddRights = True
If Mid(strTmp, 2, 1) = "1" Then EditRights = True
If Mid(strTmp, 3, 1) = "1" Then DeleteRights = True
Exit Sub
ERR_P:
    ShowError ("Rights :: " & Me.Caption)
    AddRights = False
    EditRights = False
    DeleteRights = False
End Sub

Private Sub OpenMasterTable()
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select MachineCode,MachineName from Machine Order by MachineCode", _
VstarDataEnv.cnDJConn, adOpenStatic
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Sub FillGrid()
On Error GoTo ERR_P
Dim intCounter As Integer
adrsDept1.Requery
If adrsDept1.EOF Then
    MSF1.Rows = 1
    TB1.TabEnabled(1) = False
    Exit Sub
End If
MSF1.Rows = adrsDept1.RecordCount + 1
adrsDept1.MoveFirst
For intCounter = 1 To adrsDept1.RecordCount
    With MSF1
        .TextMatrix(intCounter, 0) = adrsDept1("MachineCode")
        .TextMatrix(intCounter, 1) = IIf(IsNull(adrsDept1("MachineName")), "", adrsDept1("MachineName"))
    End With
    adrsDept1.MoveNext
Next
TB1.TabEnabled(1) = True
TB1.Tab = 0
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub Display()
On Error GoTo ERR_P
adrsDept1.MoveFirst
adrsDept1.Find "MachineCode='" & MSF1.TextMatrix(MSF1.row, 0) & "'"
If Not (adrsDept1.EOF) Then
    txtCode.Text = MSF1.TextMatrix(MSF1.row, 0)
    txtDesc.Text = MSF1.TextMatrix(MSF1.row, 1)
Else
    txtCode = ""
    txtDesc = ""
    MsgBox "Company not Found ", vbCritical
    Exit Sub
End If
Exit Sub
ERR_P:
    ShowError ("Display :: " & Me.Caption)
End Sub

Private Sub MSF1_dblClick()
If TB1.TabEnabled(1) = False Then Exit Sub
TB1.Tab = 1
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

Private Function ValidateAddmaster() As Boolean
On Error GoTo ERR_P
ValidateAddmaster = True
If Trim(txtCode.Text) = "" Then
    MsgBox "Machine Code cannot be blank ", vbExclamation
    ValidateAddmaster = False
    txtCode.SetFocus
    Exit Function
End If
If MSF1.Rows > 1 Then
    adrsDept1.MoveFirst
    adrsDept1.Find "MachineCode='" & txtCode.Text & "'"
    If Not adrsDept1.EOF Then
        MsgBox "Machine Code Already Exists ", vbExclamation
        txtCode.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If
If Trim(txtDesc.Text) = "" Then
    MsgBox "Machine Name cannot be blank ", vbExclamation
    ValidateAddmaster = False
    txtDesc.SetFocus
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster :: " & Me.Caption)
    ValidateAddmaster = False
End Function

Private Function ValidateModMaster() As Boolean
On Error GoTo ERR_P
ValidateModMaster = True
If Trim(txtDesc.Text) = "" Then
    MsgBox "Company Name cannot be blank ", vbExclamation
    ValidateModMaster = False
    txtDesc.SetFocus
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
    ValidateModMaster = False
End Function

Private Sub ViewAction()
cmdAddSave.Enabled = True
'If InVar.bytCom = "1" Then cmdAddSave.Enabled = False
cmdEditCan.Enabled = True
cmdDel.Enabled = True
txtCode.Enabled = False
txtDesc.Enabled = False
Call SetGButtonCap(Me)
TB1.Tab = 0
End Sub

Private Sub AddAction()
If TB1.TabEnabled(1) = False Then TB1.TabEnabled(1) = True
TB1.Tab = 1
txtCode.Enabled = True
txtDesc.Enabled = True
cmdDel.Enabled = False
Call SetGButtonCap(Me, 2)
txtCode.Text = ""
txtDesc.Text = ""
txtCode.SetFocus
End Sub

Private Sub EditAction()
txtDesc.Enabled = True
cmdAddSave.Enabled = True
Call SetGButtonCap(Me, 2)
cmdDel.Enabled = False
txtDesc.SetFocus
If TB1.Tab = 1 Then
    Exit Sub
Else
    If TB1.TabEnabled(1) Then TB1.Tab = 1
End If
End Sub

Private Sub cmdAddSave_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
    If Not AddRights Then
        MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
        Exit Sub
    Else
        bytMode = 2
        Call ChangeMode
    End If
    Case 2          '' Add Mode
        If Not ValidateAddmaster Then Exit Sub
        If Not SaveAddMaster Then Exit Sub
        Call SaveAddLog
        Call FillGrid
        bytMode = 1
        Call ChangeMode
    Case 3          '' Edit Mode
        If Not ValidateModMaster Then Exit Sub
        If Not SaveModMaster Then Exit Sub
        Call SaveModLog
        Call FillGrid
        bytMode = 1
        Call ChangeMode
End Select
Exit Sub
ERR_P:
    ShowError ("AddSave :: " & Me.Caption)
End Sub

Private Function SaveAddMaster() As Boolean
On Error GoTo ERR_P
SaveAddMaster = True        '' Insert
VstarDataEnv.cnDJConn.Execute "insert into machine values ('" & _
Trim(txtCode.Text) & "','" & txtDesc.Text & "')"
Exit Function
ERR_P:
    Select Case Err.Number
        Case -2147217900
            MsgBox "Machine Code Already Exists ", vbExclamation
        Case Else
            ShowError ("SaveAddMaster :: " & Me.Caption)
    End Select
    SaveAddMaster = False
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
SaveModMaster = True        '' Update
VstarDataEnv.cnDJConn.Execute "update Machine set MachineName='" & txtDesc.Text & _
"' where MachineCode='" & txtCode.Text & "'"
Exit Function
ERR_P:
    SaveModMaster = False
    ShowError ("SaveModMaster :: " & Me.Caption)
End Function

Private Sub cmdDel_Click()
On Error GoTo ERR_P
If Not DeleteRights Then
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    Exit Sub
Else
    If TB1.TabEnabled(1) = False Then Exit Sub
    If TB1.Tab = 0 Then
        If TB1.TabEnabled(1) Then TB1.Tab = 1
    End If
    If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion) _
    = vbYes Then
        Dim adrsMach As New ADODB.Recordset
        If adrsMach.State = 1 Then adrsMach.Close
        adrsMach.Open "Select distinct MachineCode from Empmst where MachineCode='" & txtCode.Text & "'", VstarDataEnv.cnDJConn, adOpenStatic
        If Not adrsMach.EOF Then
            MsgBox "Machine Cannot be deleted because employees belong to this Machine.", vbCritical, Me.Caption
            Exit Sub
        Else
            VstarDataEnv.cnDJConn.Execute "delete from Machine where MachineCode='" & txtCode.Text & "'"
        Call AddActivityLog(lgDelete_Action, 1, 3)
        Call AuditInfo("DELETE", Me.Caption, "Machine Deleted: " & txtCode.Text)
        End If
    End If
    Call FillGrid
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
ERR_P:
    If Err.Number = -2147467259 Or Err.Number = -2147217900 Then
            MsgBox "Machine Cannot be deleted because employees belong to this Machine.", vbCritical, Me.Caption
            Exit Sub
    End If
    ShowError ("Delete Record :: " & Me.Caption)
End Sub

Private Sub cmdEditCan_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        If TB1.TabEnabled(1) = False Then Exit Sub
        If Not EditRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        Else
            bytMode = 3
            Call ChangeMode
        End If
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

Private Sub TB1_Click(PreviousTab As Integer)
If bytMode = 2 Then Exit Sub
If PreviousTab = 1 Then Exit Sub
MSF1.Col = 0
If MSF1.Text = "Company Code" Then Exit Sub
Call Display
End Sub

Private Sub txtCode_GotFocus()
    Call GF(txtCode)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
'Else
'    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 2))))
End If
End Sub

Private Sub txtDesc_GotFocus()
    Call GF(txtDesc)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 7))))
End If
End Sub

Private Sub SaveAddLog()
On Error GoTo ERR_P
Call AddActivityLog(lgADD_MODE, 1, 3)
Call AuditInfo("ADD", Me.Caption, "Machine Added: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModLog()
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 3)
Call AuditInfo("UPDATE", Me.Caption, "Machine Edited: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

