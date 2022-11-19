VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDesg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Designation Master"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5205
      TabIndex        =   3
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6315
      TabIndex        =   4
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   2400
      Width           =   1100
   End
   Begin VB.Frame frComp 
      Height          =   2295
      Left            =   4080
      TabIndex        =   7
      Top             =   0
      Width           =   4500
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   1005
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desig Code"
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
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   600
         TabIndex        =   8
         Top             =   600
         Width           =   510
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2865
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   5054
      _Version        =   393216
      FixedCols       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmDesg"
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
Call GetRights
bytMode = 1
Call ChangeMode
End Sub

Private Sub RetCaptions()
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
    .TextMatrix(0, 0) = "Code"
    .TextMatrix(0, 1) = "Description"
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
adrsDept1.Open "Select DesigCode,DesigName from frmDesignation Order by DesigCode", _
ConMain, adOpenStatic
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
    Exit Sub
End If
MSF1.Rows = adrsDept1.RecordCount + 1
adrsDept1.MoveFirst
For intCounter = 1 To adrsDept1.RecordCount
    With MSF1
        .TextMatrix(intCounter, 0) = adrsDept1("DesigCode")
        .TextMatrix(intCounter, 1) = IIf(IsNull(adrsDept1("DesigName")), "", adrsDept1("DesigName"))
    End With
    adrsDept1.MoveNext
Next

Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub Display()
On Error GoTo ERR_P
adrsDept1.Requery
If adrsDept1.EOF Then Exit Sub
adrsDept1.MoveFirst
adrsDept1.Find "DesigCode='" & MSF1.TextMatrix(MSF1.Row, 0) & "'"
If Not (adrsDept1.EOF) Then
    txtCode.Text = MSF1.TextMatrix(MSF1.Row, 0)
    txtDesc.Text = MSF1.TextMatrix(MSF1.Row, 1)
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
    MsgBox "Code cannot be blank ", vbExclamation
    ValidateAddmaster = False
    txtCode.SetFocus
    Exit Function
End If
If MSF1.Rows > 1 Then
    adrsDept1.MoveFirst
    adrsDept1.Find "DesigCode='" & txtCode.Text & "'"
    If Not adrsDept1.EOF Then
        MsgBox "Code Already Exists ", vbExclamation
        txtCode.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If
If Trim(txtDesc.Text) = "" Then
    MsgBox "Name cannot be blank ", vbExclamation
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
Call Display
End Sub

Private Sub AddAction()
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
ConMain.Execute "insert into FrmDesignation values ('" & _
Trim(txtCode.Text) & "','" & txtDesc.Text & "')"
Exit Function
ERR_P:
    Select Case Err.Number
        Case -2147217900
            MsgBox "Code Already Exists ", vbExclamation
        Case Else
            ShowError ("SaveAddMaster :: " & Me.Caption)
    End Select
    SaveAddMaster = False
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
SaveModMaster = True        '' Update
ConMain.Execute "update frmDesignation set DesigName='" & txtDesc.Text & _
"' where DesigCode=" & txtCode.Text
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

    If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion) _
    = vbYes Then
        Dim adrsDesig As New ADODB.Recordset
        If adrsDesig.State = 1 Then adrsDesig.Close
        adrsDesig.Open "Select distinct designatn from Empmst where designatn= " & txtCode.Text & "", ConMain, adOpenStatic
        If Not adrsDesig.EOF Then
            MsgBox "Designation Cannot be deleted because employees belong to this Designation.", vbCritical, Me.Caption
            Exit Sub
        Else
            ConMain.Execute "delete from frmDesignation where DesigCode=" & txtCode.Text
            Call AddActivityLog(lgDelete_Action, 1, 3)
            Call AuditInfo("DELETE", Me.Caption, "Designation Deleted: " & txtCode.Text)
        End If
    End If
    Call FillGrid
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
ERR_P:
    If Err.Number = -2147467259 Or Err.Number = -2147217900 Or Err.Number = -2147217873 Then
            MsgBox "Designation Cannot be deleted because employees belong to this Designation.", vbCritical, Me.Caption
            Exit Sub
    End If
    ShowError ("Delete Record :: " & Me.Caption)
End Sub

Private Sub cmdEditCan_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode

        If Not EditRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        Else
            bytMode = 3
            Call ChangeMode
        End If
    Case 2       '' Add Mode
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

Private Sub MSF1_DblClick()
Call Display
End Sub

Private Sub txtCode_GotFocus()
    Call GF(txtCode)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 2))))
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
Call AuditInfo("ADD", Me.Caption, "Designation Added: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModLog()
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 3)
Call AuditInfo("UPDATE", Me.Caption, "Designation Edited: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

