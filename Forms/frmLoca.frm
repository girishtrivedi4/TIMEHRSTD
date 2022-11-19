VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLoca 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frComp 
      Height          =   2295
      Left            =   4080
      TabIndex        =   7
      Top             =   0
      Width           =   4500
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   2
         Top             =   690
         Width           =   3255
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   1
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Name"
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
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   570
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code No."
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
         TabIndex        =   8
         Top             =   390
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6315
      TabIndex        =   5
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5205
      TabIndex        =   4
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2400
      Width           =   1100
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2800
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   4948
      _Version        =   393216
      FixedCols       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmLoca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)        '' Set the Form Icon
Call RetCaptions            '' Retreive Captions
Call OpenMasterTable        '' Open Master Table
Call FillGrid               '' Fill Grid
Call GetRights              '' Gets Rights for the Operations
bytMode = 1
Call ChangeMode             '' Take Action on the Appropriate Mode
End Sub

Private Sub RetCaptions()
On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '63%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("63001", adrsC)              '' Form caption
Call SetGButtonCap(Me)                           '' Sets Appropriate Caption for the Buttons
Call CapGrid                                '' Captions For Grid
End Sub

Private Sub CapGrid()           '' Gives the Captions to the Grid
With MSF1
    '' Sets the Column Widhts
    .ColWidth(0) = .ColWidth(0) * 0.6
    .ColWidth(1) = .ColWidth(1) * 3.5
    '' Sets the Column Alignment
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    '' Sets the Appropriate Captions
    .TextMatrix(0, 0) = NewCaptionTxt("00047", adrsMod)   '' Location Code
    .TextMatrix(0, 1) = NewCaptionTxt("00052", adrsMod)   '' Location Description
End With
End Sub

Private Sub OpenMasterTable()             '' Open the recordset for the Display purposes
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select * from Location Order by Location", ConMain, adOpenStatic
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Sub FillGrid()          '' Fills the Lost Grid
On Error GoTo ERR_P
Dim intCounter As Integer
adrsDept1.Requery               '' Requeries the Recordset for any Updated Values
MSF1.Rows = adrsDept1.RecordCount + 1   '' Sets Rows Appropriately
If adrsDept1.EOF Then
    cmdEditCan.Enabled = False
    Exit Sub
End If

'' Put Appropriate Rows in the Grid

adrsDept1.MoveFirst
For intCounter = 1 To adrsDept1.RecordCount     '' Fills the Grid
    With MSF1
        .TextMatrix(intCounter, 0) = adrsDept1("Location")
        .TextMatrix(intCounter, 1) = adrsDept1("LocDesc")
    End With
    adrsDept1.MoveNext
Next

Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub GetRights()     '' Gets and Sets the Appropriate Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 4)
If Mid(strTmp, 1, 1) = "1" Then AddRights = True
If Mid(strTmp, 2, 1) = "1" Then EditRights = True
If Mid(strTmp, 3, 1) = "1" Then DeleteRights = True
Exit Sub
ERR_P:
    ShowError ("GetRights ::" & Me.Caption)
    AddRights = False
    EditRights = False
    DeleteRights = False
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

Private Sub ViewAction()    '' Procedure for Viewing Mode
'' Enable Needed Controls
cmdAddSave.Enabled = True       '' Enable ADD/SAVE Button
cmdEditCan.Enabled = True       '' Enable Edit/Cancel Button
cmdDel.Enabled = True           '' Enable Disable Button
'' Disable Needed Controls
txtCode.Enabled = False         '' Disable Code TextBox
txtDesc.Enabled = False         '' Disable Description TextBox
'' Give Captions to the Needed Controls
Call SetGButtonCap(Me)
Call Display
End Sub

Private Sub AddAction()     '' Procedure for Addition Mode
txtCode.Enabled = True          '' Disable Code TextBox
txtDesc.Enabled = True          '' Disable Description TextBox
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Clear Necessary Controls
txtCode.Text = ""               '' Clear Code TextBox
txtDesc.Text = ""               '' Clear Description TextBox
txtCode.SetFocus                '' Set Focus to the Location Code TextBox
End Sub

Private Sub EditAction()    '' Procedure for Edit Mode
'' Enable Necessary Controls
txtDesc.Enabled = True          '' Enable Description TextBox
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
txtDesc.SetFocus                '' Set Focus on the Description TextBox

End Sub

Private Function ValidateAddmaster() As Boolean     '' Validate If in Addition Mode
On Error GoTo ERR_P
ValidateAddmaster = True
If Trim(txtCode.Text) = "" Then
    MsgBox NewCaptionTxt("63002", adrsC), vbExclamation
    txtCode.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If MSF1.Rows > 1 Then
    adrsDept1.MoveFirst
    adrsDept1.Find "Location=" & txtCode.Text
    If Not adrsDept1.EOF Then
        MsgBox NewCaptionTxt("63003", adrsC), vbExclamation
        txtCode.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If
If Trim(txtDesc.Text) = "" Then
    MsgBox NewCaptionTxt("63004", adrsC), vbExclamation
    txtDesc.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster :: " & Me.Caption)
End Function

Private Function ValidateModMaster() As Boolean     '' Validate If in Addition Mode
On Error GoTo ERR_P
ValidateModMaster = True
If Trim(txtDesc.Text) = "" Then
    MsgBox NewCaptionTxt("63004", adrsC), vbExclamation
    txtDesc.SetFocus
    ValidateModMaster = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
End Function

Private Function SaveAddMaster() As Boolean
On Error GoTo ERR_P
SaveAddMaster = True        '' Insert
ConMain.Execute "insert into Location values (" & _
Trim(txtCode.Text) & ",'" & Trim(txtDesc.Text) & "')"
Exit Function
ERR_P:
    Select Case Err.Number
        Case -2147217900
            MsgBox NewCaptionTxt("63003", adrsC), vbExclamation
        Case Else
            ShowError ("SaveAddMaster :: " & Me.Caption)
    End Select
    SaveAddMaster = False
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
SaveModMaster = True        '' Update
ConMain.Execute "update Location set LocDesc='" & _
Trim(txtDesc.Text) & "' where Location=" & txtCode.Text
Exit Function
ERR_P:
    SaveModMaster = False
    ShowError ("SaveModMaster :: " & Me.Caption)
End Function

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

Private Sub cmdEditCan_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        If Not EditRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        bytMode = 3
        Call ChangeMode
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

Private Sub cmdDel_Click()
On Error GoTo ERR_P
'' Check for Rights
If Not DeleteRights Then
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    Exit Sub
Else
    If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion) _
    = vbYes Then        '' Delete the Record
        ConMain.Execute "delete from Location where Location=" & _
        txtCode.Text
        Call AddActivityLog(lgDelete_Action, 1, 10)     '' Delete Log
        Call AuditInfo("DELETE", Me.Caption, "Deleted Location: " & txtCode.Text)
    End If
    Call FillGrid       '' Reflect the Grid
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
ERR_P:
    If Err.Number = -2147467259 Or Err.Number = -2147217900 Or Err.Number = -2147217873 Then
            MsgBox "Location Cannot be deleted because Employees belong to this Location.", vbCritical, Me.Caption
            Exit Sub
    End If
    ShowError ("Delete Record :: " & Me.Caption)
End Sub

Private Sub Display()       '' Displays the Given Master Records
On Error GoTo ERR_P
adrsDept1.Requery
If adrsDept1.EOF Then
    cmdEditCan.Enabled = False
    Exit Sub
End If

txtCode.Text = MSF1.TextMatrix(MSF1.Row, 0) '' Location Code
txtDesc = MSF1.TextMatrix(MSF1.Row, 1)      '' Description
adrsDept1.MoveFirst
adrsDept1.Find "Location=" & txtCode.Text
If adrsDept1.EOF Then
    txtCode.Text = ""
    txtDesc.Text = ""
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
ERR_P:
    ShowError ("Display  :: " & Me.Caption)
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
    KeyAscii = KeyPressCheck(KeyAscii, 2)
End If
End Sub

Private Sub txtDesc_GotFocus()
    Call GF(txtDesc)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 5))))
End If
End Sub

Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo ERR_P
Call AddActivityLog(lgADD_MODE, 1, 10)     '' Add Activity
Call AuditInfo("ADD", Me.Caption, "Added Location: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 10)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edited Location: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub
