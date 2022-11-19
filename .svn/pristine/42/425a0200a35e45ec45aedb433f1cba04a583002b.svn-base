VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDept 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5685
      TabIndex        =   4
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6795
      TabIndex        =   5
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   2280
      Width           =   1100
   End
   Begin VB.Frame frDept 
      Height          =   2055
      Left            =   4560
      TabIndex        =   8
      Top             =   120
      Width           =   4545
      Begin VB.TextBox txtStrength 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1050
         Width           =   735
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   1
         Top             =   630
         Width           =   2985
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblStrength 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Strength"
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
         Left            =   480
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         TabIndex        =   10
         Top             =   720
         Width           =   570
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         TabIndex        =   9
         Top             =   300
         Width           =   570
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2775
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColorFixed  =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   2
      ScrollBars      =   2
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
Attribute VB_Name = "frmDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset
Dim prevtab As Integer

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

Private Sub GetRights()     '' Gets and Sets the Appropriate Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 2)
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

Private Sub OpenMasterTable()             '' Open the recordset for the Display purposes
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close

adrsDept1.Open "Select dept," & strKDesc & " from Deptdesc Order by Dept", _
ConMain, adOpenStatic

Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
    'Resume Next
End Sub

Private Sub RetCaptions()
Me.Caption = "Department Master"
lblCode.Caption = "Code"
lblName.Caption = "Name"
lblstrength.Caption = "Strength"
Call SetGButtonCap(Me)                           '' Sets Appropriate Caption for the Buttons
Call CapGrid                                '' Captions For Grid
End Sub

Private Sub CapGrid()           '' Gives the Captions to the Grid
With MSF1
    '' Sets the Column Widhts
    .ColWidth(0) = .ColWidth(0) * 0.78
    .ColWidth(1) = .ColWidth(1) * 4
    '' Sets the Column Alignment
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    '' Sets the Appropriate Captions
    .TextMatrix(0, 0) = "Code"
    .TextMatrix(0, 1) = "Description"
End With
End Sub

Private Sub FillGrid()          '' Fills the Grid
On Error GoTo ERR_P
Dim intCounter As Integer
adrsDept1.Requery               '' Requeries the Recordset for any Updated Values
'' Put Appropriate Rows in the Grid
If adrsDept1.EOF Then
    MSF1.Rows = 1
    Exit Sub
End If
MSF1.Rows = adrsDept1.RecordCount + 1   '' Sets Rows Appropriately
adrsDept1.MoveFirst
For intCounter = 1 To adrsDept1.RecordCount     '' Fills the Grid
    With MSF1
        .TextMatrix(intCounter, 0) = adrsDept1("Dept")
        .TextMatrix(intCounter, 1) = IIf(IsNull(adrsDept1("Desc")), "", adrsDept1("Desc"))
    End With
    adrsDept1.MoveNext
Next
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
'    Resume Next
End Sub

Private Sub Display()       '' Displays the Given Master Records
On Error GoTo ERR_P
adrsDept1.Requery
If adrsDept1.EOF Then Exit Sub
adrsDept1.MoveFirst
adrsDept1.Find "Dept=" & MSF1.TextMatrix(MSF1.Row, 0) & ""
If Not (adrsDept1.EOF) Then
    txtCode.Text = MSF1.TextMatrix(MSF1.Row, 0)     '' Department Code
    txtName.Text = MSF1.TextMatrix(MSF1.Row, 1)     '' Department Name
    Call DispStrength                               '' Department Strength
Else
    txtCode = ""
    txtName = ""
    txtStrength = ""
    MsgBox NewCaptionTxt("20003", adrsC), vbCritical
    Exit Sub
End If
Exit Sub
ERR_P:
    ShowError ("Display :: " & Me.Caption)
    Resume Next
End Sub

Private Sub MSF1_DblClick()
Call Display
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

Private Function ValidateAddmaster() As Boolean     '' Validate If in Addition Mode
On Error GoTo ERR_P
ValidateAddmaster = True
'' Check for Blank DepartMent Code
If Trim(txtCode.Text) = "" Then
    MsgBox NewCaptionTxt("20004", adrsC), vbExclamation
    ValidateAddmaster = False
    txtCode.SetFocus
    Exit Function
End If
'' Check for Existing Department Code
If MSF1.Rows > 1 Then
    adrsDept1.MoveFirst
    adrsDept1.Find "Dept=" & txtCode.Text & ""
    If Not adrsDept1.EOF Then
        MsgBox NewCaptionTxt("20005", adrsC), vbExclamation
        txtCode.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If
'' Check for Blank Department Name
If Trim(txtName.Text) = "" Then
    MsgBox NewCaptionTxt("20006", adrsC), vbExclamation
    ValidateAddmaster = False
    txtName.SetFocus
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster :: " & Me.Caption)
    ValidateAddmaster = False
End Function

Private Function ValidateModMaster() As Boolean     '' Validate If in Addition Mode
On Error GoTo ERR_P
ValidateModMaster = True
'' Check for Blank Department Name
If Trim(txtName.Text) = "" Then
    MsgBox NewCaptionTxt("20006", adrsC), vbExclamation
    ValidateModMaster = False
    txtName.SetFocus
    Exit Function
End If
''If Val(txtStrength.Text) = 0 Then txtStrength.Text = "0"
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
    ValidateModMaster = False
End Function

Private Sub ViewAction()    '' Procedure for Viewing Mode
'' Enable Needed Controls
cmdAddSave.Enabled = True       '' Enable ADD/SAVE Button
cmdEditCan.Enabled = True       '' Enable Edit/Cancel Button
cmdDel.Enabled = True           '' Enable Delete Button
'' Disable Needed Controls
txtCode.Enabled = False         '' Disable Code TextBox
txtName.Enabled = False         '' Disable Name TextBox

'' Give Captions to the Needed Controls
Call SetGButtonCap(Me)
Call Display
End Sub

Private Sub AddAction()     '' Procedure for Addition Mode
'' Enable Necessary Controls
txtCode.Enabled = True      '' Enable Code TextBox
txtName.Enabled = True      '' Enable Name TextBox

'txtapp.Enabled = True
'' Disable Necessary Controls
cmdDel.Enabled = False      '' Disable Delete Button
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Clear Necessary Controls
txtCode.Text = ""       '' Clear Code Control
txtName.Text = ""       '' Clear Name Control
'txtapp.Text = ""
txtStrength.Text = "0"  '' Clear the Strength Control
txtCode.SetFocus        '' Set Focus to the Code TextBox
End Sub

Private Sub EditAction()    '' Procedure for Edit Mode
'' Enable Necessary Controls
txtName.Enabled = True      '' Enable Code TextBox
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
txtName.SetFocus            '' Set Focus on the Name TextBox

End Sub

Private Sub cmdAddSave_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        '' Check for Rights
        If Not AddRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        Else
            bytMode = 2
            Call ChangeMode
        End If
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
    'Resume Next
End Sub

Private Function SaveAddMaster() As Boolean
On Error GoTo ERR_P
SaveAddMaster = True        '' Insert


ConMain.Execute "insert into Deptdesc(Dept," & strKDesc & ") values (" & _
Trim(txtCode.Text) & ",'" & txtName.Text & "')"

''
Exit Function
ERR_P:
    Select Case Err.Number
        Case -2147217900
            MsgBox NewCaptionTxt("20007", adrsC), vbExclamation
        Case Else
            ShowError ("SaveAddMaster :: " & Me.Caption)
    End Select
    SaveAddMaster = False
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
SaveModMaster = True        '' Update


ConMain.Execute "update DeptDesc set " & strKDesc & " ='" & txtName.Text & _
"',strenth='" & txtStrength.Text & "' where Dept=" & txtCode.Text & ""

''
Exit Function
ERR_P:
    SaveModMaster = False
    ShowError ("SaveModMaster :: " & Me.Caption)
End Function

Private Sub cmdDel_Click()
On Error GoTo ERR_P
Dim adrsDumb As New ADODB.Recordset
'' Check for Rights
If Not DeleteRights Then
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    Exit Sub
Else
   
    adrsDumb.Open "Select * from UserAccs where dept ='" & txtCode.Text & "'", _
    ConMain, adOpenStatic, adLockReadOnly
    If Not (adrsDumb.EOF And adrsDumb.BOF) Then
        MsgBox "This department is allotted to some user." & vbCrLf & _
        "You can not delete this department.", vbCritical
        Exit Sub
    End If
    If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion, Me.Caption) _
    = vbYes Then        '' Delete the Record
          ConMain.Execute "delete from Deptdesc where Dept=" & txtCode.Text & ""
          Call AddActivityLog(lgDelete_Action, 1, 9)      '' Delete Log Action
          Call AuditInfo("DELETE", Me.Caption, "Deleted Department: " & txtCode.Text)
    End If
    Call FillGrid       '' Reflect the Grid
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
ERR_P:
    If Err.Number = -2147467259 Or Err.Number = -2147217900 Or Err.Number = -2147217873 Then
            MsgBox "Department Cannot be deleted because employees belong to this department.", vbCritical, Me.Caption
            Exit Sub
    End If
    ShowError ("Delete Record :: " & Me.Caption)
End Sub

Private Sub cmdEditCan_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        '' Check for Rights
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


Private Sub txtName_GotFocus()
    Call GF(txtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 5))))
End If
End Sub


Private Sub txtStrength_GotFocus()
    Call GF(txtStrength)
End Sub

Private Sub txtStrength_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 2)
End If
End Sub

Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo ERR_P
Call AddActivityLog(lgADD_MODE, 1, 9)     '' Add Activity
Call AuditInfo("ADD", Me.Caption, "Added Department: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 9)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edited Department: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub DispStrength()
On Error GoTo ERR_P
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open "Select Count(*) from Empmst where Dept=" & txtCode.Text & "", ConMain _
, adOpenStatic
If (adrsPaid.EOF And adrsPaid.BOF) Then
    txtStrength.Text = "0"
Else
    txtStrength.Text = IIf(IsNull(adrsPaid(0)), "0", CStr(adrsPaid(0)))
End If
Exit Sub
ERR_P:
    ShowError ("Display Strength :: " & Me.Caption)
    txtStrength.Text = "0"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
If cmdEditCan.Caption = "&Cancel " Then

bytMode = 1
Call ChangeMode
Call SetGButtonCap(Me)
Else
Unload Me
End If
End If

End Sub

