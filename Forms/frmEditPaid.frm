VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmEditPaid 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command 1"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   4770
      Width           =   3015
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command 1"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   4770
      Width           =   3015
   End
   Begin TabDlg.SSTab TB1 
      Height          =   3705
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6535
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "List"
      TabPicture(0)   =   "frmEditPaid.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSF1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Details"
      TabPicture(1)   =   "frmEditPaid.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblPaid"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblDate"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblremark"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtremark"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtPaid"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtDate"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2700
         TabIndex        =   3
         Tag             =   "D"
         Text            =   " "
         Top             =   1320
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtPaid 
         Height          =   375
         Left            =   2700
         TabIndex        =   4
         Top             =   1830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         HideSelection   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   4035
         Left            =   -75000
         TabIndex        =   10
         Top             =   360
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   7117
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
         HighLight       =   2
         ScrollBars      =   0
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txtremark 
         Height          =   375
         Left            =   2700
         TabIndex        =   5
         Top             =   2350
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         HideSelection   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSForms.Label lblremark 
         Height          =   210
         Left            =   1020
         TabIndex        =   13
         Top             =   2400
         Width           =   1005
         VariousPropertyBits=   8388627
         Caption         =   "Remarks :"
         Size            =   "1773;370"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblDate 
         Height          =   210
         Left            =   1020
         TabIndex        =   11
         Top             =   1410
         Width           =   1245
         VariousPropertyBits=   8388627
         Caption         =   "Date :"
         Size            =   "2196;370"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblPaid 
         Height          =   210
         Left            =   1020
         TabIndex        =   12
         Top             =   1920
         Width           =   1005
         VariousPropertyBits=   8388627
         Caption         =   "Paid Days :"
         Size            =   "1773;370"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSForms.ComboBox cboDept 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   120
      Width           =   3675
      VariousPropertyBits=   612390939
      DisplayStyle    =   3
      Size            =   "6482;661"
      TextColumn      =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblDeptCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      TabIndex        =   0
      Top             =   187
      Width           =   1005
   End
   Begin MSForms.Label lblCode 
      Height          =   240
      Left            =   600
      TabIndex        =   9
      Top             =   617
      Width           =   885
      VariousPropertyBits=   276824083
      Caption         =   "Employee"
      Size            =   "1561;423"
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboCode 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   550
      Width           =   3615
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "6376;661"
      BoundColumn     =   0
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmEditPaid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim strLvFileName As String
''
Dim adrsC As New ADODB.Recordset

Private Sub CapGrid()           '' Gives the Captions to the Grid
With MSF1
    '' Sets the Column Widhts
    .ColWidth(0) = .ColWidth(0) * 1.5
    .ColWidth(1) = .ColWidth(1) * 0.95
    .ColWidth(2) = .ColWidth(2) * 0.75
    .ColWidth(3) = .ColWidth(3) * 0.75
    .ColWidth(4) = .ColWidth(4) * 0.75
    .ColWidth(5) = .ColWidth(5) * 0.75
    '' Sets the Column Alignment
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    .ColAlignment(4) = flexAlignLeftCenter
    .ColAlignment(5) = flexAlignLeftCenter
    '' Sets the Appropriate Captions
    .TextMatrix(0, 0) = NewCaptionTxt("00030", adrsMod)   '' Date
    .TextMatrix(0, 1) = NewCaptionTxt("22003", adrsC)   '' Paid Days
    .TextMatrix(0, 2) = NewCaptionTxt("22004", adrsC)   '' Present
    .TextMatrix(0, 3) = NewCaptionTxt("22005", adrsC)   '' Absent
    .TextMatrix(0, 4) = NewCaptionTxt("22006", adrsC)   '' WeekOff
    .TextMatrix(0, 5) = NewCaptionTxt("22007", adrsC)   '' Holidays
End With
End Sub

Private Sub cboDept_Click()
On Error GoTo ERR_P
If cboDept.ListIndex < 0 Then Exit Sub
If cboDept.Text = "ALL" Then
    Call ComboFill(cboCode, 1, 2, 17)
Else
Call ComboFill(cboCode, 12, 2, cboDept.List(cboDept.ListIndex, 1))
End If
MSF1.Rows = 1
Exit Sub
ERR_P:
    ShowError ("Department::" & Me.Caption)
End Sub

Private Sub cboCode_Click()
If cboCode.Text = "" Then Exit Sub
Call FillGrid
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
    Case 3      '' Edit Mode
        If Not ValidateModMaster Then Exit Sub  '' If not Valid for Edit
        If Not SaveModMaster Then Exit Sub      '' Save for Edit
        Call SaveModLog                         '' Save the Edit Log
        Call FillGrid
        bytMode = 1
        Call ViewAction
End Select
Exit Sub
ERR_P:
    ShowError ("EditSave :: " & Me.Caption)
    Resume Next
End Sub

Private Sub cmdExit_Click()
On Error GoTo ERR_P
If bytMode = 3 Then
    bytMode = 1
    Call ChangeMode
Else
    Unload Me
End If
Exit Sub
ERR_P:
    ShowError ("ExitCancel :: " & Me.Caption)
End Sub

Private Sub Form_Activate()
    Call OpenMastersTable
End Sub

Private Sub Form_Load()

txtRemark.Visible = False
lblRemark.Visible = False

Call SetFormIcon(Me)        '' Set the Form Icon
Call SetToolTipText(Me)     '' Set the ToolTipText
Call RetCaptions            '' Retreives the Captions
Call FillComboEmp           '' Fill Employee Combo
TB1.Tab = 0                 '' Set the Tab to List
Call GetRights              '' Gets Rights for the Operations
cboCode.Value = ""          '' Sets the Value to NULL
bytMode = 1
Call ChangeMode             '' Takes Action on Mode Basis
TB1.TabEnabled(1) = False   '' Keeps the Default Mode to No Records

End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '22%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("22001", adrsC)              '' Form caption
TB1.TabCaption(0) = NewCaptionTxt("00013", adrsMod)       '' List
TB1.TabCaption(1) = NewCaptionTxt("00014", adrsMod)       '' Details
lblDate.Caption = NewCaptionTxt("00030", adrsMod)         '' Date
lblPaid.Caption = NewCaptionTxt("22003", adrsC)         '' Paid Days
Call SetButtonCap                           '' Sets Appropriate Caption for the Buttons
Call CapGrid                                '' Captions For Grid
Call SetCritLabel(lblDeptCap)
Exit Sub
ERR_P:
    ShowError ("RetCaption :: " & Me.Caption)
End Sub

Private Sub SetButtonCap(Optional bytFlgCap As Byte = 1)    '' Sets Captions to the Main
If bytFlgCap = 1 Then                                       '' Buttons
    cmdEditCan.Caption = "Update" 'NewCaptionTxt("00005", adrsMod)  '' &Edit
    cmdExit.Caption = "Exit" ''NewCaptionTxt("00008", adrsMod)     '' E&xit
    cmdExit.Cancel = True
Else
    cmdEditCan.Caption = NewCaptionTxt("00007", adrsMod)  '' &Save
    cmdExit.Caption = NewCaptionTxt("00003", adrsMod)     '' Cancel
    cmdExit.Cancel = False
End If
End Sub

Private Sub FillComboEmp()
On Error GoTo ERR_P
Call SetCritCombos(cboDept)
If strCurrentUserType <> HOD Then
    cboDept.Text = "ALL"
End If
Exit Sub
ERR_P:
    ShowError ("FillComboEmp :: " & Me.Caption)
End Sub

Private Sub FillGrid()
On Error GoTo ERR_P
Dim intCounter As Integer

If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "select Empcode,lst_date,paiddays," & pVStar.PrsCode & "," & _
pVStar.AbsCode & "," & pVStar.WosCode & "," & pVStar.HlsCode & " from " & _
strLvFileName & " where EmpCode='" & cboCode.List(cboCode.ListIndex, 1) & "' Order by EmpCode,Lst_Date", _
ConMain
MSF1.Rows = 1
'' Put Appropriate Rows in the Grid
If adrsDept1.EOF Then
    TB1.TabEnabled(1) = False       '' Disables Tab 1 If no Records are Found
    Exit Sub
End If
Do While Not adrsDept1.EOF
    adrsDept1.Find "Empcode='" & cboCode.List(cboCode.ListIndex, 1) & "'"     '' Searches for the Specified
    If Not adrsDept1.EOF Then
        MSF1.Rows = MSF1.Rows + 1
    Else
        Exit Do
    End If
    With MSF1
        .TextMatrix(MSF1.Rows - 1, 0) = DateDisp(adrsDept1("lst_date"))
        .TextMatrix(MSF1.Rows - 1, 1) = IIf(IsNull(adrsDept1("paiddays")), "0.00", _
                                    Format(adrsDept1("paiddays"), "0.00"))
        .TextMatrix(MSF1.Rows - 1, 2) = IIf(IsNull(adrsDept1(Trim(pVStar.PrsCode))), "0.00", _
                                    Format(adrsDept1(Trim(pVStar.PrsCode)), "0.00"))
        .TextMatrix(MSF1.Rows - 1, 3) = IIf(IsNull(adrsDept1(Trim(pVStar.AbsCode))), "0.00", _
                                    Format(adrsDept1(Trim(pVStar.AbsCode)), "0.00"))
        .TextMatrix(MSF1.Rows - 1, 4) = IIf(IsNull(adrsDept1(Trim(pVStar.WosCode))), "0.00", _
                                    Format(adrsDept1(Trim(pVStar.WosCode)), "0.00"))
        .TextMatrix(MSF1.Rows - 1, 5) = IIf(IsNull(adrsDept1(Trim(pVStar.HlsCode))), "0.00", _
                                    Format(adrsDept1(Trim(pVStar.HlsCode)), "0.00"))
        '.TextMatrix(MSF1.Rows - 1, 6) = adrsDept1("remarks")
    
    adrsDept1.MoveNext
    End With
Loop
TB1.TabEnabled(1) = True        '' If Records Found Enables the TAB 1
TB1.Tab = 0
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub OpenMastersTable()
On Error GoTo ERR_P
strLvFileName = "lvtrn" & Right(pVStar.YearSel, 2)
If Not FindTable(strLvFileName) Then
    MsgBox NewCaptionTxt("00054", adrsMod) & pVStar.YearSel & NewCaptionTxt("00055", adrsMod) & _
    vbCrLf & NewCaptionTxt("22008", adrsC)
    Unload Me
Else
End If
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Sub GetRights()     '' Gets and Sets the Appropriate Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(4, 21, 4, 1)
If strTmp = "1" Then
    EditRights = True
Else
    EditRights = False
End If
Exit Sub
ERR_P:
    ShowError ("GetRights :: " & Me.Caption)
    EditRights = False
End Sub

Private Sub ChangeMode()
Select Case bytMode
    Case 1  '' View Mode
        Call ViewAction
    Case 3  '' Edit Mode
        Call EditAction
End Select
End Sub

Private Sub ViewAction()    '' Procedure for Viewing Mode
'' Disable Needed Controls
txtPaid.Enabled = False     '' Disable Paid Days TextBox
txtDate.Enabled = False     '' Disable Date TextBox
'' Give Captions to the Needed Controls
Call SetButtonCap
TB1.Tab = 0     '' Set the Tab to the First Tab
End Sub

Private Sub EditAction()    '' Procedure for Edit Mode
'' Enable Necessary Controls
txtPaid.Enabled = True     '' Enable Paid Days TextBox
'' Give Caption to the Needed Controls
Call SetButtonCap(2)
txtPaid.SetFocus   '' Set Focus on the Paid Days TextBox

If TB1.Tab = 1 Then
    Exit Sub
Else
    If TB1.TabEnabled(1) Then TB1.Tab = 1
End If
End Sub

Private Sub TB1_Click(PreviousTab As Integer)
If PreviousTab = 1 Then Exit Sub
MSF1.Col = 0
If MSF1.Text = NewCaptionTxt("00030", adrsMod) Then Exit Sub
Call Display
End Sub

Private Sub Display()       '' Displays the Given Master Records
On Error GoTo ERR_P
txtDate.Text = MSF1.TextMatrix(MSF1.row, 0)
txtPaid.Text = MSF1.TextMatrix(MSF1.row, 1)
'txtremark.Text = MSF1.TextMatrix(MSF1.row, 2)
Exit Sub
ERR_P:
    ShowError ("Display  :: " & Me.Caption)
End Sub

Private Function ValidateModMaster() As Boolean
ValidateModMaster = True
txtPaid.Text = IIf(Trim(txtPaid.Text) = "", "0.00", Format(txtPaid.Text, "0.00"))
Select Case Right(txtPaid.Text, 2)
    Case "00", "50"
    Case Else
        MsgBox NewCaptionTxt("00056", adrsMod), vbExclamation
        txtPaid.SetFocus
        ValidateModMaster = False
        Exit Function
End Select
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
SaveModMaster = True        '' Update

ConMain.Execute "update " & strLvFileName & "  set paiddays=" & _
Val(txtPaid.Text) & " where Empcode=" & "'" & cboCode.Text & "'" & " and lst_date=" & _
strDTEnc & DateCompStr(txtDate.Text) & strDTEnc

Exit Function
ERR_P:
    ShowError ("SaveModMaster :: " & Me.Caption)
    SaveModMaster = False
End Function

Private Sub MSF1_DblClick()
If TB1.TabEnabled(1) = False Then Exit Sub
TB1.Tab = 1
End Sub

Private Sub txtDate_GotFocus()
    Call GF(txtDate)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    Call CDK(txtDate, KeyAscii)
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
    If Not ValidDate(txtDate) Then txtDate.SetFocus: Cancel = True
End Sub

Private Sub txtPaid_GotFocus()
    Call GF(txtPaid)
End Sub

Private Sub txtPaid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtPaid)
End If
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 3, 24)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edit Paid Days Of Employee: " & cboCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub
