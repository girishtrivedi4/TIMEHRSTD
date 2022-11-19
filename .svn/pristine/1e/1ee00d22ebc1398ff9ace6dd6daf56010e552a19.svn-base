VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmHoliday 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   4500
      TabIndex        =   9
      Top             =   3870
      Width           =   1515
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   3870
      Width           =   1515
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1500
      TabIndex        =   7
      Top             =   3870
      Width           =   1515
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command1"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   3870
      Width           =   1515
   End
   Begin TabDlg.SSTab TB1 
      Height          =   3855
      Left            =   0
      TabIndex        =   10
      Top             =   -15
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   2
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
      TabCaption(0)   =   "List"
      TabPicture(0)   =   "frmHoliday.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSF1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Details"
      TabPicture(1)   =   "frmHoliday.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frDetails"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame frDetails 
         Height          =   3465
         Left            =   -75000
         TabIndex        =   5
         Top             =   300
         Width           =   5985
         Begin VB.TextBox txtToDate 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   1935
            TabIndex        =   4
            Tag             =   "D"
            Top             =   2250
            Width           =   1515
         End
         Begin VB.TextBox txtDate 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   1935
            TabIndex        =   3
            Tag             =   "D"
            Top             =   1755
            Width           =   1515
         End
         Begin VB.CheckBox chkAllCat 
            Caption         =   "Add this holiday for all categories"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   60
            TabIndex        =   0
            Top             =   150
            Width           =   3270
         End
         Begin MSMask.MaskEdBox txtDesc 
            Height          =   360
            Left            =   1950
            TabIndex        =   2
            Top             =   1185
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   49
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label lblToDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
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
            Left            =   1350
            TabIndex        =   16
            Top             =   2250
            Width           =   210
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
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
            Left            =   1440
            TabIndex        =   15
            Top             =   1800
            Width           =   450
         End
         Begin VB.Label lblDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
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
            Left            =   150
            TabIndex        =   14
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblCat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Specific Category"
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
            Left            =   90
            TabIndex        =   12
            Top             =   720
            Width           =   1530
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   990
            TabIndex        =   13
            Top             =   1800
            Width           =   405
         End
         Begin MSForms.ComboBox cboCat 
            Height          =   345
            Left            =   1950
            TabIndex        =   1
            Top             =   675
            Width           =   1515
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "2672;609"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   3495
         Left            =   30
         TabIndex        =   11
         Top             =   360
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   6165
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   12632256
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   2
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
Attribute VB_Name = "frmHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset

Private Sub cboCat_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then SendKeys Chr(9)
End Sub

Private Sub chkAllCat_Click()
If chkAllCat.Value = 1 Then
    cboCat.Enabled = False
    cboCat.Value = ""
Else
    cboCat.Enabled = True
End If
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
        If Not SaveAddMasterNew Then Exit Sub      '' Save for Add
     
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

        ConMain.Execute _
        "delete from holiday where cat='" & CatCode(cboCat.Text) & _
         "' and " & strKDesc & "='" & txtDesc.Text & "' and " & strKDate & "=" & strDTEnc & Format(DateCompDate(txtDate.Text), "DD/MMM/yy") & strDTEnc & ""
   
         Call AddActivityLog(lgDelete_Action, 1, 12)    '' Delete Log
         Call AuditInfo("DELETE", Me.Caption, "Deleted Holiday: " & txtDesc.Text)
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
        ''  on mansi's request
        'frmShiftCr.Show vbModal
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
Call FillGrid               '' Fill Grid
Call FillComboCat           '' Fill Category Combo
TB1.Tab = 0                 '' Set the Tab to List
Call GetRights              '' Gets Rights for the Operations
bytMode = 1
Call ChangeMode             '' Take Action on the Appropriate Mode

Call SetHolidayFramSet

End Sub

Private Function SetHolidayFramSet()
On Error GoTo Err
'FromAndToDate not yet used for any client

    txtDate.Visible = True
    lblDate.Visible = True

Exit Function
Err:
    Call ShowError("SetHolidayFramSet")
End Function
Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '30%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("30001", adrsC)             '' Form caption
TB1.TabCaption(0) = NewCaptionTxt("00013", adrsMod)       '' List
TB1.TabCaption(1) = NewCaptionTxt("00014", adrsMod)       '' Details

    chkAllCat.Caption = NewCaptionTxt("30004", adrsC)       '' All Categories
    lblCat.Caption = NewCaptionTxt("30005", adrsC)           '' Specific Category

lblDate.Caption = NewCaptionTxt("00030", adrsMod)          '' Date
lblDesc.Caption = NewCaptionTxt("00052", adrsMod)         '' Description
Call SetGButtonCap(Me)                           '' Sets Appropriate Caption for the Buttons
Call CapGrid                                '' Captions For Grid
End Sub

Private Sub CapGrid()           '' Gives the Captions to the Grid
With MSF1
    '' Sets the Column Widhts
    .ColWidth(0) = .ColWidth(0) * 1.45
    .ColWidth(1) = .ColWidth(1) * 1.3
    .ColWidth(2) = .ColWidth(2) * 2.45
    '' Sets the Column Alignment
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    '' Sets the Appropriate Captions

     .TextMatrix(0, 0) = NewCaptionTxt("00051", adrsMod) '' Category
    .TextMatrix(0, 1) = NewCaptionTxt("30002", adrsC) '' Holiday Date
    .TextMatrix(0, 2) = NewCaptionTxt("30003", adrsC) '' Name of Holiday
End With
End Sub

Private Sub OpenMasterTable()             '' Open the recordset for the Display purposes
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close
Dim strHolFrom As String, strHolTo As String
strHolFrom = CStr(CDate("01-" & Left(MonthName(pVStar.Yearstart), 3) & "-" & pVStar.YearSel))
strHolTo = CDate(strHolFrom) + IIf(Year(CDate(strHolFrom)) Mod 4, 364, 365)
adrsDept1.Open "Select Cat," & strKDate & "," & strKDesc & ",Hcode from Holiday where (" & strKDate & " between " & _
strDTEnc & DateCompStr(strHolFrom) & strDTEnc & " and " & strDTEnc & DateCompStr(strHolTo) & strDTEnc & _
") Order by Cat," & strKDate & "", ConMain, adOpenKeyset, adLockOptimistic
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Sub FillComboCat()      '' Fills Category Combo
On Error GoTo ERR_P
cboCat.clear
If AdrsCat.State = 1 Then AdrsCat.Close

    AdrsCat.Open "Select " & strKDesc & " from catdesc where cat <> '100' Order by Cat", ConMain

If Not (AdrsCat.BOF And AdrsCat.EOF) Then
    Do While Not AdrsCat.EOF
        cboCat.AddItem AdrsCat(0)
        AdrsCat.MoveNext
    Loop
End If
AdrsCat.Close
Exit Sub
ERR_P:
    ShowError ("FillComboCat :: " & Me.Caption)
End Sub

Private Sub FillGrid()          '' Fills the Grid
On Error GoTo ERR_P
Dim intCounter As Integer
adrsDept1.Requery               '' Requeries the Recordset for any Updated Values
'' Put Appropriate Rows in the Grid
If adrsDept1.EOF Then
    MSF1.Rows = 1
    TB1.TabEnabled(1) = False       '' Disables Tab 1 If No Records are Found
    Exit Sub
End If
MSF1.Rows = adrsDept1.RecordCount + 1   '' Sets Rows Appropriately
adrsDept1.MoveFirst
For intCounter = 1 To adrsDept1.RecordCount     '' Fills the Grid
    With MSF1
        .TextMatrix(intCounter, 0) = adrsDept1("Cat")
        .TextMatrix(intCounter, 1) = DateDisp(adrsDept1("date"))
        .TextMatrix(intCounter, 2) = IIf(IsNull(adrsDept1("Desc")), "", adrsDept1("Desc"))
    End With
    adrsDept1.MoveNext
Next
TB1.TabEnabled(1) = True        '' If Records Found Enables the TAB 1
TB1.Tab = 0
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub GetRights()     '' Gets and Sets the Appropriate Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 12)
If Mid(strTmp, 1, 1) = "1" Then AddRights = True
If Mid(strTmp, 2, 1) = "1" Then EditRights = True
If Mid(strTmp, 3, 1) = "1" Then DeleteRights = True
Exit Sub
ERR_P:
    ShowError ("GetRights :: " & Me.Caption)
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

Private Sub AddAction()     '' Procedure for Addition Mode
If TB1.TabEnabled(1) = False Then TB1.TabEnabled(1) = True
TB1.Tab = 1
'' Enable Necessary Controls
cboCat.Enabled = True       '' Enable Category TextBox
txtDate.Enabled = True      '' Enable Date TextBox
chkAllCat.Enabled = True    '' Enable Category CheckBox
txtDesc.Enabled = True      '' Enable Description TextBox
txtToDate.Enabled = True
txtToDate.Text = ""
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Clear Necessary Controls
chkAllCat.Value = 0         '' Reset All Categories CheckBox
cboCat.Value = ""           '' Reset the ComboBox To Empty
txtDate.Text = ""           '' Clear Date Control
txtDesc.Text = ""           '' Clear Description Control
chkAllCat.SetFocus          '' Set Focus on the Category CheckBox

End Sub

Private Sub EditAction()    '' Procedure for Edit Mode
'' Enable Necessary Controls
txtDesc.Enabled = True          '' Enable Description TextBox
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
txtDesc.SetFocus                '' Set Focus on the Holiday Description TextBox
If TB1.Tab = 1 Then
    Exit Sub
Else
    If TB1.TabEnabled(1) Then TB1.Tab = 1
End If
End Sub

Private Sub ViewAction()    '' Procedure for Viewing Mode
'' Enable Needed Controls
cmdAddSave.Enabled = True       '' Enable ADD/SAVE Button
cmdEditCan.Enabled = True       '' Enable Edit/Cancel Button
cmdDel.Enabled = True           '' Enable Disable Button
'' Disable Needed Controls
txtDate.Enabled = False         '' Disable Date TextBox
txtDesc.Enabled = False         '' Disable Description TextBox
chkAllCat.Enabled = False       '' Disable Category Check Box
cboCat.Enabled = False          '' Disable Category Combo
txtToDate.Enabled = False
'' Give Captions to the Needed Controls
Call SetGButtonCap(Me)
TB1.Tab = 0     '' Set the Tab to the First Tab
End Sub

Private Function ValidateAddmaster() As Boolean
On Error GoTo ERR_P
ValidateAddmaster = True
If chkAllCat.Value = False And cboCat.Text = "" Then
    MsgBox NewCaptionTxt("30006", adrsC), vbExclamation
    cboCat.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If chkAllCat.Value = 1 And cboCat.ListCount = 0 Then
    MsgBox NewCaptionTxt("30007", adrsC), vbExclamation
    chkAllCat.SetFocus
    ValidateAddmaster = False
    Exit Function
End If

    If Not CheckDate Then
        ValidateAddmaster = False
        Exit Function
Else
    If Trim(txtDate.Text) = "" Then
        MsgBox NewCaptionTxt("30008", adrsC), vbExclamation
        txtDate.SetFocus
        ValidateAddmaster = False
        Exit Function
    ElseIf Trim(txtToDate.Text) = "" Then

            MsgBox NewCaptionTxt("30008", adrsC), vbExclamation
            txtDate.SetFocus
            ValidateAddmaster = False
            Exit Function
    
    Else
        If DateDiff("m", Year_Start, CDate(txtDate.Text)) > 11 Or DateDiff("m", Year_Start, CDate(txtDate.Text)) < 0 Then
            MsgBox NewCaptionTxt("00030", adrsMod) & txtDate.Text & NewCaptionTxt("00021", adrsMod), vbExclamation
            txtDate.SetFocus
            ValidateAddmaster = False
            Exit Function
        End If
    End If
End If
If Trim(txtDesc.Text) = "" Then
    MsgBox NewCaptionTxt("30009", adrsC), vbExclamation
    txtDesc.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If chkAllCat.Value = 0 Then
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "select * from Holiday where holiday." & strKDate & "=" & strDTEnc & _
    DateCompStr(txtDate.Text) & strDTEnc & " and Cat=" & "'" & CatCode(Trim(cboCat.Text)) & "'" _
    , ConMain, adOpenKeyset, adLockOptimistic
    If Not (adrsTemp.EOF And adrsTemp.BOF) Then
        MsgBox NewCaptionTxt("30010", adrsC), vbExclamation
        txtDate.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
Else
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "select * from Holiday where holiday." & strKDate & "=" & strDTEnc & _
    DateCompStr(txtDate.Text) & strDTEnc, ConMain, adOpenKeyset, adLockOptimistic
    If Not (adrsTemp.EOF And adrsTemp.BOF) Then
        MsgBox NewCaptionTxt("30010", adrsC), vbExclamation
        txtDate.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster :: " & Me.Caption)
    ValidateAddmaster = False
End Function
Private Function CheckDate() As Boolean
On Error GoTo Err
CheckDate = True

    If DateCompDate(txtToDate.Text) < DateCompDate(txtDate.Text) Then
        MsgBox NewCaptionTxt("00018", adrsMod), vbExclamation, App.EXEName
        CheckDate = False
        txtDate.SetFocus
        Exit Function
    End If

If Trim(txtDate.Text) = "" Then
        MsgBox NewCaptionTxt("30008", adrsC), vbExclamation
        txtDate.SetFocus
        CheckDate = False
        Exit Function
    Else
        If DateDiff("m", Year_Start, CDate(txtDate.Text)) > 11 Or DateDiff("m", Year_Start, CDate(txtDate.Text)) < 0 Then
            MsgBox NewCaptionTxt("00030", adrsMod) & txtDate.Text & NewCaptionTxt("00021", adrsMod), vbExclamation
            txtDate.SetFocus
            CheckDate = False
            Exit Function
        End If
End If

If Trim(txtToDate.Text) = "" Then
        MsgBox NewCaptionTxt("30008", adrsC), vbExclamation
        txtToDate.SetFocus
        CheckDate = False
        Exit Function
    Else
        If DateDiff("m", Year_Start, CDate(txtToDate.Text)) > 11 Or DateDiff("m", Year_Start, CDate(txtToDate.Text)) < 0 Then
            MsgBox NewCaptionTxt("00030", adrsMod) & txtToDate.Text & NewCaptionTxt("00021", adrsMod), vbExclamation
            txtToDate.SetFocus
            CheckDate = False
            Exit Function
        End If
End If

Exit Function
Err:
    Call ShowError("CheckDate")
    CheckDate = False
End Function
Private Function ValidateModMaster() As Boolean
On Error GoTo ERR_P
ValidateModMaster = True
If Trim(txtDesc.Text) = "" Then
    MsgBox NewCaptionTxt("30009", adrsC), vbExclamation
    txtDesc.SetFocus
    ValidateModMaster = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
    ValidateModMaster = False
End Function

Private Function SaveAddMasterNew() As Boolean          ' 21-05-09
On Error GoTo ERR_P
Dim bytCat As Byte
SaveAddMasterNew = True        '' Insert
If chkAllCat.Value = 0 Then
    ConMain.Execute "insert into Holiday(Cat," & strKDate & "," & strKDesc & ",HCode) Values('" & _
    CatCode(cboCat.Text) & "'," & strDTEnc & DateSaveIns(txtDate.Text) & strDTEnc & ",'" & _
    Trim(txtDesc.Text) & "','" & pVStar.HlsCode & "')"
Else
    For bytCat = 0 To cboCat.ListCount - 1
        ConMain.Execute "insert into Holiday(Cat," & strKDate & "," & strKDesc & ",HCode) Values('" & _
        CatCode(cboCat.List(bytCat)) & "'," & strDTEnc & DateSaveIns(txtDate.Text) & _
        strDTEnc & ",'" & txtDesc.Text & "','" & pVStar.HlsCode & "')"
    Next
End If
If DateCompDate(txtDate.Text) <> DateCompDate(txtToDate) Then
    txtDate.Text = DateAdd("d", 1, DateCompDate(txtDate.Text))
    Call SaveAddMasterNew
End If
Exit Function
ERR_P:
    SaveAddMasterNew = False
    ShowError ("SaveAddMasterNew :: " & Me.Caption)
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
SaveModMaster = True        '' Update
ConMain.Execute "update Holiday set " & strKDesc & "=" & "'" & Trim(txtDesc.Text) & _
"'" & " where Cat=" & "'" & CatCode(cboCat.Text) & "'" & " and " & strKDate & "=" & _
strDTEnc & DateSaveIns(txtDate.Text) & strDTEnc
Exit Function
ERR_P:
    SaveModMaster = False
    ShowError ("SaveModMaster :: " & Me.Caption)
End Function

Private Sub txtDate_Click()
varCalDt = ""
varCalDt = Trim(txtDate.Text)
txtDate.Text = ""
Call ShowCalendar
End Sub

Private Sub txtDate_GotFocus()
    Call GF(txtDate)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    Call CDK(txtDate, KeyAscii)
End Sub

Private Sub Display()       '' Displays the Given Master Records
On Error GoTo ERR_P
cboCat.Value = GetDesc(MSF1.TextMatrix(MSF1.row, 0))     '' Category Code
If cboCat.Value = "" Then
    MsgBox NewCaptionTxt("30011", adrsC), vbExclamation
    bytMode = 1
    Call ChangeMode
    Exit Sub
End If
txtDate = DateDisp(MSF1.TextMatrix(MSF1.row, 1))        '' Date
txtDesc = MSF1.TextMatrix(MSF1.row, 2)                  '' Description
       ' 26-05-09
    If adrsC.State = 1 Then adrsC.Close
    adrsC.Open "Select * from Holiday Where Cat = " & "'" & CatCode(cboCat.Text) & "'" & " And " & strKDate & " = " & strDTEnc & DateSaveIns(txtDate.Text) & strDTEnc
    If Not adrsC.EOF Then
        txtToDate = DateDisp(MSF1.TextMatrix(MSF1.row, 1))
    End If

Exit Sub
ERR_P:
    ShowError ("Display  :: " & Me.Caption)
    'Resume Next
End Sub

Private Function GetDesc(ByVal strCatName As String) As String
On Error GoTo ERR_P         '' Gets the Description for a Particular Code"
If adrsTemp.State = 1 Then adrsTemp.Close
  adrsTemp.Open "Select " & strKDesc & " from Catdesc where Cat='" & strCatName & "'" _
    , ConMain
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    GetDesc = adrsTemp.Fields(0)    ' adrsTemp("Desc")
Else
    GetDesc = ""
End If
Exit Function
ERR_P:
    ShowError ("GetDesc  :: " & Me.Caption)
End Function

Private Sub MSF1_DblClick()
If TB1.TabEnabled(1) = False Then Exit Sub
TB1.Tab = 1
End Sub

Private Sub TB1_Click(PreviousTab As Integer)
If bytMode = 2 Then Exit Sub
If PreviousTab = 1 Then Exit Sub
MSF1.Col = 0
If MSF1.Text = NewCaptionTxt("00051", adrsMod) Then Exit Sub
Call Display
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
    If Not ValidDate(txtDate) Then txtDate.SetFocus: Cancel = True
End Sub

Private Sub txtToDate_Click()   ' 21-05-09
varCalDt = ""
varCalDt = Trim(txtToDate.Text)
txtToDate.Text = ""
Call ShowCalendar
End Sub

Private Sub txtToDate_GotFocus()
    Call GF(txtToDate)
End Sub

Private Sub txtToDate_KeyPress(KeyAscii As Integer)
    Call CDK(txtToDate, KeyAscii)
End Sub

Private Sub txtToDate_Validate(Cancel As Boolean)
    If Not ValidDate(txtToDate) Then txtDate.SetFocus: Cancel = True
End Sub

Private Sub txtDesc_GotFocus()
    Call GF(txtDesc)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 3))))
End If
End Sub

Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo ERR_P
Call AddActivityLog(lgADD_MODE, 1, 12)     '' Add Activity
Call AuditInfo("ADD", Me.Caption, "Added Holiday: " & txtDesc.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 12)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edited Holiday: " & txtDesc.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub
