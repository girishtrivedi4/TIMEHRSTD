VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmEncash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   435
      Left            =   4980
      TabIndex        =   5
      Top             =   3720
      Width           =   2625
   End
   Begin VB.CommandButton cmdDelCan 
      Caption         =   "Command2"
      Height          =   435
      Left            =   2490
      TabIndex        =   4
      Top             =   3720
      Width           =   2505
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command3"
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   3720
      Width           =   2505
   End
   Begin TabDlg.SSTab TB1 
      Height          =   3225
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   5689
      _Version        =   393216
      Tabs            =   2
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
      TabPicture(0)   =   "frmEncash.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSF1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmEncash.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frEncash"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame frEncash 
         Height          =   2805
         Left            =   -74940
         TabIndex        =   12
         Top             =   360
         Width           =   7500
         Begin VB.TextBox txtFrom 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2160
            TabIndex        =   1
            Tag             =   "D"
            Text            =   " "
            Top             =   1080
            Width           =   1455
         End
         Begin MSMask.MaskEdBox txtDays 
            Height          =   375
            Left            =   2160
            TabIndex        =   2
            Top             =   1590
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            AutoTab         =   -1  'True
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
         Begin MSFlexGridLib.MSFlexGrid MSF2 
            Height          =   2205
            Left            =   3900
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   540
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   3889
            _Version        =   393216
            Rows            =   1
            Cols            =   3
            FixedCols       =   0
            BackColorFixed  =   -2147483640
            ForeColorFixed  =   8454143
            GridColor       =   4194368
            ScrollBars      =   2
            Appearance      =   0
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
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "dsfsdsf"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   630
         End
         Begin VB.Label lblBal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "gdfgdfff"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3900
            TabIndex        =   17
            Top             =   225
            Width           =   720
         End
         Begin VB.Label lblLeave 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "fgfdg"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   14
            Top             =   675
            Width           =   405
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "gdf"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   15
            Top             =   1170
            Width           =   255
         End
         Begin VB.Label lblDays 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "fdgfdg"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   16
            Top             =   1680
            Width           =   510
         End
         Begin MSForms.ComboBox cboLeave 
            Height          =   375
            Left            =   2160
            TabIndex        =   0
            Top             =   570
            Width           =   1485
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "2619;661"
            ListWidth       =   6000
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   0
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1500;4500"
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2745
         Left            =   1380
         TabIndex        =   7
         Top             =   390
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   4842
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   4194368
         ForeColorFixed  =   8454143
         GridColor       =   4194368
         HighLight       =   2
         ScrollBars      =   2
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
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4650
      TabIndex        =   11
      Top             =   60
      Width           =   45
   End
   Begin MSForms.ComboBox cboCode 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   0
      Width           =   2175
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3836;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblNameCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "sfsdfsdf"
      Height          =   195
      Left            =   3960
      TabIndex        =   10
      Top             =   60
      Width           =   540
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "sdfsdffd"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmEncash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim strCatAvail As String               '' For Employee Category
Dim dtJoin As Date                      '' For Empoyee Joindate
Dim sngDaysBal As Single                '' To Get the Balance of the Particular Leave
Dim strRW As String                     '' For the Type of Leave i.e R or W
''
Dim adrsC As New ADODB.Recordset
Dim ELLeave As String, ELSubLeave As String

Private Sub cboCode_Change()
Call cboCode_Click
End Sub

Private Sub cboLeave_Click()
On Error GoTo ERR_P
If bytMode <> 3 Then Exit Sub       '' If Not Add Mode then Exit
Call ToggleType
Exit Sub
ERR_P:
    ShowError ("Leave :: " & Me.Caption)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)        '' Sets the Form Icon
Call SetToolTipText(Me)     '' Set the ToolTipText
Call RetCaptions            '' Sets the Captions for the Controls
Call GetRights              '' Get the Rights
Call FillCombo              '' Fills the Emplyee Combo
TB1.TabEnabled(1) = False   '' Disable the Tab 1
bytMode = 2                 '' Set the Mode Back to Normal or View Mode
Call ChangeMode             '' Take Action According to the Mode
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions Where ID Like '07%' or ID Like '25%'", _
ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("25001", adrsC)              '' Form Caption
Call SetButtonCap                                       '' Button Captions
Call SetOutGridCap                                      '' Msf1 Captions
Call SetInGridCap                                       '' MSF2 Captions
Call SetOtherCaps                                       '' Other Control Captions
TB1.TabCaption(0) = NewCaptionTxt("00013", adrsMod)       '' List
TB1.TabCaption(1) = NewCaptionTxt("00014", adrsMod)       '' Details
End Sub

Private Sub SetOutGridCap()
With MSF1
    .TextMatrix(0, 0) = NewCaptionTxt("07009", adrsC)   '' Leave Code
    .TextMatrix(0, 1) = NewCaptionTxt("07010", adrsC)   '' Leave From
    .TextMatrix(0, 2) = NewCaptionTxt("07011", adrsC)   '' Leave To
    .TextMatrix(0, 3) = NewCaptionTxt("07012", adrsC)   '' Leave Days
End With
End Sub

Private Sub SetInGridCap()
With MSF2
    .TextMatrix(0, 0) = NewCaptionTxt("07013", adrsC)       '' Code
    .TextMatrix(0, 1) = NewCaptionTxt("07014", adrsC)       '' Name
    .TextMatrix(0, 2) = NewCaptionTxt("07015", adrsC)       '' Balance
End With
End Sub

Private Sub SetButtonCap(Optional bytCap As Byte = 1)
Select Case bytCap
    Case 1
        cmdAddSave.Caption = "Add"
        cmdDelCan.Caption = "Delete"
        cmdExit.Caption = "Exit"
    Case 2
        cmdAddSave.Caption = "Save"
        cmdDelCan.Caption = "Cancel"
End Select
End Sub

Private Sub SetOtherCaps()
lblCode.Caption = NewCaptionTxt("07002", adrsC)         '' Employee Code
lblNameCap.Caption = NewCaptionTxt("07003", adrsC)      '' Name
lblInfo.Caption = NewCaptionTxt("07004", adrsC)         '' Leave Information
lblLeave.Caption = NewCaptionTxt("07005", adrsC, 0)     '' Leave Code
lblFrom.Caption = NewCaptionTxt("25002", adrsC)         '' Encash on
lblDays.Caption = NewCaptionTxt("07007", adrsC)         '' No of Days
lblBal.Caption = NewCaptionTxt("07008", adrsC)          '' Balance
End Sub

Private Sub ChangeMode()        '' Action to be Taken when Mode Changes
Select Case bytMode
    Case 2  '' View Mode
        Call ViewAction
    Case 3  '' Add Mode
        Call AddAction
End Select
End Sub

Private Sub ViewAction()        '' Action to be Taken when the Form is in View Mode
TB1.Tab = 0
'' Enable Necessary Controls
cboLeave.Enabled = False    '' Leave ComboBox
txtFrom.Enabled = False     '' From date TextBox
txtDays.Enabled = False     '' Days TextBox
'' Give Caption to the Needed Controls
Call SetButtonCap
End Sub

Private Sub AddAction()     '' Procedure for Addition Mode
'' Set the Tab Accordingly
TB1.Tab = 1
TB1.TabEnabled(1) = True
'' Enable Necessary Controls
cboLeave.Enabled = True     '' Leave ComboBox
txtFrom.Enabled = True      '' From date TextBox
txtDays.Enabled = True      '' Days TextBox
'' Clear Necessary Controls
cboLeave.Value = ""         '' Leave ComboBox
txtFrom.Text = ""           '' From date TextBox
txtDays.Text = "0.00"       '' Leave Days
'' Give Caption to the Needed Controls
Call SetButtonCap(2)
cboLeave.SetFocus           '' Set Focus to the Leave ComboBox
End Sub

Private Sub GetCat()                '' Gets the Category of a Particular Employee
On Error GoTo ERR_P
If cboCode.Text = "" Then
    strCatAvail = ""
    Exit Sub
End If
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open "Select Cat,JoinDate from EmpMst where EmpCode='" & cboCode.Text & "'" _
, ConMain
If (adrsPaid.EOF And adrsPaid.BOF) Then
    MsgBox NewCaptionTxt("07016", adrsC), vbExclamation
    cboCode.Value = ""
    bytMode = 4
    TB1.Tab = 0
    Exit Sub
Else
    strCatAvail = adrsPaid("Cat")
    If Not IsNull(adrsPaid("JoinDate")) Then
        dtJoin = DateCompDate(adrsPaid("JoinDate"))
    Else
        dtJoin = DateCompDate("31-December-2100")
    End If
End If
Exit Sub
ERR_P:
    bytMode = 4
    ShowError ("Getcat :: " & Me.Caption)
End Sub

Private Sub FillGrid()      '' Fills the Grid with the Leaves the Employee Has Encashed
On Error GoTo ERR_P
Dim bytCnt As Byte
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select Lcode,FromDate,ToDate,Days  From " & "LvInfo" & _
Right(pVStar.YearSel, 2) & " where Trcd=3" & " and Empcode=" & "'" & _
cboCode.Text & "' Order by LCode,Fromdate", ConMain, adOpenStatic
MSF1.Rows = 1
If (adrsDept1.EOF And adrsDept1.BOF) Then
    bytMode = 4
    TB1.Tab = 0
Else
    MSF1.Rows = adrsDept1.RecordCount + 1
    For bytCnt = 1 To adrsDept1.RecordCount
        With MSF1
            .TextMatrix(bytCnt, 0) = adrsDept1("LCode")                 '' Leave Code
            .TextMatrix(bytCnt, 1) = DateDisp(adrsDept1("FromDate"))    '' From date
            .TextMatrix(bytCnt, 2) = DateDisp(adrsDept1("ToDate"))      '' To date
            .TextMatrix(bytCnt, 3) = IIf(IsNull(adrsDept1("Days")), "0.00", _
                                        Format(adrsDept1("Days"), "0.00"))  '' Days
        End With
        adrsDept1.MoveNext
    Next
End If
Exit Sub
ERR_P:
    bytMode = 5
    ShowError ("FillGrid : Outer ::" & Me.Caption)
End Sub

Private Sub FillGridBalance()                   '' Gets the Leaves and Balances of a
On Error GoTo ERR_P                             '' Particular Employee
Dim intTmp As Integer                           '' For Field Count
Dim bytTmp As Byte                              '' For Field Count
Dim strLeaveList() As String                    '' For Leave Array
bytTmp = 0
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select * from LvBal" & Right(pVStar.YearSel, 2) & " Where Empcode='" & _
cboCode.Text & "'", ConMain, adOpenStatic
MSF2.Rows = 1
cboLeave.clear              '' Clear the Leave List Box
If SubLeaveFlag = 1 Then ' 15-10
    ELLeave = ""
    If FieldExists("LvBaL" & Right(pVStar.YearSel, 2), "EL") Then ELLeave = "EL"
    If ELLeave = "EL" Then
        If FieldExists("LvBaL" & Right(pVStar.YearSel, 2), "EN") Then ELSubLeave = ",EN"
        If FieldExists("LvBaL" & Right(pVStar.YearSel, 2), "NE") Then ELSubLeave = ELSubLeave & ",NE"
        If ELSubLeave = "" Then ELSubLeave = ","
        ELSubLeave = Right(ELSubLeave, Len(ELSubLeave) - 1)
    End If
End If
If Not (adrsDept1.EOF And adrsDept1.BOF) Then
    For intTmp = 0 To adrsDept1.Fields.Count - 1
        If UCase(adrsDept1(intTmp).name) <> "EMPCODE" Then  '' if Field is not Employee Code
            If adrsRits.State = 1 Then adrsRits.Close
            adrsRits.Open "Select Leave from LeavDesc where LvCode='" & _
            adrsDept1(intTmp).name & "' and EnCase='Y' and Cat='" & strCatAvail & "'", _
            ConMain
            If Not (adrsRits.EOF And adrsRits.BOF) Then         '' if Leave Desc is Found
                MSF2.Rows = MSF2.Rows + 1
                MSF2.TextMatrix(MSF2.Rows - 1, 0) = adrsDept1(intTmp).name
                MSF2.TextMatrix(MSF2.Rows - 1, 1) = adrsRits("Leave").Value
                MSF2.TextMatrix(MSF2.Rows - 1, 2) = IIf(IsNull(adrsDept1(intTmp).Value), "0.00", _
                Format(adrsDept1(intTmp).Value, "0.00"))
                '' Fill the Leave Array
                bytTmp = bytTmp + 1
                If SubLeaveFlag = 1 And (UCase(adrsDept1(intTmp).name) = "SL" Or UCase(adrsDept1(intTmp).name) = "EL") Then bytTmp = bytTmp - 1   ' 07-11
            End If
        End If
    Next
    If bytTmp = 0 Then Exit Sub
    ReDim Preserve strLeaveList(bytTmp - 1, 1)
    bytTmp = 0
    For intTmp = 0 To adrsDept1.Fields.Count - 1
        If UCase(adrsDept1(intTmp).name) <> "EMPCODE" Then  '' if Field is not Employee Code
            If Not (SubLeaveFlag = 1 And (UCase(adrsDept1(intTmp).name) = "SL" Or UCase(adrsDept1(intTmp).name) = "EL")) Then   ' 15-10
                If adrsRits.State = 1 Then adrsRits.Close
                adrsRits.Open "Select Leave from LeavDesc where LvCode='" & _
                adrsDept1(intTmp).name & "' and EnCase='Y' and Cat='" & strCatAvail & "'", _
                ConMain
                If Not (adrsRits.EOF And adrsRits.BOF) Then         '' if Leave Desc is Found
                    '' Fill the Leave Array
                    strLeaveList(bytTmp, 0) = adrsDept1(intTmp).name
                    strLeaveList(bytTmp, 1) = adrsRits("Leave").Value
                    bytTmp = bytTmp + 1
                End If
            End If
        End If
    Next
    cboLeave.List = strLeaveList '' Fill the Leave Box
    Erase strLeaveList
End If
Exit Sub
ERR_P:
    bytMode = 4
    ShowError ("FillGridBalance :: " & Me.Caption)
End Sub

Private Sub cboCode_Click()
On Error GoTo ERR_P
If cboCode.ListIndex < 0 Then Exit Sub
bytMode = 1
'' Displays the Employee Name
lblName.Caption = cboCode.List(cboCode.ListIndex, 1)
'' Gets the Employee Category
Call GetCat
If bytMode <> 4 Then
    '' Get All the Leaves the Employee has Encashed that Year
    bytMode = 1
    Call FillGrid
End If
If bytMode <> 5 Then
    '' Fill the Inner Grid With the Leave Balances of that Employee
    bytMode = 1
    Call FillGridBalance
End If
If bytMode = 4 Or bytMode = 5 Then
    '' If Invalid or Error in the Previous Processes
    MSF1.Rows = 1
    TB1.TabEnabled(1) = False
Else
    If MSF1.Rows > 1 Then
        TB1.TabEnabled(1) = True
    Else
        TB1.TabEnabled(1) = False
    End If
End If
If TB1.TabEnabled(1) = False And TB1.Tab = 1 Then
    TB1.Tab = 0
End If
bytMode = 2     '' Sets the Mode Back to the Normal Mode or View Mode
Exit Sub
ERR_P:
    ShowError ("Employee :: " & Me.Caption)
End Sub

Private Sub cmdAddSave_Click()
On Error GoTo ERR_P
If bytMode = 4 Then bytMode = 2
Select Case bytMode
    Case 2          '' View Mode
        If cboCode.Text = "" Then Exit Sub
        '' Check for Rights
        If Not AddRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        bytMode = 3
        Call ChangeMode
    Case 3          '' Add Mode
        If Not ValidateAddmaster Then Exit Sub      '' Validate For Add
        If Not SaveAddMaster Then Exit Sub          '' Save for Add
        Call SaveAddLog                             '' Save the Add Log
        Call FillGrid                               '' Reflect the Grid
        If bytMode <> 5 Then Call FillGridBalance   '' Fill the Balance Grid
        If MSF1.Rows > 1 Then                       '' Enable Tab 1
            TB1.TabEnabled(1) = True
        Else
            TB1.TabEnabled(1) = False
        End If
        bytMode = 2                                 '' Make Mode to View Mode
        Call ChangeMode                             '' Take Action Based on the Mode
End Select
Exit Sub
ERR_P:
    ShowError ("AddSave :: " & Me.Caption)
End Sub

Private Sub cmdDelCan_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 2          '' View Mode
        If cboCode.Text = "" Then Exit Sub
        If MSF1.Rows = 1 Then Exit Sub
        '' Check for Rights
        If Not DeleteRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        If TB1.Tab = 0 Then TB1.Tab = 1
        If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion) = vbYes Then
            Call UpdateDeleteBalance    '' Update the Balance
            '' Delete the Leave From LvInfo
            ConMain.Execute "Delete From " & "LvInfo" & Right(pVStar.YearSel, 2) & _
            " where Empcode=" & "'" & cboCode.Text & "'" & " and " & "Lcode=" & "'" & _
            cboLeave.Text & "'" & " and Trcd = 3" & " and FromDate=" & _
            strDTEnc & DateCompStr(txtFrom.Text) & strDTEnc
            Call AddActivityLog(lgDelete_Action, 3, 20)     '' Delete Log
            Call AuditInfo("DELETE", Me.Caption, "Delete Leave Encash Entry " & cboLeave.Text & " For Employee " & cboCode.Text)
            Call FillGridBalance            '' Fill the Balances Grid
            Call FillGrid                   '' Fill the Outer Grid
            If MSF1.Rows > 1 Then
                TB1.TabEnabled(1) = True
            Else
                TB1.TabEnabled(1) = False
            End If
        End If
            TB1.Tab = 0
         '' Code For Deletion
    Case 3       '' Add Mode
        bytMode = 2
        Call ChangeMode
        If MSF1.Rows < 2 Then TB1.TabEnabled(1) = False
End Select
Exit Sub
ERR_P:
    ShowError ("DeleteCancel :: " & Me.Caption)
End Sub

Private Sub GetRights()             '' Gets the Rights of for a Particular User
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(2, 3, 5)
If Mid(strTmp, 1, 1) = "1" Then AddRights = True
If Mid(strTmp, 2, 1) = "1" Then DeleteRights = True
Exit Sub
ERR_P:
    ShowError ("GetRights :: " & Me.Caption)
    AddRights = False
    DeleteRights = False
End Sub

Private Sub UpdateDeleteBalance()       '' Updates the Balance of the Particular Employee
On Error GoTo ERR_P                     '' Once the Encashed Leave is Deleted
Dim bytCntTmp As Byte
sngDaysBal = 0
If SubLeaveFlag = 1 And (cboLeave.Text = "HP" Or cboLeave.Text = "CM") Then   ' 07-11
    For bytCntTmp = 1 To MSF2.Rows - 1
        If "SL" = MSF2.TextMatrix(bytCntTmp, 0) Then
            sngDaysBal = Val(MSF2.TextMatrix(bytCntTmp, 2))
        End If
    Next
Else
    For bytCntTmp = 1 To MSF2.Rows - 1
        If cboLeave.Text = MSF2.TextMatrix(bytCntTmp, 0) Then
            sngDaysBal = Val(MSF2.TextMatrix(bytCntTmp, 2))
        End If
    Next
End If
sngDaysBal = sngDaysBal + Val(txtDays.Text)
'' Update Balance of the Employee
If SubLeaveFlag = 1 Then   ' 07-11
    If cboLeave.Text = "HP" Or cboLeave.Text = "CM" Then
        ConMain.Execute "Update LvBal" & Right(pVStar.YearSel, 2) & " Set " & _
        "SL=" & sngDaysBal & " Where Empcode='" & cboCode.Text & "'"
    Else
        ConMain.Execute "Update LvBal" & Right(pVStar.YearSel, 2) & " Set " & _
        cboLeave.Text & "=" & sngDaysBal & " Where Empcode='" & cboCode.Text & "'"
    End If
    If (cboLeave.Text = "EN" Or cboLeave.Text = "NE") Then  ' 15-10
        Dim strqry As String
        strqry = "select " & ELSubLeave & ",lvbal" & Right(pVStar.YearSel, 2) & ".EMPCODE from lvbal" & Right(pVStar.YearSel, 2) & " where empcode='" & cboCode.Text & "'"
        Call UpDateSubLeave("lvbal" & Right(pVStar.YearSel, 2), ELSubLeave, strqry, ELLeave)
    End If
Else
    ConMain.Execute "Update LvBal" & Right(pVStar.YearSel, 2) & " Set " & _
    cboLeave.Text & "=" & sngDaysBal & " Where Empcode='" & cboCode.Text & "'"
End If
Exit Sub
ERR_P:
    ShowError ("UpdateDeleteBalance :: " & Me.Caption)
End Sub

Private Sub Display()       '' Displays the Leave Details
On Error GoTo ERR_P
cboLeave.Text = MSF1.TextMatrix(MSF1.row, 0)        '' Leave Code
txtFrom.Text = MSF1.TextMatrix(MSF1.row, 1)         '' From date
txtDays.Text = Format(MSF1.TextMatrix(MSF1.row, 3), "0.00")     '' Leave Days
Exit Sub
ERR_P:
    Select Case Err.Number
        Case 380
            MsgBox NewCaptionTxt("07035", adrsC) & vbCrLf & _
            NewCaptionTxt("07036", adrsC) & vbCrLf & _
            NewCaptionTxt("07037", adrsC), vbExclamation, App.EXEName
        Case Else
            ShowError ("Display  :: " & Me.Caption)
    End Select
End Sub

Private Sub TB1_Click(PreviousTab As Integer)
If TB1.Tab = 1 Then cboCode.Enabled = False
If TB1.Tab = 0 Then
Call SetButtonCap(1)
cboCode.Enabled = True
bytMode = 2
If MSF1.Rows > 1 Then
        TB1.TabEnabled(1) = True
     Else
       TB1.TabEnabled(1) = False
    End If
Exit Sub            '' If Tab is 0
End If
If bytMode = 1 Then Exit Sub            '' If TempMode
If PreviousTab = 1 Then Exit Sub        '' if Wrong Tab then Exit sub
If bytMode = 3 Then Exit Sub            '' If Add Mode then Exit Sub
MSF1.Col = 0                            '' Set the Column to 0
If MSF1.Text = NewCaptionTxt("07009", adrsC) Then Exit Sub
Call Display
End Sub

Private Function ValidateAddmaster() As Boolean     '' Validate Details befor Encashing the
On Error GoTo ERR_P                                 '' Leave
ValidateAddmaster = True
'' Check if Any Leave is Selected or Not
If cboLeave.Text = "" Then
    MsgBox NewCaptionTxt("25003", adrsC), vbExclamation
    cboLeave.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
'' Check for Invalid Number of EncashLeave Days
If Val(txtDays.Text) <= 0 Then
    MsgBox NewCaptionTxt("25004", adrsC), vbExclamation
    txtDays.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
txtDays.Text = Format(txtDays.Text, "0.00")
If Not GetFlagStatus("HYUNDAI") Then
If Right(txtDays.Text, 2) <> "00" And Right(txtDays.Text, 2) <> "50" Then
    MsgBox NewCaptionTxt("25005", adrsC), vbExclamation
    txtDays.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
End If
'' Check for Invalid date
If Not ValidLeaveDate Then
    ValidateAddmaster = False
    Exit Function
End If
'' Check if there is Enough Balance
If Not NumOfBalance Then
    ValidateAddmaster = False
    Exit Function
End If
'' Check if Leaves are Already Encashed for the Specified Dates
If Not ALreadyAvailedDate Then
    ValidateAddmaster = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster :: " & Me.Caption)
    ValidateAddmaster = False
End Function

Private Function ValidLeaveDate() As Boolean        '' Validates the Leave Dates Specified
ValidLeaveDate = True
'' Check for EmptyDate
If txtFrom.Text = "" Then
    MsgBox NewCaptionTxt("00016", adrsMod), vbExclamation
    txtFrom.SetFocus
    ValidLeaveDate = False
    Exit Function
End If
If DateDiff("m", Year_Start, DateCompDate(txtFrom.Text)) > 11 Or _
DateDiff("m", Year_Start, DateCompDate(txtFrom.Text)) < 0 Then
    MsgBox NewCaptionTxt("00019", adrsMod) & txtFrom.Text & NewCaptionTxt("00021", adrsMod), vbExclamation
    txtFrom.SetFocus
    ValidLeaveDate = False
    Exit Function
End If
If DateCompDate(txtFrom.Text) <= dtJoin Then
    MsgBox NewCaptionTxt("00112", adrsMod), vbExclamation
    txtFrom.SetFocus
    ValidLeaveDate = False
    Exit Function
End If
End Function

Private Function NumOfBalance() As Boolean      '' Checks the Leave Balance of
On Error GoTo ERR_P                             '' the Employee
NumOfBalance = True
Dim bytCntTmp As Byte
sngDaysBal = 0
If SubLeaveFlag = 1 And (cboLeave.Text = "HP" Or cboLeave.Text = "CM") Then
    For bytCntTmp = 1 To MSF2.Rows - 1
        If "SL" = MSF2.TextMatrix(bytCntTmp, 0) Then
            sngDaysBal = Val(MSF2.TextMatrix(bytCntTmp, 2))
        End If
    Next
Else
    For bytCntTmp = 1 To MSF2.Rows - 1
        If cboLeave.Text = MSF2.TextMatrix(bytCntTmp, 0) Then
            sngDaysBal = Val(MSF2.TextMatrix(bytCntTmp, 2))
        End If
    Next
End If
If sngDaysBal <= 0 Then
    If SubLeaveFlag = 1 And cboLeave.Text = "NE" And sngDaysBal = 0 Then   ' 07-11
        MsgBox "NE Leave Balance Cannot become Negative." & vbCrLf & "Encash Remaining days from EN Leave Balance.", vbExclamation
        cboLeave.SetFocus
        NumOfBalance = False
    Else
        If MsgBox(NewCaptionTxt("25006", adrsC), vbYesNo + vbQuestion) = vbYes Then
            NumOfBalance = True
            If SubLeaveFlag = 1 And cboLeave.Text = "CM" Then  ' 15-10
                sngDaysBal = sngDaysBal - (Val(txtDays.Text) * 2)
            Else
                sngDaysBal = sngDaysBal - Val(txtDays.Text)
            End If
        Else
            txtDays.SetFocus
            NumOfBalance = False
        End If
    End If
Else
        NumOfBalance = True
        If SubLeaveFlag = 1 Then  ' 15-10
            If cboLeave.Text = "CM" Then
                sngDaysBal = sngDaysBal - (Val(txtDays.Text) * 2)
            Else
                sngDaysBal = sngDaysBal - Val(txtDays.Text)
            End If
            If cboLeave.Text = "NE" Then
                If sngDaysBal < 0 Then
                    MsgBox "NE Leave Balance Cannot become Negative." & vbCrLf & "Encash Remaining days from EN Leave Balance.", vbExclamation
                    cboLeave.SetFocus
                    NumOfBalance = False
                End If
            End If
        Else
            sngDaysBal = sngDaysBal - Val(txtDays.Text)
        End If
End If
Exit Function
ERR_P:
    ShowError ("NumOfBalance :: " & Me.Caption)
    NumOfBalance = False
End Function

Private Function ALreadyAvailedDate() As Boolean    '' Checks if Leave is Already Encashed
On Error GoTo ERR_P                                 '' by the Employee for the Same Dates
ALreadyAvailedDate = True
Dim strA_R As String
strA_R = "Select * from LvInfo" & Right(pVStar.YearSel, 2) & " where  FromDate=" & strDTEnc & _
DateCompStr(txtFrom.Text) & strDTEnc & " and ToDate=" & strDTEnc & DateCompStr(txtFrom.Text) & _
strDTEnc & " and Trcd= 3" & " and Empcode=" & "'" & cboCode.Text & "'" & " and Lcode=" & _
"'" & cboLeave.Text & "'"
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open strA_R, ConMain
If Not (adrsPaid.EOF And adrsPaid.BOF) Then
        MsgBox NewCaptionTxt("25007", adrsC), vbExclamation
        txtFrom.SetFocus
        ALreadyAvailedDate = False
        Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ALreadyAvailedDate :: " & Me.Caption)
    ALreadyAvailedDate = False
End Function

Private Sub txtDays_GotFocus()
    Call GF(txtDays)
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = KeyDecimal3(KeyAscii, txtDays)
End Select
End Sub

Private Sub txtFrom_Click()
varCalDt = ""
varCalDt = Trim(txtFrom.Text)
txtFrom.Text = ""
Call ShowCalendar
End Sub

Private Sub txtFrom_GotFocus()
    Call GF(txtFrom)
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    Call CDK(txtFrom, KeyAscii)
End Sub

Private Sub txtFrom_Validate(Cancel As Boolean)
    If Not ValidDate(txtFrom) Then txtFrom.SetFocus: Cancel = True
End Sub

Private Function SaveAddMaster() As Boolean     '' Saves Data in the Leave Infomation File
On Error GoTo ERR_P                             '' and Updates the Balances of the Employee
SaveAddMaster = True                            '' in the Leave Balance File
'' Insert Information in LvInfo
Dim NumDays As String
If SubLeaveFlag = 1 And (cboLeave.Text = "CM") Then ' 15-10
    NumDays = Format(CStr(Val(txtDays.Text) * 2), "0.00")
Else
    NumDays = txtDays.Text
End If
ConMain.Execute "insert into LvInfo" & Right(pVStar.YearSel, 2) & _
"(Empcode,trcd,fromdate,todate,lcode,days,lv_type_rw,entrydate) values" & _
"(" & "'" & cboCode.Text & "'" & "," & "3" & "," & strDTEnc & DateSaveIns(txtFrom.Text) & _
strDTEnc & "," & strDTEnc & DateSaveIns(txtFrom.Text) & strDTEnc & "," & "'" & _
cboLeave.Text & "'" & "," & NumDays & "," & "'" & strRW & _
"'" & "," & "'" & DateSaveIns(CStr(Date)) & "'" & ")"
'' Update balance in LvBal
If SubLeaveFlag = 1 Then   ' 07-11
    If cboLeave.Text = "HP" Or cboLeave.Text = "CM" Then
        ConMain.Execute "Update Lvbal" & Right(pVStar.YearSel, 2) & " Set " & _
        "SL=" & sngDaysBal & " Where Empcode='" & cboCode.Text & "'"
    Else
        ConMain.Execute "Update Lvbal" & Right(pVStar.YearSel, 2) & " Set " & _
        cboLeave.Text & "=" & sngDaysBal & " Where Empcode='" & cboCode.Text & "'"
    End If
    If (cboLeave.Text = "EN" Or cboLeave.Text = "NE") Then  ' 15-10
        Dim strqry As String
        strqry = "select " & ELSubLeave & ",lvbal" & Right(pVStar.YearSel, 2) & ".EMPCODE from lvbal" & Right(pVStar.YearSel, 2) & " where empcode='" & cboCode.Text & "'"
        Call UpDateSubLeave("lvbal" & Right(pVStar.YearSel, 2), ELSubLeave, strqry, ELLeave)
    End If
Else
    ConMain.Execute "Update Lvbal" & Right(pVStar.YearSel, 2) & " Set " & _
    cboLeave.Text & "=" & sngDaysBal & " Where Empcode='" & cboCode.Text & "'"
End If
If bytBackEnd = 2 Then Sleep (2000)
Exit Function
ERR_P:
    SaveAddMaster = False
    ShowError ("SaveAddMaster :: " & Me.Caption)
End Function

Public Sub ToggleType()         '' Checks Type of Leave
On Error GoTo ERR_P
If Trim(cboCode.Text) = "" Then Exit Sub
If Trim(cboLeave.Text) = "" Then Exit Sub
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select Run_Wrk from LeavDesc where Cat='" & strCatAvail & "' and LvCode='" & _
cboLeave.Text & "'", ConMain
If Not (adrsDept1.EOF And adrsDept1.BOF) Then
    strRW = IIf(IsNull(adrsDept1("Run_Wrk")), "", adrsDept1("Run_Wrk"))
Else
    strRW = ""
End If
adrsDept1.Close
Exit Sub
ERR_P:
    ShowError ("ToggleType")
End Sub

Private Sub MSF1_DblClick()
If TB1.TabEnabled(1) = False Then Exit Sub
TB1.Tab = 1
End Sub

Private Sub FillCombo()             '' Fill Employee Code Combo
On Error GoTo ERR_P
If strCurrentUserType = HOD Then
    Call ComboFill(cboCode, 16, 2)
Else
    Call ComboFill(cboCode, 19, 2)
End If
Exit Sub
ERR_P:
    ShowError ("Fill Employee Combo :: " & Me.Caption)
End Sub

Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo ERR_P
Call AddActivityLog(lgADD_MODE, 3, 20)     '' Add Activity
Call AuditInfo("ADD", Me.Caption, "Add Leave Encash Entry " & cboLeave.Text & " For Employee " & cboCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub
