VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmOfficialDuty 
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command1"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   1755
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1740
      TabIndex        =   2
      Top             =   3000
      Width           =   1725
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3450
      TabIndex        =   1
      Top             =   3000
      Width           =   1605
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3000
      Width           =   1545
   End
   Begin TabDlg.SSTab TB1 
      Height          =   2925
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5159
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
      TabCaption(0)   =   "List"
      TabPicture(0)   =   "frmOfficialDuty.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSF1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Details"
      TabPicture(1)   =   "frmOfficialDuty.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2295
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   6375
         Begin VB.TextBox txttotalhrs 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   9
            Tag             =   "D"
            Top             =   1440
            Width           =   1155
         End
         Begin VB.TextBox txtFrom 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   8
            Tag             =   "D"
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtfromhrs 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   960
            TabIndex        =   7
            Tag             =   "D"
            Top             =   1200
            Width           =   1155
         End
         Begin VB.TextBox txttohrs 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   960
            TabIndex        =   6
            Tag             =   "D"
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Hrs."
            Height          =   195
            Left            =   2880
            TabIndex        =   16
            Top             =   1515
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   480
            TabIndex        =   15
            Top             =   1755
            Width           =   195
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   360
            TabIndex        =   14
            Top             =   1275
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Official Duty in Hrs."
            Height          =   195
            Left            =   360
            TabIndex        =   13
            Top             =   840
            Width           =   1350
         End
         Begin VB.Label lblFromD 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   195
            Left            =   360
            TabIndex        =   12
            Top             =   285
            Width           =   345
         End
         Begin MSForms.ComboBox cbocode 
            Height          =   315
            Left            =   3840
            TabIndex        =   11
            Top             =   240
            Width           =   1365
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "2408;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empcode"
            Height          =   195
            Left            =   2880
            TabIndex        =   10
            Top             =   285
            Width           =   675
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2535
         Left            =   -74970
         TabIndex        =   17
         Top             =   360
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   1
         Cols            =   5
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
End
Attribute VB_Name = "frmOfficialDuty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prevtab As Integer
Dim rsOffduty As New ADODB.Recordset
Dim adrsForm As New ADODB.Recordset
Dim adrsC As New ADODB.Recordset

Private Sub cmdAddSave_Click()
On Error GoTo Err_P
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
        'If Not ValidateAddmaster Then Exit Sub  '' Validate For Add
        If Not SaveAddMaster Then Exit Sub      '' Save for Add
        Call SaveAddLog                         '' Save the Add Log
        Call FillGrid                           '' Reflect the Grid
        bytMode = 1
        Call ChangeMode
    Case 3          '' Edit Mode
        'If Not ValidateModMaster Then Exit Sub  '' Validate for Edit
        If Not SaveModMaster Then Exit Sub      '' Save for Edit
        Call SaveModLog                         '' Save the Edit Log
        Call FillGrid                           '' Reflect the Grid
        bytMode = 1
        Call ChangeMode
End Select
Exit Sub
Err_P:
    ShowError ("AddSave :: " & Me.Caption)
End Sub

Private Sub cmdDel_Click()
On Error GoTo Err_P
Dim adrsDumb As New ADODB.Recordset
'' Check for Rights
If Not DeleteRights Then
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    Exit Sub
Else
    If TB1.TabEnabled(1) = False Then Exit Sub
    If TB1.Tab = 0 Then                         '' Do not Display Record if
        If TB1.TabEnabled(1) Then TB1.Tab = 1   '' Already Displayed
    End If
    adrsDumb.Open "Select * from UserAccs where dept ='" & cbocode.Text & "'", _
    VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
    If Not (adrsDumb.EOF And adrsDumb.BOF) Then
        MsgBox "This right is allotted to some user." & vbCrLf & _
        "You can not delete this record.", vbCritical
        Exit Sub
    End If
    If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion, Me.Caption) _
    = vbYes Then        '' Delete the Record
          VstarDataEnv.cnDJConn.Execute "delete from OfficialDuty where empcode='" & cbocode.Text & "'"
          Call AddActivityLog(lgDelete_Action, 1, 9)      '' Delete Log Action
    End If
    Call FillGrid       '' Reflect the Grid
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
Err_P:
    If Err.Number = -2147467259 Or Err.Number = -2147217900 Then
            MsgBox "Department Cannot be deleted because employees belong to this department.", vbCritical, Me.Caption
            Exit Sub
    End If
    ShowError ("Delete Record :: " & Me.Caption)
End Sub

Private Sub cmdEditCan_Click()
On Error GoTo Err_P
Select Case bytMode
    Case 1          '' View Mode
        If TB1.TabEnabled(1) = False Then Exit Sub
        '' Check for Rights
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
Err_P:
    ShowError ("EditCancel :: " & Me.Caption)
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)        '' Set the Form Icon
Call SetGButtonCap(Me)                           '' Sets Appropriate Caption for the Buttons
Call RetCaptions
Call CapGrid
Call OpenMasterTable        '' Open Master Table
Call FillGrid               '' Fill Grid
Call FillEmpCode
Call GetRights
TB1.Tab = 0                 '' Set the Tab to List
bytMode = 1
Call ChangeMode             '' Take Action on the Appropriate Mode
End Sub

Private Sub txtFrom_Click()
varCalDt = ""
varCalDt = Trim(txtFrom.Text)
txtFrom.Text = ""
Call ShowCalendar
End Sub
Private Sub CapGrid()           '' Gives the Captions to the Grid
With MSF1
    '' Sets the Column Widhts
    '.ColWidth(0) = .ColWidth(0) * 0.73
    '.ColWidth(1) = .ColWidth(1) * 4
    '' Sets the Column Alignment
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    '' Sets the Appropriate Captions
    .TextMatrix(0, 0) = "Date"
    .TextMatrix(0, 1) = "Empcode"
    .TextMatrix(0, 2) = "From"
    .TextMatrix(0, 3) = "To"
    .TextMatrix(0, 4) = "Total Hrs"
End With
End Sub
Private Sub OpenMasterTable()             '' Open the recordset for the Display purposes
On Error GoTo Err_P
If rsOffduty.State = 1 Then rsOffduty.Close
rsOffduty.Open "Select * from OfficialDuty", VstarDataEnv.cnDJConn, adOpenStatic
Exit Sub
Err_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Sub FillGrid()          '' Fills the Grid
On Error GoTo Err_P
Dim intCounter As Integer
rsOffduty.Requery               '' Requeries the Recordset for any Updated Values
'' Put Appropriate Rows in the Grid
If rsOffduty.EOF Then
    MSF1.Rows = 1
    TB1.TabEnabled(1) = False       '' Disables Tab 1 If no Records are Found
    Exit Sub
End If
MSF1.Rows = rsOffduty.RecordCount + 1   '' Sets Rows Appropriately
rsOffduty.MoveFirst
For intCounter = 1 To rsOffduty.RecordCount     '' Fills the Grid
    With MSF1
        .TextMatrix(intCounter, 0) = IIf(IsNull(rsOffduty(1)), "", rsOffduty(1))
        .TextMatrix(intCounter, 1) = rsOffduty(0)
        .TextMatrix(intCounter, 2) = IIf(IsNull(rsOffduty(2)), "", rsOffduty(2))
        .TextMatrix(intCounter, 3) = IIf(IsNull(rsOffduty(3)), "", rsOffduty(3))
        .TextMatrix(intCounter, 4) = IIf(IsNull(rsOffduty(4)), "", rsOffduty(4))
    End With
    rsOffduty.MoveNext
Next
TB1.TabEnabled(1) = True        '' If Records Found Enables the TAB 1
TB1.Tab = 0
Exit Sub
Err_P:
    ShowError ("FillGrid :: " & Me.Caption)
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
cmdDel.Enabled = True           '' Enable Delete Button
'' Disable Needed Controls
cbocode.Enabled = True         '' Disable Code TextBox
txtfromhrs.Locked = True          '' Disable Name TextBox
txttohrs.Locked = True
txttotalhrs.Locked = True
cbocode.Value = ""
txtfromhrs.Text = ""
txttohrs.Text = ""
txttotalhrs.Text = ""
'' Give Captions to the Needed Controls
Call SetGButtonCap(Me)
TB1.Tab = 0     '' Set the Tab to the First Tab
End Sub

Private Sub FillEmpCode()       '' Fills Employee Code & Name Combos
On Error GoTo Err_P
cbocode.Clear       '' Clear Code Combo
Dim strTmp As String
If strCurrentUserType = HOD Then
    ''Original ->strTmp = " Where Dept=" & intCurrDept & " "
    strTmp = strCurrData
End If
If adrsForm.State = 1 Then adrsForm.Close
adrsForm.Open "Select Empcode from Empmst " & strTmp & " Order by Empcode", _
VstarDataEnv.cnDJConn, adOpenKeyset, adLockOptimistic       '' Fill Code Combo
If Not (adrsForm.BOF And adrsForm.EOF) Then
    Do While Not adrsForm.EOF
        cbocode.AddItem adrsForm(0)
        adrsForm.MoveNext
    Loop
End If
Exit Sub
Err_P:
    ShowError (" FillEmpCode :: " & Me.Caption)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
If cmdEditCan.Caption = "&Cancel " Then
TB1.Tab = 0
bytMode = 1
Call ChangeMode
Call SetGButtonCap(Me)
Else
Unload Me
End If
End If
If KeyCode = vbKeyF10 Then Call ShowF10("20")
If KeyCode = vbKeyF10 Then KeyCode = 0
End Sub

Private Sub MSF1_DblClick()
If TB1.TabEnabled(1) = False Then Exit Sub
TB1.Tab = 1
End Sub

Private Sub TB1_Click(PreviousTab As Integer)
prevtab = PreviousTab
If bytMode = 2 Then Exit Sub
If PreviousTab = 1 Then Exit Sub
'MSF1.Col = 0
If MSF1.Text = NewCaptionTxt("00047", adrsMod) Then Exit Sub
Call Display
End Sub
Private Sub Display()       '' Displays the Given Master Records
On Error GoTo Err_P
rsOffduty.MoveFirst
rsOffduty.Find "[Date]=" & MSF1.TextMatrix(MSF1.Row, 0) & ""
If Not (rsOffduty.EOF) Then
    txtFrom.Text = MSF1.TextMatrix(MSF1.Row, 0)
    cbocode.Value = MSF1.TextMatrix(MSF1.Row, 1)
    txtfromhrs.Text = MSF1.TextMatrix(MSF1.Row, 2)
    txttohrs.Text = MSF1.TextMatrix(MSF1.Row, 3)
    txttotalhrs.Text = MSF1.TextMatrix(MSF1.Row, 4)
Else
    txtFrom.Text = ""
    txtfromhrs.Text = ""
    txttohrs.Text = ""
    txttotalhrs.Text = ""
    'MsgBox NewCaptionTxt("20003", adrsC), vbCritical
    Exit Sub
End If
Exit Sub
Err_P:
    ShowError ("Display :: " & Me.Caption)
End Sub
Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo Err_P
Call AddActivityLog(lgADD_MODE, 1, 9)     '' Add Activity
Exit Sub
Err_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo Err_P
Call AddActivityLog(lgEdit_Mode, 1, 9)     '' Edit Activity
Exit Sub
Err_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub
Private Function SaveAddMaster() As Boolean
On Error GoTo Err_P
SaveAddMaster = True        '' Insert
''SG07
Call CommonCalc
VstarDataEnv.cnDJConn.Execute "insert into OfficialDuty(Empcode,[Date],Fromhrs,Tohrs,Totalhrs)" _
& " values ( '" & cbocode.Text & "',#" & Format(DateCompDate(txtFrom.Text), "DD-MMM-YYYY") & "#," & Trim(txtfromhrs.Text) & "," & txttohrs.Text & "," & txttotalhrs.Text & ")"
Exit Function
Err_P:
    Select Case Err.Number
        Case -2147217900
            MsgBox NewCaptionTxt("20007", adrsC), vbExclamation
        Case Else
            ShowError ("SaveAddMaster :: " & Me.Caption)
    End Select
    SaveAddMaster = False
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo Err_P
SaveModMaster = True        '' Update
''SG07
Call CommonCalc
VstarDataEnv.cnDJConn.Execute "update OfficialDuty set Fromhrs=" & Trim(txtfromhrs.Text) & ",Tohrs=" & txttohrs.Text & ",Totalhrs=" & txttotalhrs.Text & ",[date]=# " & txtFrom.Text & " # where empcode ='" & cbocode.Text & "'"
Exit Function
Err_P:
    SaveModMaster = False
    ShowError ("SaveModMaster :: " & Me.Caption)
End Function

Private Sub AddAction()     '' Procedure for Addition Mode
'' Enable Necessary Controls
If TB1.TabEnabled(1) = False Then TB1.TabEnabled(1) = True
TB1.Tab = 1
txtfromhrs.Locked = False       '' Enable Code TextBox
txttohrs.Locked = False       '' Enable Name TextBox
txttotalhrs.Locked = True
'' Disable Necessary Controls
cmdDel.Enabled = False      '' Disable Delete Button
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Clear Necessary Controls
txtfromhrs.Text = ""       '' Clear Code Control
txttohrs.Text = ""       '' Clear Name Control
txttotalhrs.Text = ""  '' Clear the Strength Control
txtFrom.Text = ""
txtfromhrs.SetFocus        '' Set Focus to the Code TextBox
End Sub
Private Sub EditAction()    '' Procedure for Edit Mode
'' Enable Necessary Controls
txtfromhrs.Locked = False       '' Enable Code TextBox
txttohrs.Locked = False      '' Enable Name TextBox
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
txtfromhrs.SetFocus            '' Set Focus on the Name TextBox
If TB1.Tab = 1 Then
    Exit Sub
Else
    If TB1.TabEnabled(1) Then TB1.Tab = 1
End If
End Sub
Private Sub GetRights()     '' Gets and Sets the Appropriate Rights
On Error GoTo Err_P
Dim strTmp As String
strTmp = RetRights(1, 2)
If Mid(strTmp, 1, 1) = "1" Then AddRights = True
If Mid(strTmp, 2, 1) = "1" Then EditRights = True
If Mid(strTmp, 3, 1) = "1" Then DeleteRights = True
Exit Sub
Err_P:
    ShowError ("Rights :: " & Me.Caption)
    AddRights = False
    EditRights = False
    DeleteRights = False
End Sub
Private Sub txttohrs_GotFocus()
    Call GF(txttohrs)
End Sub
Private Sub txtfromhrs_GotFocus()
    Call GF(txtfromhrs)
End Sub

Private Sub txtfromhrs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txttohrs.SetFocus
    'SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtfromhrs)
End If
End Sub
Private Function CommonCalc()
txttotalhrs.Text = TimDiff(Val(txttohrs.Text), Val(txtfromhrs.Text))
End Function
Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '20%'", VstarDataEnv.cnDJConn, adOpenStatic
End Sub

