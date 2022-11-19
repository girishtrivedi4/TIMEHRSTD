VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmmonthlyrule 
   Caption         =   "Monthly Rule"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   4755
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command2"
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   3360
      Width           =   1005
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command1"
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   3360
      Width           =   1125
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command1"
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   3360
      Width           =   1125
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      Top             =   3360
      Width           =   915
   End
   Begin TabDlg.SSTab TB1 
      Height          =   3075
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   5424
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
      TabPicture(0)   =   "frmmonthlyrule.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSF1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Details"
      TabPicture(1)   =   "frmmonthlyrule.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frMain"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame frMain 
         Height          =   2265
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   4725
         Begin VB.TextBox txthr 
            Height          =   375
            Left            =   3240
            TabIndex        =   9
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txttr 
            Height          =   375
            Left            =   1440
            TabIndex        =   8
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtfr 
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   735
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   3795
         Left            =   -74700
         TabIndex        =   1
         Top             =   420
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   6694
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   12632256
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
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
Attribute VB_Name = "frmmonthlyrule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adrsC As New ADODB.Recordset
Dim prevtab As Integer

Private Sub cmdAddSave_Click()        '
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
        'If Not ValidateAddmaster Then Exit Sub  '' Validate For Add
        If Not SaveAddMaster Then Exit Sub      '' Save for Add
        'Call SaveAddLog                         '' Save the Add Log
        Call FillGrid                           '' Reflect the Grid
        bytMode = 1
        Call ChangeMode
    Case 3          '' Edit Mode
        'If Not ValidateModMaster Then Exit Sub  '' Validate for Edit
        If Not SaveModMaster Then Exit Sub      '' Save for Edit
        'Call SaveModLog                         '' Save the Edit Log
        Call FillGrid                           '' Reflect the Grid
        bytMode = 1
        Call ChangeMode
End Select
Exit Sub
ERR_P:
    ShowError ("AddSave :: " & Me.Caption)
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)        '' Set the Form Icon
Call RetCaption             '' Retreive Captions
Call OpenMasterTable        '' Open Master Table
Call FillGrid               '' Fill Grid
TB1.Tab = 0                 '' Set the Tab to List
Call GetRights              '' Gets Rights for the Operations
bytMode = 1
Call ChangeMode             '' Take Action on the Appropriate Mode
End Sub
Private Sub RetCaption()    '
On Error GoTo ERR_P
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '48%'", VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
'Me.Caption = NewCaptionTxt("48001", adrsC)        '' Form caption
TB1.TabCaption(0) = NewCaptionTxt("00013", adrsMod) '' List
TB1.TabCaption(1) = NewCaptionTxt("00014", adrsMod) '' Details
'Call SetOtherCaps                           '' Set Captions for other Captions
Call SetGButtonCap(Me)                           '' Sets Appropriate Caption for the Buttons
Call CapGrid                                '' Captions For Grid
Exit Sub
ERR_P:
    ShowError ("RetCaption :: " & Me.Caption)
End Sub
Private Sub CapGrid()           '' Gives the Captions to the Grid
With MSF1
    '' Sets the Column Widhts
    .ColWidth(0) = .ColWidth(0) * 1.05
    .ColWidth(1) = .ColWidth(1) * 1.05
    .ColWidth(2) = .ColWidth(2) * 1.24
    '' Sets the Column Alignment
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    '.ColAlignment(3) = flexAlignLeftCenter
    '' Sets the Appropriate Captions
    .TextMatrix(0, 0) = "From" '' From
    .TextMatrix(0, 1) = "To" '' To
    .TextMatrix(0, 2) = "Hours" '' Hours
End With
End Sub
Private Sub OpenMasterTable()             '' Open the recordset for the Display purposes
On Error GoTo ERR_P
'If FindTable("Monthrule") Then VstarDataEnv.cnDJConn.Execute "drop table Monthrule"
'VstarDataEnv.cnDJConn.Execute "create table Monthrule(fr smallmoney ,tr smallmoney ,hr smallmoney)"
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "select *from Monthrule", _
VstarDataEnv.cnDJConn, adOpenStatic
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Function SaveAddMaster() As Boolean     '
On Error GoTo ERR_P
SaveAddMaster = True        '' Insert
VstarDataEnv.cnDJConn.Execute "insert into Monthrule(fr,tr,hr) values(" & txtfr.Text & "," & txttr.Text & "," & txthr.Text & ")"
'Format(adrsDept1("hr"), "00.00"))
Exit Function
ERR_P:
    Select Case Err.Number
        Case -2147217900
            MsgBox NewCaptionTxt("10017", adrsC), vbExclamation
        Case Else
            ShowError ("SaveAddMaster :: " & Me.Caption)
    End Select
    SaveAddMaster = False
End Function

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
Private Sub FillGrid()          '' Fills the Grid                 '
On Error GoTo ERR_P
Dim intCounter As Integer
adrsDept1.Requery               '' Requeries the Recordset for any Updated Values
'' Put Appropriate Rows in the Grid
If adrsDept1.EOF Then
    MSF1.Rows = 1
    TB1.TabEnabled(1) = False       '' Disables Tab 1 If no Records are Found
    Exit Sub
End If
MSF1.Rows = adrsDept1.RecordCount + 1  '' Sets Rows Appropriately
adrsDept1.MoveFirst
For intCounter = 1 To adrsDept1.RecordCount   '' Fills the Grid
    With MSF1
        .TextMatrix(intCounter, 0) = adrsDept1("fr")
        .TextMatrix(intCounter, 1) = IIf(IsNull(adrsDept1("tr")), "", adrsDept1("tr"))
        .TextMatrix(intCounter, 2) = IIf(IsNull(adrsDept1("hr")), "", adrsDept1("hr"))
    End With
    adrsDept1.MoveNext
Next
TB1.TabEnabled(1) = True        '' If Records Found Enables the TAB 1
TB1.Tab = 0
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub
Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
SaveModMaster = True        '' Update
''SG07
VstarDataEnv.cnDJConn.Execute "update Monthrule set fr =" & txtfr.Text & _
",tr=" & txttr.Text & ",hr=" & txthr.Text & " where hr=" & txthr.Text & " or fr =" & txtfr.Text & " or tr=" & txttr.Text & ""
''
Exit Function
ERR_P:
    SaveModMaster = False
    ShowError ("SaveModMaster :: " & Me.Caption)
End Function
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
Private Sub ViewAction()    '' Procedure for Viewing Mode
'' Enable Needed Controls
cmdAddSave.Enabled = True       '' Enable ADD/SAVE Button
cmdEditCan.Enabled = True       '' Enable Edit/Cancel Button
cmdDel.Enabled = True           '' Enable Delete Button
'' Disable Needed Control
txtfr.Enabled = False         '' Disable Code TextBox
txttr.Enabled = False         '' Disable Name TextBox
txthr.Enabled = False
'' Give Captions to the Needed Controls
Call SetGButtonCap(Me)
TB1.Tab = 0     '' Set the Tab to the First Tab
End Sub
Private Sub AddAction()     '' Procedure for Addition Mode
'' Enable Necessary Controls
If TB1.TabEnabled(1) = False Then TB1.TabEnabled(1) = True
TB1.Tab = 1
txtfr.Enabled = True       '' Enable Code TextBox
txttr.Enabled = True      '' Enable Name TextBox
txthr.Enabled = True
'' Disable Necessary Controls
cmdDel.Enabled = False      '' Disable Delete Button
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Clear Necessary Controls
txtfr.Text = ""       '' Clear Code Control
txttr.Text = ""       '' Clear Name Control
txthr.Text = "0"  '' Clear the Strength Control
txtfr.SetFocus        '' Set Focus to the Code TextBox
End Sub
Private Sub EditAction()    '' Procedure for Edit Mode
'' Enable Necessary Controls
txtfr.Enabled = True      '' Enable from TextBox
txttr.Enabled = True
txthr.Enabled = True
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
txtfr.SetFocus            '' Set Focus on the Name TextBox
If TB1.Tab = 1 Then
    Exit Sub
Else
    If TB1.TabEnabled(1) Then TB1.Tab = 1
End If
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
Private Sub cmdDel_Click()
On Error GoTo ERR_P
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
    'adrsDumb.Open "Select * from Monthrule where fr =" & txtfr.Text & "", _
    'VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
    
    If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion, Me.Caption) _
    = vbYes Then        '' Delete the Record
          VstarDataEnv.cnDJConn.Execute "delete from Monthrule where fr=" & txtfr.Text & " and tr=" & txttr.Text & " and hr=" & txthr.Text & ""
          Call AddActivityLog(lgDelete_Action, 1, 9)      '' Delete Log Action
          Call AuditInfo("DELETE", Me.Caption, "Deleted From: " & txtfr.Text)
    End If
    Call FillGrid       '' Reflect the Grid
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
ERR_P:
    If Err.Number = -2147467259 Or Err.Number = -2147217900 Then
            MsgBox " Cannot be deleted ", vbCritical, Me.Caption
            Exit Sub
    End If
    ShowError ("Delete Record :: " & Me.Caption)
End Sub
Private Sub MSF1_dblClick()
If TB1.TabEnabled(1) = False Then Exit Sub
TB1.Tab = 1
End Sub
Private Sub TB1_Click(PreviousTab As Integer)
prevtab = PreviousTab
If bytMode = 2 Then Exit Sub
If PreviousTab = 1 Then Exit Sub
MSF1.Col = 0
If MSF1.Text = NewCaptionTxt("00047", adrsMod) Then Exit Sub
Call Display
End Sub
Private Sub Display()       '' Displays the Given Master Records
On Error GoTo ERR_P
adrsDept1.MoveFirst
adrsDept1.Find "fr=" & MSF1.TextMatrix(MSF1.row, 0) & ""
If Not (adrsDept1.EOF) Then
    txtfr.Text = MSF1.TextMatrix(MSF1.row, 0)
    txttr.Text = MSF1.TextMatrix(MSF1.row, 1)
    txthr.Text = MSF1.TextMatrix(MSF1.row, 2)
    '' Display Strength
    'Call DispStrength                               '' Department Strength
Else
    txtfr = ""
    txttr = ""
    txthr = ""
    MsgBox NewCaptionTxt("20003", adrsC), vbCritical
    Exit Sub
End If
Exit Sub
ERR_P:
    ShowError ("Display :: " & Me.Caption)
End Sub


