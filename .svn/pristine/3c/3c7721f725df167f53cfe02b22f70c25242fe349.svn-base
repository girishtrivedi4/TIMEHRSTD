VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLocationIP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEditCan 
      Height          =   375
      Left            =   1005
      TabIndex        =   2
      Top             =   3840
      Width           =   1005
   End
   Begin VB.CommandButton cmdDel 
      Height          =   375
      Left            =   2010
      TabIndex        =   3
      Top             =   3840
      Width           =   1005
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3015
      TabIndex        =   4
      Top             =   3840
      Width           =   1125
   End
   Begin VB.CommandButton cmdAddSave 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3840
      Width           =   1005
   End
   Begin TabDlg.SSTab TB1 
      Height          =   3765
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   6641
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
      TabPicture(0)   =   "frmLocationIP.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cboLoc"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "MSF1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Details"
      TabPicture(1)   =   "frmLocationIP.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frLoc"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame frLoc 
         Height          =   3375
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   3945
         Begin VB.Frame fraIP 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   1080
            TabIndex        =   18
            Top             =   720
            Width           =   2775
            Begin MSMask.MaskEdBox txtIP 
               Height          =   240
               Index           =   0
               Left            =   0
               TabIndex        =   7
               Top             =   240
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   423
               _Version        =   393216
               Appearance      =   0
               AutoTab         =   -1  'True
               MaxLength       =   3
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
            Begin MSMask.MaskEdBox txtIP 
               Height          =   240
               Index           =   1
               Left            =   720
               TabIndex        =   8
               Top             =   240
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   423
               _Version        =   393216
               Appearance      =   0
               AutoTab         =   -1  'True
               MaxLength       =   3
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
            Begin MSMask.MaskEdBox txtIP 
               Height          =   240
               Index           =   2
               Left            =   1440
               TabIndex        =   9
               Top             =   240
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   423
               _Version        =   393216
               Appearance      =   0
               AutoTab         =   -1  'True
               MaxLength       =   3
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
            Begin MSMask.MaskEdBox txtIP 
               Height          =   240
               Index           =   3
               Left            =   2160
               TabIndex        =   10
               Top             =   240
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   423
               _Version        =   393216
               Appearance      =   0
               AutoTab         =   -1  'True
               MaxLength       =   3
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
            Begin VB.Label Label5 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   360
               Left            =   2040
               TabIndex        =   21
               Top             =   120
               Width           =   105
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   360
               Left            =   1320
               TabIndex        =   20
               Top             =   120
               Width           =   105
            End
            Begin VB.Label Label3 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   360
               Left            =   600
               TabIndex        =   19
               Top             =   120
               Width           =   105
            End
         End
         Begin VB.Label lblIPAdd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IP Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   14
            Top             =   930
            Width           =   900
         End
         Begin VB.Label lblLoc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   450
         End
         Begin MSForms.ComboBox CboLoc1 
            Height          =   375
            Left            =   1080
            TabIndex        =   6
            Top             =   360
            Width           =   1335
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "2355;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2415
         Left            =   -74970
         TabIndex        =   11
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4260
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedCols       =   0
         BackColorFixed  =   12632256
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
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   -73560
         TabIndex        =   17
         Top             =   960
         Width           =   2445
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Location Name : "
         Height          =   195
         Left            =   -74880
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin MSForms.ComboBox cboLoc 
         Height          =   375
         Left            =   -73560
         TabIndex        =   0
         Top             =   480
         Width           =   1815
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3201;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Location Code : "
         Height          =   195
         Left            =   -74880
         TabIndex        =   15
         Top             =   600
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmLocationIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strKey As String, strIp As String
Dim adrsC As New ADODB.Recordset

Private Sub cboLoc_Click()
    If cboLoc.ListIndex < 0 Then Exit Sub
    lblName.Caption = cboLoc.List(cboLoc.ListIndex, 1)
    Call FillGrid
End Sub

Private Sub cmdAddSave_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        If Not AddRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        bytMode = 2
        Call ChangeMode
    Case 2          '' Add Mode
        GetIP
        If Not ValidateAddmaster Then Exit Sub  '' Validate For Add
        If Not SaveAddMaster Then Exit Sub      '' Save for Add
        'Call SaveAddLog                         '' Save the Add Log
        cboLoc.ListIndex = CboLoc1.ListIndex
        Call FillGrid
        bytMode = 1
        Call ChangeMode
    Case 3          '' Edit Mode
        GetIP
        If Not ValidateModMaster Then Exit Sub  '' Validate for Edit
        If Not SaveModMaster Then Exit Sub      '' Save for Edit
'        Call SaveModLog                         '' Save the Edit Log
        Call FillGrid
        bytMode = 1
        Call ChangeMode
End Select
Exit Sub
ERR_P:
    ShowError ("AddSave :: " & Me.Caption)
End Sub

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
    If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion) = vbYes Then        '' Delete the Record
        VstarDataEnv.cnDJConn.Execute "delete from LocationIP where Location=" & _
        cboLoc.Text & " and IP='" & strIp & "'"
        Call AddActivityLog(lgDelete_Action, 1, 10)
        Call AuditInfo("DELETE", Me.Caption, "Deleted Location IP Address: " & strIp)
    End If
    Call FillGrid
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
strKey = "1234567890"
Call SetFormIcon(Me)        '' Set the Form Icon
Call RetCaptions            '' Retreive Captions
Call ComboFill(cboLoc, 11, 2)
TB1.Tab = 0                 '' Set the Tab to List
Call GetRights              '' Gets Rights for the Operations
bytMode = 1
Call ChangeMode             '' Take Action on the Appropriate Mode
TB1.TabEnabled(1) = False
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

Private Sub TB1_Click(PreviousTab As Integer)
If bytMode = 2 Then Exit Sub
If PreviousTab = 1 Then Exit Sub
MSF1.Col = 0
If MSF1.Text = NewCaptionTxt("00059", adrsMod) Then Exit Sub
Call Display
End Sub
Private Sub Display()       '' Displays the Given Master Records
On Error GoTo ERR_P
Dim i As Integer
Dim strArrIP() As String
CboLoc1.ListIndex = cboLoc.ListIndex
strArrIP = Split(MSF1.TextMatrix(MSF1.row, 0), ".")
For i = 0 To UBound(strArrIP)
    txtIP(i).Text = strArrIP(i)
Next
adrsDept1.MoveFirst
adrsDept1.Find "Location=" & CboLoc1.Text
If adrsDept1.EOF Then
    'txtIP.Text = ""
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
ERR_P:
    ShowError ("Display  :: " & Me.Caption)
End Sub
Private Sub RetCaptions()
On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '63%'", VstarDataEnv.cnDJConn, adOpenStatic
Me.Caption = NewCaptionTxt("63001", adrsC)              '' Form caption
TB1.TabCaption(0) = NewCaptionTxt("00013", adrsMod)       '' List
TB1.TabCaption(1) = NewCaptionTxt("00014", adrsMod)       '' Details
Call SetGButtonCap(Me)                           '' Sets Appropriate Caption for the Buttons
Call CapGrid                                '' Captions For Grid
End Sub

Private Sub CapGrid()
MSF1.ColWidth(0) = MSF1.ColWidth(0) * 1.3
MSF1.ColAlignment(0) = flexAlignLeftCenter
MSF1.TextMatrix(0, 0) = "IP Address"
End Sub

Private Sub FillGrid()
On Error GoTo ERR_P
Dim intCounter As Integer
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select LocationIP.Location,IP,LocDesc from LocationIP,Location where LocationIP.Location=Location.Location and LocationIP.Location=" & cboLoc.Text, VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
If adrsDept1.EOF Then
    MSF1.Rows = 1
    TB1.TabEnabled(1) = False
    Exit Sub
End If
MSF1.Rows = adrsDept1.RecordCount + 1
adrsDept1.MoveFirst
For intCounter = 1 To adrsDept1.RecordCount
    MSF1.TextMatrix(intCounter, 0) = adrsDept1("IP")
    adrsDept1.MoveNext
Next
TB1.TabEnabled(1) = True
TB1.Tab = 0
Exit Sub
ERR_P:
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

Private Sub AddAction()
If TB1.TabEnabled(1) = False Then TB1.TabEnabled(1) = True
TB1.Tab = 1
CboLoc1.Enabled = True
fraIP.Enabled = True
cmdDel.Enabled = False
Call SetGButtonCap(Me, 2)
Call ComboFill(CboLoc1, 11, 2)
CboLoc1.ListIndex = cboLoc.ListIndex
For i = 0 To 3
    txtIP(i).Text = ""
Next
If CboLoc1.Text = "" Then
    CboLoc1.SetFocus
Else
    txtIP(0).SetFocus
End If
End Sub
Private Sub ViewAction()
cmdAddSave.Enabled = True
cmdEditCan.Enabled = True
cmdDel.Enabled = True
CboLoc1.Enabled = False
fraIP.Enabled = False
Call SetGButtonCap(Me)
TB1.Tab = 0
End Sub

Private Sub EditAction()
CboLoc1.Enabled = False
fraIP.Enabled = True
Call SetGButtonCap(Me, 2)
Call ComboFill(CboLoc1, 11, 2)
CboLoc1.ListIndex = cboLoc.ListIndex
cmdDel.Enabled = False
If TB1.Tab = 1 Then
    Exit Sub
Else
    If TB1.TabEnabled(1) Then TB1.Tab = 1
End If
txtIP(0).SetFocus
End Sub
Private Function ValidateAddmaster() As Boolean
On Error GoTo ERR_P
ValidateAddmaster = True
If Trim(CboLoc1.Text) = "" Then
    MsgBox NewCaptionTxt("63002", adrsC), vbExclamation
    CboLoc1.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If MSF1.Rows > 1 Then
    adrsDept1.MoveFirst
    adrsDept1.Find "IP=" & strIp
    If Not adrsDept1.EOF Then
        MsgBox "IP Address Already Exist", vbExclamation
        txtIP(0).SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If
If strIp = "..." Then
    MsgBox "IP Address cannot be Blank", vbExclamation
    txtIP(0).SetFocus
    ValidateAddmaster = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster :: " & Me.Caption)
End Function
Private Function ValidateModMaster() As Boolean
On Error GoTo ERR_P
ValidateModMaster = True
If Trim(strIp) = "..." Then
    MsgBox "IP Address cannot be Blank", vbExclamation
    txtIP(0).SetFocus
    ValidateModMaster = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
SaveModMaster = True        '' Update
VstarDataEnv.cnDJConn.Execute "update LocationIP set IP='" & _
strIp & "' where Location=" & CboLoc1.Text
Exit Function
ERR_P:
    SaveModMaster = False
    ShowError ("SaveModMaster :: " & Me.Caption)
End Function
Private Function SaveAddMaster() As Boolean
On Error GoTo ERR_P
SaveAddMaster = True        '' Insert
VstarDataEnv.cnDJConn.Execute "insert into LocationIP values (" & Trim(CboLoc1.Text) & ",'" & strIp & "')"
Exit Function
ERR_P:
    Select Case Err.Number
        Case Else
            ShowError ("SaveAddMaster :: " & Me.Caption)
    End Select
    SaveAddMaster = False
End Function

Private Sub txtIP_Change(Index As Integer)
If Len(Trim(txtIP(Index).Text)) = 3 Then
    If Val(txtIP(Index).Text) > 255 Then
        MsgBox txtIP(Index) & " is not a valid entry. Please specify a value between 0 and 255."
        txtIP(Index).SelStart = 0
        txtIP(Index).SelLength = Len(txtIP(Index).Text)
        txtIP(Index).SetFocus
        Exit Sub
    End If
    If Index <> 3 Then
        Index = Index + 1
        txtIP(Index).SelStart = 0
        txtIP(Index).SelLength = Len(txtIP(Index))
        If fraIP.Enabled = True Then txtIP(Index).SetFocus
    End If
End If
End Sub

Private Sub txtIP_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 8 Then
        If (InStr(strKey, Chr(KeyAscii))) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtIP_Validate(Index As Integer, Cancel As Boolean)
    If Trim(txtIP(Index).Text) = "" Then
        txtIP(Index).Text = 0
    End If
End Sub
 Private Sub GetIP()
    Dim i As Integer
    strIp = ""
    For i = 0 To 3
        If Trim(txtIP(i).Text) = "" Then txtIP(i).Text = 0
        strIp = strIp & Trim(txtIP(i).Text) & "."
    Next
    strIp = Mid(strIp, 1, Len(strIp) - 1)
 End Sub

