VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form UsersFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ExitCmd 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   405
      Left            =   5100
      TabIndex        =   66
      Top             =   5940
      Width           =   1695
   End
   Begin VB.CommandButton DeleteCmd 
      Caption         =   "&Delete"
      Height          =   405
      Left            =   3390
      TabIndex        =   65
      Top             =   5940
      Width           =   1725
   End
   Begin VB.CommandButton EditCanCmd 
      Caption         =   "&Edit"
      Height          =   405
      Left            =   1710
      TabIndex        =   64
      Top             =   5940
      Width           =   1695
   End
   Begin VB.CommandButton AddSaveCmd 
      Caption         =   "&Add"
      Height          =   405
      Left            =   0
      TabIndex        =   63
      Top             =   5940
      Width           =   1725
   End
   Begin TabDlg.SSTab UserTab 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   10398
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
      TabPicture(0)   =   "UsersFrm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Details"
      TabPicture(1)   =   "UsersFrm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "RightsTab"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "PassFrame"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TabDlg.SSTab RightsTab 
         Height          =   4155
         Left            =   -74880
         TabIndex        =   9
         Top             =   1680
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7329
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
         TabCaption(0)   =   "Master File Rights"
         TabPicture(0)   =   "UsersFrm.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LeaveTranFrame"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Sel_UnSelectCmd"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "MSF1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Other Rights"
         TabPicture(1)   =   "UsersFrm.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frRep"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frBackRes"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "RulesFrame"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "frCompact"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "PaidFrame"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "otherCmd"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "LoginFrame"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "CorrectFrame"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "ShiftFrame"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "YrLvFrame"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "ProcessFrame"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "ParameterFrame"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).ControlCount=   12
         Begin VB.Frame frRep 
            Caption         =   "Reports"
            Enabled         =   0   'False
            Height          =   525
            Left            =   -72090
            TabIndex        =   51
            Top             =   2490
            Width           =   2175
            Begin VB.CheckBox UserChk 
               Caption         =   "Reports"
               Height          =   225
               Index           =   30
               Left            =   120
               TabIndex        =   52
               Top             =   210
               Width           =   975
            End
         End
         Begin VB.Frame frBackRes 
            Caption         =   "BackUp and Restore"
            Height          =   525
            Left            =   -72090
            TabIndex        =   37
            Top             =   870
            Width           =   3645
            Begin VB.CheckBox UserChk 
               Caption         =   "Restore"
               Height          =   225
               Index           =   29
               Left            =   1590
               TabIndex        =   39
               Top             =   210
               Width           =   1335
            End
            Begin VB.CheckBox UserChk 
               Caption         =   "BackUp"
               Height          =   225
               Index           =   28
               Left            =   180
               TabIndex        =   38
               Top             =   210
               Width           =   1335
            End
         End
         Begin MSFlexGridLib.MSFlexGrid MSF1 
            Height          =   3765
            Left            =   30
            TabIndex        =   10
            ToolTipText     =   "Double Click or press SPACEBARto Toggle Rights"
            Top             =   330
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   6641
            _Version        =   393216
            Rows            =   15
            Cols            =   4
            FixedCols       =   0
            AllowBigSelection=   0   'False
            TextStyle       =   4
            HighLight       =   2
            ScrollBars      =   0
         End
         Begin VB.Frame RulesFrame 
            Caption         =   "Rules"
            Enabled         =   0   'False
            Height          =   525
            Left            =   -74910
            TabIndex        =   44
            Top             =   1950
            Width           =   2805
            Begin VB.CheckBox UserChk 
               Caption         =   "Restore"
               Height          =   255
               Index           =   27
               Left            =   1620
               TabIndex        =   46
               Top             =   180
               Width           =   1065
            End
            Begin VB.CheckBox UserChk 
               Caption         =   "Edit"
               Height          =   225
               Index           =   26
               Left            =   90
               TabIndex        =   45
               Top             =   210
               Width           =   1275
            End
         End
         Begin VB.Frame frCompact 
            Caption         =   "Compact Database"
            Enabled         =   0   'False
            Height          =   525
            Left            =   -74910
            TabIndex        =   49
            Top             =   2490
            Width           =   2805
            Begin VB.CheckBox UserChk 
               Caption         =   "Permission"
               Height          =   225
               Index           =   25
               Left            =   90
               TabIndex        =   50
               Top             =   240
               Width           =   1305
            End
         End
         Begin VB.Frame PaidFrame 
            Caption         =   "Paid Days"
            Enabled         =   0   'False
            Height          =   525
            Left            =   -70620
            TabIndex        =   32
            Top             =   330
            Width           =   2175
            Begin VB.CheckBox UserChk 
               Caption         =   "Edit"
               Height          =   225
               Index           =   24
               Left            =   120
               TabIndex        =   33
               Top             =   210
               Width           =   975
            End
         End
         Begin VB.CommandButton otherCmd 
            Caption         =   "Select/Unselect All"
            Height          =   435
            Left            =   -70350
            TabIndex        =   62
            Top             =   3180
            Width           =   1875
         End
         Begin VB.CommandButton Sel_UnSelectCmd 
            Caption         =   "Select/Unselect All"
            Height          =   465
            Left            =   4230
            TabIndex        =   26
            Top             =   3150
            Width           =   2175
         End
         Begin VB.Frame LeaveTranFrame 
            Caption         =   "Leave Transaction"
            Height          =   2295
            Left            =   4200
            TabIndex        =   11
            Top             =   600
            Width           =   2295
            Begin VB.CheckBox UserChk 
               Height          =   255
               Index           =   7
               Left            =   1620
               TabIndex        =   25
               Top             =   1680
               Width           =   255
            End
            Begin VB.CheckBox UserChk 
               Height          =   255
               Index           =   6
               Left            =   1080
               TabIndex        =   24
               Top             =   1680
               Width           =   255
            End
            Begin VB.CheckBox UserChk 
               Height          =   255
               Index           =   5
               Left            =   1620
               TabIndex        =   22
               Top             =   1320
               Width           =   255
            End
            Begin VB.CheckBox UserChk 
               Height          =   255
               Index           =   4
               Left            =   1080
               TabIndex        =   21
               Top             =   1320
               Width           =   255
            End
            Begin VB.CheckBox UserChk 
               Height          =   255
               Index           =   3
               Left            =   1620
               TabIndex        =   19
               Top             =   960
               Width           =   255
            End
            Begin VB.CheckBox UserChk 
               Height          =   255
               Index           =   2
               Left            =   1080
               TabIndex        =   18
               Top             =   960
               Width           =   255
            End
            Begin VB.CheckBox UserChk 
               Height          =   255
               Index           =   1
               Left            =   1620
               TabIndex        =   16
               Top             =   600
               Width           =   255
            End
            Begin VB.CheckBox UserChk 
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   15
               Top             =   600
               Width           =   255
            End
            Begin VB.Label UserLbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Delete"
               Height          =   225
               Index           =   8
               Left            =   1500
               TabIndex        =   13
               Top             =   240
               Width           =   540
            End
            Begin VB.Label UserLbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Add"
               Height          =   225
               Index           =   7
               Left            =   1020
               TabIndex        =   12
               Top             =   240
               Width           =   315
            End
            Begin VB.Label UserLbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Avail"
               Height          =   225
               Index           =   6
               Left            =   240
               TabIndex        =   23
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label UserLbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Encash"
               Height          =   225
               Index           =   5
               Left            =   240
               TabIndex        =   20
               Top             =   1320
               Width           =   630
            End
            Begin VB.Label UserLbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Credit"
               Height          =   225
               Index           =   4
               Left            =   240
               TabIndex        =   17
               Top             =   960
               Width           =   495
            End
            Begin VB.Label UserLbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Opening"
               Height          =   225
               Index           =   3
               Left            =   240
               TabIndex        =   14
               Top             =   600
               Width           =   705
            End
         End
         Begin VB.Frame LoginFrame 
            Caption         =   "Login Users"
            Height          =   525
            Left            =   -72090
            TabIndex        =   47
            Top             =   1950
            Width           =   3645
            Begin VB.CheckBox UserChk 
               Caption         =   "Add/Edit/Delete of Login Users"
               Height          =   225
               Index           =   23
               Left            =   180
               TabIndex        =   48
               Top             =   210
               Width           =   2835
            End
         End
         Begin VB.Frame CorrectFrame 
            Caption         =   "Correction"
            Height          =   1065
            Left            =   -74910
            TabIndex        =   53
            Top             =   3030
            Width           =   4485
            Begin VB.CheckBox UserChk 
               Caption         =   "OT Authorization"
               Height          =   225
               Index           =   31
               Left            =   1620
               TabIndex        =   61
               Top             =   780
               Width           =   1665
            End
            Begin VB.CheckBox UserChk 
               Caption         =   "Time"
               Height          =   225
               Index           =   22
               Left            =   90
               TabIndex        =   60
               Top             =   780
               Width           =   735
            End
            Begin VB.CheckBox UserChk 
               Caption         =   "OT"
               Height          =   225
               Index           =   21
               Left            =   3090
               TabIndex        =   59
               Top             =   540
               Width           =   855
            End
            Begin VB.CheckBox UserChk 
               Caption         =   "Off Duty"
               Height          =   225
               Index           =   20
               Left            =   1620
               TabIndex        =   58
               Top             =   510
               Width           =   975
            End
            Begin VB.CheckBox UserChk 
               Caption         =   "On Duty"
               Height          =   225
               Index           =   19
               Left            =   90
               TabIndex        =   57
               Top             =   510
               Width           =   975
            End
            Begin VB.CheckBox UserChk 
               Caption         =   "Status"
               Height          =   225
               Index           =   18
               Left            =   3090
               TabIndex        =   56
               Top             =   240
               Width           =   855
            End
            Begin VB.CheckBox UserChk 
               Caption         =   "Record"
               Height          =   225
               Index           =   17
               Left            =   1620
               TabIndex        =   55
               Top             =   210
               Width           =   975
            End
            Begin VB.CheckBox UserChk 
               Caption         =   "shift"
               Height          =   225
               Index           =   16
               Left            =   90
               TabIndex        =   54
               Top             =   210
               Width           =   855
            End
         End
         Begin VB.Frame ShiftFrame 
            Caption         =   "Shift Schedule"
            Height          =   525
            Left            =   -74910
            TabIndex        =   40
            Top             =   1410
            Width           =   6465
            Begin VB.CheckBox UserChk 
               Caption         =   "Change"
               Height          =   225
               Index           =   15
               Left            =   3000
               TabIndex        =   43
               Top             =   210
               Width           =   975
            End
            Begin VB.CheckBox UserChk 
               Caption         =   "Create"
               Height          =   225
               Index           =   14
               Left            =   1620
               TabIndex        =   42
               Top             =   210
               Width           =   1215
            End
            Begin VB.CheckBox UserChk 
               Caption         =   "Edit"
               Height          =   225
               Index           =   13
               Left            =   90
               TabIndex        =   41
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame YrLvFrame 
            Caption         =   "Yearly Leaves"
            Height          =   525
            Left            =   -74910
            TabIndex        =   34
            Top             =   870
            Width           =   2805
            Begin VB.CheckBox UserChk 
               Caption         =   "Update"
               Height          =   225
               Index           =   12
               Left            =   1620
               TabIndex        =   36
               Top             =   180
               Width           =   1095
            End
            Begin VB.CheckBox UserChk 
               Caption         =   "Create"
               Height          =   225
               Index           =   11
               Left            =   90
               TabIndex        =   35
               Top             =   210
               Width           =   1335
            End
         End
         Begin VB.Frame ProcessFrame 
            Caption         =   "Process"
            Height          =   525
            Left            =   -73470
            TabIndex        =   29
            Top             =   330
            Width           =   2835
            Begin VB.CheckBox UserChk 
               Caption         =   "Monthly"
               Height          =   225
               Index           =   10
               Left            =   1560
               TabIndex        =   31
               Top             =   240
               Width           =   1095
            End
            Begin VB.CheckBox UserChk 
               Caption         =   "Daily"
               Height          =   225
               Index           =   9
               Left            =   180
               TabIndex        =   30
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame ParameterFrame 
            Caption         =   "Parameter"
            Height          =   525
            Left            =   -74910
            TabIndex        =   27
            Top             =   330
            Width           =   1425
            Begin VB.CheckBox UserChk 
               Caption         =   "Edit"
               Height          =   225
               Index           =   8
               Left            =   90
               TabIndex        =   28
               Top             =   210
               Width           =   765
            End
         End
      End
      Begin VB.Frame PassFrame 
         Caption         =   "Password / UserName"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   2
         Top             =   360
         Width           =   6615
         Begin VB.TextBox Passtxt 
            Appearance      =   0  'Flat
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   4470
            MaxLength       =   16
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox UserNametxt 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1020
            MaxLength       =   16
            TabIndex        =   6
            Text            =   " "
            Top             =   720
            Width           =   1875
         End
         Begin VB.TextBox UsrCodeTxt 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   1020
            TabIndex        =   4
            Text            =   " "
            Top             =   240
            Width           =   975
         End
         Begin VB.Label UserLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Opening"
            Height          =   225
            Index           =   2
            Left            =   3390
            TabIndex        =   7
            Top             =   780
            Width           =   705
         End
         Begin VB.Label UserLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   510
         End
         Begin VB.Label UserLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Code"
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   3
            Top             =   270
            Width           =   900
         End
      End
      Begin MSFlexGridLib.MSFlexGrid ListGrid 
         Height          =   4335
         Left            =   420
         TabIndex        =   1
         Top             =   720
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   7646
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         HighLight       =   2
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
Attribute VB_Name = "UsersFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddFlag As Boolean
Dim EditFlag As Boolean
Dim UsrCnt%
Dim cnt%
''
Dim adrsC As New ADODB.Recordset
''DJ Variables
Dim strRghts As String, strChkPass As String
Dim blnSELUN1 As Boolean, blnSELUN2 As Boolean, blnChk As Boolean

Private Sub Form_Activate()
On Error GoTo Err_P
If UCase(Trim(UserName)) <> strPrintUser Then
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "Select Lv_Rights from user_leave_rights where username=" & "'" & UserName & _
        "'", VstarDataEnv.cnDJConn, adOpenStatic
        If Mid(adrsTemp(0), 24, 1) <> "1" Then
                MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
                Unload Me
        End If
        adrsTemp.Close
End If
Exit Sub
Err_P:
    ShowError ("Users :: Users ")
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Err_P
Call SetFormIcon(Me)
blnSELUN1 = True
blnSELUN2 = True
blnChk = True
Call RetCaptions
ExitEditMode
Passtxt = ""
'*********check if the person has administrative login then enable ADD,DELETE,EDIT button
If Not blnBackRes Then
    frBackRes.Visible = False
    frCompact.Visible = False
Else
    frBackRes.Visible = True
    frCompact.Visible = True
End If
Exit Sub
Err_P:
    ShowError ("Users :: Users")
End Sub

Private Sub RetCaptions()
On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '56%'", VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
Me.Caption = NewCaptionTxt("56001", adrsC)
UserTab.TabCaption(0) = NewCaptionTxt("56002", adrsC)
UserTab.TabCaption(1) = NewCaptionTxt("56003", adrsC)
PassFrame.Caption = NewCaptionTxt("56004", adrsC)
RightsTab.TabCaption(0) = NewCaptionTxt("56005", adrsC)
RightsTab.TabCaption(1) = NewCaptionTxt("56006", adrsC)
LeaveTranFrame.Caption = NewCaptionTxt("56007", adrsC)
Sel_UnSelectCmd.Caption = NewCaptionTxt("56008", adrsC)
otherCmd.Caption = NewCaptionTxt("56008", adrsC)
ParameterFrame.Caption = NewCaptionTxt("56009", adrsC)
ProcessFrame.Caption = NewCaptionTxt("56010", adrsC)
YrLvFrame.Caption = NewCaptionTxt("56011", adrsC)
UserChk(24).Caption = NewCaptionTxt("56012", adrsC)
UserChk(8).Caption = NewCaptionTxt("56012", adrsC)
UserChk(9).Caption = NewCaptionTxt("56013", adrsC)
UserChk(10).Caption = NewCaptionTxt("56014", adrsC)
UserChk(11).Caption = NewCaptionTxt("56015", adrsC)
UserChk(12).Caption = NewCaptionTxt("56016", adrsC)
UserChk(13).Caption = NewCaptionTxt("56012", adrsC)
UserChk(14).Caption = NewCaptionTxt("56015", adrsC)
UserChk(15).Caption = NewCaptionTxt("56017", adrsC)
UserChk(16).Caption = NewCaptionTxt("00031", adrsMod)
UserChk(17).Caption = NewCaptionTxt("56018", adrsC)
UserChk(18).Caption = NewCaptionTxt("00033", adrsMod)
UserChk(19).Caption = NewCaptionTxt("56019", adrsC)
UserChk(20).Caption = NewCaptionTxt("56020", adrsC)
UserChk(21).Caption = NewCaptionTxt("56021", adrsC)
UserChk(22).Caption = NewCaptionTxt("56022", adrsC)
UserChk(23).Caption = NewCaptionTxt("56023", adrsC)
UserChk(26).Caption = NewCaptionTxt("56012", adrsC)
UserChk(28).Caption = NewCaptionTxt("56024", adrsC)
UserChk(29).Caption = NewCaptionTxt("56025", adrsC)
LoginFrame.Caption = NewCaptionTxt("56026", adrsC)
CorrectFrame.Caption = NewCaptionTxt("56027", adrsC)
ShiftFrame.Caption = NewCaptionTxt("56028", adrsC)
PaidFrame.Caption = NewCaptionTxt("56029", adrsC)
frRep.Caption = NewCaptionTxt("00105", adrsMod)
UserLbl(0).Caption = NewCaptionTxt("56030", adrsC)
UserLbl(1).Caption = NewCaptionTxt("00048", adrsMod)
UserLbl(2).Caption = NewCaptionTxt("56031", adrsC)
UserLbl(4).Caption = NewCaptionTxt("56032", adrsC)
UserLbl(5).Caption = NewCaptionTxt("56033", adrsC)
UserLbl(6).Caption = NewCaptionTxt("56034", adrsC)
UserLbl(7).Caption = NewCaptionTxt("56035", adrsC)
UserLbl(8).Caption = NewCaptionTxt("56036", adrsC)
'' New
UserLbl(3).Caption = NewCaptionTxt("56037", adrsC)
RulesFrame.Caption = NewCaptionTxt("56038", adrsC)
UserChk(27).Caption = NewCaptionTxt("56025", adrsC)
frCompact.Caption = NewCaptionTxt("56039", adrsC)
UserChk(25).Caption = NewCaptionTxt("56040", adrsC)
UserChk(30).Caption = NewCaptionTxt("00105", adrsMod)
UserChk(31).Caption = NewCaptionTxt("00104", adrsMod)

frBackRes.Caption = NewCaptionTxt("56058", adrsC)
'' End
'all the rights for yearly and monthly tables will be placed int the userpermission table
FillListGrid
AddSaveCmd.Caption = NewCaptionTxt("00004", adrsMod)
EditCanCmd.Caption = NewCaptionTxt("00005", adrsMod)
DeleteCmd.Caption = NewCaptionTxt("00006", adrsMod)
ExitCmd.Caption = NewCaptionTxt("00008", adrsMod)
MSF1.ToolTipText = NewCaptionTxt("56059", adrsC)

End Sub

Private Sub AddSaveCmd_Click()
On Error GoTo Err_P
If AddFlag = False And EditFlag = False Then
    If (ListGrid.Rows - 1) >= CByte(InVar.bytUse) Then
        MsgBox NewCaptionTxt("56041", adrsC) & InVar.bytUse, vbInformation
        Exit Sub
    End If
    AddSaveCmd.Caption = NewCaptionTxt("00007", adrsMod)
    AddSaveCmd.Enabled = False
    EditCanCmd.Enabled = True
    EditCanCmd.Caption = NewCaptionTxt("00003", adrsMod)
    DeleteCmd.Enabled = False
    ''ExitCmd.Enabled = False
    EditMode
    UserTab.Tab = 1
    RightsTab.Tab = 0
    UserClearCtls
    AddFlag = True
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "select usercode from UserInfo"
    If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            adrsTemp.MoveLast
            UsrCnt% = adrsTemp(0) + 1
            UsrCodeTxt = UsrCnt
    Else
            UsrCnt% = 1
            UsrCodeTxt = UsrCnt
    End If
    Passtxt.Visible = True
    UserNametxt.Text = ""
    Passtxt.Text = ""
    UserNametxt.SetFocus
    Exit Sub
ElseIf EditFlag = True Or AddFlag = True Then
    If AddFlag Then
            If Trim(UserNametxt) = "" Then
                    MsgBox NewCaptionTxt("56042", adrsC), vbExclamation
                    UserNametxt.SetFocus
                    Exit Sub
            End If
            If Trim(Passtxt) = "" Then
                    MsgBox NewCaptionTxt("56043", adrsC), vbExclamation
                    Passtxt.SetFocus
                    Exit Sub
            End If
            If UCase(Trim(UserNametxt.Text)) = UCase(strPrintUser) Then
                MsgBox "Black Code :: Unrecognized Operation", vbCritical
                Unload Me
                Exit Sub
            End If
    End If
    If InStr(DEncryptDat(UCase(Passtxt.Text), 1), "'") > 0 Then
        MsgBox NewCaptionTxt("56044", adrsC), vbExclamation
        Passtxt.SetFocus
    End If
    If InStr(DEncryptDat(UCase(Passtxt.Text), 1), Chr(34)) > 0 Then
        MsgBox NewCaptionTxt("56044", adrsC), vbExclamation
        Passtxt.SetFocus
    End If
    If AddFlag = True Then
            DoEvents
            Call AddUser
    End If
    If EditFlag Then
            Call SaveModLog                         '' Save the Edit Log
            Call UpdateUserRits
    End If
    AddSaveCmd.Caption = NewCaptionTxt("00004", adrsMod)
    AddSaveCmd.Enabled = True
    EditCanCmd.Enabled = True
    EditCanCmd.Caption = NewCaptionTxt("00005", adrsMod)
    DeleteCmd.Enabled = True
    ExitCmd.Enabled = True
    ExitEditMode
    UserClearCtls
    FillListGrid
    If AddFlag = True Then AddFlag = False
    If EditFlag = True Then EditFlag = False
    Passtxt.Visible = True
    UserTab.Tab = 0
    Exit Sub
End If
Exit Sub
Err_P:
Select Case Err.Number
    Case 1002
            MsgBox NewCaptionTxt("56045", adrsC), vbExclamation, NewCaptionTxt("56046", adrsC)
            UserTab.Tab = 0
    Case 5
            Resume Next
    Case 3033  'incorrect password
            MsgBox Err.Description, vbCritical
    Case Else
            MsgBox Err.Number & vbCrLf & Err.Description, vbCritical
End Select
End Sub

Private Sub DeleteCmd_Click()
On Error GoTo ERR_Particular
If ListGrid.Rows = 1 Then Exit Sub
ListGrid.Col = 1
If UserTab.Tab = 0 Then UserTab.Tab = 1
If UCase(UserNametxt.Text) = UCase(UserName) Then
        MsgBox NewCaptionTxt("56047", adrsC), vbExclamation
Else
        If UserNametxt.Text <> "" Then
                If (MsgBox(NewCaptionTxt("56048", adrsC) & " " & UserNametxt.Text & " ?", vbYesNo + vbQuestion)) = vbYes Then
                    Dim strCaseConv As String
                    Select Case bytBackEnd
                        Case 1  '' MS-SQL Server
                            strCaseConv = "Upper("
                        Case 2  '' Ms-Access
                            strCaseConv = "Ucase("
                    End Select
                        Call AddActivityLog(lgDelete_Action, 1, 29)     '' Delete Log
                        ''Delete From UserInfo
                        VstarDataEnv.cnDJConn.Execute "Delete from UserInfo where " & strCaseConv & "username)='" & UCase(UserNametxt.Text) & _
                        "'"
                        'Delete From UserLeave Rights
                        VstarDataEnv.cnDJConn.Execute "Delete from User_Leave_Rights where " & strCaseConv & "username)='" & _
                        UCase(UserNametxt.Text) & "'"
                        UCase (UserNametxt.Text) & "'"
                        ''Delete From UserPermission
                        VstarDataEnv.cnDJConn.Execute "Delete from Userpermission where " & strCaseConv & "username)='" & _
                        UCase(UserNametxt.Text) & "'"
                        Call FillListGrid
                        UserTab.Tab = 0
                  End If
        Else
                MsgBox NewCaptionTxt("56049", adrsC), vbExclamation
        End If
End If
Exit Sub
ERR_Particular:
    ShowError ("Delete :: " & Me.Caption)
End Sub

Private Sub EditCanCmd_Click()
On Error GoTo Err_P
If EditFlag = False And AddFlag = False Then
        If ListGrid.Rows = 1 Then Exit Sub
        EditMode
        AddSaveCmd.Caption = NewCaptionTxt("00007", adrsMod)
        AddSaveCmd.Enabled = False
        EditCanCmd.Enabled = True
        EditCanCmd.Caption = NewCaptionTxt("00003", adrsMod)
        DeleteCmd.Enabled = False
        ExitCmd.Enabled = False
        UserTab.Tab = 1
        EditFlag = True
ElseIf EditFlag = True Or AddFlag = True Then
        ExitEditMode
        AddSaveCmd.Caption = NewCaptionTxt("00004", adrsMod)
        AddSaveCmd.Enabled = True
        EditCanCmd.Enabled = True
        EditCanCmd.Caption = NewCaptionTxt("00005", adrsMod)
        DeleteCmd.Enabled = True
        ExitCmd.Enabled = True
        If EditFlag = True Then EditFlag = False
        If AddFlag = True Then AddFlag = False
        UserClearCtls
        UserTab.Tab = 0
End If
Exit Sub
Err_P:
    ShowError ("Edit / Cancel :: " & Me.Caption)
End Sub

Private Sub FillUserGrid()
On Error GoTo Err_P
With MSF1
        .ColWidth(0) = 1200
        .ColWidth(1) = 950
        .ColWidth(2) = 950
        .ColWidth(3) = 950
        .Row = 0
        .Col = 0
        .Text = NewCaptionTxt("56056", adrsC)
        .Col = 1
        .Text = NewCaptionTxt("56035", adrsC)
        .Col = 2
        .Text = NewCaptionTxt("56036", adrsC)
        .Col = 3
        .Text = NewCaptionTxt("56012", adrsC)
End With
Call UpdateUserGrid
Exit Sub
Err_P:
    ShowError ("FillUserGrid :: " & Me.Caption)
End Sub
        
Private Sub AddUser()       '' Procedure to Add User
On Error GoTo ERR_Particular
Dim intErr As Integer, strERR As String, strUSER_ALREADY As String
strUSER_ALREADY = NewCaptionTxt("56050", adrsC)
If adrsRits.State = 1 Then adrsRits.Close
adrsRits.Open "Select * from userinfo where username='" & UCase(UserNametxt.Text) & "'"
If Not (adrsRits.EOF And adrsRits.BOF) Then
        intErr = intErr + 1
        strERR = strERR & vbCrLf & strUSER_ALREADY
End If
adrsRits.Close
If intErr > 0 Then
        MsgBox strERR, vbExclamation
Else
        Call SaveAddLog                         '' Save the Add Log
        '' insert into User Info Table
        VstarDataEnv.cnDJConn.Execute "insert into UserInfo Values(" & UsrCodeTxt.Text & ",'" & UserNametxt.Text & "','" & _
        DEncryptDat(UCase(Trim(Passtxt.Text)), 1) & "',1)"
        '' insert into userLeave rights Table
        VstarDataEnv.cnDJConn.Execute "insert into User_leave_rights(usercode,username) Values(" & UsrCodeTxt.Text & _
        ",'" & UserNametxt.Text & "')"
        ''insert into UserPermission table
        Dim bytCnt As Byte, strMast(1 To 14) As String
        For bytCnt = 1 To UBound(strMast)
            strMast(bytCnt) = Choose(bytCnt, "SHIFT", "CATEGORY", "OT RULES", "CO RULES", _
            "LEAVE", "DEPARTMENT", "HOLIDAY", "DECLARE", "EMPLOYEE", "LOST", "ROTATION", _
            "GROUP", "LOCATION", "COMPANY")
            VstarDataEnv.cnDJConn.Execute "insert into userpermission values('" & _
            UserNametxt.Text & "','" & strMast(bytCnt) & _
            "',1,1,1)"
        Next
        Call UpdateUserRits
End If
Exit Sub
ERR_Particular:
    ShowError ("Add User :: " & Me.Caption)
    If adrsRits.State = 1 Then adrsRits.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
AddFlag = False
EditFlag = False
End Sub

Private Sub ListGrid_Click()
If ListGrid.RowSel > cnt% Then
        ListGrid.HighLight = flexHighlightNever
        ListGrid.FocusRect = flexFocusNone
Else
        ListGrid.HighLight = flexHighlightWithFocus
        ListGrid.FocusRect = flexFocusLight
End If
End Sub

Private Sub ListGrid_DblClick()
If ListGrid.Row > 0 And ListGrid.Text <> "" Then
        RetUserFields
        UserTab.Tab = 1
End If
End Sub

Private Sub MSF1_DblClick()
If EditFlag = True Or AddFlag = True Then
        Select Case MSF1.Row
                Case 0
                        If MSF1.Col > 0 And MSF1.Text = "" Then
                                Dim bytCnt As Byte
                                For bytCnt = 1 To 13
                                        If blnChk Then
                                                MSF1.TextMatrix(bytCnt, MSF1.Col) = "Yes"
                                        Else
                                                MSF1.TextMatrix(bytCnt, MSF1.Col) = "No"
                                        End If
                                Next
                                blnChk = Not blnChk
                                AddSaveCmd.Enabled = True
                        End If
                Case 1 To 14
                        If MSF1.Text = "Yes" Then
                                MSF1.Text = "No"
                                AddSaveCmd.Enabled = True
                                Exit Sub
                        End If
                        If MSF1.Text = "No" Then
                                MSF1.Text = "Yes"
                                AddSaveCmd.Enabled = True
                        End If
        End Select
End If
End Sub

Private Sub FillListGrid()
On Error GoTo ERR_Particular
With ListGrid
        .Clear
        .Row = 0
        .Col = 0
        .Text = NewCaptionTxt("56030", adrsC)
        .Col = 1
        .Text = NewCaptionTxt("56057", adrsC)
End With
'*********8
With ListGrid
        .ColWidth(0) = .Width / 2
        .ColWidth(1) = .Width / 2
End With
cnt = 1
With ListGrid
        cnt = 1
        .Rows = 1
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "select usercode,username from userinfo ", VstarDataEnv.cnDJConn, adOpenStatic
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
                .Rows = adrsTemp.RecordCount + 1
                Do Until adrsTemp.EOF
                        .Row = cnt
                        .Col = 0
                        .Text = adrsTemp!usercode
                        .Col = 1
                        .Text = adrsTemp!UserName
                        adrsTemp.MoveNext
                        cnt = cnt + 1
                Loop
        Else
                MsgBox NewCaptionTxt("56051", adrsC) & vbCrLf & NewCaptionTxt("56052", adrsC), vbInformation
                End
        End If
End With
Exit Sub
ERR_Particular:
    ShowError ("FillListGrid :: " & Me.Caption)
End Sub

Private Sub RetUserFields()
On Error GoTo ERR_Particular
ListGrid.Col = 1
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "select *  from userinfo where username=" & "'" & ListGrid.Text & "'", VstarDataEnv.cnDJConn
UsrCodeTxt = adrsTemp!usercode
UserNametxt.Text = adrsTemp!UserName
Passtxt.Text = UCase(DEncryptDat(adrsTemp!Password, 1))
strChkPass = UCase(DEncryptDat(adrsTemp!Password, 1))
UserChk(23).Value = adrsTemp!UserRights
adrsTemp.Close
adrsTemp.Open "select *  from user_leave_rights where username=" & "'" & ListGrid.Text & "'", VstarDataEnv.cnDJConn
strRghts = IIf(IsNull(adrsTemp("lv_rights")), "", adrsTemp("lv_rights"))
''Leave transaction '8 characters
UserChk(0).Value = IIf(Mid(strRghts, 1, 1) = "1", 1, 0)
UserChk(1).Value = IIf(Mid(strRghts, 2, 1) = "1", 1, 0)
UserChk(2).Value = IIf(Mid(strRghts, 3, 1) = "1", 1, 0)
UserChk(3).Value = IIf(Mid(strRghts, 4, 1) = "1", 1, 0)
UserChk(4).Value = IIf(Mid(strRghts, 5, 1) = "1", 1, 0)
UserChk(5).Value = IIf(Mid(strRghts, 6, 1) = "1", 1, 0)
UserChk(6).Value = IIf(Mid(strRghts, 7, 1) = "1", 1, 0)
UserChk(7).Value = IIf(Mid(strRghts, 8, 1) = "1", 1, 0)
UserChk(8).Value = IIf(Mid(strRghts, 9, 1) = "1", 1, 0)
UserChk(9).Value = IIf(Mid(strRghts, 10, 1) = "1", 1, 0)
UserChk(10).Value = IIf(Mid(strRghts, 11, 1) = "1", 1, 0)
UserChk(11).Value = IIf(Mid(strRghts, 12, 1) = "1", 1, 0)
UserChk(12).Value = IIf(Mid(strRghts, 13, 1) = "1", 1, 0)
UserChk(13).Value = IIf(Mid(strRghts, 14, 1) = "1", 1, 0)
UserChk(14).Value = IIf(Mid(strRghts, 15, 1) = "1", 1, 0)
UserChk(15).Value = IIf(Mid(strRghts, 16, 1) = "1", 1, 0)
UserChk(16).Value = IIf(Mid(strRghts, 17, 1) = "1", 1, 0)
UserChk(17).Value = IIf(Mid(strRghts, 18, 1) = "1", 1, 0)
UserChk(18).Value = IIf(Mid(strRghts, 19, 1) = "1", 1, 0)
UserChk(19).Value = IIf(Mid(strRghts, 20, 1) = "1", 1, 0)
UserChk(20).Value = IIf(Mid(strRghts, 21, 1) = "1", 1, 0)
UserChk(21).Value = IIf(Mid(strRghts, 22, 1) = "1", 1, 0)
UserChk(22).Value = IIf(Mid(strRghts, 23, 1) = "1", 1, 0)
UserChk(23).Value = IIf(Mid(strRghts, 24, 1) = "1", 1, 0)
UserChk(24).Value = IIf(Mid(strRghts, 25, 1) = "1", 1, 0)
UserChk(25).Value = IIf(Mid(strRghts, 26, 1) = "1", 1, 0)
UserChk(26).Value = IIf(Mid(strRghts, 27, 1) = "1", 1, 0)
UserChk(27).Value = IIf(Mid(strRghts, 28, 1) = "1", 1, 0)
UserChk(28).Value = IIf(Mid(strRghts, 29, 1) = "1", 1, 0)
UserChk(29).Value = IIf(Mid(strRghts, 30, 1) = "1", 1, 0)
UserChk(30).Value = IIf(Mid(strRghts, 31, 1) = "1", 1, 0)
UserChk(31).Value = IIf(Mid(strRghts, 32, 1) = "1", 1, 0)
'**********
ListGrid.Col = 1
FillUserGrid
strRghts = ""
Exit Sub
ERR_Particular:
    ShowError ("RetUserFields :: " & Me.Caption)
End Sub

Private Sub MSF1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then Call MSF1_DblClick
End Sub

Private Sub otherCmd_Click()
If AddFlag = True Or EditFlag = True Then
    Dim bytCnt As Byte
    If blnSELUN2 = True Then
        For bytCnt = 8 To 31
                UserChk(bytCnt).Value = 1
        Next
    End If
    If blnSELUN2 = False Then
        For bytCnt = 8 To 31
                UserChk(bytCnt).Value = 0
        Next
    End If
    blnSELUN2 = Not blnSELUN2
End If
End Sub

Private Sub PassTxt_Change()
        If Trim(Passtxt) <> "" Then AddSaveCmd.Enabled = True
End Sub

Private Sub Passtxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 6)
End If
End Sub

Private Sub RightsTab_Click(PreviousTab As Integer)
If RightsTab.Tab = 1 And Not (EditFlag Or AddFlag) Then
            RightsTab.Tab = PreviousTab
End If
End Sub

Private Sub Sel_UnSelectCmd_Click()
If AddFlag = True Or EditFlag = True Then
        Dim bytCnt As Byte
        If blnSELUN1 = True Then
                For bytCnt = 1 To 14
                        With MSF1
                                .TextMatrix(bytCnt, 1) = "Yes"
                                .TextMatrix(bytCnt, 2) = "Yes"
                                .TextMatrix(bytCnt, 3) = "Yes"
                        End With
                Next
                For bytCnt = 0 To 7
                        UserChk(bytCnt).Value = 1
                Next
        Else
                For bytCnt = 1 To 14
                        With MSF1
                                .TextMatrix(bytCnt, 1) = "No"
                                .TextMatrix(bytCnt, 2) = "No"
                                .TextMatrix(bytCnt, 3) = "No"
                        End With
                Next
                For bytCnt = 0 To 7
                        UserChk(bytCnt).Value = 0
                Next
        End If
        blnSELUN1 = Not blnSELUN1
End If
End Sub

Private Sub UserChk_Click(Index As Integer)
        AddSaveCmd.Enabled = True
End Sub

Private Sub UserNametxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 6)
End If
End Sub

Private Sub UserTab_Click(PreviousTab As Integer)
''ADO
If UserTab.Tab = 1 And (Not AddFlag And Not EditFlag) Then
        ListGrid.Col = 1
        If ListGrid.Text <> "" Then
                RetUserFields
        Else
                UserTab.Tab = 0
        End If
End If
End Sub

Private Sub EditMode()
UserNametxt.Enabled = True
Passtxt.Enabled = True
LeaveTranFrame.Enabled = True
ParameterFrame.Enabled = True
ProcessFrame.Enabled = True
YrLvFrame.Enabled = True
ShiftFrame.Enabled = True
CorrectFrame.Enabled = True
LoginFrame.Enabled = True
Sel_UnSelectCmd.Enabled = True
PaidFrame.Enabled = True
frRep.Enabled = True
frCompact.Enabled = True
RulesFrame.Enabled = True
End Sub

Private Sub ExitEditMode()
UsrCodeTxt.Enabled = False
UserNametxt.Enabled = False
Passtxt.Enabled = False
LeaveTranFrame.Enabled = False
ParameterFrame.Enabled = False
ProcessFrame.Enabled = False
YrLvFrame.Enabled = False
ShiftFrame.Enabled = False
CorrectFrame.Enabled = False
LoginFrame.Enabled = False
Sel_UnSelectCmd.Enabled = False
PaidFrame.Enabled = False
frRep.Enabled = False
frCompact.Enabled = False
RulesFrame.Enabled = False
End Sub

Public Sub EnableAllMaster()
UserChk(0).Value = 1
UserChk(1).Value = 1
UserChk(2).Value = 1
UserChk(3).Value = 1
UserChk(4).Value = 1
UserChk(5).Value = 1
UserChk(6).Value = 1
UserChk(7).Value = 1
End Sub

Public Sub DisableAllMaster()
UserChk(0).Value = 0
UserChk(1).Value = 0
UserChk(2).Value = 0
UserChk(3).Value = 0
UserChk(4).Value = 0
UserChk(5).Value = 0
UserChk(6).Value = 0
UserChk(7).Value = 0
End Sub

Public Sub UserClearCtls()
For i = 0 To 31
        On Error Resume Next
        UserChk(i).Value = 0
Next i
Call ClearTextGrid
UsrCodeTxt = ""
UserNametxt = ""
Passtxt.Text = ""
End Sub

Private Sub UpdateUserGrid()
On Error GoTo ERR_Particular
Dim strArrMst()  As String, bytCnt As Byte
For bytCnt = 0 To 12
        ReDim Preserve strArrMst(bytCnt)
        strArrMst(bytCnt) = Choose(bytCnt + 1, "SHIFT", "CATEGORY", "OT RULES", _
        "CO RULES", "LEAVE", "DEPARTMENT", "HOLIDAY", "DECLARE", "EMPLOYEE", "LOST", _
        "ROTATION", "GROUP", "LOCATION", "COMPANY")
Next
'' Code to Fill the Grid in a Loop
For bytCnt = 0 To 13
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "Select * from userpermission where UserName='" & UserNametxt.Text & _
    "'", VstarDataEnv.cnDJConn, adOpenKeyset
    Dim bytCntGrid As Byte
    For bytCntGrid = 1 To 14
        Dim strYN As String
        With MSF1
                .TextMatrix(bytCntGrid, 0) = adrsTemp(1)
                If adrsTemp(2) = 0 Then
                        strYN = "No"
                Else
                        strYN = "Yes"
                End If
                .TextMatrix(bytCntGrid, 1) = strYN
                If adrsTemp(3) = 0 Then
                        strYN = "No"
                Else
                        strYN = "Yes"
                End If
                .TextMatrix(bytCntGrid, 2) = strYN
                If adrsTemp(4) = 0 Then
                        strYN = "No"
                Else
                        strYN = "Yes"
                End If
                .TextMatrix(bytCntGrid, 3) = strYN
                If bytCntGrid <> adrsTemp.RecordCount Then adrsTemp.MoveNext
        End With
    Next
Next
Exit Sub
ERR_Particular:
    ShowError ("UpdateUserGrid :: " & Me.Caption)
    Resume Next
End Sub

Private Sub UpdateUserRits()
On Error GoTo ERR_Particular
Err.Clear
Dim Rts As String
'' Change Password
If EditFlag Then
        If Trim(Passtxt) <> "" Then
                If UCase(strChkPass) <> UCase(Trim(Passtxt.Text)) Then
                        Select Case MsgBox(NewCaptionTxt("56053", adrsC) & vbCrLf & vbTab & NewCaptionTxt("56054", adrsC), _
                                vbYesNo + vbQuestion)
                                Case vbYes
                                        VstarDataEnv.cnDJConn.Execute "Update Userinfo set Password='" & DEncryptDat(UCase(Trim(Passtxt.Text)), 1) & _
                                        "' Where UserName='" & UserNametxt.Text & "'"
                                        MsgBox NewCaptionTxt("56055", adrsC), vbInformation
                        End Select
                End If
        End If
End If
''Check for User Administration Rights
If UserChk(23).Value = 1 Then
        VstarDataEnv.cnDJConn.Execute "Update Userinfo set UserRights=1 Where UserName='" & UserNametxt.Text & "'"
ElseIf UserChk(23).Value = 0 Then
        VstarDataEnv.cnDJConn.Execute "Update Userinfo set UserRights=0 Where UserName='" & UserNametxt.Text & "'"
End If
Call GridReflect
''Check for Other Rights
Rts = UserChk(0).Value & UserChk(1).Value & UserChk(2).Value & UserChk(3).Value & _
UserChk(4).Value & UserChk(5).Value & UserChk(6).Value & UserChk(7).Value & _
UserChk(8).Value & UserChk(9).Value & UserChk(10).Value & UserChk(11).Value & _
UserChk(12).Value & UserChk(13).Value & UserChk(14).Value & UserChk(15).Value & _
UserChk(16).Value & UserChk(17).Value & UserChk(18).Value & UserChk(19).Value & _
UserChk(20).Value & UserChk(21).Value & _
UserChk(22).Value & UserChk(23).Value & UserChk(24).Value & _
UserChk(25).Value & UserChk(26).Value & UserChk(27).Value & _
UserChk(28).Value & UserChk(29).Value & UserChk(30).Value & _
UserChk(31).Value
VstarDataEnv.cnDJConn.Execute "Update user_leave_rights set lv_rights='" & Rts & "' Where Username='" & _
UserNametxt.Text & "'"
Exit Sub
ERR_Particular:
    ShowError ("UpdateUserRits :: " & Me.Caption)
End Sub

Private Sub GridReflect()
On Error GoTo Err_P
Dim bytCnt As Byte
Dim bytBit(2) As Byte
For bytCnt = 1 To 14
        If UCase(MSF1.TextMatrix(bytCnt, 1)) = "YES" Then
                bytBit(0) = 1
        Else
                bytBit(0) = 0
        End If
        If UCase(MSF1.TextMatrix(bytCnt, 2)) = "YES" Then
                bytBit(1) = 1
        Else
                bytBit(1) = 0
        End If
        If UCase(MSF1.TextMatrix(bytCnt, 3)) = "YES" Then
                bytBit(2) = 1
        Else
                bytBit(2) = 0
        End If
        VstarDataEnv.cnDJConn.Execute "Update userpermission set [Add]=" & bytBit(0) & ", [Delete]=" & bytBit(1) & _
        ", [Edit]=" & bytBit(2) & _
        " Where UserName='" & UserNametxt.Text & "' and [Menu Items] ='" & MSF1.TextMatrix(bytCnt, 0) & "'"
Next
Exit Sub
Err_P:
    ShowError ("Grid Reflect :: Users")
End Sub

Private Sub UsrCodeTxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys Chr(9)
End Sub

Private Sub ClearTextGrid()
Dim bytCnt As Byte
For bytCnt = 1 To 14
    MSF1.TextMatrix(bytCnt, 1) = "No"
    MSF1.TextMatrix(bytCnt, 2) = "No"
    MSF1.TextMatrix(bytCnt, 3) = "No"
Next
End Sub

Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo Err_P
Call AddActivityLog(lgADD_MODE, 1, 29)     '' Add Activity
Exit Sub
Err_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo Err_P
Call AddActivityLog(lgEdit_Mode, 1, 29)     '' Edit Activity
Exit Sub
Err_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub ExitCmd_Click()
        Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then Call ShowF10("56")
If KeyCode = vbKeyF10 Then KeyCode = 0
End Sub
