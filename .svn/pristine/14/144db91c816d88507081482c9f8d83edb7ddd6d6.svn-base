VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAvail 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraRemark 
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   0
      TabIndex        =   28
      Top             =   4815
      Visible         =   0   'False
      Width           =   8115
      Begin VB.TextBox txtRemark 
         Height          =   915
         Left            =   855
         MaxLength       =   256
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   29
         Top             =   0
         Width           =   7260
      End
      Begin VB.Label lblRemark 
         AutoSize        =   -1  'True
         Caption         =   "Remark::"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   45
         TabIndex        =   30
         Top             =   90
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   435
      Left            =   5430
      TabIndex        =   11
      Top             =   4320
      Width           =   2715
   End
   Begin VB.CommandButton cmdDelCan 
      Caption         =   "Command2"
      Height          =   435
      Left            =   2730
      TabIndex        =   10
      Top             =   4320
      Width           =   2715
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command3"
      Height          =   435
      Left            =   0
      TabIndex        =   9
      Top             =   4320
      Width           =   2745
   End
   Begin TabDlg.SSTab TB1 
      Height          =   3555
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6271
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
      TabPicture(0)   =   "frmAvail.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSF1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrColour"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmAvail.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frAvail"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame FrColour 
         Caption         =   "Cut Leave By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   -68475
         TabIndex        =   31
         Top             =   495
         Width           =   1590
         Begin VB.Label lvlG 
            BackColor       =   &H0000FF00&
            Height          =   240
            Left            =   90
            TabIndex        =   35
            Top             =   585
            Width           =   420
         End
         Begin VB.Label Label1 
            Caption         =   "Early Rule"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   585
            TabIndex        =   34
            Top             =   585
            Width           =   915
         End
         Begin VB.Label lvlR 
            BackColor       =   &H000000FF&
            Height          =   240
            Left            =   90
            TabIndex        =   33
            Top             =   270
            Width           =   420
         End
         Begin VB.Label Label3 
            Caption         =   "Late Rule"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   630
            TabIndex        =   32
            Top             =   270
            Width           =   870
         End
      End
      Begin VB.Frame frAvail 
         Height          =   3180
         Left            =   30
         TabIndex        =   17
         Top             =   330
         Width           =   8070
         Begin VB.TextBox txtFrom 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   1770
            TabIndex        =   2
            Tag             =   "D"
            Top             =   1020
            Width           =   1215
         End
         Begin VB.TextBox txtTo 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   1770
            TabIndex        =   4
            Tag             =   "D"
            Text            =   "  "
            Top             =   1410
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid MSF2 
            Height          =   2535
            Left            =   4410
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   600
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   4471
            _Version        =   393216
            Rows            =   1
            Cols            =   3
            FixedCols       =   0
            BackColorFixed  =   4194368
            ForeColorFixed  =   8454143
            GridColor       =   4194368
            Appearance      =   0
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
         Begin MSMask.MaskEdBox txtDays 
            Height          =   300
            Left            =   1770
            TabIndex        =   8
            Top             =   2565
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
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
         Begin MSForms.ComboBox cboCOEntry 
            Height          =   345
            Left            =   2280
            TabIndex        =   6
            Top             =   1800
            Visible         =   0   'False
            Width           =   1935
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "3413;609"
            ColumnCount     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblCOEntry 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CO for extra work done on"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   1860
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Label lblBal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " vxcvx"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4395
            TabIndex        =   25
            Top             =   255
            Width           =   570
         End
         Begin VB.Label lblDays 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "hfhf"
            Height          =   195
            Left            =   135
            TabIndex        =   24
            Top             =   2625
            Width           =   270
         End
         Begin VB.Label lblRW 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "hfhf"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   2280
            Width           =   270
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "hhfhf"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1515
            Width           =   360
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "hfh"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   225
         End
         Begin VB.Label lblLeave 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "hh"
            Height          =   195
            Left            =   105
            TabIndex        =   19
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " fgfg"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   18
            Top             =   270
            Width           =   405
         End
         Begin MSForms.ComboBox cboLeave 
            Height          =   345
            Left            =   1770
            TabIndex        =   1
            Top             =   630
            Width           =   2505
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "4419;609"
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
         Begin MSForms.ComboBox cboFrom 
            Height          =   315
            Left            =   3060
            TabIndex        =   3
            Top             =   1020
            Width           =   1215
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "2143;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboTo 
            Height          =   345
            Left            =   3060
            TabIndex        =   5
            Top             =   1410
            Width           =   1215
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "2143;609"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboRW 
            Height          =   345
            Left            =   1770
            TabIndex        =   7
            Top             =   2190
            Width           =   1215
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "2143;609"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   3105
         Left            =   -73200
         TabIndex        =   12
         Top             =   390
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   5477
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   4194368
         ForeColorFixed  =   8454143
         GridColor       =   4194368
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
   Begin VB.Label lblJoin 
      AutoSize        =   -1  'True
      Caption         =   "Join Date"
      Height          =   195
      Left            =   3840
      TabIndex        =   27
      Top             =   480
      Width           =   675
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "sdfsdffd"
      Height          =   225
      Left            =   420
      TabIndex        =   13
      Top             =   90
      Width           =   660
   End
   Begin VB.Label lblNameCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "sfsdfsdf"
      Height          =   225
      Left            =   3840
      TabIndex        =   15
      Top             =   60
      Width           =   660
   End
   Begin MSForms.ComboBox cboCode 
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   0
      Width           =   1935
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3413;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   4680
      TabIndex        =   16
      Top             =   60
      Width           =   45
   End
End
Attribute VB_Name = "frmAvail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'' CO Constant
Private Const LV_CO = "CO"
Dim strCatAvail As String               '' For Employee Category
Dim dtJoin As Date                      '' For Empoyee Joindate
Dim strF As String, strT As String      '' For From & To Status
Dim strFH As String, strSH As String    '' For From & To Status
Dim strHf_Opt As String                 '' For Half Option Status
Dim sngDays As Single                   '' For the No of Days Availed
Dim sngDiff As Single                   '' For the Difference between From & To
Dim strRW As String                     '' For the Type of Leave i.e R or W
Dim sngDaysBal As Single                '' To Get the Balance of the Particular Leave
Dim bytMin As Byte                      '' Minimum No. to be Availed
Dim bytMax As Byte                      '' Maximum No. to be Availed
Dim intTimes As Integer                 '' Maximum No of Times to be Availed
Dim strShiftDel As String               '' Get the Shift
Dim sngHDend As Single                  '' HalfDayEnd for that Shift
                                        '' shftstr$, HfEnd!
Dim blnUnPaid As Boolean                '' Flag for Unpaid Leave
Dim blnNoBal As Boolean                 '' Flag for No Balance to be kept
'' For CO Availment
Dim blnCOChecks As Boolean              '' If CO Checks are to be performed or not
Dim bytCOCode As Byte                   '' CO Rule Code
Dim bytCOLimit As Byte                  '' For COAvailment Limit
Dim sngCOBal As Single                  '' For balance of any particular Transaction
Dim sngCOAvail As Single                '' For balance of Already Availed
''
Dim adrsC As New ADODB.Recordset
Dim ELLeave As String, ELSubLeave As String

Private Sub CriteriaOneAdd(ByRef adrsRec As ADODB.Recordset, _
ByVal strTmpTrn As String, ByVal strTmpShf As String)
On Error GoTo ERR_P
Select Case Left(strHf_Opt, (Len(strHf_Opt) / 2))
    Case "FF"
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(cboLeave.Text, 2) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn _
                & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            ConMain.Execute "update " & strTmpTrn & _
            " set Presabs=" & "'" & ReplicateVal(cboLeave.Text, 2) & "'" & " where Empcode=" & _
            "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
            DateCompStr(adrsRec!Date) & strDTEnc
        End If
        
    Case "FS"
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 3, 2, cboLeave.Text) & _
                "'" & " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
            "'" & StuffVal(adrsRec!presabs, 3, 2, cboLeave.Text) & "'" & " where Empcode=" & _
            "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
            DateCompStr(adrsRec!Date) & strDTEnc
        End If
        
        ''This If Condition Add By
    Case "F "
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(cboLeave.Text, 2) & "'" & " where Empcode=" & _
                "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
                DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
            "'" & ReplicateVal(cboLeave.Text, 2) & "'" & " where Empcode=" & "'" & cboCode.Text & _
            "'" & " and  " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        End If
        

End Select

Exit Sub
ERR_P:
    ShowError ("CriteriaOneAdd ::" & Me.Caption)
End Sub

'------
Private Sub CriteriaTwoAdd(ByRef adrsRec As ADODB.Recordset, _
ByVal strTmpTrn As String, ByVal strTmpShf As String)
On Error GoTo ERR_P
Select Case Right(strHf_Opt, Len(strHf_Opt) / 2)
    Case "TF"
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 1, 2, cboLeave.Text) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
            "'" & StuffVal(adrsRec!presabs, 1, 2, cboLeave.Text) & "'" & " where Empcode=" & _
            "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
            DateCompStr(adrsRec!Date) & strDTEnc
        End If
        ''This If Condition Add By
    Case "TS"
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(cboLeave.Text, 2) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
            "'" & ReplicateVal(cboLeave.Text, 2) & "'" & " where Empcode=" & "'" & cboCode.Text & _
            "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        End If
        ''This If Condition Add By
    Case "T "
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(cboLeave.Text, 2) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
            "'" & ReplicateVal(cboLeave.Text, 2) & "'" & " where Empcode=" & "'" & cboCode.Text & "'" & _
            " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        End If
        ''This If Condition Add By
End Select
Exit Sub
ERR_P:
    ShowError ("CriteriaTwoAdd :: " & Me.Caption)
End Sub
Private Sub CriteriaThreeAdd(ByRef adrsRec As ADODB.Recordset, _
ByVal strTmpTrn As String, ByVal strTmpShf As String)
On Error GoTo ERR_P
If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
    If strRW = "R" Then
        ConMain.Execute "update " & strTmpTrn & _
        " set Presabs=" & "'" & ReplicateVal(cboLeave.Text, 2) & "'" & " where Empcode=" & _
        "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
        DateCompStr(adrsRec!Date) & strDTEnc
    End If
Else
    ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
    "'" & ReplicateVal(cboLeave.Text, 2) & "'" & " where Empcode=" & "'" & cboCode.Text & _
    "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
End If
''This If Condition Add By
Exit Sub
ERR_P:
    ShowError ("CriteriaThreeAdd :: " & Me.Caption)
End Sub
Private Sub CriteriaFourAdd(ByRef adrsRec As ADODB.Recordset, _
ByVal strTmpTrn As String, ByVal strTmpShf As String)
On Error GoTo ERR_P
Select Case strHf_Opt
    Case "FFTF"
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 1, 2, cboLeave.Text) & _
                "'" & " where Empcode=" & "'" & cboCode.Text & "'" & " and " & _
                strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            ConMain.Execute "update " & strTmpTrn & _
            " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 1, 2, cboLeave.Text) & "'" & _
            " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
            "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        End If
        ''This If Condition Add By
    Case "FSTS"
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 3, 2, cboLeave.Text) & _
                "'" & " where Empcode=" & "'" & cboCode.Text & "'" & " and " & _
                strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            ConMain.Execute "update " & strTmpTrn & _
            " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 3, 2, cboLeave.Text) & _
            "'" & " where Empcode=" & "'" & cboCode.Text & "'" & " and " & _
            strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        End If
        ''This If Condition Add By
    Case "F T "
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
            ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(cboLeave.Text, 2) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & _
                strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            ConMain.Execute "update " & strTmpTrn & _
            " set Presabs=" & "'" & ReplicateVal(cboLeave.Text, 2) & "'" & " where Empcode=" & _
            "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
            DateCompStr(adrsRec!Date) & strDTEnc
        End If
        ''This If Condition Add By
End Select
Exit Sub
ERR_P:
    ShowError ("CriteriaFourAdd :: " & Me.Caption)
End Sub
Private Sub CriteriaOneDelete(ByRef adrsRec As ADODB.Recordset, _
ByVal strTmpTrn As String, ByVal strTmpShf As String)
On Error GoTo ERR_P
Select Case Left(strHf_Opt, Len(strHf_Opt) / 2)
    Case "FF"
        If strShiftDel = pVStar.WosCode Or strShiftDel = pVStar.HlsCode Then
            ConMain.Execute "update " & strTmpTrn & " set Presabs=" & "'" & _
            StuffVal(adrsRec!presabs, 1, 2, strShiftDel) & "'" & " where Empcode=" & "'" & cboCode.Text _
            & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        Else
            If adrsRec("EntReq") = 0 Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & "'" & _
                StuffVal(adrsRec!presabs, 1, 2, pVStar.PrsCode) & "'" & " where Empcode=" & _
                "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
                Exit Sub
            End If
            If adrsRec!arrtim > 0 And adrsRec!arrtim < sngHDend Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & "'" & _
                StuffVal(adrsRec!presabs, 1, 2, pVStar.PrsCode) & "'" & " where Empcode=" & _
                "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            ElseIf (adrsRec!arrtim > 0 And adrsRec!arrtim > sngHDend) Or (adrsRec!arrtim = 0) Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & "'" & _
                StuffVal(adrsRec!presabs, 1, 2, pVStar.AbsCode) & "'" & " where Empcode=" & _
                "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        End If
    Case "FS"
        If strShiftDel = pVStar.WosCode Or strShiftDel = pVStar.HlsCode Then
            ConMain.Execute "update " & strTmpTrn & " set Presabs=" & "'" & _
            StuffVal(adrsRec!presabs, 3, 2, strShiftDel) & "'" & " where Empcode=" & "'" & _
            cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        Else
            If adrsRec("EntReq") = 0 Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & "'" & _
                StuffVal(adrsRec!presabs, 3, 2, pVStar.PrsCode) & "'" & " where Empcode=" & _
                "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & _
                strDTEnc
                Exit Sub
            End If
            If adrsRec!arrtim > 0 And adrsRec!arrtim < sngHDend Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & "'" & _
                StuffVal(adrsRec!presabs, 3, 2, pVStar.PrsCode) & "'" & " where Empcode=" & _
                "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & _
                strDTEnc
            ElseIf (adrsRec!arrtim > 0 And adrsRec!arrtim > sngHDend) Or (adrsRec!arrtim = 0) Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & "'" _
                & StuffVal(adrsRec!presabs, 3, 2, pVStar.AbsCode) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & _
                strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        End If
    Case "F "
        If strShiftDel = pVStar.WosCode Or strShiftDel = pVStar.HlsCode Then
            ConMain.Execute "update " & strTmpTrn & " set Presabs=" & "'" & _
            ReplicateVal(strShiftDel, 2) & "'" & " where Empcode=" & "'" & cboCode.Text & "'" & _
            " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        Else
            If adrsRec("EntReq") = 0 Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
                "'" & ReplicateVal(pVStar.PrsCode, 2) & "'" & " where Empcode=" & _
                "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
                DateCompStr(adrsRec!Date) & strDTEnc
                Exit Sub
            End If
            If adrsRec!arrtim > 0 And adrsRec!arrtim < sngHDend Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
                "'" & ReplicateVal(pVStar.PrsCode, 2) & "'" & " where Empcode=" & _
                "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
                DateCompStr(adrsRec!Date) & strDTEnc
            ElseIf (adrsRec!arrtim > 0 And adrsRec!arrtim > sngHDend) Or (adrsRec!arrtim = 0) Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
                "'" & ReplicateVal(pVStar.AbsCode, 2) & "'" & " where Empcode=" & "'" & _
                cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        End If
End Select
Exit Sub
ERR_P:
    ShowError ("CriteriaOneDelete :: " & Me.Caption)
End Sub

Private Sub CriteriaTwoDelete(ByRef adrsRec As ADODB.Recordset, _
ByVal strTmpTrn As String, ByVal strTmpShf As String)
On Error GoTo ERR_P
Select Case Right(strHf_Opt, Len(strHf_Opt) / 2)
    Case "TF"
        If strShiftDel = pVStar.WosCode Or strShiftDel = pVStar.HlsCode Then
            ConMain.Execute "update " & strTmpTrn & " set Presabs=" & "'" _
            & StuffVal(adrsRec!presabs, 1, 2, strShiftDel) & "'" & " where Empcode=" & "'" & _
            cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        Else
            If adrsRec("EntReq") = 0 Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
                "'" & StuffVal(adrsRec!presabs, 1, 2, pVStar.PrsCode) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
                Exit Sub
            End If
            If adrsRec!arrtim > 0 And adrsRec!arrtim < sngHDend Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
                "'" & StuffVal(adrsRec!presabs, 1, 2, pVStar.PrsCode) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            ElseIf (adrsRec!arrtim > 0 And adrsRec!arrtim > sngHDend) Or (adrsRec!arrtim = 0) Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
                "'" & StuffVal(adrsRec!presabs, 1, 2, pVStar.AbsCode) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        End If
    Case "TS"
        If strShiftDel = pVStar.WosCode Or strShiftDel = pVStar.HlsCode Then
            ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
            "'" & StuffVal(adrsRec!presabs, 3, 2, strShiftDel) & "'" & " where Empcode=" & _
            "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
            DateCompStr(adrsRec!Date) & strDTEnc
        Else
            If adrsRec("EntReq") = 0 Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 3, 2, pVStar.PrsCode) & _
                "'" & " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
                Exit Sub
            End If
            If adrsRec!arrtim > 0 And adrsRec!arrtim < sngHDend Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 3, 2, pVStar.PrsCode) & _
                "'" & " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            ElseIf (adrsRec!arrtim > 0 And adrsRec!arrtim > sngHDend) Or (adrsRec!arrtim = 0) Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 3, 2, pVStar.AbsCode) & _
                "'" & " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        End If
    Case "T "
        If strShiftDel = pVStar.WosCode Or strShiftDel = pVStar.HlsCode Then
            ConMain.Execute "update " & strTmpTrn & _
            " set Presabs=" & "'" & ReplicateVal(strShiftDel, 2) & _
            "'" & " where Empcode=" & "'" & cboCode.Text & "'" & _
            " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
            DateCompStr(adrsRec!Date) & strDTEnc
        Else
            If adrsRec("EntReq") = 0 Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(pVStar.PrsCode, 2) & _
                "'" & " where Empcode=" & "'" & cboCode.Text & "'" & _
                " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
                Exit Sub
            End If
            If adrsRec!arrtim > 0 And adrsRec!arrtim < sngHDend Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(pVStar.PrsCode, 2) & _
                "'" & " where Empcode=" & "'" & cboCode.Text & "'" & _
                " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            ElseIf (adrsRec!arrtim > 0 And adrsRec!arrtim > sngHDend) Or (adrsRec!arrtim = 0) Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(pVStar.AbsCode, 2) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        End If
End Select
Exit Sub
ERR_P:
    ShowError ("CriteriaTwoDelete :: " & Me.Caption)
End Sub

Private Sub CriteriaThreeDelete(ByRef adrsRec As ADODB.Recordset, _
ByVal strTmpTrn As String, ByVal strTmpShf As String)
On Error GoTo ERR_P
Select Case strHf_Opt
    Case "FFTF"
        If strShiftDel = pVStar.WosCode Or strShiftDel = pVStar.HlsCode Then
            ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
            "'" & StuffVal(adrsRec!presabs, 1, 2, strShiftDel) & "'" & " where Empcode=" & _
            "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
            DateCompStr(adrsRec!Date) & strDTEnc
        Else
            If adrsRec("EntReq") = 0 Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
                "'" & StuffVal(adrsRec!presabs, 1, 2, pVStar.PrsCode) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
                Exit Sub
            End If
            Sleep (500)
            If adrsRec!arrtim > 0 And adrsRec!arrtim < sngHDend Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
                "'" & StuffVal(adrsRec!presabs, 1, 2, pVStar.PrsCode) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            ElseIf (adrsRec!arrtim > 0 And adrsRec!arrtim >= sngHDend) Or (adrsRec!arrtim = 0) Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
                "'" & StuffVal(adrsRec!presabs, 1, 2, pVStar.AbsCode) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        End If
    Case "FSTS"
        If strShiftDel = pVStar.WosCode Or strShiftDel = pVStar.HlsCode Then
            ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
            "'" & StuffVal(adrsRec!presabs, 3, 2, strShiftDel) & "'" & " where Empcode=" & _
            "'" & cboCode.Text & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
            DateCompStr(adrsRec!Date) & strDTEnc
        Else
            If adrsRec("EntReq") = 0 Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
                "'" & StuffVal(adrsRec!presabs, 3, 2, pVStar.PrsCode) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
                Exit Sub
            End If
            If adrsRec!arrtim > 0 And adrsRec!arrtim < sngHDend Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
                "'" & StuffVal(adrsRec!presabs, 3, 2, pVStar.AbsCode) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            ElseIf (adrsRec!arrtim > 0 And adrsRec!arrtim >= sngHDend) Or (adrsRec!arrtim = 0) Then
                ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
                "'" & StuffVal(adrsRec!presabs, 3, 2, pVStar.AbsCode) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        End If
    Case "F T "
        If strShiftDel = pVStar.WosCode Or strShiftDel = pVStar.HlsCode Then
            ConMain.Execute "update " & strTmpTrn & " set Presabs=" & _
            "'" & ReplicateVal(strShiftDel, 2) & "'" & " where Empcode=" & "'" & cboCode.Text & _
            "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        Else
            If adrsRec("EntReq") = 0 Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(pVStar.PrsCode, 2) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
                Exit Sub
            End If
            If adrsRec!arrtim > 0 And adrsRec!arrtim < sngHDend Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(pVStar.PrsCode, 2) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            ElseIf (adrsRec!arrtim > 0 And adrsRec!arrtim > sngHDend) Or (adrsRec!arrtim = 0) Then
                ConMain.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(pVStar.AbsCode, 2) & "'" & _
                " where Empcode=" & "'" & cboCode.Text & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        End If
End Select
Exit Sub
ERR_P:
    ShowError ("CriteriaThreeDelete :: " & Me.Caption)
End Sub

Private Sub CriteriaFourDelete(ByRef adrsRec As ADODB.Recordset, _
ByVal strTmpTrn As String, ByVal strTmpShf As String)
On Error GoTo ERR_P
If strShiftDel = pVStar.WosCode Or strShiftDel = pVStar.HlsCode Then
    ConMain.Execute "update " & strTmpTrn & " set Presabs=" & "'" & _
    ReplicateVal(strShiftDel, 2) & "'" & " where Empcode=" & "'" & cboCode.Text & "'" & " and " & _
    strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
Else
    If adrsRec("EntReq") = 0 Then
        ConMain.Execute "update " & strTmpTrn & " set Presabs=" & "'" & _
        ReplicateVal(pVStar.PrsCode, 2) & "'" & " where Empcode=" & "'" & cboCode.Text & _
        "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        Exit Sub
    End If
    If (adrsRec!arrtim > 0 Or IsNull(adrsRec!arrtim)) And (adrsRec!arrtim < sngHDend Or IsNull(adrsRec!arrtim)) Then
        ConMain.Execute "update " & strTmpTrn & " set Presabs=" & "'" & _
        ReplicateVal(pVStar.PrsCode, 2) & "'" & " where Empcode=" & "'" & cboCode.Text & _
        "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
    ElseIf ((adrsRec!arrtim > 0 Or IsNull(adrsRec!arrtim)) And (adrsRec!arrtim > sngHDend Or IsNull(adrsRec!arrtim))) Or _
    (adrsRec!arrtim = 0 Or IsNull(adrsRec!arrtim)) Then
        ConMain.Execute "update " & strTmpTrn & " set Presabs=" & "'" _
        & ReplicateVal(pVStar.AbsCode, 2) & "'" & " where Empcode=" & "'" & cboCode.Text & _
        "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
    End If
End If
Exit Sub
ERR_P:
    ShowError ("CriteriaFourDelete :: " & Me.Caption)
End Sub

Private Sub cboCode_Change()
Call cboCode_Click
FrColour.Visible = False
End Sub

Private Sub cboCode_Click()
On Error GoTo ERR_P
If cboCode.ListIndex < 0 Then Exit Sub
bytMode = 1
'' Displays the Employee Name
lblName.Caption = cboCode.List(cboCode.ListIndex, 1)
''For Mauritius 14-08-2003
lblJoin.Caption = "Join Date   " & GetJoinDate
''
'' Gets the Employee Category
Call GetCat
Call GetCOLimitDetails
If bytMode <> 4 Then
    '' Get All the Leaves the Employee has Availed that Year
    bytMode = 1
    Call FillGrid
End If
If bytMode <> 5 Then
    '' Fill the Inner Grid With the Leave Balances of that Employee
    bytMode = 1
    Call FillGridBalance
End If
If bytMode = 4 Or bytMode = 5 Then
    '' if Invalid or Error in the Previous Processes
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

Private Function GetJoinDate() As String
On Error GoTo ERR_P
If adrsEmp.State = 0 Then Exit Function
If adrsEmp.EOF Then Exit Function
adrsEmp.MoveFirst
adrsEmp.Find "Empcode = '" & cboCode.List(cboCode.ListIndex, 0) & "'"
If adrsEmp.EOF Then Exit Function
If Not IsNull(adrsEmp("JoinDate")) Then
    GetJoinDate = DateDisp(adrsEmp("Joindate"))
End If
Exit Function
ERR_P:
    ShowError ("GetJoinDate :: " & Me.Caption)
    GetJoinDate = ""
End Function

Private Sub cboFrom_Click()
    Call CalculateDays
End Sub

Private Sub cboFrom_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then SendKeys Chr(9)
End Sub

Private Sub cboLeave_Change()
Call cboLeave_Click
End Sub
''
Private Sub cboTo_Click()
    Call CalculateDays
End Sub

Private Sub cboLeave_Click()
On Error GoTo ERR_P
If bytMode <> 3 Then Exit Sub       '' If Not Add Mode then Exit
Call ToggleType
Call CalculateDays
'' IF CO Take necessary Actions
Call InputCOEntryDate
Exit Sub
ERR_P:
    ShowError ("Leave :: " & Me.Caption)
End Sub

Private Sub cboRW_Click()
If bytMode <> 3 Then Exit Sub       '' If not Add Mode then Exit
If cboRW.Text <> "" Then
    If cboRW.Text = "Running" Then
        strRW = "R"
    Else
        strRW = "W"
    End If
End If
Call CalculateDays
End Sub

Private Sub cboTo_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then SendKeys Chr(9)
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
        Else
            bytMode = 3
            Call ChangeMode
        End If
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
AutoLeaveDelete:
            Call UpdateDeleteBalance    '' Update the Balance
            '' Delete the Leave From LvInfo

            If (MSF1.TextMatrix(MSF1.row, 4) = "6" Or MSF1.TextMatrix(MSF1.row, 4) = "7") Then
                ConMain.Execute "Delete from " & "LvInfo" & _
                Right(pVStar.YearSel, 2) & " Where Empcode=" & "'" & _
                cboCode.Text & "'" & " and " & "LCode=" & "'" & _
                cboLeave.Text & "'" & " and Trcd = " & MSF1.TextMatrix(MSF1.row, 4) & " " & " and fromdate=" & _
                strDTEnc & DateCompStr(txtFrom.Text) & strDTEnc & " "
            Else
                ConMain.Execute "Delete from " & "LvInfo" & _
                Right(pVStar.YearSel, 2) & " Where Empcode=" & "'" & _
                cboCode.Text & "'" & " and " & "LCode=" & "'" & _
                cboLeave.Text & "'" & " and Trcd IN (4,6,7) " & " and fromdate=" & _
                strDTEnc & DateCompStr(txtFrom.Text) & strDTEnc & " and Hf_Option='" & _
                MSF1.TextMatrix(MSF1.row, 4) & "'"
            End If
  
            Call AddActivityLog(lgDelete_Action, 2, 21)     '' Delete Log
            Call AuditInfo("DELETE", Me.Caption, "Delete Leave Avail Entry " & cboLeave.Text & " For Employee " & cboCode.Text)
            Call UpdateStatusOnDelete       '' Update Status on Delete
            Call FillGridBalance            '' Fill the Balances Grid
            Call FillGrid                   '' Fill the Outer Grid
            If MSF1.Rows > 1 Then
                TB1.TabEnabled(1) = True
            Else
                TB1.TabEnabled(1) = False
            End If
        End If
        TB1.Tab = 0
    Case 3       '' Add Mode
        bytMode = 2
        Call ChangeMode
        If MSF1.Rows < 2 Then TB1.TabEnabled(1) = False
End Select
Exit Sub
ERR_P:
    ShowError ("DeleteCancel :: " & Me.Caption)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)        '' Sets the Form Icon.
Call SetToolTipText(Me)     '' Set the ToolTipText
Call RetCaptions            '' Sets the Captions for the Controls
Call GetRights              '' Get the Rights
Call FillCombo              '' Fills the Emplyee Combo
Call FillFromToCombo        '' Fills the FromToCombo
Call FillTypeCombo          '' Fills the Type Combo
''14-08-2003 For Mauritius
Call OpenEmpForJoinDate     '' Opens Recordset for join date
''
TB1.TabEnabled(1) = False   '' Disable the Tab 1
bytMode = 2                 '' Set the Mode Back to Normal or View Mode
Call ChangeMode             '' Take Action According to the Mode
End Sub

Private Sub MSF1_DblClick()
If TB1.TabEnabled(1) = False Then Exit Sub
TB1.Tab = 1
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions Where ID Like '07%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("07001", adrsC)
Call SetButtonCap                           '' Button Captions
Call SetOutGridCap                          '' Msf1 Captions
Call SetInGridCap                           '' MSF2 Captions
Call SetOtherCaps                           '' Other Control Captions
TB1.TabCaption(0) = NewCaptionTxt("00013", adrsMod)       '' List
TB1.TabCaption(1) = NewCaptionTxt("00014", adrsMod)       '' Details
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

Private Sub SetOutGridCap()
With MSF1
    .TextMatrix(0, 0) = NewCaptionTxt("07009", adrsC)   '' Leave Code
    .TextMatrix(0, 1) = NewCaptionTxt("07010", adrsC)   '' Leave From
    .TextMatrix(0, 2) = NewCaptionTxt("07011", adrsC)   '' Leave To
    .TextMatrix(0, 3) = NewCaptionTxt("07012", adrsC)   '' Leave Days
    .ColWidth(4) = 0                        '' Set the Col Widh of the HF to 0
End With
End Sub

Private Sub SetInGridCap()
With MSF2
    .TextMatrix(0, 0) = NewCaptionTxt("07013", adrsC)       '' Code
    .TextMatrix(0, 1) = NewCaptionTxt("07014", adrsC)       '' Name
    .TextMatrix(0, 2) = NewCaptionTxt("07015", adrsC)       '' Balance
End With
End Sub

Private Sub SetOtherCaps()
lblCode.Caption = NewCaptionTxt("07002", adrsC)         '' Employee Code
lblNameCap.Caption = NewCaptionTxt("07003", adrsC)      '' Name
lblInfo.Caption = NewCaptionTxt("07004", adrsC)         '' Leave Information
lblLeave.Caption = NewCaptionTxt("07005", adrsC)        '' Leave Code
lblFrom.Caption = NewCaptionTxt("00010", adrsMod)         '' From
lblTo.Caption = NewCaptionTxt("00011", adrsMod)           '' To
lblCOEntry.Caption = NewCaptionTxt("07038", adrsC)      '' CO for extra....
lblRW.Caption = NewCaptionTxt("07006", adrsC)           '' Leave Type
lblDays.Caption = NewCaptionTxt("07007", adrsC)         '' No of Days
lblBal.Caption = NewCaptionTxt("07008", adrsC)          '' Balance
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

Private Sub FillFromToCombo()       '' Fills the From & To ComboBoxes
cboFrom.AddItem "Full Day"
cboFrom.AddItem "First Half"
cboFrom.AddItem "Second Half"
cboTo.AddItem "Full Day"
cboTo.AddItem "First Half"
cboTo.AddItem "Second Half"
End Sub

Private Sub GetCat()                '' Gets the Category of a Particular Employee
On Error GoTo ERR_P
If cboCode.Text = "" Then
    strCatAvail = ""
    Exit Sub
End If
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open "Select Cat,JoinDate,COCode from EmpMst where EmpCode='" & cboCode.Text & "'" _
, ConMain
If (adrsPaid.EOF And adrsPaid.BOF) Then
    MsgBox NewCaptionTxt("07016", adrsC), vbExclamation
    cboCode.Value = ""
    bytMode = 4
    TB1.Tab = 0
    bytCOCode = 100
    Exit Sub
Else
    strCatAvail = adrsPaid("Cat")
    If Not IsNull(adrsPaid("JoinDate")) Then
        dtJoin = DateCompDate(adrsPaid("JoinDate"))
    Else
        dtJoin = DateCompDate("31-December-2100")
    End If
    bytCOCode = IIf(IsNull(adrsPaid("COCode")), 100, adrsPaid("COCode"))
End If
Exit Sub
ERR_P:
    bytMode = 4
    ShowError ("Getcat :: " & Me.Caption)
End Sub

Private Sub GetCOLimitDetails()
On Error GoTo ERR_P
'' Check if any Code exists
If bytCOCode = 100 Then
    bytCOLimit = 0
    Exit Sub
End If
'' Check if Given COCOde Exists in CORul Master
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open "Select COAvail from CORul Where COCOde=" & bytCOCode, ConMain, adOpenStatic, adLockReadOnly
If adrsPaid.EOF Then
    bytCOCode = 100
    bytCOLimit = 0
    Exit Sub
End If
'' Get COLimit
bytCOLimit = IIf(IsNull(adrsPaid("COAvail")), 0, adrsPaid("COAvail"))
Exit Sub
ERR_P:
    ShowError ("GetCOLimitDetails::" & Me.Caption)
End Sub

Private Sub OpenEmpForJoinDate()
On Error GoTo ERR_P
If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "Select Empcode,joindate from empmst order by Empcode", ConMain, adOpenStatic, adLockOptimistic
Exit Sub
ERR_P:
    ShowError ("OpenEmpForJoinDate :" & Me.Caption)
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
If bytMode = 3 Then
    Exit Sub            '' If Add Mode then Exit Sub
End If
MSF1.Col = 0                            '' Set the Column to 0
If MSF1.Text = NewCaptionTxt("07009", adrsC) Then Exit Sub
Call Display
End Sub

Private Sub FillGrid()      '' Fills the Grid with the Leaves the Employee Has Availed
On Error GoTo ERR_P
Dim bytCnt As Byte
If adrsDept1.State = 1 Then adrsDept1.Close
    adrsDept1.Open "Select LCode,FromDate,ToDate,Days,Hf_Option,trcd from lvinfo" & _
    Right(pVStar.YearSel, 2) & " where trcd IN(4,6,7) " & " and Empcode=" & "'" & cboCode.Text & "'" & _
    " Order by  LCode,Fromdate", ConMain, adOpenStatic
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
            If IsNull(adrsDept1("HF_Option")) Then     ' 23-05-09
                .row = bytCnt
                .Col = 3
                If adrsDept1.Fields!trcd = 6 Then
                    .CellForeColor = &HFF&
                ElseIf adrsDept1.Fields!trcd = 7 Then
                    .CellForeColor = &HFF00&
                End If
                .CellFontBold = True
                .TextMatrix(bytCnt, 4) = adrsDept1("trcd")
                FrColour.Visible = True
            Else
            .TextMatrix(bytCnt, 4) = adrsDept1("HF_Option")
            End If
        End With
        adrsDept1.MoveNext
    Next
End If
Exit Sub
ERR_P:
    bytMode = 5
    ShowError ("FillGrid : Outer ::" & Me.Caption)
End Sub

Private Sub GetRights()             '' Gets the Rights of for a Particular User
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(2, 4, 5)
If Mid(strTmp, 1, 1) = "1" Then AddRights = True
If Mid(strTmp, 2, 1) = "1" Then DeleteRights = True
Exit Sub
ERR_P:
    ShowError ("GetRights :: " & Me.Caption)
    AddRights = False
    DeleteRights = False
End Sub

Private Sub FillGridBalance()                   '' Gets the Leaves and Balances of a
On Error GoTo ERR_P                             '' Particular Employee
Dim strLeaveList() As String                    '' For Leave Array
Dim bytTmp As Byte, intTmp As Integer           '' Temporary Variables
'' Fill LeaveCombo
cboLeave.clear
If adrsDept1.State = 1 Then adrsDept1.Close
If SubLeaveFlag = 1 Then ' 15-10
    If FieldExists("LvBaL" & Right(pVStar.YearSel, 2), "EL") Then ELLeave = "EL"   ' 07-11
    If ELLeave = "EL" Then
        If FieldExists("LvBaL" & Right(pVStar.YearSel, 2), "EN") Then ELSubLeave = ",EN"
        If FieldExists("LvBaL" & Right(pVStar.YearSel, 2), "NE") Then ELSubLeave = ELSubLeave & ",NE"
        ELSubLeave = Right(ELSubLeave, Len(ELSubLeave) - 1)
    End If
    adrsDept1.Open "Select LvCode,Leave from LeavDesc where LvCode Not in('" & pVStar.AbsCode & _
    "','" & pVStar.PrsCode & "','" & pVStar.HlsCode & "','" & pVStar.WosCode & "','SL','EL') and Cat='" & _
    strCatAvail & "'"
Else
    adrsDept1.Open "Select LvCode,Leave from LeavDesc where LvCode Not in('" & pVStar.AbsCode & _
    "','" & pVStar.PrsCode & "','" & pVStar.HlsCode & "','" & pVStar.WosCode & "') and Cat='" & _
    strCatAvail & "'"
End If
If adrsDept1.RecordCount <= 0 Then Exit Sub
ReDim strLeaveList(adrsDept1.RecordCount - 1, 1): bytTmp = 0
Do While Not adrsDept1.EOF
    strLeaveList(bytTmp, 0) = adrsDept1("LvCode")
    strLeaveList(bytTmp, 1) = adrsDept1("Leave")
    bytTmp = bytTmp + 1
    adrsDept1.MoveNext
Loop
If adrsDept1.RecordCount > 0 Then cboLeave.List = strLeaveList
'' Fill Grid
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select * from LvBal" & Right(pVStar.YearSel, 2) & " Where Empcode='" & _
cboCode.Text & "'", ConMain, adOpenStatic
MSF2.Rows = 1
If (adrsDept1.EOF And adrsDept1.BOF) Then Exit Sub
For intTmp = 0 To adrsDept1.Fields.Count - 1
    If UCase(adrsDept1(intTmp).name) <> "EMPCODE" Then  '' if Field is not Employee Code
        If adrsRits.State = 1 Then adrsRits.Close
        adrsRits.Open "Select Leave from LeavDesc where LvCode='" & _
        adrsDept1(intTmp).name & "' and Cat='" & strCatAvail & "'", ConMain
        If Not (adrsRits.EOF And adrsRits.BOF) Then         '' if Leave Desc is Found
            MSF2.Rows = MSF2.Rows + 1
            MSF2.TextMatrix(MSF2.Rows - 1, 0) = adrsDept1(intTmp).name
            MSF2.TextMatrix(MSF2.Rows - 1, 1) = adrsRits("Leave").Value
            MSF2.TextMatrix(MSF2.Rows - 1, 2) = IIf(IsNull(adrsDept1(intTmp).Value), "0.00", _
            Format(adrsDept1(intTmp).Value, "0.00"))
        End If
    End If
Next
Exit Sub
ERR_P:
    bytMode = 4
    ShowError ("FillGridBalance :: " & Me.Caption)
    Resume Next
End Sub

Private Sub txtDays_GotFocus()
   Call GF(txtDays)
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtDays)
End If
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
If Not ValidDate(txtFrom) Then
    txtFrom.SetFocus
    Cancel = True
Else
    Call CalculateDays
    Cancel = False
End If
End Sub

Private Sub txtCOEntry_Click()
varCalDt = ""
If cboCode.Visible = True Then varCalDt = cboCOEntry.Value
Call ShowCalendar
End Sub
'
'Private Sub txtCOEntry_Validate(Cancel As Boolean)
'If Not ValidDate(txtCOEntry) Then
'    txtCOEntry.SetFocus
'    Cancel = True
'Else
'    Call CalculateDays
'    Cancel = False
'End If
'End Sub

Private Sub txtTo_Click()
varCalDt = ""
varCalDt = Trim(txtTo.Text)
txtTo.Text = ""
Call ShowCalendar
End Sub

Private Sub txtTo_GotFocus()
    Call GF(txtTo)
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    Call CDK(txtTo, KeyAscii)
End Sub

Private Sub txtTo_Validate(Cancel As Boolean)
If Not ValidDate(txtTo) Then
    txtTo.SetFocus
    Cancel = True
Else
    Call CalculateDays
End If
End Sub

Private Sub CalculateDays()         '' Calculates the Number of the Days the Leaves to be
On Error GoTo ERR_P                 '' Availed
If Trim(txtFrom.Text) = "" Then Exit Sub
If Trim(txtTo.Text) = "" Then Exit Sub
If Trim(cboFrom.Text) = "" Then Exit Sub
If Trim(cboTo.Text) = "" Then Exit Sub
''Start Code for Calculation strHf_Opt based on Status
Select Case cboFrom.Text
        Case "Full Day"
                strF = "F "
        Case "First Half"
                strF = "FF"
        Case "Second Half"
                strF = "FS"
End Select
Select Case cboTo.Text
        Case "Full Day"
                strT = "T "
        Case "First Half"
                strT = "TF"
        Case "Second Half"
                strT = "TS"
End Select
strHf_Opt = strF & strT
''End Code for Calculation strHf_Opt based on Status

''Start Code for Calculation of the days based on Status
If DateCompDate(txtFrom.Text) = DateCompDate(txtTo.Text) Then
        Select Case strHf_Opt
                Case "F T "                             'XX
                        sngDays = 1
                        strHf_Opt = "F T "
                Case "FFTF"                             'YY
                        sngDays = 0.5
                Case "FFT "                             'YX
                        sngDays = 1
                        strHf_Opt = "F T "
                Case "F TF"                             'XY
                        sngDays = 1
                        strHf_Opt = "F T "
                Case "FST "                             'ZX
                        sngDays = 1
                        strHf_Opt = "F T "
                Case "FSTS"                             'ZZ
                        sngDays = 0.5
                Case "F TS"                             'XZ
                        sngDays = 1
                        strHf_Opt = "F T "
                Case "FSTF"                             'ZY
                        sngDays = 1
                        strHf_Opt = "F T "
                Case "FFTS"                             'YZ
                        sngDays = 1
                        strHf_Opt = "F T "
        End Select
Else
                sngDiff = DateDiff("d", DateCompDate(txtFrom.Text), DateCompDate(txtTo.Text))
                If cboFrom.Text = cboTo.Text Then ''Same
                        If cboFrom.Text = "Full Day" Then
                                sngDays = sngDiff + 1
                        Else
                                sngDays = sngDiff + 0.5
                        End If
                Else                                                                         ''Different
                        ''X Z
                        If cboFrom.Text = "Full Day" And cboTo.Text = "Second Half" Then _
                        sngDays = sngDiff + 1
                        ''Y X
                        If cboFrom.Text = "First Half" And cboTo.Text = "Full Day" Then _
                        sngDays = sngDiff + 1
                        ''Y Z
                        If cboFrom.Text = "First Half" And cboTo.Text = "Second Half" Then _
                        sngDays = sngDiff + 1
                        ''Z Y
                        If cboFrom.Text = "Second Half" And cboTo.Text = "First Half" Then _
                        sngDays = sngDiff
                        ''X Y
                        If cboFrom.Text = "Full Day" And cboTo.Text = "First Half" Then _
                        sngDays = sngDiff + 0.5
                        ''Z X
                         If cboFrom.Text = "Second Half" And cboTo.Text = "Full Day" Then _
                        sngDays = sngDiff + 0.5
                End If
End If
''End Code for Calculation of the days based on Status
If SubLeaveFlag = 1 And (cboLeave.Text = "CM") Then ' 15-10
    txtDays.Text = Format(Abs(sngDays) * 2, "0.00")
Else
    txtDays.Text = Format(Abs(sngDays), "0.00")
End If
'' Start Code for Calulation of days based on the Type of Leave(R/W)
If Trim(cboLeave.Text) = "" Then Exit Sub
Call TypeLeave
'' End Code for Calulation of days based on the Type of Leave(R/W)
Exit Sub
ERR_P:
    ShowError ("CalculateDays :: " & Me.Caption)
    txtDays.Text = "0.00"
End Sub

Private Sub TypeLeave()             '' Adjusts the Number of Leaves Depending on the
On Error GoTo ERR_P                 '' Leave Type

Dim strShf As String, adrsDD As New ADODB.Recordset

''
If strRW = "W" Then
        Dim bytCtr As Byte, dttmp As Date
        bytCtr = 0
        dttmp = DateCompDate(txtFrom.Text)
        Do While dttmp <= DateCompDate(txtTo.Text)
                Dim strDay As String
                Select Case WeekDay(dttmp)
                        Case 1
                                strDay = "SU"
                        Case 2
                                strDay = "MO"
                        Case 3
                                strDay = "TU"
                        Case 4
                                strDay = "WE"
                        Case 5
                                strDay = "TH"
                        Case 6
                                strDay = "FR"
                        Case 7
                                strDay = "SA"
                End Select
                
                strShf = GetshiftFile(dttmp)
                If FindTable(strShf) Then
                    If adrsDD.State = 1 Then adrsDD.Close
                    adrsDD.Open "Select d" & Day(dttmp) & " from " & strShf & " where empcode = '" & cboCode.Text & "'", ConMain, adOpenStatic, adLockReadOnly
                    If Not (adrsDD.EOF And adrsDD.BOF) Then
                        If adrsDD(0) = pVStar.HlsCode Or adrsDD(0) = pVStar.WosCode Then
                            bytCtr = bytCtr + 1
                        End If
                        GoTo EndOfLoop
                    End If
                End If
                
                ''
                If adrsRits.State = 1 Then adrsRits.Close
                adrsRits.Open "select " & strKOff & " , off2,wo_1_3,wo_2_4 from empmst where Empcode='" & cboCode.Text & "'" _
                , ConMain
                If UCase(adrsRits(0)) = strDay Then
                        bytCtr = bytCtr + 1
                        GoTo EndOfLoop
                End If
                If UCase(adrsRits("off2")) = strDay Then
                        bytCtr = bytCtr + 1
                        GoTo EndOfLoop
                End If
                If UCase(adrsRits("wo_1_3")) = strDay Then
                        bytCtr = bytCtr + 1
                        GoTo EndOfLoop
                End If
                If UCase(adrsRits("wo_2_4")) = strDay Then
                        bytCtr = bytCtr + 1
                        GoTo EndOfLoop
                End If
                If adrsRits.State = 1 Then adrsRits.Close
                adrsRits.Open "Select * from holiday where " & strKDate & " =" & strDTEnc & _
                DateCompStr(dttmp) & strDTEnc & " and cat='" & strCatAvail & "'" _
                , ConMain
                If Not (adrsRits.EOF And adrsRits.BOF) Then bytCtr = bytCtr + 1
EndOfLoop:
                dttmp = dttmp + 1
        Loop
        Dim sngDaysTmp As Single
        sngDaysTmp = txtDays.Text
        sngDaysTmp = sngDaysTmp - bytCtr
        txtDays.Text = Format(Abs(sngDaysTmp), "0.00")
End If
Exit Sub
ERR_P:
    ShowError ("TypeLeave ::" & Me.Caption)
End Sub

Private Sub InputCOEntryDate()
On Error GoTo ERR_P
'' Variable to check if CO Availment operations are to be carried out or not
If UCase(cboLeave.Text) = LV_CO And bytCOLimit > 0 Then
    blnCOChecks = True
Else
    blnCOChecks = False
End If
If blnCOChecks Then
    lblCOEntry.Visible = True
    cboCOEntry.Visible = True
    Call FillCboCo
    cboCOEntry.ListIndex = -1
    'txtCOEntry.Visible = True
Else
    lblCOEntry.Visible = False
    cboCOEntry.Visible = False
End If
Exit Sub
ERR_P:
    ShowError ("InputCOEntryDate::" & Me.Caption)
End Sub

Public Sub ToggleType()         '' Checks Type of Leave & Adjusts the Type ComboBox
On Error GoTo ERR_P
If Trim(cboCode.Text) = "" Then Exit Sub
If Trim(cboLeave.Text) = "" Then Exit Sub
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select run_wrk from leavdesc where cat='" & strCatAvail & "' and lvcode='" & _
cboLeave.Text & "'", ConMain
If Not (adrsDept1.EOF And adrsDept1.BOF) Then
    If adrsDept1(0) = "O" Then
        cboRW.Enabled = True
        cboRW.ListIndex = 0
        strRW = "R"
    Else
        cboRW.Enabled = False
        strRW = adrsDept1(0)
    End If
Else
    cboLeave.RemoveItem cboLeave.ListIndex
End If
If adrsDept1.State = 1 Then adrsDept1.Close
Exit Sub
ERR_P:
    ShowError ("ToggleType")
End Sub

Private Sub FillTypeCombo()     '' Fills Leave Type ComboBox
cboRW.AddItem "Running"
cboRW.AddItem "Working"
End Sub

Private Function ValidateAddmaster() As Boolean     '' Validate Details befor Availing the
On Error GoTo ERR_P                                 '' Leave
ValidateAddmaster = True
'' Check if Any Leave is Selected or Not
If cboLeave.Text = "" Then
    MsgBox NewCaptionTxt("07017", adrsC), vbExclamation
    cboLeave.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
'' Check for Invalid Number of AvailLeave Days
If Val(txtDays.Text) <= 0 Then
    MsgBox NewCaptionTxt("07018", adrsC), vbExclamation
    ValidateAddmaster = False
    Exit Function
End If
'' Check for Invalid date
If Not ValidLeaveDate Then
    ValidateAddmaster = False
    Exit Function
End If
If blnCOChecks Then
    If Not COChecks Then
        ValidateAddmaster = False
        Exit Function
    End If
End If
'' Get the Details of the Specified Leave from the LeaveDesc Table
Call GetLeaveDetails
If (Not blnUnPaid) Or (Not blnNoBal) Then
    '' Check if Any Number of Leaves are Allowed or Not.
    
    'If Not NumOfLeaves Then
    '    ValidateAddmaster = False
    '    Exit Function
    'End If
    '' Check if there is Enough Balance
    If cboLeave <> "OD" Then
        If Not NumOfBalance Then
            ValidateAddmaster = False
            Exit Function
        End If
    End If
End If
'' Check if Leaves are Already Availed for the Specified Dates
If Not ALreadyAvailedDate Then
    ValidateAddmaster = False
    Exit Function
End If
'' Check if he has Availed Leaves for More than Allowed Times
If Not NumOfTimesAvailed Then
    ValidateAddmaster = False
    Exit Function
End If
'' Check Minimum & Maximum Limits
If Not CheckMinMaxLeave Then
    ValidateAddmaster = False
    Exit Function
End If
'' Check if the Person is Immidiate Absent before or or After the Leave Dates
If Not ImmediateAbsent Then
    ValidateAddmaster = False
    Exit Function
End If
If Not ImmediateLeave Then
    ValidateAddmaster = False
    Exit Function
End If

Exit Function
ERR_P:
    ShowError ("ValidateAddMaster :: " & Me.Caption)
    ValidateAddmaster = False
End Function

Private Function COChecks() As Boolean
On Error GoTo ERR_P
'' CO Validations
' If Dates Entered regarding CO are Correct or not
If Not CODateVals Then Exit Function
'' If COEntry date is really for a CO Credited on that Date
If Not COCredited Then Exit Function
'' IF Date Limit Is Already Availed for that Date
If Not CODateLimit Then Exit Function
'' If Balance is already Over for that entry date
If Not COBalanceOver Then Exit Function
COChecks = True
Exit Function
ERR_P:
    ShowError ("COChecks::" & Me.Caption)
End Function

Private Function COBalanceOver() As Boolean
On Error GoTo ERR_P
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open "Select sum(days) from LVINFO" & Right(pVStar.YearSel, 2) & " Where " & _
"Empcode='" & cboCode.Text & "' and LCode='" & LV_CO & "' and Trcd=4 and " & _
"EntryDate=" & strDTEnc & cboCOEntry.Value & strDTEnc
If adrsPaid.EOF Then
    sngCOAvail = 0
Else
    sngCOAvail = IIf(IsNull(adrsPaid(0)), 0, adrsPaid(0))
End If
If Val(txtDays.Text) > (sngCOBal - sngCOAvail) Then
    MsgBox NewCaptionTxt("07043", adrsC), vbExclamation
'    txtCOEntry.SetFocus
    Exit Function
End If
COBalanceOver = True
Exit Function
ERR_P:
    ShowError ("COBalanceOver::" & Me.Caption)
End Function

Private Function CODateLimit() As Boolean
On Error GoTo ERR_P
Dim dttmp As Date, bytTmp As Byte
dttmp = cboCOEntry.Value
Select Case bytCOLimit
    Case 1
        dttmp = DateCompDate(FdtLdt(Month(dttmp), CStr(Year(dttmp)), "L"))
    Case Else
        bytTmp = bytCOLimit - 1
     
        bytTmp = bytTmp * 15
    
        dttmp = DateAdd("D", bytTmp, dttmp)
End Select
If dttmp < DateCompDate(txtFrom.Text) Then
    If MsgBox(NewCaptionTxt("07042", adrsC), vbYesNo + vbQuestion) = vbNo Then
        cboCOEntry.SetFocus
        Exit Function
    End If
 
End If
CODateLimit = True
Exit Function
ERR_P:
    ShowError ("CODateLimit::" & Me.Caption)
End Function

Private Function COCredited() As Boolean
On Error GoTo ERR_P
Dim strTmp As String
If Val(pVStar.Yearstart) > cboCOEntry.Value Then
    strTmp = "LVINFO" & Right(CStr(Year(DateCompDate(cboCOEntry.Value)) - 1), 2)
Else
    strTmp = "LVINFO" & Right(CStr(Year(DateCompDate(cboCOEntry.Value))), 2)
End If
If adrsPaid.State = 1 Then adrsPaid.Close

adrsPaid.Open "Select days from " & strTmp & " Where EntryDate=" & strDTEnc & _
DateCompStr(cboCOEntry.Value) & strDTEnc & " and Empcode='" & cboCode.Text & "'" & _
" and LCode='" & LV_CO & "' and Trcd=2", ConMain, adOpenStatic, _
adLockReadOnly

If adrsPaid.EOF Then
    sngCOBal = 0
Else
    sngCOBal = IIf(IsNull(adrsPaid("Days")), 0, adrsPaid("Days"))
End If
If sngCOBal <= 0 Then
    MsgBox NewCaptionTxt("07041", adrsC), vbExclamation
    cboCOEntry.SetFocus
    Exit Function
End If
COCredited = True
Exit Function
ERR_P:
    ShowError ("COCredited::" & Me.Caption)
End Function

Private Function CODateVals() As Boolean
On Error GoTo ERR_P
If Trim(cboCOEntry.Text) = "" Then
    MsgBox NewCaptionTxt("07039", adrsC), vbExclamation
    'txtCOEntry.SetFocus
    cboCOEntry.SetFocus
    Exit Function
End If
If Trim(txtFrom.Text) <> Trim(txtTo.Text) Then
    MsgBox NewCaptionTxt("07040", adrsC), vbExclamation
    txtTo.SetFocus
    Exit Function
End If
If Trim(txtFrom.Text) = Trim(cboCOEntry.Value) Then
    MsgBox NewCaptionTxt("07044", adrsC), vbExclamation
    cboCOEntry.SetFocus
    Exit Function
End If
CODateVals = True
Exit Function
ERR_P:
    ShowError ("CODateVals::" & Me.Caption)
End Function

Private Sub ChangeMode()        '' Action to be Taken when Mode Changes
Select Case bytMode
    Case 2  '' View Mode
        Call ViewAction
    Case 3  '' Add Mode
        Call AddAction
End Select
End Sub

Private Sub AddAction()     '' Procedure for Addition Mode
'' Set the Tab Accordingly
TB1.Tab = 1
TB1.TabEnabled(1) = True
'' Enable Necessary Controls
cboLeave.Enabled = True     '' Leave ComboBox
lblCOEntry.Visible = False
txtFrom.Enabled = True      '' From date TextBox
txtTo.Enabled = True        '' To Date TextBox
cboFrom.Enabled = True      '' From ComboBox
cboTo.Enabled = True        '' To ComboBox
txtRemark.Enabled = True
txtRemark.Text = ""

''txtDays.Enabled = True
''
'' Clear Necessary Controls
cboLeave.Value = ""         '' Leave ComboBox
txtFrom.Text = ""           '' From date TextBox
txtTo.Text = ""             '' To Date TextBox
cboFrom.Value = "Full Day"  '' From ComboBox
cboTo.Value = "Full Day"    '' To ComboBox
cboRW.Value = ""            '' Type Combo
txtDays.Text = "0.00"       '' Leave Days
'' Give Caption to the Needed Controls
Call SetButtonCap(2)
cboLeave.SetFocus           '' Set Focus to the Leave ComboBox
End Sub

Private Sub ViewAction()        '' Action to be Taken when the Form is in View Mode
TB1.Tab = 0
'' Enable Necessary Controls
cboLeave.Enabled = False    '' Leave ComboBox
lblCOEntry.Visible = False
'txtCOEntry.Visible = False
txtFrom.Enabled = False     '' From date TextBox
txtTo.Enabled = False       '' To Date TextBox
cboFrom.Enabled = False     '' From ComboBox
cboTo.Enabled = False       '' To ComboBox
cboRW.Enabled = False       '' Type ComboBox
txtDays.Enabled = False
txtRemark.Enabled = False
'' Give Caption to the Needed Controls
Call SetButtonCap
End Sub

Private Function ValidLeaveDate() As Boolean        '' Validates the Leave Dates Specified
ValidLeaveDate = True
'' Check for EmptyDate
If Trim(txtFrom.Text) = "" Then
    MsgBox NewCaptionTxt("00016", adrsMod), vbExclamation
    txtFrom.SetFocus
    ValidLeaveDate = False
    Exit Function
End If
If txtTo.Text = "" Then
    MsgBox NewCaptionTxt("00017", adrsMod), vbExclamation
    txtTo.SetFocus
    ValidLeaveDate = False
    Exit Function
End If
If DateCompDate(txtFrom.Text) > DateCompDate(txtTo.Text) Then
    MsgBox NewCaptionTxt("00018", adrsMod), vbExclamation
    txtTo.SetFocus
    ValidLeaveDate = False
    Exit Function
End If
If DateDiff("m", Year_Start, DateCompDate(txtFrom.Text)) > 11 Or _
DateDiff("m", Year_Start, DateCompDate(txtFrom.Text)) < 0 Then
    MsgBox NewCaptionTxt("00019", adrsMod) & txtFrom.Text & NewCaptionTxt("00021", adrsMod), _
    vbExclamation
    txtFrom.SetFocus
    ValidLeaveDate = False
    Exit Function
End If
If DateDiff("m", Year_Start, DateCompDate(txtTo.Text)) > 11 Or _
DateDiff("m", Year_Start, DateCompDate(txtTo.Text)) < 0 Then
    MsgBox NewCaptionTxt("00020", adrsMod) & txtTo.Text & NewCaptionTxt("00021", adrsMod), _
    vbExclamation
    txtTo.SetFocus
    ValidLeaveDate = False
    Exit Function
End If

If DateCompDate(txtFrom.Text) < dtJoin Then
    MsgBox NewCaptionTxt("00112", adrsMod), vbExclamation
    txtFrom.SetFocus
    ValidLeaveDate = False
    Exit Function
End If
End Function
 
Private Function NumOfTimesAvailed() As Boolean     '' Checks if How Many times the Employee
On Error GoTo ERR_P                                 '' is Allowed to Avail the Leave
NumOfTimesAvailed = True
Dim bytCntTmp As Byte, bytTmp As Byte
bytTmp = 0
'' Number of Times Availed
For bytCntTmp = 1 To MSF1.Rows - 1
    If cboLeave.Text = MSF1.TextMatrix(bytCntTmp, 0) Then bytTmp = bytTmp + 1
Next
If intTimes > 0 Then
    If bytTmp >= intTimes Then
        If MsgBox(NewCaptionTxt("07019", adrsC) & intTimes & NewCaptionTxt("07020", adrsC), _
        vbQuestion + vbYesNo) = vbYes Then
            NumOfTimesAvailed = True
        Else
            txtFrom.SetFocus
            NumOfTimesAvailed = False
        End If
    End If
End If
Exit Function
ERR_P:
    ShowError ("NumOfTimesAvailed :: " & Me.Caption)
    NumOfTimesAvailed = False
End Function

Private Function CheckMinMaxLeave() As Boolean  '' Checks the Minimum & the Maximum
CheckMinMaxLeave = True                         '' Leaves the Employee is Allowed to Avail
If bytMax <> 0 Then
    If Val(txtDays.Text) > bytMax Then
        If MsgBox(NewCaptionTxt("07021", adrsC) & bytMax & NewCaptionTxt("07022", adrsC), _
        vbQuestion + vbYesNo) = vbYes Then
            CheckMinMaxLeave = True
            Exit Function
        Else
            txtFrom.SetFocus
            CheckMinMaxLeave = False
            Exit Function
        End If
    End If
End If
If bytMin <> 0 Then
    If Val(txtDays.Text) < bytMin Then
        If MsgBox(NewCaptionTxt("07023", adrsC) & bytMin & NewCaptionTxt("07022", adrsC), _
        vbQuestion + vbYesNo) = vbYes Then
            CheckMinMaxLeave = True
            Exit Function
        Else
            txtFrom.SetFocus
            CheckMinMaxLeave = False
            Exit Function
        End If
    End If
End If
End Function


Private Function NumOfBalance() As Boolean      '' Checks the Leave Balance of
On Error GoTo ERR_P                             '' the Employee
NumOfBalance = True
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
If sngDaysBal <= 0 Then
    If SubLeaveFlag = 1 And cboLeave.Text = "NE" And sngDaysBal = 0 Then   ' 07-11
        MsgBox "NE Leave Balance Cannot become Negative." & vbCrLf & "Avail Remaining days from EN Leave Balance.", vbExclamation
        cboLeave.SetFocus
        NumOfBalance = False
    Else
        If MsgBox(NewCaptionTxt("07026", adrsC), vbYesNo + vbQuestion) = vbYes Then
            NumOfBalance = True
            sngDaysBal = sngDaysBal - Val(txtDays.Text)
        Else
            txtFrom.SetFocus
            NumOfBalance = False
        End If
    End If
Else
        NumOfBalance = True
        sngDaysBal = sngDaysBal - Val(txtDays.Text)
End If
If SubLeaveFlag = 1 And cboLeave.Text = "NE" And sngDaysBal < 0 Then   ' 07-11
    MsgBox "NE Leave Balance Cannot become Negative." & vbCrLf & "Avail Remaining days from EN Leave Balance.", vbExclamation
    cboLeave.SetFocus
    NumOfBalance = False
End If
Exit Function
ERR_P:
    ShowError ("NumOfBalance :: " & Me.Caption)
    NumOfBalance = False
End Function


Private Function ALreadyAvailedDate() As Boolean    '' Checks if Leave is Already Availed
On Error GoTo ERR_P                                 '' by the Employee for the Same Dates
ALreadyAvailedDate = True
Dim strA_R As String, strHFOPT As String
Dim bytCtr As Byte
strA_R = "select * from LvInfo" & Right(pVStar.YearSel, 2) & " where ((" & strDTEnc & _
DateCompStr(txtFrom.Text) & strDTEnc & " between fromdate and todate ) or (" & _
strDTEnc & DateCompStr(txtTo.Text) & strDTEnc & _
" between fromdate and todate)) and trcd=4 and Empcode='" & cboCode.Text & "'"
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open strA_R, ConMain
bytCtr = 0
If Not (adrsPaid.EOF And adrsPaid.BOF) Then
        Do While Not adrsPaid.EOF
                If strF = "F " Then bytCtr = bytCtr + 1
                If strT = "T " Then bytCtr = bytCtr + 1
                strHFOPT = Left(adrsPaid("hf_option"), 2)
                If strHFOPT = strF Or strHFOPT = "F " Then bytCtr = bytCtr + 1
                strHFOPT = Right(adrsPaid("HF_Option"), 2)
                If strHFOPT = strT Or strHFOPT = "T " Then bytCtr = bytCtr + 1
                adrsPaid.MoveNext
        Loop
End If
If bytCtr > 0 Then
        MsgBox NewCaptionTxt("07027", adrsC), vbExclamation
        txtFrom.SetFocus
        ALreadyAvailedDate = False
        Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ALreadyAvailedDate :: " & Me.Caption)
    ALreadyAvailedDate = False
End Function

Private Sub GetLeaveDetails()       '' Gets the Other Primary Details of a Leave
On Error GoTo ERR_P                 '' from the Leave Master
If adrsPaid.State = 1 Then adrsPaid.Close
If SubLeaveFlag = 1 And (cboLeave.Text = "HP" Or cboLeave.Text = "CM") Then   ' 07-11
    adrsPaid.Open "Select AllowDays,MinAllowDays,No_OfTimes,Paid,Type from leavdesc where LvCode=" & _
    "'SL' and Cat='" & strCatAvail & "'", ConMain
ElseIf SubLeaveFlag = 1 And cboLeave.Text = "LT" Then
     adrsPaid.Open "Select AllowDays,MinAllowDays,No_OfTimes,Paid,Type from leavdesc where LvCode=" & _
    "'PL' and Cat='" & strCatAvail & "'", ConMain
Else
    adrsPaid.Open "Select AllowDays,MinAllowDays,No_OfTimes,Paid,Type from leavdesc where LvCode=" & _
    "'" & cboLeave.Text & "' and Cat='" & strCatAvail & "'", ConMain
End If
bytMax = IIf(IsNull(adrsPaid("AllowDays")), 0, adrsPaid("AllowDays"))
bytMin = IIf(IsNull(adrsPaid("MinAllowDays")), 0, adrsPaid("MinAllowDays"))
intTimes = IIf(IsNull(adrsPaid("No_OfTimes")), 0, adrsPaid("No_OfTimes"))
blnUnPaid = IIf(adrsPaid("Paid") = "N", True, False)
blnNoBal = IIf(adrsPaid("Type") = "N", True, False)
Exit Sub
ERR_P:
    ShowError ("GetLeaveDetails :: " & Me.Caption)
    bytMax = 0
    bytMin = 0
    intTimes = 0
End Sub

Private Function ImmediateAbsent() As Boolean       '' Checks if the Employee is Absent on
On Error GoTo ERR_P                                 '' Consecutive Days
Dim strTmp As String, strPATmp As String, bytPATmp As Byte
ImmediateAbsent = True
strTmp = ""
strPATmp = ""
bytPATmp = 0
strTmp = GetMnlTrnFile((DateCompDate(txtFrom.Text) - 1))
If FindTable(strTmp) Then
    If adrsPaid.State = 1 Then adrsPaid.Close
    adrsPaid.Open "Select Presabs from " & strTmp & " where " & strKDate & " =" & strDTEnc & _
    DateCompStr(CStr(DateCompDate(txtFrom.Text) - 1)) & strDTEnc & " and EmpCode=" & _
    "'" & cboCode.Text & "'", ConMain, adOpenKeyset
    If Not (adrsPaid.BOF And adrsPaid.EOF) Then
        strPATmp = adrsPaid("Presabs")
        bytPATmp = Len(strPATmp)
        If Mid(strPATmp, 1, bytPATmp / 2) = pVStar.AbsCode Then
            Select Case MsgBox(NewCaptionTxt("07028", adrsC) & vbCrLf & NewCaptionTxt("07029", adrsC) & _
            cboLeave.Text & NewCaptionTxt("07030", adrsC), vbYesNo + vbQuestion)
                Case 7 'no
                    txtFrom.SetFocus
                    ImmediateAbsent = False
                    Exit Function
            End Select
        End If
    End If
End If

strTmp = ""
strPATmp = ""
bytPATmp = 0
strTmp = GetMnlTrnFile((DateCompDate(txtTo.Text) + 1))
If FindTable(strTmp) Then
    If adrsPaid.State = 1 Then adrsPaid.Close
    adrsPaid.Open "Select Presabs from " & strTmp & " where " & strKDate & " =" & strDTEnc & _
    DateCompStr(CStr(DateCompDate(txtTo.Text) + 1)) & strDTEnc & " and Empcode=" & _
    "'" & cboCode.Text & "'", ConMain
    If Not (adrsPaid.BOF And adrsPaid.EOF) Then
        strPATmp = adrsPaid("Presabs")
        bytPATmp = Len(strPATmp)
        If Mid(strPATmp, 1, bytPATmp / 2) = pVStar.AbsCode Then
            Select Case MsgBox(NewCaptionTxt("07028", adrsC) & vbCrLf & NewCaptionTxt("07029", adrsC) & _
            cboLeave.Text & NewCaptionTxt("07030", adrsC), vbYesNo + vbQuestion)
                Case 7 'no
                    txtFrom.SetFocus
                    ImmediateAbsent = False
                    Exit Function
            End Select
        End If
    End If
End If
Exit Function
ERR_P:
    ShowError ("ImmediateAbsent :: " & Me.Caption)
    ImmediateAbsent = False
End Function


Private Function ImmediateLeave() As Boolean        '' Checks if Immidiate Leaves are Taken
On Error GoTo ERR_P                                 '' or not
ImmediateLeave = True
Dim strA_R As String, strHFOPT As String
Dim bytCtr As Byte
strA_R = "select * from LvInfo" & Right(pVStar.YearSel, 2) & " where ((" & strDTEnc & _
DateCompStr(CStr(DateCompDate(txtFrom.Text) - 1)) & strDTEnc & " between fromdate and todate ) or (" & _
strDTEnc & DateCompStr(CStr(DateCompDate(txtTo.Text) + 1)) & strDTEnc & _
" between fromdate and todate)) and trcd=4 and Empcode='" & cboCode.Text & "'"
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open strA_R, ConMain
bytCtr = 0
If Not (adrsPaid.EOF And adrsPaid.BOF) Then
        Do While Not adrsPaid.EOF
                If strF = "F " Then bytCtr = bytCtr + 1
                If strT = "T " Then bytCtr = bytCtr + 1
                strHFOPT = Left(adrsPaid("hf_option"), 2)
                If strHFOPT = strF Or strHFOPT = "F " Then bytCtr = bytCtr + 1
                strHFOPT = Right(adrsPaid("HF_Option"), 2)
                If strHFOPT = strT Or strHFOPT = "T " Then bytCtr = bytCtr + 1
                adrsPaid.MoveNext
        Loop
End If
If bytCtr > 0 Then
    If MsgBox(NewCaptionTxt("07031", adrsC), vbYesNo + vbQuestion) = vbYes Then
        ImmediateLeave = True
    Else
        txtFrom.SetFocus
        ImmediateLeave = False
    End If
End If
Exit Function
ERR_P:
    ShowError ("ImmediateLeave :: " & Me.Caption)
    ImmediateLeave = False
End Function

Private Function SaveAddMaster() As Boolean     '' Saves Data in the Leave Infomation File
On Error GoTo ERR_P                             '' and Updates the Balances of the Employee
Dim strTmp As String, strTmpTrn As String
Dim dtTemp As Date
If blnCOChecks Then
    strTmp = cboCOEntry.Value
Else
    strTmp = CStr(Date)
End If

SaveAddMaster = True                            '' in the Leave Balance File
ConMain.BeginTrans
'' Insert Information in LvInfo

    If blnCOChecks = False Then
        ConMain.Execute "insert into LvInfo" & Right((pVStar.YearSel), 2) & _
        " (EmpCode,LCode,Fromdate,Todate,Trcd,Days,Lv_Type_rw,Hf_Option,Entrydate )  values" & _
        "(" & "'" & cboCode.Text & "'" & "," & "'" & cboLeave.Text & "'" & "," & strDTEnc & DateSaveIns(txtFrom.Text) & _
        strDTEnc & "," & strDTEnc & DateSaveIns(txtTo.Text) & strDTEnc & "," & 4 & "," & _
        txtDays.Text & "," & "'" & strRW & "'" & "," & "'" & _
        strHf_Opt & "'" & "," & strDTEnc & DateSaveIns(strTmp) & strDTEnc & ")"
    Else
        ConMain.Execute "insert into LvInfo" & Right((pVStar.YearSel), 2) & _
        " (EmpCode,LCode,Fromdate,Todate,Trcd,Days,Lv_Type_rw,Hf_Option,Entrydate, fordate)  values" & _
        "(" & "'" & cboCode.Text & "'" & "," & "'" & cboLeave.Text & "'" & "," & strDTEnc & DateSaveIns(txtFrom.Text) & _
        strDTEnc & "," & strDTEnc & DateSaveIns(txtTo.Text) & strDTEnc & "," & 4 & "," & _
        txtDays.Text & "," & "'" & strRW & "'" & "," & "'" & _
        strHf_Opt & "'" & "," & strDTEnc & DateSaveIns(strTmp) & strDTEnc & ", " & strDTEnc & cboCOEntry.Value & strDTEnc & ")"
    End If
'' Update balance in LvBal
If (Not blnNoBal) Then
    If SubLeaveFlag = 1 Then   ' 07-11
        If (cboLeave.Text = "HP" Or cboLeave.Text = "CM") Then
            ConMain.Execute "Update Lvbal" & Right(pVStar.YearSel, 2) & " Set " & _
            "SL=" & sngDaysBal & " Where Empcode='" & cboCode.Text & "'"
        ElseIf cboLeave.Text = "LT" Then
            ConMain.Execute "Update Lvbal" & Right(pVStar.YearSel, 2) & " Set " & _
            "PL=" & sngDaysBal & " Where Empcode='" & cboCode.Text & "'"
        Else
            ConMain.Execute "Update Lvbal" & Right(pVStar.YearSel, 2) & " Set " & _
            cboLeave.Text & "=" & sngDaysBal & " Where Empcode='" & cboCode.Text & "'"
        End If
        If (cboLeave.Text = "EN" Or cboLeave.Text = "NE") Then ' 15-10
            Dim strqry As String
            strqry = "select " & ELSubLeave & ",lvbal" & Right(pVStar.YearSel, 2) & ".EMPCODE from lvbal" & Right(pVStar.YearSel, 2) & " where empcode='" & cboCode.Text & "'"
            Call UpDateSubLeave("lvbal" & Right(pVStar.YearSel, 2), ELSubLeave, strqry, ELLeave)

        End If
    Else
        ConMain.Execute "Update Lvbal" & Right(pVStar.YearSel, 2) & " Set " & _
        cboLeave.Text & "=" & sngDaysBal & " Where Empcode='" & cboCode.Text & "'"
    End If
End If
'' Update Status in the Monthly Transaction File
Call UpdateStatusOnAdd
ConMain.CommitTrans
Exit Function
ERR_P:
    ConMain.RollbackTrans
    SaveAddMaster = False
    ShowError ("SaveAddMaster :: " & Me.Caption)
End Function

Private Sub UpdateStatusOnAdd()         '' Updates the Status in the Monthly Transaction
On Error GoTo ERR_P                     '' File After Leave is Availed by the Employee
Dim bytCnt As Byte, bytCntTmp As Byte   '' Temporary Variables
Dim strTmpTrn As String, strTmpShf As String    '' Temporary Trn and Shf File Variables
Dim dtTemp As Date                      '' Temporary Variables
dtTemp = DateCompDate(txtFrom.Text)

Do While Month(dtTemp) <= Month(DateCompDate(txtTo.Text)) And Year(dtTemp) <= Year(DateCompDate(txtTo.Text))
    strTmpTrn = MakeName(MonthName(Month(dtTemp)), Year(dtTemp), "Trn")
    strTmpShf = MakeName(MonthName(Month(dtTemp)), Year(dtTemp), "shf")
    '' Only If Monthly Trn File Is Found
    If FindTable(strTmpTrn) Then
        If adrsPaid.State = 1 Then adrsPaid.Close
        ''This If Condition Add By
            adrsPaid.Open "Select " & strKDate & " ,Presabs from " & strTmpTrn & " where Empcode= " & "'" & _
            cboCode.Text & "'" & " order by " & strKDate & " ", _
            ConMain, adOpenKeyset, adLockOptimistic
        ''
        If Not (adrsPaid.BOF And adrsPaid.EOF) Then
            Do
                If adrsPaid!Date > DateCompDate(txtTo.Text) Then Exit Sub
                If adrsPaid("date") >= DateCompDate(txtFrom.Text) And _
                adrsPaid("date") <= DateCompDate(txtTo.Text) Then
                    '' For Date=From date and <>To Date
                    If adrsPaid("date") = DateCompDate(txtFrom.Text) And _
                    adrsPaid("date") <> DateCompDate(txtTo.Text) Then
                        Call CriteriaOneAdd(adrsPaid, strTmpTrn, strTmpShf)
                    End If
                    '' For Date=To date and <> From date
                    If adrsPaid("date") = DateCompDate(txtTo.Text) And _
                    adrsPaid("date") <> DateCompDate(txtFrom.Text) Then
                        Call CriteriaTwoAdd(adrsPaid, strTmpTrn, strTmpShf)
                    End If
                    '' For Date > From date and < To Date
                    If adrsPaid("date") > DateCompDate(txtFrom.Text) And _
                    adrsPaid("date") < DateCompDate(txtTo.Text) Then
                        Call CriteriaThreeAdd(adrsPaid, strTmpTrn, strTmpShf)
                    End If
                    '' For Date=From date and =To date
                    If adrsPaid("date") = DateCompDate(txtFrom.Text) And _
                    adrsPaid("date") = DateCompDate(txtTo.Text) Then
                        Call CriteriaFourAdd(adrsPaid, strTmpTrn, strTmpShf)
                    End If
                End If
                adrsPaid.MoveNext
            Loop Until adrsPaid.EOF
            adrsPaid.Close
        End If
    End If
    dtTemp = DateAdd("m", 1, dtTemp)
Loop
Exit Sub
ERR_P:
    ShowError ("UpdateStatusOnAdd :: " & Me.Caption)
End Sub

Private Sub UpdateStatusOnDelete()      '' Updates the Status in the Monthly Transaction
On Error GoTo ERR_P                     '' After Availed Leave is Deleted for the Employee
Dim bytCnt As Byte, bytCntTmp As Byte   '' Temporary Variables
Dim strTmpTrn As String, strTmpShf As String        '' Temporary Trn & Shf File Variables
Dim dtTemp As Date                      '' Temporary Variables

    dtTemp = DateCompDate(txtFrom.Text)
''
Do While Month(dtTemp) <= Month(DateCompDate(txtTo.Text))
    strTmpTrn = MakeName(MonthName(Month(dtTemp)), Year(dtTemp), "Trn")
    strTmpShf = MakeName(MonthName(Month(dtTemp)), Year(dtTemp), "shf")
    If Not FindTable(strTmpTrn) Then
            MsgBox NewCaptionTxt("07032", adrsC) & vbCrLf & NewCaptionTxt("07033", adrsC) & _
            NewCaptionTxt("07034", adrsC), vbInformation
        GoTo ChangeStatus
    Else
        '' Only if Monthly Transaction File is Found
        If adrsPaid.State = 1 Then adrsPaid.Close
        adrsPaid.Open "select * from " & strTmpTrn & " where Empcode=" & "'" & cboCode.Text & _
        "'" & " order by Empcode," & strKDate & " ", ConMain, adOpenKeyset, adLockOptimistic
        If Not (adrsPaid.BOF And adrsPaid.EOF) Then
            Do
                If adrsPaid!Date > DateCompDate(txtFrom.Text) And (adrsPaid.EOF And adrsPaid.BOF) Then Exit Sub
                '' get the Shift on That Day
                strShiftDel = RetShiftStat(strTmpShf, adrsPaid!Date, cboCode.Text)
                '' get HalfDay Timing on that day
                sngHDend = HalfDay(adrsPaid("shift"))
                '' if date= From date and <> To Date
                If adrsPaid!Date = DateCompDate(txtFrom.Text) And _
                adrsPaid!Date <> DateCompDate(txtTo.Text) Then
                    Call CriteriaOneDelete(adrsPaid, strTmpTrn, strTmpShf)
                '' If Date<> From Date and = To Date
                ElseIf adrsPaid!Date <> DateCompDate(txtFrom.Text) And _
                adrsPaid!Date = DateCompDate(txtTo.Text) Then
                    Call CriteriaTwoDelete(adrsPaid, strTmpTrn, strTmpShf)
                '' If Date= From Date and = To Date
                ElseIf adrsPaid!Date = DateCompDate(txtFrom.Text) And _
                adrsPaid!Date = DateCompDate(txtTo.Text) Then
                    Call CriteriaThreeDelete(adrsPaid, strTmpTrn, strTmpShf)
                '' If Date > From Date and < To Date
                ElseIf adrsPaid!Date > DateCompDate(txtFrom.Text) And _
                adrsPaid!Date < DateCompDate(txtTo.Text) Then
                    Call CriteriaFourDelete(adrsPaid, strTmpTrn, strTmpShf)
                End If
                adrsPaid.MoveNext
             Loop Until adrsPaid.EOF
            adrsPaid.Close
        End If
    End If
    dtTemp = DateAdd("m", 1, dtTemp)
Loop
ChangeStatus:
Exit Sub
ERR_P:
    ShowError ("UpdateStatusOnDelete :: " & Me.Caption)
End Sub

Private Sub UpdateDeleteBalance()       '' Updates the Balance of the Particular Employee
On Error GoTo ERR_P                     '' Once the Availed Leave is Deleted
Dim bytCntTmp As Byte, blnTmp As Boolean
sngDaysBal = 0: blnTmp = False
If SubLeaveFlag = 1 And (cboLeave.Text = "HP" Or cboLeave.Text = "CM") Then   ' 07-11
    For bytCntTmp = 1 To MSF2.Rows - 1
        If "SL" = MSF2.TextMatrix(bytCntTmp, 0) Then
            blnTmp = True
            sngDaysBal = Val(MSF2.TextMatrix(bytCntTmp, 2))
        End If
    Next
ElseIf SubLeaveFlag = 1 And cboLeave.Text = "LT" Then
    For bytCntTmp = 1 To MSF2.Rows - 1
        If "PL" = MSF2.TextMatrix(bytCntTmp, 0) Then
            blnTmp = True
            sngDaysBal = Val(MSF2.TextMatrix(bytCntTmp, 2))
        End If
    Next
Else
    For bytCntTmp = 1 To MSF2.Rows - 1
        If cboLeave.Text = MSF2.TextMatrix(bytCntTmp, 0) Then
            blnTmp = True
            sngDaysBal = Val(MSF2.TextMatrix(bytCntTmp, 2))
        End If
    Next
End If
If blnTmp Then
    sngDaysBal = sngDaysBal + Val(txtDays.Text)
    '' Update Balance of the Employee
    If SubLeaveFlag = 1 Then   ' 07-11
        If cboLeave.Text = "HP" Or cboLeave.Text = "CM" Then
            ConMain.Execute "Update LvBal" & Right(pVStar.YearSel, 2) & " Set " & _
            "SL=" & sngDaysBal & " Where Empcode='" & cboCode.Text & "'"
        ElseIf cboLeave.Text = "LT" Then
            ConMain.Execute "Update LvBal" & Right(pVStar.YearSel, 2) & " Set " & _
            "PL=" & sngDaysBal & " Where Empcode='" & cboCode.Text & "'"
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
End If
Exit Sub
ERR_P:
    ShowError ("UpdateDeleteBalance :: " & Me.Caption)
End Sub

Private Sub Display()       '' Displays the Leave Details
On Error GoTo ERR_P
cboLeave.Text = MSF1.TextMatrix(MSF1.row, 0)        '' Leave Code
txtFrom.Text = MSF1.TextMatrix(MSF1.row, 1)         '' From date
txtTo.Text = MSF1.TextMatrix(MSF1.row, 2)           '' To Date
Select Case Left(MSF1.TextMatrix(MSF1.row, 4), 2)
        Case "F "
                strFH = "Full Day"
        Case "FF"
                strFH = "First Half"
        Case "FS"
                strFH = "Second Half"
        Case "F"                '' For data from Import
                strFH = "Full Day"
        Case "H"                '' For data from Import
                strFH = "First Half"
        Case " ", "", "  "      '' For data from Import
                strFH = "Full Day"
End Select
Select Case Right(MSF1.TextMatrix(MSF1.row, 4), 2)
        Case "T "
                strSH = "Full Day"
        Case "TF"
                strSH = "First Half"
        Case "TS"
                strSH = "Second Half"
        Case "F"                '' For data from Import
                strSH = "Full Day"
        Case "H"                '' For data from Import
                strSH = "First Half"
        Case " ", "", "  "      '' For data from Import
                strSH = "Full Day"
End Select
If Not (MSF1.TextMatrix(MSF1.row, 4) = "6" Or MSF1.TextMatrix(MSF1.row, 4) = "7") Then   ' 23-05-09
cboFrom.Text = strFH                        '' From Status
cboTo.Text = strSH                          '' To Status
End If

'' Type of Leave
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open "Select Run_Wrk from LeavDesc where Cat='" & strCatAvail & _
"' and LvCode='" & cboLeave.Text & "'", ConMain
If Not (adrsPaid.EOF And adrsPaid.BOF) Then
    If adrsPaid("Run_Wrk") = "R" Then
        cboRW.Value = "Running"
    ElseIf adrsPaid("Run_Wrk") = "W" Then
        cboRW.Value = "Working"
    Else
        cboRW.Value = ""
    End If
Else
    cboRW.Value = ""
End If
txtDays.Text = Format(MSF1.TextMatrix(MSF1.row, 3), "0.00")     '' Leave Days
Exit Sub
ERR_P:
    Select Case Err.Number
        Case 380
            MsgBox NewCaptionTxt("07035", adrsC) & vbCrLf & _
            NewCaptionTxt("07036", adrsC) & vbCrLf & _
            NewCaptionTxt("07037", adrsC), vbExclamation
        Case Else
            ShowError ("Display  :: " & Me.Caption)
            'Resume Next
    End Select
End Sub

Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo ERR_P
Call AddActivityLog(lgADD_MODE, 3, 21)     '' Add Activity
Call AuditInfo("ADD", Me.Caption, "Add Leave Avail Entry " & cboLeave.Text & " For Employee " & cboCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Function HalfDay(ByVal shftName As String) As Single
On Error GoTo Err_particular
HalfDay = 0
If adRsintshft.State = 1 Then adRsintshft.Close
adRsintshft.Open "select hdend  from instshft where shift=" & "'" & shftName & "'", ConMain
If Not (adRsintshft.EOF And adRsintshft.BOF) Then HalfDay = IIf(IsNull(adRsintshft(0)), 0, adRsintshft(0))
Exit Function
Err_particular:
    ShowError ("Half Day :: Common")
End Function

Private Function RetShiftStat(ByVal tbName As String, ByVal dt As Date, ByVal strEmpCode) As String
On Error GoTo Err_particular
' Returns the shift for the day and month passed as argument
If adRsintshft.State = 1 Then adRsintshft.Close
adRsintshft.Open "select " & "D" & Day(dt) & " from " & tbName & " where Empcode='" & strEmpCode & _
"'", ConMain
If Not (adRsintshft.EOF And adRsintshft.BOF) Then RetShiftStat = IIf(IsNull(adRsintshft(0)), "", adRsintshft(0))
Exit Function
Err_particular:
    ShowError ("Return Shift Status :: Common")
End Function

Private Function GetMnlTrnFile(ByVal dt As String) As String
On Error GoTo ERR_P
Dim Mon_trn As String
Mon_trn = MakeName(MonthName(Month(CDate(dt))), Year(CDate(dt)), "trn")
If FindTable(Mon_trn) Then
    GetMnlTrnFile = Mon_trn
Else
    GetMnlTrnFile = ""
End If
Exit Function
ERR_P:
    ShowError ("GetMnlTrnFile :: Common ")
    GetMnlTrnFile = ""
End Function

Public Sub FillCboCo()
Dim strArrTmp() As String, bytTmp As Byte
 
 bytTmp = 0
    cboCOEntry.ColumnCount = 2
    cboCOEntry.ListWidth = "4 cm"
    cboCOEntry.ColumnWidths = "2 cm;1.5 cm"
    If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "Select distinct * from lvinfo" & Right(pVStar.YearSel, 2) & " where empcode='" & cboCode.Text & "' and lcode='" & LV_CO & "' and trcd=2", ConMain
    If adrsTemp.EOF = False Then
        adrsTemp.MoveFirst
        ReDim strArrTmp(adrsTemp.RecordCount - 1, 1)
        Do While adrsTemp.EOF = False
        If adrsCO.State = 1 Then adrsCO.Close
            'change by  from "MM/DD/YYYY" to "DD/MMM/YYYY"
            adrsCO.Open "SELECT * FROM lvinfo" & Right(pVStar.YearSel, 2) & _
            " where empcode='" & cboCode.Text & _
            "' and lcode='" & LV_CO & "' and trcd=4 and entrydate=" & _
            strDTEnc & Format(adrsTemp!FromDate, "DD/MMM/YYYY") & _
            strDTEnc, ConMain
            If adrsCO.EOF = True Then
                
                  strArrTmp(bytTmp, 0) = Format(adrsTemp("EntryDate"), "DD/MMM/YYYY")
                  strArrTmp(bytTmp, 1) = adrsTemp("Days")         '' Type
                   bytTmp = bytTmp + 1
            End If
            adrsTemp.MoveNext
            If adrsTemp.EOF = True Then Exit Do
        Loop
        cboCOEntry.List = strArrTmp
        Erase strArrTmp
    End If
End Sub



