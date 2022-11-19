VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLunchRule 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lunch Rules"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frLateL 
      Caption         =   " "
      Height          =   3645
      Left            =   9000
      TabIndex        =   23
      Top             =   4200
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Frame frLL 
         Caption         =   "asdad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   4035
         Begin VB.OptionButton optpaidLL 
            Caption         =   "Paid Days"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   330
            Width           =   1380
         End
         Begin VB.OptionButton optLeavesLL 
            Caption         =   "Leaves"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1935
            TabIndex        =   25
            Top             =   285
            Width           =   975
         End
      End
      Begin MSMask.MaskEdBox txtDL 
         Height          =   330
         Left            =   3060
         TabIndex        =   27
         Top             =   765
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
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
      Begin MSMask.MaskEdBox txtCL 
         Height          =   315
         Left            =   810
         TabIndex        =   28
         Top             =   780
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
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
      Begin MSMask.MaskEdBox txtLL 
         Height          =   345
         Left            =   3090
         TabIndex        =   29
         Top             =   300
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
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
      Begin VB.Label lblLL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Late Allowed in a Month"
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
         TabIndex        =   40
         Top             =   360
         Width           =   2565
      End
      Begin VB.Label lblCL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cut"
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
         Left            =   135
         TabIndex        =   39
         Top             =   810
         Width           =   300
      End
      Begin VB.Label lblDL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day for Every"
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
         Left            =   1560
         TabIndex        =   38
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Late"
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
         Left            =   3720
         TabIndex        =   37
         Top             =   810
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leaves"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2100
         TabIndex        =   36
         Top             =   2010
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbl1LL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1st Preference"
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
         Left            =   300
         TabIndex        =   35
         Top             =   2340
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lbl2LL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Preference"
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
         Left            =   300
         TabIndex        =   34
         Top             =   2820
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl3LL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3rd Preference"
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
         Left            =   300
         TabIndex        =   33
         Top             =   3300
         Visible         =   0   'False
         Width           =   1260
      End
      Begin MSForms.ComboBox cbo1LL 
         Height          =   375
         Left            =   2100
         TabIndex        =   32
         Top             =   2310
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo2LL 
         Height          =   375
         Left            =   2100
         TabIndex        =   31
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo3LL 
         Height          =   375
         Left            =   2100
         TabIndex        =   30
         Top             =   3210
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CheckBox chkL 
      Caption         =   "Check1"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2625
   End
   Begin VB.CommandButton cmdEditSave 
      Caption         =   " "
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   1065
   End
   Begin VB.CommandButton cmdCanReset 
      Caption         =   " "
      Height          =   435
      Left            =   1440
      TabIndex        =   2
      Top             =   4440
      Width           =   1185
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   " "
      Height          =   435
      Left            =   2880
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CheckBox chkLL 
      Caption         =   "Check1"
      Height          =   255
      Left            =   9120
      TabIndex        =   0
      Top             =   4200
      Width           =   2625
   End
   Begin VB.Frame frLate 
      Caption         =   " "
      Height          =   3645
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Frame frL 
         Caption         =   "asdad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   90
         TabIndex        =   6
         Top             =   1200
         Width           =   4035
         Begin VB.OptionButton optLeavesL 
            Caption         =   "Leaves"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1935
            TabIndex        =   8
            Top             =   285
            Width           =   975
         End
         Begin VB.OptionButton optPaidL 
            Caption         =   "Paid Days"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   330
            Width           =   1380
         End
      End
      Begin MSMask.MaskEdBox txtLate 
         Height          =   330
         Left            =   3060
         TabIndex        =   9
         Top             =   765
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
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
      Begin MSMask.MaskEdBox txtDaysL 
         Height          =   315
         Left            =   810
         TabIndex        =   10
         Top             =   780
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
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
      Begin MSMask.MaskEdBox txtTotalL 
         Height          =   345
         Left            =   3090
         TabIndex        =   11
         Top             =   300
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
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
      Begin MSForms.ComboBox cbo3L 
         Height          =   375
         Left            =   2100
         TabIndex        =   22
         Top             =   3210
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo2L 
         Height          =   375
         Left            =   2100
         TabIndex        =   21
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo1L 
         Height          =   375
         Left            =   2100
         TabIndex        =   20
         Top             =   2310
         Visible         =   0   'False
         Width           =   1095
         VariousPropertyBits=   1820346395
         DisplayStyle    =   7
         Size            =   "1931;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lbl3L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3rd Preference"
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
         Left            =   300
         TabIndex        =   19
         Top             =   3300
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lbl2L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Preference"
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
         Left            =   300
         TabIndex        =   18
         Top             =   2820
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl1L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1st Preference"
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
         Left            =   300
         TabIndex        =   17
         Top             =   2340
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lblCapLL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leaves"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2100
         TabIndex        =   16
         Top             =   2010
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblLate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Late"
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
         Left            =   3720
         TabIndex        =   15
         Top             =   810
         Width           =   375
      End
      Begin VB.Label lblDayL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day for Every"
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
         Left            =   1500
         TabIndex        =   14
         Top             =   810
         Width           =   1155
      End
      Begin VB.Label lblCutL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cut"
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
         Left            =   135
         TabIndex        =   13
         Top             =   810
         Width           =   300
      End
      Begin VB.Label lblTotalL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Late Allowed in a Month"
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
         TabIndex        =   12
         Top             =   360
         Width           =   2565
      End
   End
   Begin MSForms.ComboBox cboCat 
      Height          =   375
      Left            =   2205
      TabIndex        =   42
      Top             =   0
      Width           =   2055
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "3625;661"
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
   Begin VB.Label lblCat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ewrwerwerw"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      TabIndex        =   41
      Top             =   60
      Width           =   1245
   End
End
Attribute VB_Name = "frmLunchRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adrsC As New ADODB.Recordset

Private Sub RetCaption()
On Error GoTo ERR_P
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '44%'", VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
lblCat.Caption = NewCaptionTxt("44002", adrsC)
chkL.Caption = NewCaptionTxt("44003", adrsC)
lblTotalL.Caption = NewCaptionTxt("44004", adrsC)
lblCutL.Caption = NewCaptionTxt("44005", adrsC)
lblDayL.Caption = NewCaptionTxt("44006", adrsC)
lblLate.Caption = NewCaptionTxt("00035", adrsMod)
frL.Caption = NewCaptionTxt("44007", adrsC)
optPaidL.Caption = NewCaptionTxt("44008", adrsC)
optLeavesL.Caption = NewCaptionTxt("44009", adrsC)
lblCapLL.Caption = NewCaptionTxt("44009", adrsC)
lbl1L.Caption = NewCaptionTxt("44010", adrsC)
lbl2L.Caption = NewCaptionTxt("44011", adrsC)
lbl3L.Caption = NewCaptionTxt("44012", adrsC)
chkLL.Caption = "LunchLate"
Call SetButtonCap
Exit Sub
ERR_P:
    ShowError ("RetCaption :: " & Me.Caption)
End Sub

Private Sub SetButtonCap(Optional bytFlgCap As Byte = 1)    '' Sets Captions to the
Select Case bytFlgCap
    Case 1
        cmdEditSave.Caption = NewCaptionTxt("00005", adrsMod)
        cmdCanReset.Caption = NewCaptionTxt("44015", adrsC)
        cmdExit.Caption = NewCaptionTxt("00008", adrsMod)
    Case 2
        cmdEditSave.Caption = NewCaptionTxt("00007", adrsMod)
        cmdCanReset.Caption = NewCaptionTxt("00003", adrsMod)
End Select
End Sub

Private Sub cboCat_Change()
''SG07
Call cboCat_Click
''
End Sub

Private Sub cboCat_Click()
If cboCat.Text = "" Then Exit Sub
Call FillLeaves             '' Fills the Leaves Combo
Call Display                '' Displays All the Details of the Selected Category
End Sub

Private Sub chkL_Click()
If chkL.Value = 1 Then
    frLate.Visible = True
Else
    frLate.Visible = False
End If
End Sub

Private Sub chkLL_Click()
If chkLL.Value = 1 Then
    frLateL.Visible = True
Else
    frLateL.Visible = False
End If
End Sub

Private Sub cmdCanReset_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 3
        bytMode = 1
        Call Display
        Call ChangeMode
    Case 1
        If Not DeleteRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        Else
            If cboCat.Text = "" Then Exit Sub
            If MsgBox(NewCaptionTxt("44016", adrsC), vbQuestion + vbYesNo) = vbYes Then
                '' Set All the Rules to Default Values
                '' Update
                VstarDataEnv.cnDJConn.Execute "Update CatDesc Set LunchLtRule='N'," & _
                "LunchLtInMnth=0.00,LunchLtCut=0.00,EverLunchLt=0.00,DedLunchLt='PD',FstLunchLtPr='',SecLunchLtPr=''," & _
                "TrdLunchLtPr='' Where Cat='" & cboCat.Text & "'"
                Call AddActivityLog(lgReset_Action, 1, 8)     '' Reset Activity
                Call AuditInfo("RESET", Me.Caption, "Reset Late/Early Rule Of Category: " & cboCat.Text)
                adrsDept1.Requery
                Call Display
            End If
        End If
End Select
Exit Sub
ERR_P:
    ShowError ("Cancel Reset :: " & Me.Caption)
End Sub

Private Sub cmdEditSave_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        If cboCat.Text = "" Then Exit Sub
        '' Check for Rights
        If Not EditRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        bytMode = 3
        Call ChangeMode
    Case 3      '' Edit Mode
        If Not ValidateModMaster Then Exit Sub  '' If not Valida for Edit
        If Not SaveModMaster Then Exit Sub      '' Save for Edit
        Call SaveModLog                         '' Save the Edit Log
        adrsDept1.Requery
        Call Display
        bytMode = 1
        Call ViewAction
End Select
Exit Sub
ERR_P:
    ShowError ("EditSave :: " & Me.Caption)
    'Resume Next
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()

chkLL.Visible = False

Call SetFormIcon(Me)        '' Sets the Form Icon
Call RetCaption             '' Sets the Captions
Call GetRights              '' Gets the Rights
Call OpenMasterTable        '' Opens the Master Table for Category
Call FillCombo              '' Fills the Category Combo
bytMode = 1                 '' Sets the Mode to the Default Mode
Call ChangeMode             '' Take Action Accordingly
End Sub

Private Sub ChangeMode()
Select Case bytMode
    Case 1          '' View Mode
        Call ViewAction
    Case 3          '' Edit Mde
        Call EditAction
End Select
End Sub

Private Sub ViewAction()
'' Disable Necessary Controls
frLate.Enabled = False          '' Disable Late Frame
chkL.Enabled = False            '' Disable Late CheckBox
frLateL.Enabled = False
chkLL.Enabled = False
'' Set Necessary Captions
Call SetButtonCap
End Sub

Private Sub EditAction()
'' Enable Necessary Controls
frLate.Enabled = True       '' Enable Late Frame
chkL.Enabled = True         '' Enable Late CheckBox
frLateL.Enabled = True
chkLL.Enabled = True
'' Change Necesary Captions
Call SetButtonCap(2)        '' SetCaptions of the Buttons
If chkL.Value = 1 Then
    txtTotalL.SetFocus          '' Sets the Focus to Total Late Allowed TextBox
Else
    chkL.SetFocus
End If
End Sub


Private Sub optLeavesL_Click()
cbo1L.Visible = True
cbo2L.Visible = True
cbo3L.Visible = True
lbl1L.Visible = True
lbl2L.Visible = True
lbl3L.Visible = True
lblCapLL.Visible = True
End Sub

Private Sub optPaidL_Click()
cbo1L.Visible = False
cbo2L.Visible = False
cbo3L.Visible = False
lbl1L.Visible = False
lbl2L.Visible = False
lbl3L.Visible = False
lblCapLL.Visible = False
End Sub

Private Sub OpenMasterTable()       '' Opens the Category Master Table
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select Cat," & strKDesc & ",LunchLtRule,LunchLtInMnth,LunchLtCut,EverLunchLt,DedLunchLt,FstLunchLtPr,SecLunchLtPr," & _
"TrdLunchLtPr From CatDesc" & _
" where cat <> '100' Order by Cat", VstarDataEnv.cnDJConn, adOpenStatic
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Function ValidateModMaster() As Boolean
On Error GoTo ERR_P                 '' Validates the Record Befor Updating it to the
ValidateModMaster = True            '' Table
'' For Late
Call FormatAll
If chkL.Value = 1 Then
    If Right(txtDaysL.Text, 2) <> "00" And Right(txtDaysL.Text, 2) <> "50" Then
        MsgBox NewCaptionTxt("44017", adrsC)
        txtDaysL.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
    If Right(txtLate.Text, 2) <> "00" Then
        MsgBox NewCaptionTxt("44018", adrsC)
        txtLate.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
    If Right(txtTotalL.Text, 2) <> "00" Then
        MsgBox NewCaptionTxt("44018", adrsC)
        txtTotalL.SetFocus
        ValidateModMaster = False
        Exit Function
    End If
    If optLeavesL.Value = True Then
        If cbo1L.Text = "" And cbo2L.Text = "" And cbo3L.Text = "" Then
            MsgBox NewCaptionTxt("44019", adrsC), vbExclamation
            cbo1L.SetFocus
            ValidateModMaster = False
            Exit Function
        End If
        If cbo3L.Text <> "" And cbo2L.Text = "" Then
            MsgBox NewCaptionTxt("44020", adrsC), vbExclamation
            cbo2L.SetFocus
            ValidateModMaster = False
            Exit Function
        End If
        If cbo2L.Text <> "" And cbo1L.Text = "" Then
            MsgBox NewCaptionTxt("44020", adrsC), vbExclamation
            cbo1L.SetFocus
            ValidateModMaster = False
            Exit Function
        End If
    End If
End If
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
    ValidateModMaster = False
End Function

Private Sub FillCombo()     '' Fills the Category Combo
On Error GoTo ERR_P
Dim strCatArr() As String   '' Array for Categories
Dim bytCntTmp As Byte       '' Count of Records

If adrsDept1.EOF Then Exit Sub
cboCat.Clear                '' Clear the ComboBox
adrsDept1.MoveFirst         '' Move to the First Record
ReDim strCatArr(adrsDept1.RecordCount - 1, 1)       '' Redimension the Array

For bytCntTmp = 0 To adrsDept1.RecordCount - 1
    strCatArr(bytCntTmp, 0) = adrsDept1("Cat")
    strCatArr(bytCntTmp, 1) = adrsDept1("Desc")
    adrsDept1.MoveNext
Next
adrsDept1.MoveFirst         '' Move to the First Record

cboCat.List = strCatArr     '' Assign the Array to the ComboBox
Exit Sub
ERR_P:
    ShowError ("FillCombo ::" & Me.Caption)
End Sub

Private Sub FillLeaves()            '' Fills the Combo with the Leaves Alloted to the
On Error GoTo ERR_P                 '' Employee
Dim strArrLeaves() As String        '' For the Leaves
Dim bytCntTmp As Byte               '' For the Count of Leaves
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open "Select Distinct(lvcode),Leave  from LeavDesc where Lvcode not in (" & "'" & _
pVStar.PrsCode & "'" & "," & "'" & pVStar.AbsCode & "'" & "," & "'" & pVStar.WosCode & _
"'" & "," & "'" & pVStar.HlsCode & "'" & ")" & " and type='Y'" & " and cat=" & "'" & _
cboCat.Text & "'", VstarDataEnv.cnDJConn, adOpenStatic
If (adrsPaid.EOF And adrsPaid.BOF) Then
    Exit Sub
Else
    bytCntTmp = adrsPaid.RecordCount - 1
    ReDim strArrLeaves(bytCntTmp + 1, 1)
    For bytCntTmp = 0 To adrsPaid.RecordCount - 1
        strArrLeaves(bytCntTmp, 0) = adrsPaid("LvCode")
        strArrLeaves(bytCntTmp, 1) = adrsPaid("Leave")
        adrsPaid.MoveNext
    Next
    strArrLeaves(bytCntTmp, 0) = ""
    strArrLeaves(bytCntTmp, 1) = ""
End If
'' Clear All the Combos
cbo1L.Clear
cbo2L.Clear
cbo3L.Clear
cbo1LL.Clear
cbo2LL.Clear
cbo3LL.Clear

'' Fill All the Combos
cbo1L.List = strArrLeaves
cbo2L.List = strArrLeaves
cbo3L.List = strArrLeaves
cbo1LL.List = strArrLeaves
cbo2LL.List = strArrLeaves
cbo3LL.List = strArrLeaves
Exit Sub
ERR_P:
    ShowError ("FillLeaves :: " & Me.Caption)
End Sub

Private Sub Display()
On Error GoTo ERR_P
adrsDept1.MoveFirst
adrsDept1.Find "Cat='" & cboCat.Text & "'"
If adrsDept1.EOF Then Exit Sub
'' Late
'' Others
chkL.Value = IIf(IsNull(adrsDept1("LunchLtRule")) = True Or adrsDept1("LunchLtRule") = "N", 0, 1)

txtTotalL.Text = IIf(IsNull(adrsDept1("LunchLtInMnth")), "0.00", _
                Format(adrsDept1("LunchLtInMnth"), "0.00"))
txtDaysL.Text = IIf(IsNull(adrsDept1("LunchLtCut")), "0.00", _
                Format(adrsDept1("LunchLtCut"), "0.00"))
txtLate.Text = IIf(IsNull(adrsDept1("EverLunchLt")), "0.00", _
                Format(adrsDept1("EverLunchLt"), "0.00"))
optPaidL.Value = IIf(IsNull(adrsDept1("DedLunchLt")) Or adrsDept1("DedLunchLt") = "PD", True, False)
optLeavesL.Value = IIf(adrsDept1("DedLunchLt") = "LV", True, False)
'' Leaves
cbo1L.Value = GetValidLeave(IIf(IsNull(adrsDept1("FstLunchLtPr")), "", adrsDept1("FstLunchLtPr")))
cbo2L.Value = GetValidLeave(IIf(IsNull(adrsDept1("SecLunchLtPr")), "", adrsDept1("SecLunchLtPr")))
cbo3L.Value = GetValidLeave(IIf(IsNull(adrsDept1("TrdLunchLtPr")), "", adrsDept1("TrdLunchLtPr")))
Exit Sub
ERR_P:
    ShowError ("Display :: " & Me.Caption)
End Sub

Private Function GetValidLeave(ByVal strLeave As String) As String
On Error GoTo ERR_P             '' Checks if the Leave is Valid to be Displayed inth ComboBox
If IsNull(strLeave) Or strLeave = "" Then
GetValidLeave = ""
Exit Function
End If
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open "Select LvCode from LeavDesc Where LvCode='" & strLeave & "'", _
VstarDataEnv.cnDJConn, adOpenStatic
If (adrsPaid.EOF And adrsPaid.BOF) Then
    GetValidLeave = ""
Else
    GetValidLeave = adrsPaid("LvCode")
End If
Exit Function
ERR_P:
    ShowError ("GetValidLeave :: " & Me.Caption)
    GetValidLeave = ""
End Function

Private Function SaveModMaster() As Boolean     '' Saves the Data in the Edit Mode
On Error GoTo ERR_P
Dim strLate As String
If optPaidL.Value = True Then
    strLate = "PD"
Else
    strLate = "LV"
End If
SaveModMaster = True
'' Update
VstarDataEnv.cnDJConn.Execute "Update CatDesc Set LunchLtRule='" & IIf(chkL.Value = 1, "Y", "N") & _
"',LunchLtInMnth=" & txtTotalL.Text & ",LunchLtCut=" & txtDaysL.Text & ",EverLunchLt=" & txtLate.Text & _
",DedLunchLt='" & strLate & "',FstLunchLtPr='" & cbo1L.Value & "',SecLunchLtPr='" & cbo2L.Value & "'," & _
"TrdLunchLtPr='" & cbo3L.Value & "' Where Cat='" & cboCat.Text & "'"
Exit Function
ERR_P:
    ShowError ("SaveModMaster :: " & Me.Caption)
    SaveModMaster = False
End Function

Private Sub FormatAll()     '' Formats All the Numerical data to the 0.00 Format
txtTotalL.Text = IIf(txtTotalL.Text = "", "0.00", Format(txtTotalL.Text, "0.00"))
txtDaysL.Text = IIf(txtDaysL.Text = "", "0.00", Format(txtDaysL.Text, "0.00"))
txtLate.Text = IIf(txtLate.Text = "", "0.00", Format(txtLate.Text, "0.00"))
End Sub

Private Sub GetRights()         '' Gets and Sets the Particular Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 1)
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


Private Sub txtDaysL_GotFocus()
    Call GF(txtDaysL)
End Sub

Private Sub txtDaysL_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = keycheck(KeyAscii, txtDaysL)
End Select
End Sub



Private Sub txtLate_GotFocus()
    Call GF(txtLate)
End Sub

Private Sub txtLate_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = keycheck(KeyAscii, txtLate)
End Select
End Sub


Private Sub txtTotalL_GotFocus()
    Call GF(txtTotalL)
End Sub

Private Sub txtTotalL_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        SendKeys Chr(9)
    Case Else
        KeyAscii = keycheck(KeyAscii, txtTotalL)
End Select
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 8)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edit Late/Early Rule Of Category: " & cboCat.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub
