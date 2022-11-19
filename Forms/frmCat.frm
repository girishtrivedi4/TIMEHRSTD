VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Minimum Work Hour"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frMisc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3000
      TabIndex        =   23
      Top             =   0
      Width           =   6135
      Begin VB.CheckBox ChkWOffPaid 
         Alignment       =   1  'Right Justify
         Caption         =   "Weekly Off Paid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   345
         Width           =   1845
      End
      Begin MSMask.MaskEdBox txtName 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
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
      Begin MSMask.MaskEdBox txtCode 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   315
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
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
         PromptChar      =   "_"
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category  Code"
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
         TabIndex        =   25
         Top             =   352
         Width           =   1350
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category Name"
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
         TabIndex        =   24
         Top             =   810
         Width           =   1350
      End
   End
   Begin VB.Frame frHours 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   3000
      TabIndex        =   16
      Top             =   1215
      Width           =   8340
      Begin MSMask.MaskEdBox txtCutE 
         Height          =   315
         Left            =   7365
         TabIndex        =   9
         Top             =   874
         Width           =   795
         _ExtentX        =   1402
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
      Begin MSMask.MaskEdBox txtCutL 
         Height          =   315
         Left            =   7365
         TabIndex        =   8
         Top             =   420
         Width           =   795
         _ExtentX        =   1402
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
      Begin MSMask.MaskEdBox txtLateG 
         Height          =   315
         Left            =   3165
         TabIndex        =   7
         Top             =   1782
         Width           =   795
         _ExtentX        =   1402
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
      Begin MSMask.MaskEdBox txtEarlyA 
         Height          =   315
         Left            =   3165
         TabIndex        =   6
         Top             =   1328
         Width           =   795
         _ExtentX        =   1402
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
      Begin MSMask.MaskEdBox txtGEarly 
         Height          =   315
         Left            =   3165
         TabIndex        =   5
         Top             =   874
         Width           =   795
         _ExtentX        =   1402
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
      Begin MSMask.MaskEdBox txtHalfDay 
         Height          =   315
         Left            =   7365
         TabIndex        =   11
         Top             =   1782
         Width           =   795
         _ExtentX        =   1402
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
      Begin MSMask.MaskEdBox txtFullDay 
         Height          =   315
         Left            =   7365
         TabIndex        =   10
         Top             =   1328
         Width           =   795
         _ExtentX        =   1402
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
      Begin MSMask.MaskEdBox txtCLate 
         Height          =   315
         Left            =   3165
         TabIndex        =   4
         Top             =   420
         Width           =   795
         _ExtentX        =   1402
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
         PromptChar      =   "0"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hours"
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
         Left            =   7440
         TabIndex        =   29
         Top             =   195
         Width           =   570
      End
      Begin VB.Label lblCLate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Employee to come late by"
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
         Left            =   225
         TabIndex        =   28
         Top             =   450
         Width           =   2805
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum half day present hours"
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
         Left            =   4440
         TabIndex        =   27
         Top             =   1815
         Width           =   2775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum full day present hours"
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
         Left            =   4500
         TabIndex        =   26
         Top             =   1365
         Width           =   2715
      End
      Begin VB.Label lblGEarly 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allow Employee to go early by"
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
         Left            =   390
         TabIndex        =   22
         Top             =   915
         Width           =   2640
      End
      Begin VB.Label lblEarlyA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ignore early arrival before shift by"
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
         Left            =   195
         TabIndex        =   21
         Top             =   1365
         Width           =   2835
      End
      Begin VB.Label lblLateG 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ignore late going after shift by"
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
         TabIndex        =   20
         Top             =   1815
         Width           =   2550
      End
      Begin VB.Label lblCutL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cut half day if Late coming  by"
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
         Left            =   4575
         TabIndex        =   19
         Top             =   450
         Width           =   2640
      End
      Begin VB.Label lblCutE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cut half day if early going by"
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
         Left            =   4755
         TabIndex        =   18
         Top             =   915
         Width           =   2460
      End
      Begin VB.Label lblHrs1 
         AutoSize        =   -1  'True
         Caption         =   "Hours"
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
         Left            =   3270
         TabIndex        =   17
         Top             =   195
         Width           =   570
      End
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command1"
      Height          =   405
      Left            =   3000
      TabIndex        =   12
      Top             =   3720
      Width           =   1725
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command1"
      Height          =   405
      Left            =   4680
      TabIndex        =   13
      Top             =   3720
      Width           =   1725
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command1"
      Height          =   405
      Left            =   6360
      TabIndex        =   14
      Top             =   3720
      Width           =   1725
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   405
      Left            =   8040
      TabIndex        =   15
      Top             =   3720
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   6324
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColorFixed  =   12632256
      AllowBigSelection=   0   'False
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
Attribute VB_Name = "frmCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset


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
'
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '10%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("10001", adrsC)              '' Form caption
Call SetOtherCaps                           '' Set the Captions for the Other Controls
Call SetGButtonCap(Me)                           '' Sets Appropriate Caption for the Buttons
Call CapGrid                                '' Captions For Grid
Exit Sub
ERR_P:
    ShowError ("RetCaption :: " & Me.Caption)
End Sub

Private Sub SetOtherCaps()
'' Misc
frMisc.Caption = NewCaptionTxt("10004", adrsC)      '' Info
lblCode.Caption = NewCaptionTxt("10002", adrsC)     '' Category Code
lblName.Caption = NewCaptionTxt("10003", adrsC)     '' Category Name
'' Hours
frHours.Caption = NewCaptionTxt("10005", adrsC)     '' Late/Early Rules
lblHrs1.Caption = NewCaptionTxt("00023", adrsMod)     '' Hours
lblCLate.Caption = NewCaptionTxt("10006", adrsC)    '' Come Late
lblGEarly.Caption = NewCaptionTxt("10007", adrsC)   '' Go Early
lblEarlyA.Caption = NewCaptionTxt("10008", adrsC)   '' Early Arrival
lblLateG.Caption = NewCaptionTxt("10009", adrsC)    '' Late Going
lblCutL.Caption = NewCaptionTxt("10010", adrsC)     '' Cut Late
lblCutE.Caption = NewCaptionTxt("10011", adrsC)     '' Cut Early
End Sub

Private Sub CapGrid()           '' Gives the Captions to the Grid
With MSF1
    '' Sets the Column Widhts
    .ColWidth(0) = .ColWidth(0) * 0.7
    .ColWidth(1) = .ColWidth(1) * 2
    '' Sets the Column Alignment
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    '' Sets the Appropriate Captions
    .TextMatrix(0, 0) = "Code"
    .TextMatrix(0, 1) = "Description"
End With
End Sub

Private Sub OpenMasterTable()             '' Open the recordset for the Display purposes
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select * from CatDesc where cat <> '100' Order by Cat", ConMain, adOpenStatic
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
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
        .TextMatrix(intCounter, 0) = adrsDept1("Cat")
        .TextMatrix(intCounter, 1) = IIf(IsNull(adrsDept1("Desc")), "", adrsDept1("Desc"))
    End With
    adrsDept1.MoveNext
Next

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

Private Sub ViewAction()    '' Procedure for Viewing Mode
'' Enable Needed Controls
cmdAddSave.Enabled = True       '' Enable ADD/SAVE Button
cmdEditCan.Enabled = True       '' Enable Edit/Cancel Button
cmdDel.Enabled = True           '' Enable Delete Button
'' Disable Needed Controls
frMisc.Enabled = False          '' Disable Info Frame
frHours.Enabled = False         '' Disable Late Early Frame

'' Give Captions to the Needed Controls
Call SetGButtonCap(Me)
Call Display
End Sub

Private Sub EditAction()    '' Procedure for Edit Mode
'' Enable Necessary Controls
frMisc.Enabled = True       '' Enable Info Frame
frHours.Enabled = True      '' Enable Late / Early Frame
Call SetGButtonCap(Me, 2)
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
txtCode.Enabled = False         '' Disable Code TextBox
txtName.SetFocus                '' Set Focus on the Name TextBox

End Sub

Private Sub AddAction()     '' Procedure for Addition Mode
frMisc.Enabled = True       '' Enable Info Frame
frHours.Enabled = True      '' Enable Late / Early Frame
txtCode.Enabled = True      '' Enable Code TextBox

cmdDel.Enabled = False      '' Disable Delete Button
Call SetGButtonCap(Me, 2)
Call clear
End Sub

Private Sub clear()
txtCode.Text = ""       '' Clear Code Control
txtName.Text = ""       '' Clear Name Control
'' Hours
txtCLate.Text = ""      '' Allow Coming Late By
txtGEarly.Text = ""     '' Allow Going Early By
txtEarlyA.Text = ""     '' Ignore Late Going
txtLateG.Text = ""      '' Ignore early Coming
txtCutL.Text = ""       '' Cut Half Day if Came Late More Than
txtCutE.Text = ""       '' Cut Half Days if Early Gone More Than
txtCode.SetFocus      '' Set Focus to the Code TextBox
ChkWOffPaid.Value = 1
txtFullDay.Text = ""
txtHalfDay.Text = ""
End Sub

Private Sub Display()       '' Displays the Given Master Records
On Error GoTo ERR_P
adrsDept1.Requery
If adrsDept1.EOF Then
    cmdEditCan.Enabled = False
    Exit Sub
End If
adrsDept1.MoveFirst
adrsDept1.Find "Cat='" & MSF1.TextMatrix(MSF1.Row, 0) & "'"
If Not (adrsDept1.EOF) Then
    '' Category Code
    txtCode.Text = adrsDept1("Cat")
    '' Category Name
    txtName.Text = adrsDept1("Desc")
    '' Allow to Come late by
    txtCLate.Text = IIf(IsNull(adrsDept1("Lt_Allow")), "00.00", _
                    Format(adrsDept1("Lt_Allow"), "00.00"))
    '' Allow to go Early by
    txtGEarly.Text = IIf(IsNull(adrsDept1("Erl_Allow")), "00.00", _
                    Format(adrsDept1("Erl_Allow"), "00.00"))
    '' Ignore Early Arrival
    txtEarlyA.Text = IIf(IsNull(adrsDept1("Erl_Ignore")), "00.00", _
                    Format(adrsDept1("Erl_Ignore"), "00.00"))
    '' Ignore Late Going
    txtLateG.Text = IIf(IsNull(adrsDept1("Lt_Ignore")), "00.00", _
                    Format(adrsDept1("Lt_Ignore"), "00.00"))
    '' Cut Half Day if Came Late more than
    txtCutL.Text = IIf(IsNull(adrsDept1("HalfCutLt")), "00.00", _
                    Format(adrsDept1("HalfCutLt"), "00.00"))
    '' Cut Half Day ig Early Gone More Than
    txtCutE.Text = IIf(IsNull(adrsDept1("HalfCutEr")), "00.00", _
                    Format(adrsDept1("HalfCutEr"), "00.00"))
    txtFullDay.Text = IIf(IsNull(adrsDept1("FullDayHr")), "00.00", Format(adrsDept1("FullDayHr"), "00.00"))
    txtHalfDay.Text = IIf(IsNull(adrsDept1("HalfDayHr")), "00.00", Format(adrsDept1("HalfDayHr"), "00.00"))
    ChkWOffPaid.Value = IIf(adrsDept1!WeekOffPaid = "Y", 1, 0) ''  26-12

End If
Exit Sub
ERR_P:
    ShowError ("Display :: " & Me.Caption)
    'Resume Next
End Sub

Private Function ValidateAddmaster() As Boolean     '' Validate If in Addition Mode
On Error GoTo ERR_P
ValidateAddmaster = True
'' Check for Blank Category Code
If Trim(txtCode.Text) = "" Then
    MsgBox NewCaptionTxt("10016", adrsC), vbExclamation
    ValidateAddmaster = False
    txtCode.SetFocus
    Exit Function
End If
'' Check for Existing Category Code
If MSF1.Rows > 1 Then
    '' Category Code
    adrsDept1.MoveFirst
    adrsDept1.Find "Cat='" & txtCode.Text & "'"
    If Not adrsDept1.EOF Then
        MsgBox NewCaptionTxt("10017", adrsC), vbExclamation
        txtCode.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
    '' Category Name
    adrsDept1.MoveFirst
    adrsDept1.Find "Desc='" & txtName.Text & "'"
    If Not adrsDept1.EOF Then
        MsgBox NewCaptionTxt("10018", adrsC), vbExclamation
        txtName.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If
'' Check for Blank Category Name
If Trim(txtName.Text) = "" Then
    MsgBox NewCaptionTxt("10019", adrsC), vbExclamation
    ValidateAddmaster = False
    txtName.SetFocus
    Exit Function
End If
Call FormatTODecimal            '' Get all the Text in the Equal 0.00 Format
If Not CheckDecimal(txtCLate) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtGEarly) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtEarlyA) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtLateG) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtCutL) Then
    ValidateAddmaster = False
    Exit Function
End If
If Not CheckDecimal(txtCutE) Then
    ValidateAddmaster = False
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
'' Check for Blank Category Name
If Trim(txtName.Text) = "" Then
    MsgBox NewCaptionTxt("10019", adrsC), vbExclamation
    ValidateModMaster = False
    txtName.SetFocus
    Exit Function
End If
Call FormatTODecimal            '' Get all the Text in the Equal 0.00 Format
If Not CheckDecimal(txtCLate) Then
    ValidateModMaster = False
    Exit Function
End If

If Not CheckDecimal(txtGEarly) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtEarlyA) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtLateG) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtCutL) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtCutE) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtFullDay) Then
    ValidateModMaster = False
    Exit Function
End If
If Not CheckDecimal(txtHalfDay) Then
    ValidateModMaster = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
    ValidateModMaster = False
End Function

Private Sub FormatTODecimal()
On Error GoTo ERR_P
txtCLate.Text = IIf(Trim(txtCLate.Text) = "", "0.00", Format(txtCLate.Text, "0.00"))
txtGEarly.Text = IIf(Trim(txtGEarly.Text) = "", "0.00", Format(txtGEarly.Text, "0.00"))
txtEarlyA.Text = IIf(Trim(txtEarlyA.Text) = "", "0.00", Format(txtEarlyA.Text, "0.00"))
txtLateG.Text = IIf(Trim(txtLateG.Text) = "", "0.00", Format(txtLateG.Text, "0.00"))
txtCutL.Text = IIf(Trim(txtCutL.Text) = "", "0.00", Format(txtCutL.Text, "0.00"))
txtCutE.Text = IIf(Trim(txtCutE.Text) = "", "0.00", Format(txtCutE.Text, "0.00"))
txtFullDay.Text = IIf(Trim(txtFullDay.Text) = "", "0.00", Format(txtFullDay.Text, "0.00"))  ' 27-05-09
txtHalfDay.Text = IIf(Trim(txtHalfDay.Text) = "", "0.00", Format(txtHalfDay.Text, "0.00"))
Exit Sub
ERR_P:
    ShowError ("FormatToDecimal :: " & Me.Caption)
End Sub

Private Function CheckDecimal(ByRef txt As Object) As Boolean
If Val(Right(txt.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod), vbExclamation
    txt.SetFocus
    CheckDecimal = False
    Exit Function
Else
    CheckDecimal = True
End If
If Val(txt.Text) > 23.59 Then
    MsgBox NewCaptionTxt("00025", adrsMod), vbExclamation
    txt.SetFocus
    CheckDecimal = False
Else
    CheckDecimal = True
End If
End Function

Private Sub GetRights()     '' Gets and Sets the Appropriate Rights
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
    Resume Next
End Sub

Private Sub cmdDel_Click()
On Error GoTo ERR_P
'' Check for Rights
If Not DeleteRights Then
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    Exit Sub
Else
    If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion) = vbYes Then        '' Delete the Record
        ConMain.Execute "delete from CatDesc where Cat='" & txtCode.Text & "'"
        Call AddActivityLog(lgDelete_Action, 1, 4)  '' Delete Log
        Call AuditInfo("DELETE", Me.Caption, "Category Deleted: " & txtCode.Text)
    End If
    Call FillGrid       '' Reflect the Grid
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
ERR_P:
    If Err.Number = -2147467259 Or Err.Number = -2147217900 Or Err.Number = -2147217873 Then
            MsgBox "Cannot delete, Dependent records are existing..", vbCritical, Me.Caption
            Exit Sub
    End If
    ShowError ("Delete Record :: " & Me.Caption)
End Sub

Private Sub cmdEditCan_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
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

Private Function SaveAddMaster() As Boolean
On Error GoTo ERR_P
SaveAddMaster = True        '' Insert
Dim strFld As String
strFld = "Cat," & strKDesc & ",Lt_Allow,Erl_Allow,Erl_Ignore,Lt_Ignore,HalfCutLt,HalfCutEr,  WeekOffPaid"

ConMain.Execute "insert into Catdesc(" & strFld & ") values('" & txtCode.Text & _
"','" & txtName.Text & "'," & txtCLate.Text & "," & txtGEarly.Text & "," & txtEarlyA.Text & _
"," & txtLateG.Text & "," & txtCutL.Text & "," & txtCutE.Text & "," & IIf(ChkWOffPaid.Value = 1, "'Y'", "'N'") & ")"

Exit Function
ERR_P:
    Select Case Err.Number
        Case -2147217900
            MsgBox NewCaptionTxt("10017", adrsC), vbExclamation
        Case Else
            ShowError ("SaveAddMaster :: " & Me.Caption)
    End Select
    'Resume Next
    SaveAddMaster = False
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
SaveModMaster = True        '' Update
ConMain.Execute "Update Catdesc Set " & strKDesc & "='" & txtName.Text & "',Lt_Allow=" & _
txtCLate.Text & ",Erl_Allow=" & txtGEarly.Text & ",Erl_Ignore=" & txtEarlyA.Text & _
",Lt_Ignore=" & txtLateG.Text & ",HalfCutLt=" & txtCutL.Text & ",HalfCutEr=" & txtCutE.Text & _
",  WeekOffPaid = '" & IIf(ChkWOffPaid.Value = 1, "Y", "N") & "',fstletpr= '',secletpr= '',trdletpr='', FullDayHr = " & txtFullDay.Text & ", HalfDayHr = " & txtHalfDay.Text & "  Where Cat='" & txtCode.Text & "'"     ''  26-12

Exit Function
ERR_P:
    SaveModMaster = False
    ShowError ("SaveModMaster :: " & Me.Caption)
End Function

Private Sub MSF1_DblClick()
    Call Display
End Sub

Private Sub txtCLate_GotFocus()
     Call GF(txtCLate)
End Sub

Private Sub txtCLate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtCLate)
End If
End Sub

Private Sub txtCode_GotFocus()
    Call GF(txtCode)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 6))))
End If
End Sub

Private Sub txtCode_LostFocus()
 If txtCode.Text = "100" Then
 MsgBox " This Category is reserved for Application"
 txtCode.Text = ""
 txtCode.SetFocus
 End If
End Sub

Private Sub txtCutE_GotFocus()
    Call GF(txtCutE)
End Sub

Private Sub txtCutE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtCutE)
End If
End Sub

Private Sub txtCutL_GotFocus()
    Call GF(txtCutL)
End Sub

Private Sub txtCutL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtCutL)
End If
End Sub

Private Sub txtEarlyA_GotFocus()
    Call GF(txtEarlyA)
End Sub

Private Sub txtEarlyA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtEarlyA)
End If
End Sub

Private Sub txtFullDay_GotFocus()
    Call GF(txtFullDay)
End Sub

Private Sub txtFullDay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtFullDay)
End If
End Sub

Private Sub txtGEarly_GotFocus()
    Call GF(txtGEarly)
End Sub

Private Sub txtGEarly_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtGEarly)
End If
End Sub

Private Sub txtHalfDay_GotFocus()
    Call GF(txtHalfDay)
End Sub

Private Sub txtHalfDay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtHalfDay)
End If
End Sub

Private Sub txtLateG_GotFocus()
    Call GF(txtLateG)
End Sub

Private Sub txtLateG_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtLateG)
End If
End Sub

Private Sub txtName_GotFocus()
    Call GF(txtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 1))))
End If
End Sub

Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo ERR_P
Call AddActivityLog(lgADD_MODE, 1, 4)     '' Add Activity
Call AuditInfo("ADD", Me.Caption, "Category Added: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 4)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Category Edited: " & txtCode.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub
