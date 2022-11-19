VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmWeekOfSelection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Week Of Selection in month"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5490
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   3975
      Left            =   120
      TabIndex        =   49
      Top             =   120
      Width           =   5295
      Begin VB.Frame Frame3 
         Height          =   2415
         Left            =   120
         TabIndex        =   52
         Top             =   960
         Width           =   5055
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshWeekOff 
            Height          =   2055
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   3625
            _Version        =   393216
            Rows            =   6
            Cols            =   7
            _NumberOfBands  =   1
            _Band(0).Cols   =   7
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   5055
         Begin MSForms.ComboBox cboCode 
            Height          =   315
            Left            =   1200
            TabIndex        =   56
            Top             =   240
            Width           =   3795
            VariousPropertyBits=   612390939
            MaxLength       =   10
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "6694;556"
            TextColumn      =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            Caption         =   "Enter Id"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   240
            Width           =   855
         End
      End
      Begin LVbuttons.LaVolpeButton cmdOk 
         Height          =   405
         Left            =   2880
         TabIndex        =   54
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "&Ok"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   12648447
         FCOL            =   4210752
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmWeekOffSelection.frx":0000
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   405
         Left            =   4080
         TabIndex        =   55
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   12648447
         FCOL            =   4210752
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmWeekOffSelection.frx":001C
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmdDelete 
         Height          =   405
         Left            =   120
         TabIndex        =   57
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "&Delete"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   12648447
         FCOL            =   4210752
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmWeekOffSelection.frx":0038
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmdClear 
         Height          =   405
         Left            =   1320
         TabIndex        =   58
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "C&lear"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   12648447
         FCOL            =   4210752
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmWeekOffSelection.frx":0054
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   1080
      TabIndex        =   0
      Top             =   5280
      Width           =   5895
      Begin VB.CheckBox chkFifth 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   6
         Left            =   5400
         TabIndex        =   48
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkFifth 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   5
         Left            =   4800
         TabIndex        =   47
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkFifth 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   4200
         TabIndex        =   46
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkFifth 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   3600
         TabIndex        =   45
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkFifth 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   3000
         TabIndex        =   44
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkFifth 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   2400
         TabIndex        =   43
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkFifth 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   1800
         TabIndex        =   42
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkFourth 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   6
         Left            =   5400
         TabIndex        =   41
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox chkFourth 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   5
         Left            =   4800
         TabIndex        =   40
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox chkFourth 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   4200
         TabIndex        =   39
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox chkFourth 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   3600
         TabIndex        =   38
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox chkFourth 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   3000
         TabIndex        =   37
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox chkFourth 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   2400
         TabIndex        =   36
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox chkFourth 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   1800
         TabIndex        =   35
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox chkThird 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   6
         Left            =   5400
         TabIndex        =   34
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkThird 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   5
         Left            =   4800
         TabIndex        =   33
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkThird 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   4200
         TabIndex        =   32
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkThird 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   3600
         TabIndex        =   31
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkThird 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   3000
         TabIndex        =   30
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkThird 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   2400
         TabIndex        =   29
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkThird 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   1800
         TabIndex        =   28
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkSecond 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   6
         Left            =   5400
         TabIndex        =   27
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chkSecond 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   5
         Left            =   4800
         TabIndex        =   26
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chkSecond 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   4200
         TabIndex        =   25
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chkSecond 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   3600
         TabIndex        =   24
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chkSecond 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   3000
         TabIndex        =   23
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chkSecond 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   2400
         TabIndex        =   22
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chkSecond 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   1800
         TabIndex        =   21
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chkFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   6
         Left            =   5400
         TabIndex        =   20
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   5
         Left            =   4800
         TabIndex        =   19
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   4200
         TabIndex        =   18
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   3600
         TabIndex        =   17
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   3000
         TabIndex        =   16
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   2400
         TabIndex        =   15
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   1800
         TabIndex        =   14
         Top             =   720
         Width           =   255
      End
      Begin VB.Line Line11 
         X1              =   1560
         X2              =   5760
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line10 
         X1              =   1560
         X2              =   5760
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line9 
         X1              =   1440
         X2              =   5640
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line8 
         X1              =   1440
         X2              =   5640
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line7 
         X1              =   5160
         X2              =   5160
         Y1              =   600
         Y2              =   2400
      End
      Begin VB.Line Line6 
         X1              =   4560
         X2              =   4560
         Y1              =   600
         Y2              =   2400
      End
      Begin VB.Line Line5 
         X1              =   3960
         X2              =   3960
         Y1              =   600
         Y2              =   2400
      End
      Begin VB.Line Line4 
         X1              =   3360
         X2              =   3360
         Y1              =   600
         Y2              =   2400
      End
      Begin VB.Line Line3 
         X1              =   2760
         X2              =   2760
         Y1              =   600
         Y2              =   2400
      End
      Begin VB.Line Line1 
         X1              =   2160
         X2              =   2160
         Y1              =   600
         Y2              =   2400
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   1815
         Left            =   1560
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Selection"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sun"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5160
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wed"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sat"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4560
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fri"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Thu"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tue"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mon"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fifth Week"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fourth Week"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Week"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Second Week"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Third Week"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmWeekOfSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCode_Change()
    Call cboCode_Click
End Sub

Private Sub cboCode_Click()
'
    Dim adrsTemp As Recordset
    Dim mRow As Integer
   On Error GoTo cboCode_Click_Error

    Set adrsTemp = OpenRecordSet("SELECT * FROM WeekOff WHERE code='" & _
        cbocode.Text & "'")
'    If Not (adrsTemp.EOF And adrsTemp.BOF) Then
'        'mshWeekOff.Clear
'        Exit Sub
'    End If
    Do While Not adrsTemp.EOF
        For mRow = 1 To 6
            With mshWeekOff
                .TextMatrix(mRow, 1) = adrsTemp.Fields("sun")
                .TextMatrix(mRow, 2) = adrsTemp.Fields("mon")
                .TextMatrix(mRow, 3) = adrsTemp.Fields("the")
                .TextMatrix(mRow, 4) = adrsTemp.Fields("wed")
                .TextMatrix(mRow, 5) = adrsTemp.Fields("thu")
                .TextMatrix(mRow, 6) = adrsTemp.Fields("fri")
                .TextMatrix(mRow, 7) = adrsTemp.Fields("sat")
            End With
            adrsTemp.MoveNext
        Next
    Loop

   On Error GoTo 0
   Exit Sub

cboCode_Click_Error:

    ShowError "Error in procedure cboCode_Click of Form frmWeekOfSelection"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim mRow As Integer
    Dim mColumn As Integer
    For mRow = 1 To 6
        For mColumn = 1 To 7
            mshWeekOff.TextMatrix(mRow, mColumn) = strUnChecked
        Next
    Next
End Sub

Private Sub cmdDelete_Click()
   On Error GoTo cmdDelete_Click_Error

    If MsgBox("Are you sure delete a selected item", vbInformation + vbYesNo) Then
        VstarDataEnv.cnDJConn.Execute "DELETE FROM WeekOff WHERE code='" & _
        cbocode.Text & "'"
        MsgBox "Deleted", vbInformation
    End If
    Call FillWOCombo(cbocode)
   On Error GoTo 0
   Exit Sub

cmdDelete_Click_Error:

    ShowError "Error in procedure cmdDelete_Click of Form frmWeekOfSelection"
End Sub

Private Sub cmdOK_Click()
    Dim mRow As Integer
    Dim mCol As Integer
    Dim UpdateInsertFlag As Boolean
   On Error GoTo cmdOK_Click_Error
    
    If Trim(cbocode.Text) = "" Then
        MsgBox "Enter Proper WeekOF Selection Id", vbExclamation, "Error Saving ......."
        cbocode.SetFocus
        Exit Sub
    End If

    With mshWeekOff
        UpdateInsertFlag = FindCode("code", "WeekOFF", cbocode.Text)
        For mRow = 1 To 6
            If Not UpdateInsertFlag Then
                VstarDataEnv.cnDJConn.Execute "INSERT INTO WeekOff(code, " & _
                "weekno,sun,mon,the,wed,thu,fri,sat) VALUES('" & cbocode.Text & _
                "','" & .TextMatrix(mRow, 0) & "','" & .TextMatrix(mRow, 1) & _
                "','" & .TextMatrix(mRow, 2) & "','" & .TextMatrix(mRow, 3) & _
                "','" & .TextMatrix(mRow, 4) & "','" & .TextMatrix(mRow, 5) & _
                "','" & .TextMatrix(mRow, 6) & "','" & .TextMatrix(mRow, 7) & "')"
            Else
                VstarDataEnv.cnDJConn.Execute "UPDATE WeekOff SET code='" & cbocode.Text & _
                "',weekno='" & .TextMatrix(mRow, 0) & "',sun='" & .TextMatrix(mRow, 1) & _
                "',mon='" & .TextMatrix(mRow, 2) & "',the='" & .TextMatrix(mRow, 3) & _
                "',wed='" & .TextMatrix(mRow, 4) & "',thu='" & .TextMatrix(mRow, 5) & _
                "',fri='" & .TextMatrix(mRow, 6) & "',sat='" & .TextMatrix(mRow, 7) & _
                "' WHERE code='" & cbocode.Text & "' AND weekno='" & .TextMatrix(mRow, 0) & "'"
            End If
        Next
    End With
    If Not UpdateInsertFlag Then Call FillWOCombo(cbocode)
    MsgBox "Transaction Completed", vbInformation
   On Error GoTo 0
   Exit Sub

cmdOK_Click_Error:

    ShowError "Error in procedure cmdOK_Click of Form frmWeekOfSelection"
End Sub

Private Function FindCode(strColumnName As String, _
    strTableName As String, strValue) As Boolean
    
    Dim adrsT As Recordset
    
   On Error GoTo FindCode_Error

    Set adrsT = OpenRecordSet("SELECT COUNT(" & strColumnName & _
    ") FROM " & strTableName & " WHERE " & strColumnName & " = '" & strValue & "'")
    
    If Not (adrsT.EOF And adrsT.BOF) Then
        If adrsT.Fields(0) > 0 Then
            FindCode = True
        Else
            FindCode = False
        End If
    End If

   On Error GoTo 0
   Exit Function

FindCode_Error:

    ShowError "Error in procedure FindCode of Form frmWeekOfSelection"
End Function

'---------------------------------------------------------------------------------------
' Procedure : FillCombo
' DateTime  : 26/06/2008 10:16
' Author    :
' Purpose   : Fill Combo From Database
' Pre       :
' Post      :
' Return    :
'---------------------------------------------------------------------------------------
'
Public Sub FillWOCombo(ByRef cmbI As MSForms.ComboBox)
    Dim adrsTemp As Recordset
On Error GoTo FillCombo_Error
    cmbI.Clear
    Set adrsTemp = OpenRecordSet("SELECT DISTINCT Code FROM WeekOff")
    Do While Not adrsTemp.EOF
        cmbI.AddItem adrsTemp.Fields("Code")
        adrsTemp.MoveNext
    Loop
    cmbI.AddItem ""
On Error GoTo 0
Exit Sub
FillCombo_Error:
   If Erl = 0 Then
      ShowError "Error in procedure FillCombo of Form frmWeekOfSelection"
   Else
      ShowError "Error in procedure FillCombo of Form frmWeekOfSelection And Line:" & Erl
   End If
End Sub
Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    Call SetFormIcon(Me)
    Call SetGrid(mshWeekOff)
    Call FillWOCombo(cbocode)

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    ShowError "Error in procedure Form_Load of Form frmWeekOfSelection"
End Sub

Private Sub SetGrid(mhFlex As MSHFlexGrid)
   On Error GoTo SetGrid_Error

    With mhFlex
       .Cols = 8
       .TextMatrix(0, 1) = "Sun"
       .TextMatrix(0, 2) = "Mon"
       .TextMatrix(0, 3) = "Tue"
       .TextMatrix(0, 4) = "Wed"
       .TextMatrix(0, 5) = "Thu"
       .TextMatrix(0, 6) = "Fri"
       .TextMatrix(0, 7) = "Sat"
       .Rows = 7
       .TextMatrix(1, 0) = "First Week"
       .TextMatrix(2, 0) = "Second Week"
       .TextMatrix(3, 0) = "Third Week"
       .TextMatrix(4, 0) = "Fourth Week"
       .TextMatrix(5, 0) = "Fifth Week"
       .TextMatrix(6, 0) = "Six Week"
       .ColWidth(1) = 500
       .ColWidth(2) = 500
       .ColWidth(3) = 500
       .ColWidth(4) = 500
       .ColWidth(5) = 500
       .ColWidth(6) = 500
       .ColWidth(7) = 500
        Dim mRow As Integer
        Dim mColumn As Integer
        For mRow = 1 To 6
            For mColumn = 1 To 7
                .row = mRow
                .Col = mColumn
                .CellFontName = "Wingdings"
                .CellFontSize = 14
                .CellAlignment = flexAlignCenterCenter
                .TextMatrix(mRow, mColumn) = strUnChecked
            Next
        Next
    End With

   On Error GoTo 0
   Exit Sub

SetGrid_Error:

    ShowError "Error in procedure SetGrid of Form frmWeekOfSelection"
End Sub

Private Sub mshWeekOff_KeyPress(KeyAscii As Integer)
   On Error GoTo mshWeekOff_KeyPress_Error

If KeyAscii = 13 Or KeyAscii = 32 Then
    With mshWeekOff
        If Not (.row = 0 And .Col = 0) Then
            Call TriggerCheckbox(.row, .Col)
        End If
    End With
End If

   On Error GoTo 0
   Exit Sub

mshWeekOff_KeyPress_Error:

    ShowError "Error in procedure mshWeekOff_KeyPress of Form frmWeekOfSelection"
End Sub



Private Sub TriggerCheckbox(iRow As Integer, iCol As Integer)
   On Error GoTo TriggerCheckbox_Error

    With mshWeekOff
        If .TextMatrix(iRow, iCol) = strUnChecked Then
            .TextMatrix(iRow, iCol) = strChecked
        Else
            .TextMatrix(iRow, iCol) = strUnChecked
        End If
    End With

   On Error GoTo 0
   Exit Sub

TriggerCheckbox_Error:

    ShowError "Error in procedure TriggerCheckbox of Form frmWeekOfSelection"
End Sub

Private Sub mshWeekOff_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo mshWeekOff_MouseUp_Error

    If Button = 1 Then
        With mshWeekOff
            'MsgBox .MouseRow & .MouseCol
            If Not (.MouseCol = 0) Then
                If Not (.MouseRow = 0) Then
                    Call TriggerCheckbox(.MouseRow, .MouseCol)
                End If
            End If
        End With
    End If

   On Error GoTo 0
   Exit Sub

mshWeekOff_MouseUp_Error:

    ShowError "Error in procedure mshWeekOff_MouseUp of Form frmWeekOfSelection"
End Sub
