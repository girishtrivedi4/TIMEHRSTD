VERSION 5.00
Begin VB.Form frmTasker 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frTasker 
      Height          =   3795
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   5085
      Begin VB.CheckBox chkSch 
         Caption         =   "Reminder Disabled "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1110
         TabIndex        =   14
         Top             =   2850
         Width           =   2415
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2490
         TabIndex        =   13
         Top             =   3180
         Width           =   1635
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   840
         TabIndex        =   12
         Top             =   3180
         Width           =   1635
      End
      Begin VB.Frame Frame1 
         Caption         =   "On"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   840
         TabIndex        =   4
         Top             =   870
         Width           =   3285
         Begin VB.ComboBox cboWeek 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   690
            Width           =   615
         End
         Begin VB.ComboBox cboDay 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1380
            Width           =   615
         End
         Begin VB.OptionButton optSel 
            Caption         =   "Day"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   270
            TabIndex        =   9
            Top             =   1380
            Width           =   705
         End
         Begin VB.OptionButton optSel 
            Caption         =   "Weekly"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   270
            TabIndex        =   6
            Top             =   690
            Width           =   1035
         End
         Begin VB.OptionButton optSel 
            Caption         =   "Monthly(1st Day)"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   270
            TabIndex        =   8
            Top             =   1050
            Width           =   2055
         End
         Begin VB.OptionButton optSel 
            Caption         =   "Everyday"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   270
            TabIndex        =   5
            Top             =   360
            Width           =   1245
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "(1st Day)"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2070
            TabIndex        =   15
            Top             =   690
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Of Month"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1770
            TabIndex        =   11
            Top             =   1380
            Width           =   1065
         End
      End
      Begin VB.ComboBox cboTask 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   540
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Remind Me To Perform :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         TabIndex        =   2
         Top             =   510
         Width           =   3090
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Task Reminder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1590
         TabIndex        =   1
         Top             =   150
         Width           =   2145
      End
   End
End
Attribute VB_Name = "frmTasker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bytCboValue As Byte
Private bytCboSel As Byte

Private Sub cboDay_Click()
    bytCboValue = Val(cboDay.Text)
End Sub

Private Sub cboWeek_Click()
    bytCboValue = Val(Left(cboWeek.Text, 1))
End Sub

Private Sub chkSch_Click()
If chkSch.Value = 1 Then
    chkSch.Caption = "Reminder Enabled "
    Call SaveSetting(App.EXEName, "Reminder", "On Or Off", chkSch.Value)
Else
    chkSch.Caption = "Reminder Disabled"
    Call SaveSetting(App.EXEName, "Reminder", "On Or Off", chkSch.Value)
End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ERR_P
Select Case cboTask.ListIndex
    Case 0      '' Tasknum 1 Daily Process
        Call AddActivityLog(lgRDaily_Action, 1, 32)     '' Daily Log
        VstarDataEnv.cnDJConn.Execute "delete from tasker where tasknum = 1 "
        VstarDataEnv.cnDJConn.Execute "insert into tasker values('" & cboTask.ListIndex + 1 & "'," & _
        bytCboSel & "," & bytCboValue & ")"
    Case 1      '' Tasknum 2 Monthly Process
        Call AddActivityLog(lgRMonthly_Action, 1, 32)   '' Monthly Log
        VstarDataEnv.cnDJConn.Execute "delete from tasker where tasknum = 2 "
        VstarDataEnv.cnDJConn.Execute "insert into tasker values('" & cboTask.ListIndex + 1 & "'," & _
        bytCboSel & "," & bytCboValue & ")"
    Case 2      '' Tasknum 3 Reports
        Call AddActivityLog(lgRReports_Action, 1, 32)   '' Reports Log
        VstarDataEnv.cnDJConn.Execute "delete from tasker where tasknum = 3 "
        VstarDataEnv.cnDJConn.Execute "insert into tasker values('" & cboTask.ListIndex + 1 & "'," & _
        bytCboSel & "," & bytCboValue & ")"
End Select
MsgBox "Saved Successfully"
Exit Sub
ERR_P:
    ShowError ("Save :: " & Me.Caption)
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Dim bytCnt As Byte
'' Add Types of Reminders
cboTask.AddItem "Daily Process"
cboTask.AddItem "Monthly Process"
cboTask.AddItem "Reports"
'' Add Weeks
cboWeek.AddItem "1st"
cboWeek.AddItem "2nd"
cboWeek.AddItem "3rd"
cboWeek.AddItem "4th"
cboWeek.AddItem "5th"
'' Add Days
For bytCnt = 1 To 31
    cboDay.AddItem bytCnt
Next
cboTask.ListIndex = 0
'' Get Settings for E/D Reminder
If GetSetting(App.EXEName, "Reminder", "On Or Off", 0) = 0 Then
    chkSch.Value = 0        '' Disable (Default)
Else
    chkSch.Value = 1        '' Enable
End If
Exit Sub
ERR_P:
    ShowError ("Load :: Reminder")
End Sub

Private Sub optSel_Click(Index As Integer)
bytCboSel = Index + 1
Select Case Index
    Case 0, 2
        cboDay.Enabled = False
        cboWeek.Enabled = False
    Case 1
        cboDay.Enabled = False
        cboWeek.Enabled = True
    Case 3
        cboDay.Enabled = True
        cboWeek.Enabled = False
End Select
End Sub
