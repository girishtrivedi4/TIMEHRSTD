VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frData 
      Height          =   6735
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   11415
      Begin VB.OptionButton opt 
         Caption         =   "Dept"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "Cat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   2040
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   2400
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   2760
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "Division"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   3120
         Width           =   1575
      End
      Begin VB.OptionButton opt 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   3480
         Width           =   1575
      End
      Begin VB.ListBox lst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   0
         Left            =   2040
         TabIndex        =   7
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ListBox lst2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   0
         Left            =   5760
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddD 
         Caption         =   " &Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   14
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   8
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "Re&move All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   10
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddAll 
         Caption         =   "Ad&d All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   9
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ListBox lst2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   1
         Left            =   8400
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin VB.ListBox lst2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   2
         Left            =   8760
         TabIndex        =   23
         Top             =   600
         Width           =   1695
      End
      Begin VB.ListBox lst2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   3
         Left            =   9120
         TabIndex        =   22
         Top             =   600
         Width           =   1695
      End
      Begin VB.ListBox lst2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   4
         Left            =   9480
         TabIndex        =   21
         Top             =   600
         Width           =   1695
      End
      Begin VB.ListBox lst2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   5
         Left            =   9840
         TabIndex        =   20
         Top             =   600
         Width           =   1695
      End
      Begin VB.ListBox lst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   1
         Left            =   0
         TabIndex        =   19
         Top             =   4920
         Width           =   1695
      End
      Begin VB.ListBox lst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   2
         Left            =   360
         TabIndex        =   18
         Top             =   4920
         Width           =   1695
      End
      Begin VB.ListBox lst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   3
         Left            =   720
         TabIndex        =   17
         Top             =   4920
         Width           =   1695
      End
      Begin VB.ListBox lst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   4
         Left            =   1080
         TabIndex        =   16
         Top             =   4920
         Width           =   1695
      End
      Begin VB.ListBox lst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Index           =   5
         Left            =   1440
         TabIndex        =   15
         Top             =   4920
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bytcnt As Byte, bytRec As Byte
Dim bytIndex As Byte, strMaster(6, 1) As String
Dim Rs As New ADODB.Recordset

Private Sub Form_Load()
Call FillArray
For bytcnt = 0 To lst.Count - 1
    Call FillList(bytcnt)
Next
End Sub

Private Sub FillArray()
strMaster(0, 0) = "Deptdesc": strMaster(0, 1) = "Dept,[Desc]"
strMaster(1, 0) = "Catdesc": strMaster(1, 1) = "Cat,[desc]"
strMaster(2, 0) = "Company": strMaster(2, 1) = "Company,cname"
strMaster(3, 0) = "Groupmst": strMaster(3, 1) = "Group,grupdesc"
strMaster(4, 0) = "Division": strMaster(4, 1) = "Div,divdesc"
strMaster(5, 0) = "Location": strMaster(5, 1) = "Location,LocDesc"
End Sub

Private Sub FillList(bytTmp As Byte)
If Rs.State = 1 Then Rs.Close
Rs.Open "select " & strMaster(bytTmp, 1) & " from " & strMaster(bytTmp, 0), VstarDataEnv.cnDJConn, adOpenStatic, adLockOptimistic
If Not (Rs.EOF And Rs.BOF) Then
    Do While Not Rs.EOF
        lst(bytTmp).AddItem Rs(0) & ":: " & Rs(1)
        Rs.MoveNext
    Loop
End If
End Sub

Private Sub opt_Click(Index As Integer)
bytIndex = Index
For bytcnt = 0 To lst.Count - 1
    lst(bytcnt).Visible = False
    lst2(bytcnt).Visible = False
Next
lst(Index).Left = lst(0).Left
lst(Index).Top = lst(0).Top
lst(Index).Visible = True

lst2(Index).Left = lst2(0).Left
lst2(Index).Top = lst2(0).Top
lst2(Index).Visible = True
End Sub

Private Sub cmdAddD_Click()
If Trim(lst(bytIndex).Text) <> "" Then
    lst2(bytIndex).AddItem lst(bytIndex).Text
End If
If lst(bytIndex).ListIndex >= 0 Then
    lst(bytIndex).RemoveItem lst(bytIndex).ListIndex
End If
End Sub

Private Sub cmdAddAll_Click()
If lst(bytIndex).ListCount > 0 Then
    For bytcnt = 0 To lst(bytIndex).ListCount - 1
        lst2(bytIndex).AddItem lst(bytIndex).List(bytcnt)
    Next
    lst(bytIndex).Clear
End If
End Sub

Private Sub cmdRemove_Click()
If Trim(lst2(bytIndex).Text) <> "" Then
    lst(bytIndex).AddItem lst2(bytIndex).Text
End If
If lst2(bytIndex).ListIndex >= 0 Then
    lst2(bytIndex).RemoveItem lst2(bytIndex).ListIndex
End If
End Sub

Private Sub cmdRemoveAll_Click()
If lst2(bytIndex).ListCount > 0 Then
    For bytcnt = 0 To lst2(bytIndex).ListCount - 1
        lst(bytIndex).AddItem lst2(bytIndex).List(bytcnt)
    Next
    lst2(bytIndex).Clear
End If
End Sub

Private Sub cmdSave_Click()
Dim strFrom As String, strWhere As String
For bytcnt = 0 To lst2.Count - 1
    If lst(bytcnt).ListCount = 0 Or lst2(bytcnt).ListCount = 0 Then
        ''if all selected or no selected
    Else
        strFrom = strFrom & "," & opt(bytcnt).Caption
        strWhere = strWhere & opt(bytcnt).Caption & " in( "
        For bytRec = 0 To lst2(bytcnt).ListCount - 1
            strWhere = strWhere & IIf(bytcnt = 1, "'", "")
            strWhere = strWhere & Left(lst2(bytcnt).List(bytRec), InStr(1, lst2(bytcnt).List(bytRec), ":") - 1)
            strWhere = strWhere & IIf(bytcnt = 1, "'", "")
            strWhere = strWhere & ","
        Next
        strWhere = Left(strWhere, Len(strWhere) - 1) & ") And "
    End If
Next
If Len(Trim(strWhere)) > 0 Then strWhere = " Where " & Left(strWhere, Len(strWhere) - 4)
MsgBox strFrom & " " & strWhere
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

