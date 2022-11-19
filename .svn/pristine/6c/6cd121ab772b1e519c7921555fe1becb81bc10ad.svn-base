VERSION 5.00
Begin VB.Form YRCRFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Yearly file creation"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton FinishCmd 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   3090
      Width           =   1215
   End
   Begin VB.CommandButton CreateCmd 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   3090
      Width           =   1335
   End
   Begin VB.ComboBox YrCombo 
      Height          =   345
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2610
      Width           =   1905
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2505
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4125
      Begin VB.Label YrLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   225
         Index           =   4
         Left            =   540
         TabIndex        =   8
         Top             =   1800
         Width           =   570
      End
      Begin VB.Label YrLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   225
         Index           =   3
         Left            =   540
         TabIndex        =   7
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label YrLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   225
         Index           =   2
         Left            =   540
         TabIndex        =   6
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label YrLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   570
      End
      Begin VB.Label YrLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yearly File Creation for the Slected Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3435
      End
   End
   Begin VB.Label YrLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   5
      Left            =   420
      TabIndex        =   9
      Top             =   2670
      Width           =   645
   End
End
Attribute VB_Name = "YRCRFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LvFileName$
Private ShiftcreateYr As String
Dim adrsC As New ADODB.Recordset    ''L

 
Private Sub Form_Activate()
    YrCombo.Text = Year(Year_Start)
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Call SetYearlyFlag(1)
Call SetFormIcon(Me)
Call RetCaptions
' Populate the Year Combo Box
Dim YrLen%
For YrLen% = 0 To 99
    YrCombo.AddItem (1997 + YrLen%)
Next YrLen%
'' For Rights
Dim strTmp As String
strTmp = RetRights(3, 2, , 1)
If strTmp = "1" Then
    CreateCmd.Enabled = True
Else
    CreateCmd.Enabled = False
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
End If
YrLen = 0
Exit Sub
ERR_P:
    ShowError ("Load :: " & Me.Caption)
End Sub
Private Sub RetCaptions()
'On Error resume next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '58%'", ConMain, adOpenStatic, adLockReadOnly
Me.Caption = NewCaptionTxt("58001", adrsC)
YrLbl(0).Caption = NewCaptionTxt("58002", adrsC)
YrLbl(1).Caption = NewCaptionTxt("58003", adrsC)
YrLbl(2).Caption = NewCaptionTxt("58004", adrsC)
YrLbl(3).Caption = NewCaptionTxt("58005", adrsC)
YrLbl(4).Caption = NewCaptionTxt("58006", adrsC)
YrLbl(5).Caption = NewCaptionTxt("00027", adrsMod)
Frame1.Caption = NewCaptionTxt("58007", adrsC)
CreateCmd.Caption = "Create" 'NewCaptionTxt("58008", adrsC)
FinishCmd.Caption = "Finish" 'NewCaptionTxt("00039", adrsMod)
End Sub

Private Sub CreateCmd_Click()
On Error GoTo Err_particular
Dim StrFieldsN As String
ShiftcreateYr = Trim(YrCombo.Text)
Dim MsgRet%, j%
Dim n%
n% = 1
Call AddActivityLog(lg_NoModeAction, 2, 6)      '' Leave Create Activity Log
Call AuditInfo("YEARLY LEAVE CREATION", Me.Caption, "Created Yearly Leave File For The Year: " & YrCombo.Text)
Call UpdateLvFile
Do
    Select Case n%
        Case 1
            LvFileName$ = MakeYrName("LVTRN", ShiftcreateYr)
            If FindTable(LvFileName$) Then   ' When Table found
                MsgRet = MsgBox(NewCaptionTxt("58009", adrsC) & LvFileName$ & vbCrLf & _
                NewCaptionTxt("58010", adrsC), vbYesNo + vbQuestion)

                If MsgRet = 6 Then        ' Overwrite
                If FindTable("oldleavtrn") Then ConMain.Execute "drop table oldleavtrn"
                    ''conmain.Execute "select * into oldLeavtrn from " & LvFileName
                    Call CreateTableIntoAs("*", LvFileName, "oldLeavtrn")
                    ConMain.Execute "drop table " & LvFileName
                    ''conmain.Execute "select * into " & LvFileName & " from leavtrn"
                    Call CreateTableIntoAs("*", "leavtrn", LvFileName)
                    If adrsTemp.State = 1 Then adrsTemp.Close
                    StrFieldsN = ""
                    If adrsTemp.State = 1 Then adrsTemp.Close
                    adrsTemp.Open "Select * from " & LvFileName, ConMain
                    For j = 0 To adrsTemp.Fields.Count - 1
                        If FieldExists("oldleavtrn", adrsTemp.Fields(j).name) Then
                            If j = adrsTemp.Fields.Count - 1 Then
                                StrFieldsN = StrFieldsN & adrsTemp.Fields(j).name
                            Else
                                StrFieldsN = StrFieldsN & adrsTemp.Fields(j).name & ","
                            End If
                        End If
                    Next
                    If Right(StrFieldsN, 1) = "," Then StrFieldsN = Left(StrFieldsN, Len(StrFieldsN) - 1)
                    ConMain.Execute "insert into " & _
                    LvFileName & "(" & StrFieldsN & ") Select " & StrFieldsN & _
                    " from oldleavtrn"
                    ' Changes to add 0
                    If adrsTemp.State = 1 Then adrsTemp.Close
                    adrsTemp.Open "Select * from " & LvFileName, ConMain
                    For j = 0 To adrsTemp.Fields.Count - 1
                        If FieldExists("oldleavtrn", adrsTemp.Fields(j).name) = False Then
                            ConMain.Execute "Update " & LvFileName & " Set " & _
                            adrsTemp.Fields(j).name & "=0"
                        End If
                    Next
                    ' End Changes
                    Call CreateTableIndexAs("LVTRNYY", , Right(LvFileName, 2))
                    If FindTable("oldleavtrn") Then ConMain.Execute "drop table oldleavtrn"
                End If     ' Overwrite
            Else   ' LvTrn not found
                ''conmain.Execute "select * into " & LvFileName & " from leavtrn"
                Call CreateTableIntoAs("*", "leavtrn", LvFileName)
                Call CreateTableIndexAs("LVTRNYY", , Right(LvFileName, 2))
            End If     ' FindTable
            
        Case 2:
            LvFileName$ = MakeYrName("LVINFO", ShiftcreateYr)
            If FindTable(LvFileName$) Then  ' When Table  found
                MsgRet = MsgBox(NewCaptionTxt("58009", adrsC) & LvFileName$ & vbCrLf & _
                NewCaptionTxt("58010", adrsC), vbYesNo + vbQuestion, "Warning")
                If MsgRet = 6 Then ' Overwrite
                    If FindTable("oldLeavInfo") Then ConMain.Execute "drop table oldLeavInfo"
                    ''conmain.Execute "select * into oldLeavInfo from " & LvFileName
                    Call CreateTableIntoAs("*", LvFileName, "oldLeavInfo")
                    ConMain.Execute "drop table " & LvFileName
                    ''conmain.Execute "select * into " & LvFileName & " from leavinfo"
                    Call CreateTableIntoAs("*", "leavinfo", LvFileName)
                    If adrsTemp.State = 1 Then adrsTemp.Close
                    StrFieldsN = ""
                    If adrsTemp.State = 1 Then adrsTemp.Close
                    adrsTemp.Open "Select * from " & LvFileName, ConMain
                    For j = 0 To adrsTemp.Fields.Count - 1
                        If FieldExists("oldLeavInfo", adrsTemp.Fields(j).name) Then
                            If j = adrsTemp.Fields.Count - 1 Then
                                StrFieldsN = StrFieldsN & adrsTemp.Fields(j).name
                            Else
                                StrFieldsN = StrFieldsN & adrsTemp.Fields(j).name & ","
                            End If
                        End If
                    Next
                    If Right(StrFieldsN, 1) = "," Then StrFieldsN = Left(StrFieldsN, Len(StrFieldsN) - 1)
                    ConMain.Execute "insert into " & _
                    LvFileName & "(" & StrFieldsN & ") Select " & StrFieldsN & _
                    " from oldleavinfo"
                    Call CreateTableIndexAs("LVINFOYY", , Right(LvFileName, 2))
                End If
                If adrsTemp.State = 1 Then adrsTemp.Close
                If FindTable("oldLeavInfo") Then ConMain.Execute "drop table oldLeavInfo"
            Else
                ''conmain.Execute "select * into " & LvFileName & " from leavinfo"
                Call CreateTableIntoAs("*", "leavinfo", LvFileName)
                Call CreateTableIndexAs("LVINFOYY", , Right(LvFileName, 2))
            End If  ' Find Table
        Case 3:
            LvFileName$ = MakeYrName("LVBAL", ShiftcreateYr)
            If FindTable(LvFileName$) Then       ' Table found
                MsgRet = MsgBox(NewCaptionTxt("58009", adrsC) & LvFileName$ & vbCrLf & _
                NewCaptionTxt("58010", adrsC), vbYesNo + vbQuestion, "Warning")
                If MsgRet = 6 Then     ' Overwrite
                    If FindTable("oldLvBal") Then ConMain.Execute "drop table oldLvBal"
                    ''conmain.Execute "select * into oldLvBal from " & LvFileName
                    Call CreateTableIntoAs("*", LvFileName, "oldLvBal")
                    ConMain.Execute "drop table " & LvFileName
                    ''conmain.Execute "select * into " & LvFileName & " from leavbal"
                    Call CreateTableIntoAs("*", "leavbal", LvFileName)
                    If adrsTemp.State = 1 Then adrsTemp.Close
                    StrFieldsN = ""
                    If adrsTemp.State = 1 Then adrsTemp.Close
                    adrsTemp.Open "Select * from " & LvFileName, ConMain
                    For j = 0 To adrsTemp.Fields.Count - 1
                        If FieldExists("oldLvBal", adrsTemp.Fields(j).name) Then
                            If j = adrsTemp.Fields.Count - 1 Then
                                StrFieldsN = StrFieldsN & adrsTemp.Fields(j).name
                            Else
                                StrFieldsN = StrFieldsN & adrsTemp.Fields(j).name & ","
                            End If
                        End If
                    Next
                    If Right(StrFieldsN, 1) = "," Then StrFieldsN = Left(StrFieldsN, Len(StrFieldsN) - 1)
                    ''conmain.Execute "ALTER TABLE " & LvFileName & "  ADD  constraint  fkey  FOREIGN KEY (EMPCODE) REFERENCES   EMPMST (Empcode)"
                    ConMain.Execute "insert into " & _
                    LvFileName & "(" & StrFieldsN & ") Select " & StrFieldsN & _
                    " from oldlvbal"
                    Call CreateTableIndexAs("LVBALYY", , Right(LvFileName, 2))
                End If
                If adrsTemp.State = 1 Then adrsTemp.Close
                adrsTemp.Open "Select * from " & LvFileName, ConMain
                If adrsTemp.EOF And adrsTemp.BOF Then
                    ConMain.Execute "insert into " & LvFileName & "(Empcode) select distinct(Empcode) from empmst"
                End If
                If FindTable("oldLvBal") Then ConMain.Execute "drop table oldLvBal"
            Else
                ''conmain.Execute "select * into " & LvFileName & " from leavbal"
                Call CreateTableIntoAs("*", "leavbal", LvFileName)
                Call CreateTableIndexAs("LVBALYY", , Right(LvFileName, 2))
                ''SG07
                ''conmain.Execute "ALTER TABLE " & LvFileName & "  ADD  constraint  fkey  FOREIGN KEY (EMPCODE) REFERENCES  EMPMST (Empcode)"
                ConMain.Execute "insert into " & LvFileName & "(Empcode) select distinct(Empcode) from empmst"
                
        End If     ' Find Table
    End Select
    n% = n% + 1
Loop Until n% > 3

MsgBox NewCaptionTxt("58011", adrsC), vbQuestion, App.EXEName
YrCombo.SetFocus
MsgRet = 0
n = 0
Exit Sub
Err_particular:
    If Err.Number = -2147217900 Then Resume Next
    ShowError ("Create Leaves :: " & Me.Caption)
    'Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetYearlyFlag(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LvFileName = ""
End Sub

Private Function MakeYrName(ByVal strA$, ByVal Yr$) As String
    MakeYrName = Trim(strA$) & Mid(Yr$, 3, 2)
End Function

Private Sub UpdateLvFile()
      Dim FlName$
10    On Error GoTo UpdateLvFile_Error
20    If adrsTemp.State = 1 Then adrsTemp.Close
30    adrsTemp.Open "select distinct(lvcode),leave,type from Leavdesc where lvcode not in (" & "'" & _
      pVStar.PrsCode & "'" & "," & "'" & pVStar.AbsCode & "'" & "," & "'" & pVStar.WosCode & "'" & "," & "'" & _
      pVStar.HlsCode & "'" & ")", ConMain
      ' Copy the Structure of Old LeavBal
'40    conmain.BeginTrans
'40        transaction = conmain.BeginTransaction()
50    If FindTable("leavold") Then ConMain.Execute "drop table leavold"
60    If FindTable("LvbTmp") Then ConMain.Execute "drop table LvbTmp"
'70      conmain.CommitTrans

      ''conmain.Execute "select Empcode into Leavold from Leavbal"
80    Call CreateTableIntoAs("Empcode", "Leavbal", "Leavold")
      ''conmain.Execute "select * into LvbTmp from leavold"
90    Call CreateTableIntoAs("*", "Leavold", "LvbTmp")
'        conmain.CommitTrans
100   If Not adrsTemp.EOF Then
110       adrsTemp.MoveFirst
120       Do
130           FlName = adrsTemp(0)
140           If UCase(adrsTemp!Type) = "Y" Then
150               Select Case FieldExists("lvbtmp", FlName$)
                      Case False
                          ''For Mauritius 18-08-2003
160                       Select Case bytBackEnd
                              Case 1, 2 ''SQL Server,MSAccess
170                               ConMain.Execute "alter table lvbtmp add " & adrsTemp(0) & " FLoat"
200
180                           Case 3 ''Oracle
190                               ConMain.Execute "alter table lvbtmp add " & adrsTemp(0) & " Number(6,2)"
                               ConMain.Execute "Commit"
210                       End Select
220                       ConMain.Execute "update lvbtmp set " & adrsTemp(0) & "=0"
230               End Select
240           End If
250           adrsTemp.MoveNext
260       Loop Until adrsTemp.EOF
270   End If
280   ConMain.Execute "drop table leavold"
290   Call DropAddTb("leavbal", "lvbtmp")
300   If adrsTemp.State = 1 Then adrsTemp.Close
310   adrsTemp.Open "select distinct(lvcode) from Leavdesc ", ConMain
320   adrsTemp.MoveFirst     ' adrsTemp is the Recordset from the LeavBal
      ' Copy the Structure of LeavtrmPerm into Leavtrn
330   ConMain.Execute "Drop Table Leavtrn"
340   If FindTable("lvtrntmp") Then ConMain.Execute "drop table lvtrntmp"
      'conmain.Execute "select * into leavtrn from Lvtrnpermt"
350   Call CreateTableIntoAs("*", "Lvtrnpermt", "leavtrn")
      'conmain.Execute "select * into LvTrnTmp from Leavtrn"
360   Call CreateTableIntoAs("*", "leavtrn", "LvTrnTmp")
370   Do
          ' Check if the Field already Exists
380       FlName = adrsTemp(0)
390       Select Case FieldExists("lvtrntmp", FlName$)
              Case False
                  ''For Mauritius 18-08-2003
400               Select Case bytBackEnd
                      Case 1, 2 ''SQL Server,MSAccess
410                       ConMain.Execute "alter table lvtrntmp add [" & adrsTemp(0) & "] float"
420                   Case 3 ''Oracle
430                       ConMain.Execute "alter table lvtrntmp add " & adrsTemp(0) & " Number(6,2)"
440                       ConMain.Execute "Commit"
450               End Select
460               ConMain.Execute "update lvtrntmp set " & adrsTemp(0) & " = 0"
470       End Select
480       adrsTemp.MoveNext
490   Loop Until adrsTemp.EOF

500   adrsTemp.Close
510   Call DropAddTb("leavtrn", "lvtrntmp")
520   FlName$ = ""
530   On Error GoTo 0
540   Exit Sub
UpdateLvFile_Error:
550      If Erl = 0 Then
560         ShowError "Error in procedure UpdateLvFile of Form YRCRFrm"
570      Else
580         ShowError "Error in procedure UpdateLvFile of Form YRCRFrm And Line:" & Erl
590      End If
Resume Next
End Sub

Private Sub DropAddTb(ByVal DrpTable As String, ByVal AddTable As String)
On Error GoTo Err_particular
With ConMain
    .Execute "drop table " & DrpTable
    '.Execute "select * into " & DrpTable & " from " & AddTable
    Call CreateTableIntoAs("*", AddTable, DrpTable)
    .Execute "drop table " & AddTable
End With
Exit Sub
Err_particular:
    ShowError ("Drop Add Table :: Common")
End Sub

Private Sub FinishCmd_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF10 Then Call ShowF10("58")
If KeyCode = vbKeyF10 Then KeyCode = 0
End Sub
