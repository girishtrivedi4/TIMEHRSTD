VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCORul 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CO Rules"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkEarly 
      Caption         =   "Deduct Early Hours"
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   2820
      Width           =   1785
   End
   Begin VB.CheckBox chkLate 
      Caption         =   "Deduct Late Hours"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2820
      Width           =   1755
   End
   Begin VB.TextBox txtRDesc 
      Height          =   285
      Left            =   3300
      MaxLength       =   50
      TabIndex        =   1
      Top             =   90
      Width           =   4035
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   495
      Left            =   5040
      TabIndex        =   18
      Top             =   3180
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3786
      TabIndex        =   17
      Top             =   3180
      Width           =   1245
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   2534
      TabIndex        =   16
      Top             =   3180
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   1282
      TabIndex        =   15
      Top             =   3180
      Width           =   1245
   End
   Begin VB.Frame frAvail 
      Caption         =   "CO must be availed within"
      Height          =   645
      Left            =   5010
      TabIndex        =   26
      Top             =   510
      Width           =   2685
      Begin VB.ComboBox cboAvail 
         Height          =   315
         ItemData        =   "frmCORul.frx":0000
         Left            =   630
         List            =   "frmCORul.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   210
         Width           =   1515
      End
   End
   Begin VB.Frame frRule 
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   2025
      Begin MSForms.ComboBox cboRule 
         Height          =   285
         Left            =   1140
         TabIndex        =   0
         Top             =   150
         Width           =   795
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1402;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblRules 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CO Rules"
         Height          =   195
         Left            =   90
         TabIndex        =   20
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.Frame frWorkhrs 
      Height          =   1545
      Left            =   0
      TabIndex        =   27
      Top             =   1170
      Width           =   7695
      Begin VB.TextBox txtHLF 
         Height          =   285
         Left            =   6330
         MaxLength       =   5
         TabIndex        =   11
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtWOF 
         Height          =   285
         Left            =   4530
         MaxLength       =   5
         TabIndex        =   10
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtWDF 
         Height          =   285
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   9
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtHLH 
         Height          =   285
         Left            =   6330
         MaxLength       =   5
         TabIndex        =   8
         Top             =   600
         Width           =   555
      End
      Begin VB.TextBox txtWOH 
         Height          =   285
         Left            =   4530
         MaxLength       =   5
         TabIndex        =   7
         Top             =   600
         Width           =   555
      End
      Begin VB.TextBox txtWDH 
         Height          =   285
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   6
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lblFull 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum for Full Day"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   1050
         Width           =   1455
      End
      Begin VB.Label lblHalf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum for 1/2 Day"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   630
         Width           =   1470
      End
      Begin VB.Label lblHL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Holidays"
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
         Left            =   6240
         TabIndex        =   31
         Top             =   270
         Width           =   735
      End
      Begin VB.Label lblWO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weekoffs"
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
         Left            =   4440
         TabIndex        =   30
         Top             =   270
         Width           =   825
      End
      Begin VB.Label lblWD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weekdays"
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
         Left            =   2580
         TabIndex        =   29
         Top             =   270
         Width           =   900
      End
      Begin VB.Label lblHrs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hours"
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
         Left            =   780
         TabIndex        =   28
         Top             =   270
         Width           =   510
      End
      Begin VB.Line Line2 
         X1              =   90
         X2              =   7560
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   7560
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Shape Shape4 
         Height          =   1245
         Left            =   5730
         Top             =   210
         Width           =   1845
      End
      Begin VB.Shape Shape3 
         Height          =   1245
         Left            =   3960
         Top             =   210
         Width           =   1785
      End
      Begin VB.Shape Shape2 
         Height          =   1245
         Left            =   2130
         Top             =   210
         Width           =   1845
      End
      Begin VB.Shape Shape1 
         Height          =   1245
         Left            =   90
         Top             =   210
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   30
      TabIndex        =   14
      Top             =   3180
      Width           =   1245
   End
   Begin VB.Frame frCheck 
      Caption         =   "Give CO on "
      Height          =   645
      Left            =   0
      TabIndex        =   25
      Top             =   510
      Width           =   4995
      Begin VB.CheckBox chkWD 
         Caption         =   "Weekdays"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   1155
      End
      Begin VB.CheckBox chkWO 
         Caption         =   "Weekoff"
         Height          =   195
         Left            =   1920
         TabIndex        =   3
         Top             =   300
         Width           =   1005
      End
      Begin VB.CheckBox chkHL 
         Caption         =   "Holiday"
         Height          =   195
         Left            =   3480
         TabIndex        =   4
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame frNewRule 
      Height          =   495
      Left            =   5700
      TabIndex        =   21
      Top             =   0
      Width           =   1635
      Begin VB.TextBox txtRNo 
         Height          =   285
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   23
         Top             =   150
         Width           =   465
      End
      Begin VB.Label lblRNo 
         AutoSize        =   -1  'True
         Caption         =   "CO Rule No."
         Height          =   195
         Left            =   60
         TabIndex        =   22
         Top             =   180
         Width           =   900
      End
   End
   Begin VB.Label lblRDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   2100
      TabIndex        =   24
      Top             =   150
      Width           =   795
   End
End
Attribute VB_Name = "frmCORul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset

Private Sub cboRule_Click()
Call Display
End Sub

Private Sub cmdAdd_Click()
On Error GoTo ERR_P
If Not AddRights Then
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    Exit Sub
Else
    bytMode = 2
    Call ChangeMode
End If
Exit Sub
ERR_P:
    ShowError ("Add::" & Me.Caption)
End Sub

Private Sub cmdCancel_Click()
Select Case bytMode
    Case 1  '' Edit / View
         Call Display
    Case 2  '' Add
        If cboRule.ListCount > 0 Then
            bytMode = 1
            Call ChangeMode
            Call Display
        End If
End Select
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ERR_P
If Not DeleteRights Then
    MsgBox NewCaptionTxt("00001", adrsMod)
    Exit Sub
Else
    If cboRule.Visible Then
        If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion) = vbYes Then
            ConMain.Execute "Delete From CORul Where COCode=" & cboRule.Text
            Call AuditInfo("DELETE", Me.Caption, "Delete CO Rule no.: " & cboRule.Text)
            Call OpenMasterTable
            Call FillCombo
            If cboRule.ListCount = 0 Then
                If Not AddRights Then
                    cmdAdd.Enabled = False
                    cmdSave.Enabled = False
                    cmdCancel.Enabled = False
                    cmdDelete.Enabled = False
                Else
                    bytMode = 2     '' Add
                    Call ChangeMode
                End If
            Else
                bytMode = 1     '' Edit / View
                cboRule.ListIndex = 0
            End If
        End If
    End If
End If
Exit Sub
ERR_P:
    If Err.Number = -2147467259 Or Err.Number = -2147217900 Or Err.Number = -2147217873 Then
            MsgBox "CORule Cannot be deleted because employees belong to this CORule.", vbCritical, Me.Caption
            Exit Sub
    End If
    ShowError ("Delete Record :: " & Me.Caption)
End Sub

Private Sub cmdsave_Click()
Select Case bytMode
    Case 1  '' Edit / View
        If Not EditRights Then
            MsgBox NewCaptionTxt("00001", adrsMod)
        Else
            If Not ValidateModMaster Then Exit Sub
            If Not SaveModMaster Then Exit Sub
            Call AuditInfo("UPDATE", Me.Caption, "Edit CO Rule no.: " & cboRule.Text)
            Call OpenMasterTable
        End If
    Case 2  '' Add
        If Not ValidateAddmaster Then Exit Sub
        If Not SaveAddMaster Then Exit Sub
        Call AuditInfo("ADD", Me.Caption, "Added CO Rule no.: " & cboRule.Text)
        Call OpenMasterTable
        Call FillCombo
        bytMode = 1
        Call ChangeMode
        If cboRule.ListCount > 0 Then cboRule.ListIndex = 0
End Select
End Sub

Private Function ValidateModMaster() As Boolean
On Error GoTo ERR_P
If cboRule.Text = "" Then Exit Function
If adrsDept1.RecordCount <= 0 Then
    MsgBox NewCaptionTxt("60010", adrsC)
    Exit Function
End If
adrsDept1.MoveFirst
adrsDept1.Find "COCode=" & cboRule.Text
If adrsDept1.EOF Then
    MsgBox NewCaptionTxt("60010", adrsC)
    Exit Function
End If
If Trim(txtRDesc.Text) = "" Then
    MsgBox NewCaptionTxt("60008", adrsC)
    txtRDesc.SetFocus
    Exit Function
End If
If Not CorrectDecs Then Exit Function
ValidateModMaster = True
Exit Function
ERR_P:
    ShowError ("ValidateModMaster::" & Me.Caption)
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
ConMain.Execute "Update CORul Set CODesc='" & Trim(txtRDesc.Text) & "'," & _
"COWD=" & chkWD.Value & ",COWO=" & chkWO.Value & ",COHL=" & chkHL.Value & ",COAvail=" & _
cboAvail.ListIndex & ",WDH=" & txtWDH.Text & ",WOH=" & txtWOH.Text & ",HLH=" & _
txtHLH.Text & ",WDF=" & txtWDF.Text & ",WOF=" & txtWOF.Text & ",HLF=" & txtHLF.Text & "," & _
"DedLate=" & chkLate.Value & ",DedEarl=" & chkEarly.Value & " Where COCode=" & cboRule.Text
SaveModMaster = True
Exit Function
ERR_P:
    ShowError ("SaveModMaster::" & Me.Caption)
End Function

Private Function SaveAddMaster() As Boolean
On Error GoTo ERR_P
ConMain.Execute "insert into CORul values(" & txtRNo.Text & ",'" & _
Trim(txtRDesc.Text) & "'," & chkWD.Value & "," & chkWO.Value & "," & chkHL.Value & "," & _
cboAvail.ListIndex & "," & txtWDH.Text & "," & txtWOH.Text & "," & txtHLH.Text & "," & _
txtWDF.Text & "," & txtWOF.Text & "," & txtHLF.Text & "," & chkLate.Value & "," & _
chkEarly.Value & ")"
SaveAddMaster = True
Exit Function
ERR_P:
    ShowError ("SaveAddMaster::" & Me.Caption)
End Function

Private Function ValidateAddmaster() As Boolean
On Error GoTo ERR_P
If Trim(txtRNo.Text) = "" Then
    MsgBox NewCaptionTxt("60006", adrsC)
    txtRNo.SetFocus
    Exit Function
Else
    If adrsDept1.RecordCount > 0 Then
        adrsDept1.MoveFirst
        adrsDept1.Find "COCode=" & Trim(txtRNo.Text)
        If Not adrsDept1.EOF Then
            MsgBox NewCaptionTxt("60007", adrsC)
            txtRNo.SetFocus
            Exit Function
        End If
    End If
End If
If Trim(txtRDesc.Text) = "" Then
    MsgBox NewCaptionTxt("60008", adrsC)
    txtRDesc.SetFocus
    Exit Function
End If
If Not CorrectDecs Then Exit Function
ValidateAddmaster = True
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster::" & Me.Caption)
End Function

Private Function CorrectDecs() As Boolean
On Error GoTo ERR_P
txtWDH.Text = IIf(Trim(txtWDH.Text) = "", "0.00", Format(txtWDH.Text, "0.00"))
If Val(Right(txtWDH.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtWDH.SetFocus
    Exit Function
End If
txtWOH.Text = IIf(Trim(txtWOH.Text) = "", "0.00", Format(txtWOH.Text, "0.00"))
If Val(Right(txtWOH.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtWOH.SetFocus
    Exit Function
End If
txtHLH.Text = IIf(Trim(txtHLH.Text) = "", "0.00", Format(txtHLH.Text, "0.00"))
If Val(Right(txtHLH.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtHLH.SetFocus
    Exit Function
End If
txtWDF.Text = IIf(Trim(txtWDF.Text) = "", "0.00", Format(txtWDF.Text, "0.00"))
If Val(Right(txtWDF.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtWDF.SetFocus
    Exit Function
End If
txtWOF.Text = IIf(Trim(txtWOF.Text) = "", "0.00", Format(txtWOF.Text, "0.00"))
If Val(Right(txtWOF.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtWOF.SetFocus
    Exit Function
End If
txtHLF.Text = IIf(Trim(txtHLF.Text) = "", "0.00", Format(txtHLF.Text, "0.00"))
If Val(Right(txtHLF.Text, 2)) > 59 Then
    MsgBox NewCaptionTxt("00024", adrsMod)
    txtHLF.SetFocus
    Exit Function
End If
'' Full Day Hours should be greater than Half Day Hours.
If Val(txtWDF.Text) <= Val(txtWDH.Text) And Val(txtWDF.Text) <> 0 Then
    MsgBox NewCaptionTxt("60009", adrsC)
    txtWDF.SetFocus
    Exit Function
End If
If Val(txtWOF.Text) <= Val(txtWOH.Text) And Val(txtWOF.Text) <> 0 Then
    MsgBox NewCaptionTxt("60009", adrsC)
    txtWOF.SetFocus
    Exit Function
End If
If Val(txtHLF.Text) <= Val(txtHLH.Text) And Val(txtHLF.Text) <> 0 Then
    MsgBox NewCaptionTxt("60009", adrsC)
    txtHLF.SetFocus
    Exit Function
End If
CorrectDecs = True
Exit Function
ERR_P:
    ShowError ("CorrectDecs::" & Me.Caption)
End Function

Private Sub Form_Load()
On Error GoTo ERR_P
Call SetFormIcon(Me)
Call GetRights
Call OpenMasterTable
Call RetCaptions
Call FillCombo
If cboRule.ListCount = 0 Then
    If Not AddRights Then
        MsgBox NewCaptionTxt("00001", adrsMod)
        cmdAdd.Enabled = False
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        cmdDelete.Enabled = False
        Exit Sub
    Else
        bytMode = 2     '' Add
    End If
Else
    bytMode = 1     '' Edit / View
    cboRule.ListIndex = 0
End If
Call ChangeMode
Exit Sub
ERR_P:
    ShowError (" Load  :: " & Me.Caption)
End Sub

Private Sub OpenMasterTable()             '' Open the Recordset for the Display purposes
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select * from CORul order by COCode,CODesc", ConMain, adOpenStatic
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Sub ChangeMode()
Select Case bytMode
    Case 1  '' Edit / View
        Call EditAction
    Case 2  '' Add
        Call AddAction
End Select
End Sub

Private Sub EditAction()
'' Enable the Necessary controls
cmdAdd.Enabled = True
cmdDelete.Enabled = True
'' Visible necessary Controls
frRule.Visible = True
'' Invisible necessary controls
frNewRule.Visible = False
End Sub

Private Sub AddAction()
On Error Resume Next
'' Enable the Necessary controls
cmdAdd.Enabled = False
cmdDelete.Enabled = False
'' Visible necessary Controls
frRule.Visible = False
'' Invisible necessary controls
frNewRule.Left = frRule.Left
frNewRule.Top = frRule.Top
frNewRule.Visible = True
txtRNo.Text = ""
Call ClearControls
txtRNo.SetFocus
End Sub

Private Sub ClearControls()
txtRDesc.Text = ""
chkWD.Value = 1
chkWO.Value = 1
chkHL.Value = 1
cboAvail.ListIndex = 0
txtWDH.Text = "0.00"
txtWOH.Text = "0.00"
txtHLH.Text = "0.00"
txtWDF.Text = "0.00"
txtWOF.Text = "0.00"
txtHLF.Text = "0.00"
End Sub

Private Sub GetRights()     '' Gets and Sets the Appropriate Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 10)
If Mid(strTmp, 1, 1) = "1" Then AddRights = True
If Mid(strTmp, 2, 1) = "1" Then EditRights = True
If Mid(strTmp, 3, 1) = "1" Then DeleteRights = True
Exit Sub
ERR_P:
    ShowError ("Rights ::" & Me.Caption)
    AddRights = False
    EditRights = False
    DeleteRights = False
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '60%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("00089", adrsMod)
lblRules.Caption = NewCaptionTxt("00089", adrsMod)
lblRNo.Caption = NewCaptionTxt("60001", adrsC)
lblRDesc.Caption = NewCaptionTxt("00052", adrsMod)
'' frCheck
frCheck.Caption = NewCaptionTxt("60002", adrsC)
chkWD.Caption = NewCaptionTxt("00092", adrsMod)
chkWO.Caption = NewCaptionTxt("00093", adrsMod)
chkHL.Caption = NewCaptionTxt("00094", adrsMod)
'' frAvail
frAvail.Caption = NewCaptionTxt("60003", adrsC)
'' Others
lblHrs.Caption = NewCaptionTxt("00023", adrsMod)
lblWD.Caption = NewCaptionTxt("00092", adrsMod)
lblWO.Caption = NewCaptionTxt("00093", adrsMod)
lblHL.Caption = NewCaptionTxt("00094", adrsMod)

chkLate.Caption = NewCaptionTxt("60011", adrsC)
chkEarly.Caption = NewCaptionTxt("60012", adrsC)

End Sub

Private Sub FillCombo()
On Error GoTo ERR_P
Dim bytTmp As Byte
'' Fill the OT Combo
Call ComboFill(cboRule, 28, 2)
'' fill the Round-off Combo

For bytTmp = 1 To 8
    cboAvail.AddItem Choose(bytTmp, "No Limits", "Same Month", "15 days", "30 days", "45 days", _
        "60 days", "75 days", "90 days")
Next

cboAvail.ListIndex = 0
Exit Sub
ERR_P:
    ShowError ("FillCombo::" & Me.Caption)
End Sub

Private Sub Display()
On Error GoTo ERR_P
If cboRule.Text = "" Then Exit Sub
If adrsDept1.RecordCount <= 0 Then Exit Sub
adrsDept1.MoveFirst
adrsDept1.Find "COCode=" & cboRule.Text
If Not adrsDept1.EOF Then
    txtRDesc.Text = IIf(IsNull(adrsDept1("CODesc")), "", adrsDept1("CODesc"))
    '' Give CO on
    chkWD.Value = IIf(adrsDept1("COWD") = 1, 1, 0)
    chkWO.Value = IIf(adrsDept1("COWO") = 1, 1, 0)
    chkHL.Value = IIf(adrsDept1("COHL") = 1, 1, 0)
    '' Continuous Slab
    cboAvail.ListIndex = IIf(IsNull(adrsDept1("COAvail")), 0, adrsDept1("COAvail"))
    '' Others
    txtWDH.Text = IIf(IsNull(adrsDept1("WDH")), "0.00", Format(adrsDept1("WDH")))
    txtWOH.Text = IIf(IsNull(adrsDept1("WOH")), "0.00", Format(adrsDept1("WOH")))
    txtHLH.Text = IIf(IsNull(adrsDept1("HLH")), "0.00", Format(adrsDept1("HLH")))
    txtWDF.Text = IIf(IsNull(adrsDept1("WDF")), "0.00", Format(adrsDept1("WDF")))
    txtWOF.Text = IIf(IsNull(adrsDept1("WOF")), "0.00", Format(adrsDept1("WOF")))
    txtHLF.Text = IIf(IsNull(adrsDept1("HLF")), "0.00", Format(adrsDept1("HLF")))
    '' Late Early
    chkLate.Value = IIf(adrsDept1("DedLate") = 1, 1, 0)
    chkEarly.Value = IIf(adrsDept1("DedEarl") = 1, 1, 0)
Else
    MsgBox NewCaptionTxt("60010", adrsC)
End If
Exit Sub
ERR_P:
    ShowError ("Display")
End Sub

Private Sub txtRDesc_GotFocus()
Call GF(txtRDesc)
End Sub

Private Sub txtRNo_GotFocus()
Call GF(txtRNo)
End Sub

Private Sub txtRNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 2))))
End If
End Sub

Private Sub txtRDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 5))))
End If
End Sub

Private Sub txtRNo_LostFocus()
If txtRNo.Text = "100" Then
 MsgBox " This CO Rule is reserved for Application"
 txtRNo.Text = ""
 txtRNo.SetFocus
 End If
End Sub

Private Sub txtWDH_GotFocus()
Call GF(txtWDH)
End Sub

Private Sub txtWDH_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtWDH)
End If
End Sub

Private Sub txtWOH_GotFocus()
Call GF(txtWOH)
End Sub

Private Sub txtWOH_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtWOH)
End If
End Sub

Private Sub txtHLH_GotFocus()
Call GF(txtHLH)
End Sub

Private Sub txtHLH_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtHLH)
End If
End Sub

Private Sub txtWDF_GotFocus()
Call GF(txtWDF)
End Sub

Private Sub txtWDF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtWDF)
End If
End Sub

Private Sub txtWOF_GotFocus()
Call GF(txtWOF)
End Sub

Private Sub txtWOF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtWOF)
End If
End Sub

Private Sub txtHLF_GotFocus()
Call GF(txtHLF)
End Sub

Private Sub txtHLF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = keycheck(KeyAscii, txtHLF)
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
