VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPE 
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLog 
      Caption         =   "&Log"
      Height          =   525
      Left            =   8250
      TabIndex        =   20
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "&Path"
      Height          =   525
      Left            =   9450
      TabIndex        =   21
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   525
      Left            =   10650
      TabIndex        =   22
      Top             =   7680
      Width           =   1155
   End
   Begin VB.Frame Frame3 
      Height          =   2415
      Left            =   4680
      TabIndex        =   23
      Top             =   5130
      Width           =   7155
      Begin VB.CommandButton cmdCurrent 
         Caption         =   "Get &Current"
         Height          =   375
         Left            =   5820
         TabIndex        =   50
         Top             =   780
         Width           =   1275
      End
      Begin VB.TextBox txtPass 
         Height          =   315
         Left            =   5340
         TabIndex        =   49
         Top             =   2040
         Width           =   1125
      End
      Begin VB.TextBox txtUName 
         Height          =   315
         Left            =   3180
         TabIndex        =   47
         Top             =   2040
         Width           =   1125
      End
      Begin VB.CheckBox chkWeb 
         Caption         =   "Web Enabled"
         Height          =   195
         Left            =   4290
         TabIndex        =   43
         Top             =   1740
         Width           =   1335
      End
      Begin VB.ComboBox cboVer 
         Height          =   315
         Left            =   4470
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1320
         Width           =   915
      End
      Begin VB.ComboBox cboBack 
         Height          =   315
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   2010
         Width           =   1335
      End
      Begin VB.TextBox txtCom 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   780
         TabIndex        =   26
         Top             =   570
         Width           =   4755
      End
      Begin VB.CheckBox chkVer 
         Caption         =   "Demo Version"
         Height          =   225
         Left            =   90
         TabIndex        =   27
         Top             =   990
         Width           =   1335
      End
      Begin VB.TextBox txtEmp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2850
         TabIndex        =   29
         Text            =   "0"
         Top             =   960
         Width           =   555
      End
      Begin VB.CheckBox chkUser 
         Caption         =   "Multiuser"
         Height          =   225
         Left            =   90
         TabIndex        =   32
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1005
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   4470
         TabIndex        =   31
         Top             =   960
         Width           =   555
      End
      Begin VB.TextBox txtComL 
         Height          =   285
         Left            =   2850
         TabIndex        =   34
         Top             =   1320
         Width           =   555
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Settings"
         Height          =   375
         Left            =   5820
         TabIndex        =   51
         Top             =   1230
         Width           =   1275
      End
      Begin VB.TextBox txtD 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   870
         TabIndex        =   38
         Top             =   1680
         Width           =   345
      End
      Begin VB.TextBox txtM 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   39
         Top             =   1680
         Width           =   345
      End
      Begin VB.TextBox txtY 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1530
         TabIndex        =   40
         Top             =   1680
         Width           =   345
      End
      Begin VB.CheckBox chkAssum 
         Caption         =   "Assum"
         Height          =   195
         Left            =   3330
         TabIndex        =   42
         Top             =   1740
         Width           =   825
      End
      Begin VB.TextBox txtIniPath 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   210
         Width           =   5415
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5160
         Top             =   570
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   195
         Left            =   4560
         TabIndex        =   48
         Top             =   2100
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         Height          =   195
         Left            =   2400
         TabIndex        =   46
         Top             =   2100
         Width           =   720
      End
      Begin VB.Label lblVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         Height          =   195
         Left            =   3660
         TabIndex        =   35
         Top             =   1350
         Width           =   525
      End
      Begin VB.Label lblBack 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database"
         Height          =   195
         Left            =   90
         TabIndex        =   44
         Top             =   2070
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   195
         Left            =   60
         TabIndex        =   25
         Top             =   630
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Limit"
         Height          =   195
         Left            =   1680
         TabIndex        =   28
         Top             =   990
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Limit"
         Height          =   195
         Left            =   3660
         TabIndex        =   30
         Top             =   990
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Limit"
         Height          =   195
         Left            =   1650
         TabIndex        =   33
         Top             =   1350
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lock date"
         Height          =   195
         Left            =   90
         TabIndex        =   37
         Top             =   1710
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(DD/MM/YY)"
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
         Left            =   2010
         TabIndex        =   41
         Top             =   1740
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2385
      Left            =   0
      TabIndex        =   11
      Top             =   5130
      Width           =   4665
      Begin VB.TextBox txtAccPass 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   4020
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1860
         Width           =   585
      End
      Begin VB.TextBox txtAccPath 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   30
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   1860
         Width           =   3975
      End
      Begin VB.CommandButton cmdSHL 
         Caption         =   "SHL"
         Height          =   405
         Left            =   4020
         TabIndex        =   17
         Top             =   1230
         Width           =   585
      End
      Begin VB.TextBox txtSHL 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   30
         TabIndex        =   16
         Top             =   1230
         Width           =   3975
      End
      Begin VB.CommandButton cmdDen 
         Caption         =   "DEN"
         Height          =   405
         Left            =   4020
         TabIndex        =   15
         Top             =   630
         Width           =   585
      End
      Begin VB.TextBox txtDenStr 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   30
         TabIndex        =   14
         Top             =   630
         Width           =   3975
      End
      Begin VB.TextBox txtXOR 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   4020
         TabIndex        =   13
         Text            =   "0"
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txtStr 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   30
         TabIndex        =   12
         Top             =   180
         Width           =   3975
      End
   End
   Begin VB.Frame frDetails 
      Height          =   5115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      Begin VB.Frame Frame1 
         Height          =   1245
         Left            =   30
         TabIndex        =   2
         Top             =   3810
         Width           =   2295
         Begin VB.CheckBox chkLog 
            Caption         =   "&Delog"
            Height          =   345
            Left            =   1350
            TabIndex        =   7
            Top             =   540
            Width           =   825
         End
         Begin VB.OptionButton optALL 
            Caption         =   "ALL"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   180
            Width           =   945
         End
         Begin VB.OptionButton optL 
            Caption         =   "ALL LV's"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   990
            Width           =   1335
         End
         Begin VB.OptionButton optS 
            Caption         =   "ALL SHF's"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   690
            Width           =   1095
         End
         Begin VB.OptionButton optT 
            Caption         =   "ALL TRN's"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   420
            Width           =   1095
         End
      End
      Begin VB.TextBox txtU 
         Height          =   585
         Left            =   2340
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   4470
         Width           =   9435
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   3735
         Left            =   2340
         TabIndex        =   9
         Top             =   750
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   6588
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         WordWrap        =   -1  'True
         AllowUserResizing=   3
      End
      Begin VB.TextBox txtQ 
         Height          =   585
         Left            =   2340
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   9435
      End
      Begin VB.ListBox lstTables 
         Height          =   3570
         Left            =   30
         TabIndex        =   1
         Top             =   180
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Print User Module
Option Explicit

Private Sub chkLog_Click()
Call AddTables
End Sub

Private Sub chkUser_Click()
    If chkUser.Value = 0 Then
        txtUser.Text = "1"
        txtUser.Enabled = False
    Else
        txtUser.Text = ""
        txtUser.Enabled = True
        txtUser.SetFocus
    End If
End Sub

Private Sub chkVer_Click()
On Error Resume Next
If chkVer.Value = 1 Then
    txtEmp.Enabled = True
    txtEmp.Text = ""
    txtEmp.SetFocus
Else
    txtEmp.Text = "0"
    txtEmp.Enabled = False
End If
End Sub

Private Sub cmdCurrent_Click()
Call ShowDefaultSettings
End Sub

Private Sub cmdDen_Click()
Dim intTmp As Integer
Select Case Val(txtXOR.Text)
    Case 0
        intTmp = 11
    Case -1
        intTmp = 12
    Case Is <= -2
        intTmp = 11
    Case Is > 0
        intTmp = Val(txtXOR.Text)
End Select
txtDenStr.Text = NewDEN(Trim(txtStr.Text), intTmp)
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdLog_Click()
Call Shell("Notepad.Exe " & App.path & "\VSLog.Log", vbMaximizedFocus)
End Sub

Private Sub cmdPath_Click()
If MsgBox("In SHELL ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
    MsgBox strBackEndPath
Else
    txtSHL.Text = txtSHL.Text & "   " & strBackEndPath
    txtSHL.SetFocus
End If
End Sub

Private Sub cmdsave_Click()
On Error GoTo ERR_P:
'' Procedure to write to the the INI File
Dim strData As String
Dim strERR As String, bytErr As Byte, strLokDate As String
If Trim(txtIniPath.Text) = "" Then Exit Sub
strERR = "The Following Errors have Encountered": bytErr = 0
If Trim(txtCom.Text) = "" Then
    strERR = strERR & vbCrLf & "Please Enter Company Name"
    bytErr = bytErr + 1
End If
If Val(txtEmp.Text) <= 0 And chkVer.Value = 1 Then
    strERR = strERR & vbCrLf & "Please Enter Employee Limit"
    bytErr = bytErr + 1
End If
If Val(txtUser.Text) < 1 Then
    strERR = strERR & vbCrLf & "Please Enter User Limit"
    bytErr = bytErr + 1
End If
If Val(txtComL.Text) < 1 Then
    strERR = strERR & vbCrLf & "Please Enter Company Limit"
    bytErr = bytErr + 1
End If
If Len(txtD.Text) < 2 Then
    strERR = strERR & vbCrLf & "Please Enter Day"
    bytErr = bytErr + 1
End If
If Len(txtM.Text) < 2 Then
    strERR = strERR & vbCrLf & "Please Enter Month"
    bytErr = bytErr + 1
End If
If Len(txtY.Text) < 2 Then
    strERR = strERR & vbCrLf & "Please Enter Year"
    bytErr = bytErr + 1
End If
Select Case Val(txtM.Text)
    Case 2
        If Val(txtD.Text) > 29 Or Val(txtD.Text) <= 0 Then
            bytErr = bytErr + 1
            strERR = strERR & vbCrLf & "Invalid Number of Days"
        End If
    Case 1, 3, 5, 7, 8, 10, 12
        If Val(txtD.Text) > 31 Or Val(txtD.Text) <= 0 Then
            bytErr = bytErr + 1
            strERR = strERR & vbCrLf & "Invalid Number of Days"
        End If
    Case 4, 6, 9, 11
        If Val(txtD.Text) > 30 Or Val(txtD.Text) <= 0 Then
            bytErr = bytErr + 1
            strERR = strERR & vbCrLf & "Invalid Number of Days"
        End If
    Case Else
        bytErr = bytErr + 1
        strERR = strERR & vbCrLf & "Invalid Month Number"
End Select
If bytErr > 0 Then
    MsgBox strERR, vbCritical, "Data Initialization"
    Exit Sub
End If
'' Process to write Data
strData = ""
'' Company Name
strData = DEncryptDat(txtCom.Text & "|", 1)
'' Version Type
strData = strData & DEncryptDat(CStr(chkVer.Value) & "|", 1)
'' Employee Limit
strData = strData & DEncryptDat(txtEmp.Text & "|", 1)
'' Net Type
strData = strData & DEncryptDat(CStr(chkUser.Value) & "|", 1)
'' User Limit
strData = strData & DEncryptDat(txtUser.Text & "|", 1)
'' Company Limit
strData = strData & DEncryptDat(txtComL.Text & "|", 1)
'' Version Number
strData = strData & DEncryptDat(cboVer.Text & "|", 1)
'' Lock Date
strLokDate = Format(txtY.Text, "00") & Format(txtM.Text, "00") & _
Format(txtD.Text, "00")
strData = strData & DEncryptDat(strLokDate & "|", 1)
'' Assum
strData = strData & DEncryptDat(CStr(chkAssum.Value) & "|", 1)
'' Back End
strData = strData & DEncryptDat(cboBack.ListIndex + 1 & "|", 1)
'' Web Enabled
strData = strData & DEncryptDat(CStr(chkWeb.Value) & "|", 1)
'' User Name
strData = strData & DEncryptDat(Trim(txtUName.Text) & "|", 1)
'' Password
strData = strData & DEncryptDat(Trim(txtPass.Text), 1)
''    '' Write
If Dir(txtIniPath.Text) <> "" Then Kill txtIniPath.Text
strData = strData & vbCrLf & DEncryptDat(strData, 2)
Open txtIniPath.Text For Output As #1
Print #1, strData
Close #1
If Err.Number = 0 Then MsgBox "File " & txtIniPath.Text & " saved successfully"
Exit Sub
ERR_P:
    ShowError ("Save::" & Me.Caption)
    Resume Next
End Sub

Private Sub cmdSHL_Click()
On Error GoTo ERR_P
If Trim(txtSHL.Text) = "" Then Exit Sub
Call Shell(Trim(txtSHL.Text), vbMaximizedFocus)
Exit Sub
ERR_P:
    ShowError ("PE::Shell::")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    If KeyAscii = 13 Then
        Select Case UCase(ActiveControl.name)
            Case "TXTQ"
                Call DoSelect
            Case "TXTU"
                Call DoUpdate
            Case "TXTSTR"
                SendKeys Chr(9)
            Case "TXTXOR"
                SendKeys Chr(9)
            Case "TXTDENSTR"
                Call cmdDen_Click
            Case "TXTSHL"
                Call cmdSHL_Click
            Case "TXTINIPATH"
                Call txtIniPath_DblClick
            Case "TXTACCPATH"
                Call txtAccPath_DblClick
        End Select
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Call SetFormIcon(Me, True)
optALL.Value = True
chkLog.Value = 0
With cboBack
    .AddItem "SQL Server"
    .AddItem "MS-Access"
    .AddItem "Oracle"
    .ListIndex = 1
End With
With cboVer
    .AddItem "2.0"
    .AddItem "3.0"
    .AddItem "4.0"
    .AddItem "5.0"
    .ListIndex = 0
End With
Call ShowDefaultSettings
If bytBackEnd = 1 Then cmdPath.Enabled = False
Exit Sub
ERR_P:
    ShowError ("PE :: " & Me.Caption)
End Sub

Private Sub lstTables_DblClick()
If lstTables.ListCount <= 0 Then Exit Sub
If lstTables.ListIndex < 0 Then Exit Sub
txtQ.Text = "Select * from " & lstTables.List(lstTables.ListIndex)
Call DoSelect
End Sub

Private Sub lstTables_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call lstTables_DblClick
End Sub

Private Sub AddTables()         '' Add User Tables to the List
On Error GoTo ERR_P
Dim adTmp As New ADODB.Recordset, bytTmp As Integer
If bytFormToLoad = 10 Then
    lstTables.Enabled = False
    txtQ.Enabled = False
    txtU.Enabled = False
    MSF1.Enabled = False
    Exit Sub
End If
adTmp.CursorType = adOpenStatic
adTmp.LockType = adLockOptimistic
If chkLog.Value = 1 Then
    Set adTmp = ConLog.OpenSchema(adSchemaTables)
Else
    Select Case bytBackEnd
        Case 1, 2
            Set adTmp = ConMain.OpenSchema(adSchemaTables)
        Case 3
            Set adTmp = ConMain.Execute("select 1,2,UPPER(Table_Name),'TABLE' from Tabs ")
    End Select
End If
lstTables.clear
If adTmp.RecordCount = 0 Then Exit Sub
For bytTmp = 0 To adTmp.RecordCount - 1
    If adTmp(3) = "TABLE" Then
        If chkLog.Value = 1 Then
            lstTables.AddItem UCase(adTmp(2))
        ElseIf optALL.Value = True Then
            lstTables.AddItem UCase(adTmp(2))
        ElseIf optT.Value = True Then
            Select Case UCase(Trim(adTmp(2)))
                Case "DTRN", "MONTRN", "LEAVTRN"
                Case Else
                    If Right(UCase(Trim(adTmp(2))), 3) = "TRN" Then
                        lstTables.AddItem UCase(adTmp(2))
                    End If
            End Select
        ElseIf optS.Value = True Then
            If Right(UCase(Trim(adTmp(2))), 3) = "SHF" Then lstTables.AddItem UCase(adTmp(2))
        ElseIf optL.Value = True Then
            Select Case UCase(Trim(adTmp(2)))
                Case "LVTRNPERMT"
                Case Else
                    If Len(Trim(adTmp(2))) > 5 Then
                        Select Case Left(UCase(Trim(adTmp(2))), 5)
                            Case "LVBAL", "LVTRN", "LVINF"
                                lstTables.AddItem UCase(adTmp(2))
                        End Select
                    End If
            End Select
        Else
            lstTables.AddItem UCase(adTmp(2))
        End If
    End If
    adTmp.MoveNext
Next
Exit Sub
ERR_P:
    ShowError (" Add Tables :: " & Me.Caption)
    
End Sub

Private Sub optALL_Click()
Call AddTables
End Sub

Private Sub optL_Click()
Call AddTables
End Sub

Private Sub optS_Click()
Call AddTables
End Sub

Private Sub optT_Click()
Call AddTables
End Sub

Private Sub DoSelect()          '' Process  the Select Query
If Len(Trim(txtQ.Text)) < 10 Then Exit Sub
If InStr(UCase(Trim(txtQ.Text)), "SELECT") <= 0 Then
    MsgBox "Please Write the Select Query", vbExclamation, App.EXEName
    Exit Sub
End If
Select Case ExeSel
    Case 0      '' Error
        MSF1.Rows = 1
        MSF1.Cols = 1
        Me.Caption = ""
    Case 1      '' Succesfull
        MSF1.Rows = 1
        MSF1.Cols = 1
        Call FillGrid
    Case 2      '' No Records
End Select
End Sub

Private Function ExeSel() As Byte       '' Execute Select Statement
On Error GoTo ERR_P
ExeSel = 1
If adrsTemp.State = 1 Then adrsTemp.Close
If chkLog.Value = 1 Then
    adrsTemp.Open Trim(txtQ.Text), ConLog, adOpenKeyset
Else
    adrsTemp.Open Trim(txtQ.Text), ConMain, adOpenKeyset
End If
If (adrsTemp.EOF And adrsTemp.BOF) Then
    MsgBox "No Records Found", vbExclamation, App.EXEName
    ExeSel = 2
End If
Exit Function
ERR_P:
    ShowError ("Execute Select :: " & Me.Caption)
    ExeSel = 0
End Function

Private Sub FillGrid()                  '' Fill the Grid with the Execute Statement
On Error GoTo ERR_P
Dim bytTmp As Byte, intTmp As Integer
'' Get All Column Names
MSF1.Cols = adrsTemp.Fields.Count
For bytTmp = 0 To adrsTemp.Fields.Count - 1
    MSF1.TextMatrix(0, bytTmp) = UCase(adrsTemp.Fields(CInt(bytTmp)).name)
Next
'' Get All Rows
MSF1.Rows = adrsTemp.RecordCount + 1
adrsTemp.MoveFirst
For intTmp = 1 To adrsTemp.RecordCount
    For bytTmp = 0 To adrsTemp.Fields.Count - 1
        MSF1.TextMatrix(intTmp, bytTmp) = IIf(IsNull(adrsTemp.Fields(CInt(bytTmp)).Value), "", _
        adrsTemp.Fields(CInt(bytTmp)).Value)
    Next
    adrsTemp.MoveNext
Next
Me.Caption = adrsTemp.RecordCount
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub DoUpdate()                      '' Process the Update Query
If Len(Trim(txtU.Text)) < 10 Then Exit Sub
If MsgBox("Are You Sure to Execute this Query", vbYesNo + vbQuestion, App.EXEName) = vbYes Then
Else
    Exit Sub
End If
Select Case ExeUpd
    Case 0      '' Error
    Case 1      '' Succesfull
        MsgBox "Operation Successfull", vbExclamation, App.EXEName
    Case 2      '' Other Consequences
End Select
End Sub

Private Function ExeUpd() As Byte       '' Execute Update Statement
On Error GoTo ERR_P
ExeUpd = 1
If chkLog.Value = 1 Then
    ConLog.Execute Trim(txtU.Text)
Else
    ConMain.Execute Trim(txtU.Text)
End If
Exit Function
ERR_P:
    ShowError ("Execute Update :: " & Me.Caption)
    ExeUpd = 0
End Function

Private Function NewDEN(pwd As String, intTimes As Integer) As String
Dim pwdStr As String, i As Integer
For i = 1 To Len(pwd)
        pwdStr = pwdStr & Chr(Asc(Mid(pwd, i, 1)) Xor intTimes)
Next i
NewDEN = pwdStr
End Function

Private Sub txtAccPath_DblClick()
CommonDialog1.FileName = ""
CommonDialog1.Filter = "*.MDB|*.MDB"
CommonDialog1.Flags = cdlOFNFileMustExist
If Trim(txtAccPath.Text) = "" Then
    CommonDialog1.InitDir = App.path
Else
    CommonDialog1.InitDir = Left(txtAccPath.Text, 3)
End If
CommonDialog1.ShowOpen
If Trim(CommonDialog1.FileName) = "" Then Exit Sub
If Dir(CommonDialog1.FileName) = "" Then
    MsgBox "Access File is Missing", vbCritical
    Exit Sub
End If
txtAccPath.Text = CommonDialog1.FileName
Call BreakIT
End Sub

Private Sub txtIniPath_DblClick()
CommonDialog1.FileName = ""
CommonDialog1.Filter = "*.INI|*.INI"
CommonDialog1.Flags = cdlOFNFileMustExist
If Trim(txtIniPath.Text) = "" Then
    CommonDialog1.InitDir = App.path
Else
    CommonDialog1.InitDir = Left(txtIniPath.Text, 3)
End If
CommonDialog1.ShowOpen
If Trim(CommonDialog1.FileName) = "" Then Exit Sub
If Dir(CommonDialog1.FileName) = "" Then
    MsgBox "Initialization File is Missing", vbCritical
    Exit Sub
End If
txtIniPath.Text = CommonDialog1.FileName
Call ShowFileSettings
End Sub

Private Sub txtSHL_GotFocus()
Call GF(txtSHL)
End Sub

Private Sub ShowDefaultSettings()
'' Procedure to show default settings
On Error Resume Next
txtIniPath.Text = App.path & "\Data\TimeHR.ini"
'' Company Name
txtCom.Text = InVar.strCOM
'' Demo Version
chkVer.Value = CByte(InVar.blnVerType)
'' Employee Limit
txtEmp.Text = InVar.lngEmp
'' Net Type
chkUser.Value = CByte(InVar.blnNetType)
'' User Limit
txtUser.Text = InVar.bytUse
'' Company Limit
txtComL.Text = InVar.bytCom
'' Version Number
cboVer.ListIndex = CByte(InVar.strVer) - 2
'' Lock date
txtY.Text = Left(InVar.strLok, 2)
txtM.Text = Mid(InVar.strLok, 3, 2)
txtD.Text = Right(InVar.strLok, 2)
'' Assum
chkAssum.Value = CByte(InVar.blnAssum)
'' Back End
cboBack.ListIndex = CByte(InVar.strSer) - 1
'' Wen Enabled
chkWeb.Value = CByte(InVar.blnWeb)
'' User Name
txtUName.Text = InVar.strUser
'' Psssword
txtPass.Text = InVar.strPass
End Sub

Private Sub ShowFileSettings()
On Error GoTo ERR_P
Dim strDataL As String, strData As String
Dim strArrDec() As String, strArrTmp() As String, strLokDate As String
strDataL = "": strData = ""
Open Trim(txtIniPath.Text) For Binary As #2
strData = Trim(Input(LOF(2), #2))
Close #2
strArrTmp = Split(strData, vbCrLf)
strArrTmp(1) = DEncryptDat(strArrTmp(1), 2)
If strArrTmp(0) = strArrTmp(1) Then
    strArrTmp(0) = DEncryptDat(strArrTmp(0), 1)
    strArrDec = Split(strArrTmp(0), "|")
    '' Company Name
    txtCom.Text = strArrDec(0)
    '' Demo Version
    chkVer.Value = CByte(strArrDec(1))
    '' Employee Limit
    txtEmp.Text = strArrDec(2)
    '' Net Type
    chkUser.Value = CByte(strArrDec(3))
    '' User Limit
    txtUser.Text = strArrDec(4)
    '' Company Limit
    txtComL.Text = strArrDec(5)
    '' Version Number
    cboVer.ListIndex = CByte(strArrDec(6)) - 2
    '' Lock Date
    strLokDate = strArrDec(7)
    '' Day
    txtD.Text = Right(strLokDate, 2)
    txtM.Text = Mid(strLokDate, 3, 2)
    txtY.Text = Left(strLokDate, 2)
    '' Month
    '' Year
    '' Assum
    chkAssum.Value = CByte(strArrDec(8))
    '' Back End
    cboBack.ListIndex = CByte(strArrDec(9)) - 1
    '' Web Enabled
    chkWeb.Value = CByte(strArrDec(10))
    '' User Name
    txtUName.Text = strArrDec(11)
    '' Password
    txtPass.Text = strArrDec(12)
Else
    MsgBox "The File has Been tampered Outside this Application", vbCritical
End If
Exit Sub
ERR_P:
    ShowError ("ShowFileSettings::" & Me.Caption)
End Sub

Private Sub BreakIT()
On Error GoTo ERR_P
Dim start, Letter, NR, AllText, TempText
NR = "134,251,236,55,93,68,156,250,198,94,40,230,19"
NR = Split(NR, ",")
Open Trim(txtAccPath.Text) For Binary As #2
AllText = Trim(Input(80, #2))
Close #2
TempText = Mid(AllText, 67, 13)
Letter = ""
For start = 1 To 13
    If Asc(Mid(TempText, start, 1)) <> NR(start - 1) Then
        Letter = Letter & Chr(Asc(Mid(TempText, start, 1)) Xor NR(start - 1))
    End If
Next
txtAccPass.Text = Letter
Exit Sub
ERR_P:
    ShowError ("BreakIt::" & Me.Caption)
End Sub
