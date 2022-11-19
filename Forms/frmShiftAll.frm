VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmShiftAll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Shift Details for all"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSet 
      Caption         =   "OK"
      Height          =   405
      Left            =   210
      TabIndex        =   11
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   405
      Left            =   2430
      TabIndex        =   12
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Frame frSet 
      Caption         =   "Set Details For"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   4725
      Begin VB.CheckBox chkDaily 
         Caption         =   "Details Regarding Daily Process"
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
         Left            =   90
         TabIndex        =   3
         Top             =   1230
         Width           =   4485
      End
      Begin VB.CheckBox chkAWO 
         Caption         =   "Details of  Additional Week Off"
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
         Left            =   90
         TabIndex        =   2
         Top             =   900
         Width           =   4485
      End
      Begin VB.CheckBox chkWO 
         Caption         =   "Details of Week Off"
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
         Left            =   90
         TabIndex        =   1
         Top             =   570
         Width           =   4485
      End
      Begin VB.CheckBox chkSInfo 
         Caption         =   "Details of Shift Info"
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
         Left            =   90
         TabIndex        =   0
         Top             =   240
         Width           =   4485
      End
   End
   Begin VB.Frame frEmp 
      Caption         =   "Employee Selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   14
      Top             =   1590
      Width           =   4725
      Begin VB.CommandButton cmdUA 
         Caption         =   "Unselect All"
         Height          =   435
         Left            =   3390
         TabIndex        =   10
         Top             =   3000
         Width           =   1305
      End
      Begin VB.CommandButton cmdSA 
         Caption         =   "Select All"
         Height          =   435
         Left            =   3390
         TabIndex        =   9
         Top             =   2580
         Width           =   1305
      End
      Begin VB.CommandButton cmdUR 
         Caption         =   "Unselect Range"
         Height          =   465
         Left            =   3390
         TabIndex        =   8
         Top             =   1980
         Width           =   1305
      End
      Begin VB.CommandButton cmdSR 
         Caption         =   "Select Range"
         Height          =   435
         Left            =   3390
         TabIndex        =   7
         Top             =   1560
         Width           =   1305
      End
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   2925
         Left            =   30
         TabIndex        =   13
         ToolTipText     =   "Click to Toggle Selection"
         Top             =   1020
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   5159
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSForms.ComboBox cboTo 
         Height          =   315
         Left            =   2790
         TabIndex        =   6
         Top             =   630
         Width           =   1215
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2143;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1500;4500"
      End
      Begin MSForms.ComboBox cboFrom 
         Height          =   315
         Left            =   930
         TabIndex        =   5
         Top             =   630
         Width           =   1215
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2143;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1500;4500"
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   2520
         TabIndex        =   17
         Top             =   690
         Width           =   195
      End
      Begin VB.Label lblEmp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From "
         Height          =   195
         Left            =   510
         TabIndex        =   16
         Top             =   690
         Width           =   390
      End
      Begin VB.Label lblDeptCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   330
         Width           =   825
      End
      Begin MSForms.ComboBox cboDept 
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   3105
         VariousPropertyBits=   612390939
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5477;556"
         TextColumn      =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmShiftAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset
Dim strArr, strSelEmp As String

Private Sub Form_Load()
strArr = Split(strRotPass, "|")
Call SetFormIcon(Me)            '' Sets the Forms Icon
Call RetCaption                 '' Sets the Captios on the Form Controls
Call FillCombos                 '' Fill Month and Year ComboBoxes
If strCurrentUserType <> HOD Then cboDept.Text = "ALL"
End Sub

Private Sub RetCaption()
On Error GoTo ERR_P
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "select * from NewCaptions where ID like '70%'", ConMain, adOpenStatic, adLockReadOnly
Me.Caption = NewCaptionTxt("70001", adrsC)

frSet.Caption = NewCaptionTxt("70002", adrsC) ''Set Details For
''Check captions
chkSInfo.Caption = NewCaptionTxt("70003", adrsC)    ''Details of Shift Info
chkWO.Caption = NewCaptionTxt("70004", adrsC)       ''Details of Week Off
chkAWO.Caption = NewCaptionTxt("70005", adrsC)      ''Details of  Additional Week Off
chkDaily.Caption = NewCaptionTxt("70006", adrsC)    ''Details Regarding Daily Process
''Set other captions
Call SetCritLabel(lblDeptCap)
'' MSF3
MSF3.ColWidth(1) = MSF3.ColWidth(1) * 2.2
Exit Sub
ERR_P:
    ShowError ("RetCaption :: " & Me.Caption)
End Sub

Private Sub FillCombos()
Call SetCritCombos(cboDept)
'If strCurrentUserType <> HOD Then cbodept.AddItem "ALL"
End Sub

Private Sub cmdSA_Click()
Call SelUnselAll(SELECTED_COLOR, MSF3)
End Sub

Private Sub cmdSR_Click()
Call SelUnsel(SELECTED_COLOR, MSF3, cboFrom, cboTo)
End Sub

Private Sub cmdUA_Click()
Call SelUnselAll(UNSELECTED_COLOR, MSF3)
End Sub

Private Sub cmdUR_Click()
Call SelUnsel(UNSELECTED_COLOR, MSF3, cboFrom, cboTo)
End Sub

Private Sub cboFrom_Click()
    If cboFrom.ListIndex >= 0 Then
        If cboTo.ListCount > 0 Then cboTo.ListIndex = cboFrom.ListIndex
    End If
End Sub

Private Sub cboDept_Click()
On Error GoTo ERR_P
If cboDept.ListIndex < 0 Then Exit Sub               '' If No Department
Call FillEmpCombos
Call SelUnselAll(UNSELECTED_COLOR, MSF3)

With MSF3
    .TextMatrix(0, 0) = "Code"
    .TextMatrix(0, 1) = "Name"
End With
Exit Sub
ERR_P:
    ShowError ("Department::" & Me.Caption)
End Sub

Private Sub FillEmpCombos()
On Error GoTo ERR_P
Dim adrsTmp As New ADODB.Recordset, intEmpCnt As Integer, intTmpCnt As Integer
Dim strDeptTmp As String, strTempforCF As String
Dim strArrEmp() As String


If cboDept.Text = "ALL" Then
    strDeptTmp = "ALL"
Else
    strDeptTmp = cboDept.List(cboDept.ListIndex, 1)
    strDeptTmp = EncloseQuotes(strDeptTmp)
End If

Select Case UCase(Trim(strDeptTmp))
    Case "", "ALL"
        strTempforCF = "select Empcode,name from empmst order by Empcode"     'Empcode,name
    Case Else
        ''For Mauritius 09-08-2003
        ''Original ->strTempforCF = "select Empcode,name from empmst Where " & SELCRIT & "=" & _
        strDeptTmp & " order by Empcode"                               'Empcode,name
        If strCurrentUserType = HOD Then
            strTempforCF = "select Empcode,name from empmst " & strCurrData & " And Empmst." & SELCRIT & "=" & _
                strDeptTmp & " order by Empcode"    'Empcode,name
        Else
            strTempforCF = "select Empcode,name from empmst Where " & SELCRIT & "=" & strDeptTmp & " order by Empcode"                    'Empcode,name
        End If
End Select

If adrsTmp.State = 1 Then adrsTmp.Close
adrsTmp.Open strTempforCF, ConMain, adOpenStatic, adLockReadOnly
If (adrsTmp.EOF And adrsTmp.BOF) Then
    cboFrom.clear
    cboTo.clear
    MSF3.Rows = 1
    Exit Sub
End If
intEmpCnt = adrsTmp.RecordCount
intTmpCnt = intEmpCnt
MSF3.Rows = intEmpCnt + 1
ReDim strArrEmp(intTmpCnt - 1, 1)
For intEmpCnt = 0 To intTmpCnt - 1
    strArrEmp(intEmpCnt, 0) = adrsTmp(0)
    strArrEmp(intEmpCnt, 1) = adrsTmp(1)
    MSF3.TextMatrix(intEmpCnt + 1, 0) = adrsTmp(0)
    MSF3.TextMatrix(intEmpCnt + 1, 1) = adrsTmp(1)
    adrsTmp.MoveNext
Next
cboFrom.List = strArrEmp
cboTo.List = strArrEmp
cboFrom.ListIndex = 0
cboTo.ListIndex = cboTo.ListCount - 1
Erase strArrEmp
Exit Sub
ERR_P:
    ShowError ("Fill Employee Combos :: " & Me.Caption)
End Sub

Private Sub MSF3_Click()
If MSF3.Rows = 1 Then Exit Sub
If MSF3.CellBackColor = SELECTED_COLOR Then
    With MSF3
        .Col = 0
        .CellBackColor = UNSELECTED_COLOR
        .Col = 1
        .CellBackColor = UNSELECTED_COLOR
    End With
Else
    With MSF3
        .Col = 0
        .CellBackColor = SELECTED_COLOR
        .Col = 1
        .CellBackColor = SELECTED_COLOR
    End With
End If
End Sub

Private Sub MSF3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then Call MSF3_Click
End Sub

Private Sub cmdSet_Click()
If Not ValidateModMater Then Exit Sub
If Not SaveModMaster Then Exit Sub
MsgBox NewCaptionTxt("70007", adrsC), vbInformation
End Sub

Private Function ValidateModMater() As Boolean
On Error GoTo ERR_P
Dim intTmp As Integer
intTmp = 0
If MSF3.Rows = 1 Then
    MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation
    Exit Function
End If
MSF3.Col = 0: strSelEmp = ""
For i = 1 To MSF3.Rows - 1
    MSF3.row = i
    If MSF3.CellBackColor = SELECTED_COLOR Then
        intTmp = intTmp + 1
        strSelEmp = strSelEmp & "'" & MSF3.Text & "'" & ","
    End If
Next
If intTmp = 0 Then
    MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation
    Exit Function
End If
strSelEmp = Left(strSelEmp, Len(strSelEmp) - 1)
ValidateModMater = True
Exit Function
ERR_P:
    ShowError ("FoundTable :: " & Me.Caption)
    ValidateModMater = False
End Function

Private Function SaveModMaster() As Boolean     '' Saves Data After Validations
On Error GoTo ERR_P
SaveModMaster = True
Dim strSInfo As String, strWO As String, strAWO As String, strDaily As String
''Week off details
strWO = " " & strKOff & "='" & Shft.WO & "'"
''Additional Week off details
strAWO = " Off2='" & Shft.WO1 & "',WO_1_3='" & Shft.WO2 & "',WO_2_4='" & Shft.WO3 & "'"
''Details regarding Daily process
strDaily = " Update EmpMst Set WOHLAction=" & Shft.WOHLAction & ",Action3Shift='" & _
            Shft.Action3Shift & "',AutoForPunch=" & IIf(Shft.AutoOnPunch, 1, 0) & _
            ",ActionBlank='" & Shft.ActionBlank & "' "
''Shift info
If Shft.ShiftType = "F" Then
     strSInfo = " STyp='F',F_Shf='" & Shft.ShiftCode & "',SCode='100'"
Else
     strSInfo = " STyp='R',F_Shf='',SCode='" & Shft.ShiftCode & "'"
End If
strSInfo = strSInfo & " ,Shf_Date=" & strDTEnc & DateCompStr(Shft.startdate) & strDTEnc

If chkSInfo.Value = 0 Then strSInfo = ""
If chkWO.Value = 0 Then
    strWO = ""
Else
    If Trim(strSInfo) <> "" Then
        strSInfo = strSInfo & ", " & strWO
    Else
        strSInfo = strWO
    End If
End If
If chkAWO.Value = 0 Then
    strAWO = ""
Else
    If Trim(strSInfo) <> "" Then
        strSInfo = strSInfo & ", " & strAWO
    Else
        strSInfo = strAWO
    End If
End If
If chkDaily.Value = 0 Then strDaily = ""

'' For Shift Details
If strSInfo <> "" Then ConMain.Execute " Update EmpMst Set " & _
        strSInfo & " Where Empcode in  (" & strSelEmp & ")"
'' For Details Regarding Daily Processing
If strDaily <> "" Then ConMain.Execute strDaily & _
        " Where Empcode in  (" & strSelEmp & ")"

''
Exit Function
ERR_P:
    ShowError (Err.Description & " :: " & Me.Caption)
    SaveModMaster = False
End Function

Private Sub cmdExit_Click()
Unload Me
End Sub
