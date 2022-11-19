VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExpTimeCard 
   Caption         =   "Time Card"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   2040
      Visible         =   0   'False
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   661
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
   End
   Begin VB.ListBox lstFields 
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   20
      Top             =   6000
      Width           =   1485
   End
   Begin VB.ListBox lstFields 
      Height          =   255
      Index           =   0
      Left            =   3960
      TabIndex        =   19
      Top             =   6000
      Width           =   1245
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   435
      Left            =   2040
      TabIndex        =   13
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   435
      Left            =   3720
      TabIndex        =   12
      Top             =   1560
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5880
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frSap 
      Height          =   795
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   5925
      Begin VB.TextBox txtFrom 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "D"
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox txtTo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4470
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "D"
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblFrPeri 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Export for the period from"
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
         Left            =   120
         TabIndex        =   18
         Top             =   270
         Width           =   2460
      End
      Begin VB.Label lblToPeri 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   3990
         TabIndex        =   17
         Top             =   270
         Width           =   300
      End
   End
   Begin VB.Frame frEmp 
      Caption         =   "Employee Selection"
      Height          =   5025
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7965
      Begin VB.CommandButton cmdUA 
         Caption         =   "U&nselect All"
         Height          =   435
         Left            =   4110
         TabIndex        =   4
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSA 
         Caption         =   "&Select All"
         Height          =   435
         Left            =   4110
         TabIndex        =   3
         Top             =   2100
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdUR 
         Caption         =   "&Unselect Range"
         Height          =   465
         Left            =   4110
         TabIndex        =   2
         Top             =   1500
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSR 
         Caption         =   "Select &Range"
         Height          =   435
         Left            =   4110
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2625
         Left            =   -600
         TabIndex        =   5
         ToolTipText     =   "Click to Toggle Selection"
         Top             =   2520
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4630
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.Label lblDeptCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         Height          =   195
         Left            =   540
         TabIndex        =   11
         Top             =   270
         Width           =   825
      End
      Begin MSForms.ComboBox cboDept 
         Height          =   315
         Left            =   1830
         TabIndex        =   10
         Top             =   210
         Width           =   3315
         VariousPropertyBits=   612390939
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5847;556"
         TextColumn      =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboTo 
         Height          =   315
         Left            =   3660
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   1365
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2408;556"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboFrom 
         Height          =   315
         Left            =   1710
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   1365
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2408;556"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T&o"
         Height          =   195
         Left            =   3150
         TabIndex        =   7
         Top             =   1140
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fro&m"
         Height          =   195
         Left            =   900
         TabIndex        =   6
         Top             =   1140
         Visible         =   0   'False
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmExpTimeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTempPath Lib "KERNEL32" Alias _
 "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


Dim adrsForm As New ADODB.Recordset
Dim strTrnData(1 To 12) As String, strTrnFields(1 To 12) As String
Dim strLvTrnData() As String, strLvTrnFields() As String
Dim strfile1 As String, strFile2 As String
Public ExpType As String
Private Type EmpPunch
    Empcode As String
    PunchTime As Date
    InOut As String
    PDate As Date
End Type
Dim rs As New ADODB.Recordset
Dim sql As String
Dim Ein As Boolean
Dim i As Integer, j As Integer
Dim CurrentInout As String
Dim InwO As Integer
Dim OutwI As Integer
Dim rngIn As Date
Dim rngOut As Date
Dim rngDiff As Date
Dim HoldSeconds As Long


Private Function GetExportData() As String
    On Error GoTo ERR_P
      '' Data String
      Dim strLinedata As String, strData As String
      '' Temporary Current Arrays
      Dim strTmpDataArr() As String, strTmpFieldsArr() As String
      '' Temporary Variables
      Dim bytTmp As Byte, strTmp As String
      
      strfile1 = MakeName(MonthName(Month(txtFrom.Text)), Year(txtFrom.Text), "trn")
      strFile2 = MakeName(MonthName(Month(txtTo.Text)), Year(txtTo.Text), "trn")
      
     strSql = "SELECT " & strfile1 & ".*, empmst.name, empmst.card, deptdesc.desc  FROM " & strfile1 & " , empmst, deptdesc"
     strSql = strSql + " WHERE " & strfile1 & ".Empcode=[empmst].[Empcode] AND empmst.dept=[deptdesc].[dept] And"
     strSql = strSql + " empmst.dept=[deptdesc].[dept] AND " & strfile1 & ".Date >= #" & Format(txtFrom.Text, "dd/MMM/yyyy") & "# AND " & strfile1 & ".Date <= #" & Format(txtTo.Text, "dd/MMM/yyyy") & "# "

    If strfile1 <> strFile2 Then
        strSql = strSql + " UNION SELECT " & strFile2 & ".*, empmst.name, empmst.card, deptdesc.desc  FROM " & strFile2 & " , empmst, deptdesc"
        strSql = strSql + " WHERE " & strFile2 & ".Empcode=[empmst].[Empcode] AND empmst.dept=[deptdesc].[dept] And"
        strSql = strSql + " empmst.dept=[deptdesc].[dept] AND " & strFile2 & ".Date >= #" & Format(txtFrom.Text, "dd/MMM/yyyy") & "# AND " & strFile2 & ".Date <= #" & Format(txtTo.Text, "dd/MMM/yyyy") & "# "
    End If
     
      If cboDept.Text <> "ALL" Then
         strSql = strSql + "  AND deptdesc.dept = " & cboDept.List(cboDept.ListIndex, 1)
      End If
    
    If adrsForm.State = 1 Then adrsForm.Close
     adrsForm.Open strSql, ConMain, adOpenDynamic, adLockOptimistic
      
     If adrsForm.RecordCount < 1 Then
       MsgBox "No Record found"
       Exit Function
     End If
     adrsForm.Sort = "Empcode, Date"
    strTmpDataArr = GetCurrentArray(1)
    strTmpFieldsArr = GetCurrentArray(2)
    adrsForm.MoveFirst
    For bytTmp = 1 To 12
        strTmp = strTmpDataArr(bytTmp)
        If strTmp <> "" Then
                strData = strData & strTmp & vbTab
       Else
                 strData = strData & vbTab
       End If
   Next
   strData = strData & vbCrLf



    Dim i As Integer

    
    If bytRepMode = 1 Then
         
        For i = 0 To adrsForm.RecordCount - 1
             With adrsForm
             strLinedata = .Fields("EmpCode") & vbTab & .Fields("Name") & vbTab & .Fields("card") & vbTab & .Fields("Desc") & vbTab & Format(.Fields("Date"), "dd/MMM/yyyy") & vbTab & Format(.Fields("Arrtim"), "0.00") & vbTab & Format(IIf(Val(.Fields("Deptim")) > 24, TimDiff(.Fields("Deptim"), 24), .Fields("Deptim")), "0.00") & vbTab & Format(.Fields("WrkHrs"), "0.00") & vbTab & Format(.Fields("Ovtim"), "0.00") & vbTab & Format(.Fields("LateHrs"), "0.00") & vbTab & Format(.Fields("EarlHrs"), "0.00") & vbTab & .Fields("Presabs")
              strData = strData & strLinedata & vbCrLf
             End With
             adrsForm.MoveNext
        Next
    Else
    
    
         For i = 0 To adrsForm.RecordCount - 1
             With adrsForm
             strLinedata = .Fields("EmpCode") & vbTab & .Fields("Name") & vbTab & .Fields("card") & vbTab & .Fields("Desc") & vbTab & Format(.Fields("Date"), "dd/MMM/yyyy") & vbTab & .Fields("Arrtim") & vbTab & IIf(Val(.Fields("Deptim")) > 24, TimDiff(.Fields("Deptim"), 24), .Fields("Deptim")) & vbTab & .Fields("WrkHrs") & vbTab & .Fields("Ovtim") & vbTab & .Fields("LateHrs") & vbTab & .Fields("EarlHrs") & vbTab & .Fields("Presabs")
              strData = strData & strLinedata & vbCrLf
             End With
             adrsForm.MoveNext
        Next
    End If

   GetExportData = strData
   Exit Function
ERR_P:
      ShowError ("GetExportData::" & Me.Caption & "Erl:" & Erl)

End Function

Private Function GetCurrentArray(Optional bytDataFields As Byte = 1) As String()
'' On Error Resume Next
Select Case bytDataFields
    Case 1      '' Data
        Select Case bytMode
            Case 1      '' Daily Data
                GetCurrentArray = strTrnData
            Case 2      '' Monthly Data
                GetCurrentArray = strLvTrnData
        End Select
    Case 2      '' Fields
        Select Case bytMode
            Case 1      '' Daily Data
                GetCurrentArray = strTrnFields
            Case 2      '' Monthly Data
                GetCurrentArray = strLvTrnFields
        End Select
End Select
End Function


Private Function GetFieldName(ByRef strTmpDataArr() As String, _
ByRef strTmpFieldsArr() As String, ByVal strTmp As String) As String
'' On Error Resume Next
Dim bytTmp As Byte
For bytTmp = 1 To UBound(strTmpDataArr)
    If Trim(UCase(strTmp)) = Trim(UCase(strTmpDataArr(bytTmp))) Then
        strTmp = strTmpFieldsArr(bytTmp)
        Exit For
    End If
Next
GetFieldName = strTmp
End Function

Private Sub FillCombos()
On Error GoTo ERR_P
Dim intTmp As Integer

    Call SetCritCombos(cboDept)

Exit Sub
ERR_P:
    ShowError ("FillCombos::" & Me.Caption)
End Sub

Private Sub LoadSpecifics()
On Error GoTo ERR_P

Call FillCombos

bytMode = 1
Call FillFieldsArrays
'' Set Recordset
If adrsForm.State = 1 Then adrsForm.Close
With adrsForm
    .ActiveConnection = ConMain
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
End With
'' Set Common Dialog Control
With CD1
    '.Filter = "Text Files|*.Txt|CSV (Comma Separated Value)|*.CSV"     'commented by  25-05
    .Flags = cdlOFNOverwritePrompt
End With
Exit Sub
ERR_P:
    ShowError ("LoadSpecifics::" & Me.Caption)
    Resume Next
End Sub

Private Sub FillFieldsArrays()
On Error Resume Next
Dim bytTmp As Byte
For bytTmp = 1 To 12
    '' Display Names
    strTrnData(bytTmp) = Choose(bytTmp, "Code", "Employee Name", "Card Code", "Department Name", "Date", "In Time", "Out Time", _
    "Worked Hrs.", "O.T. Hours", "Late Hours", "Early Hours", "Status")
    '' Actual Field Names
    strTrnFields(bytTmp) = Choose(bytTmp, "Empcode", "Name", "card", "Desc", "Date", "Arrtim", "Deptim", _
    "WrkHrs", "Ovtim", "LateHrs", "EarlHrs", "Presabs")
Next

End Sub

Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub cmdExport_Click()
  Dim strmain As String
'  If Not FindDataFile Then Exit Sub
'  If Not CheckEmployees Then Exit Sub
'  If Not DataFound Then Exit Sub
  
    CD1.Filter = "Excel Files|*.xlsx;*.xlsx"
    CD1.ShowSave
    If Trim(CD1.FileName) = "" Then Exit Sub
            strmain = GetINOUT
    Select Case UCase(ExpType)
    Case "TIMECARD"
        strmain = GetExportData
    Case "INOUT", "EXTENDED"
        strmain = GetINOUT
    Case "BIOINOUT"
        strmain = GetBioData()
    Case "ABSENTEE"
        strmain = GetBioAttendance()
    End Select
    If strmain = "False" Then Exit Sub
     strmain = GetINOUT
    SaveExportData (strmain)
    MsgBox "Data  Exported"
End Sub

Private Function GetBioData() As String
    On Error GoTo ERR_P
      '' Data String
      Dim bytTmp As Byte, strTmp As String
      Dim strData As String
      Dim strLinedata As String
      Dim rs As New ADODB.Recordset
      
      strfile1 = MakeName(MonthName(Month(txtFrom.Text)), Year(txtFrom.Text), "trnB")
      strFile2 = MakeName(MonthName(Month(txtTo.Text)), Year(txtTo.Text), "trnB")
      If Not (FindTable(strfile1) Or FindTable(strFile2)) Then
        MsgBox "Processing Not Done For Selected Month"
            GetBioData = "False"
            Exit Function
      End If
     strSql = "SELECT " & strfile1 & ".*, empmst.name, empmst.card, deptdesc.desc  FROM " & strfile1 & " , empmst, deptdesc"
     strSql = strSql + " WHERE " & strfile1 & ".Empcode=[empmst].[Empcode] AND empmst.dept=[deptdesc].[dept] And"
     strSql = strSql + " empmst.dept=[deptdesc].[dept] AND " & strfile1 & ".Date >= #" & Format(txtFrom.Text, "dd/MMM/yyyy") & "# AND " & strfile1 & ".Date <= #" & Format(txtTo.Text, "dd/MMM/yyyy") & "# "

    If strfile1 <> strFile2 Then
        strSql = strSql + " UNION SELECT " & strFile2 & ".*, empmst.name, empmst.card, deptdesc.desc  FROM " & strFile2 & " , empmst, deptdesc"
        strSql = strSql + " WHERE " & strFile2 & ".Empcode=[empmst].[Empcode] AND empmst.dept=[deptdesc].[dept] And"
        strSql = strSql + " empmst.dept=[deptdesc].[dept] AND " & strFile2 & ".Date >= #" & Format(txtFrom.Text, "dd/MMM/yyyy") & "# AND " & strFile2 & ".Date <= #" & Format(txtTo.Text, "dd/MMM/yyyy") & "# "
    End If
     
      If cboDept.Text <> "ALL" Then
         strSql = strSql + "  AND deptdesc.dept = " & cboDept.List(cboDept.ListIndex, 1)
      End If
    
    If rs.State = 1 Then rs.Close
     rs.Open strSql, ConMain, adOpenDynamic, adLockOptimistic
      
     If rs.RecordCount < 1 Then
       MsgBox "No Record found"
       GetBioData = False
       Exit Function
     End If
     rs.Sort = "Empcode, Date"
   strData = "Date" & vbTab & "Emp Code" & vbTab & "Emp Name" & vbTab & "Card Code" & vbTab & "Absent/Present" & vbTab & "In Time" & vbTab & "Out Time"
   strData = strData & vbCrLf

    Dim i As Integer
    
    If bytRepMode = 1 Then
         
        For i = 0 To rs.RecordCount - 1
             With rs
             strLinedata = Format(.Fields("Date"), "dd/MMM/yyyy") & vbTab & .Fields("EmpCode") & vbTab & .Fields("Name") & vbTab & .Fields("card") & vbTab & .Fields("Presabs") & vbTab & .Fields("Arrtim") & vbTab & IIf(Val(.Fields("Deptim")) > 24, TimDiff(.Fields("Deptim"), 24), .Fields("Deptim"))
              strData = strData & strLinedata & vbCrLf
             End With
             rs.MoveNext
        Next
    Else
    
    
         For i = 0 To rs.RecordCount - 1
             With rs
             strLinedata = Format(.Fields("Date"), "dd/MMM/yyyy") & vbTab & .Fields("EmpCode") & vbTab & .Fields("Name") & vbTab & .Fields("card") & vbTab & .Fields("Presabs") & vbTab & .Fields("Arrtim") & vbTab & IIf(Val(.Fields("Deptim")) > 24, TimDiff(.Fields("Deptim"), 24), .Fields("Deptim"))
              strData = strData & strLinedata & vbCrLf
             End With
             rs.MoveNext
        Next
    End If

   GetBioData = strData
   Exit Function
ERR_P:
      ShowError ("GetBioData" & Me.Caption & "Erl:" & Erl)

End Function

Private Function GetBioAttendance() As String
    On Error GoTo ERR_P
      '' Data String
      Dim bytTmp As Byte, strTmp As String
      Dim strData As String, strLinedata As String
      Dim strFile3 As String, strFile4 As String
      
      strfile1 = MakeName(MonthName(Month(txtFrom.Text)), Year(txtFrom.Text), "trn")
      strFile2 = MakeName(MonthName(Month(txtTo.Text)), Year(txtTo.Text), "trn")
      strFile3 = MakeName(MonthName(Month(txtFrom.Text)), Year(txtFrom.Text), "trnB")
      strFile4 = MakeName(MonthName(Month(txtTo.Text)), Year(txtTo.Text), "trnB")
      
      If Not (FindTable(strfile1) And FindTable(strFile2) And FindTable(strFile3) And FindTable(strFile4)) Then
        MsgBox "Processing Not Done For Selected Month"
            GetBioAttendance = "False"
            Exit Function
      End If
        strSql = "SELECT " & strfile1 & ".Date, Empmst.Empcode, empmst.name, " & strfile1 & ".PresAbs  FROM " & strfile1 & " , empmst, deptdesc"
        strSql = strSql + " WHERE " & strfile1 & ".Empcode=[empmst].[Empcode] AND empmst.dept=[deptdesc].[dept] And"
        strSql = strSql + " empmst.dept=[deptdesc].[dept] AND " & strfile1 & ".Date >= #" & Format(txtFrom.Text, "dd/MMM/yyyy") & "# AND " & strfile1 & ".Date <= #" & Format(txtTo.Text, "dd/MMM/yyyy") & "# "

        If strfile1 <> strFile2 Then
            strSql = strSql + " UNION SELECT " & strFile2 & ".Date, Empmst.Empcode, empmst.name, " & strFile2 & ".PresAbs  FROM " & strFile2 & " , empmst, deptdesc"
            strSql = strSql + " WHERE " & strFile2 & ".Empcode=[empmst].[Empcode] AND empmst.dept=[deptdesc].[dept] And"
            strSql = strSql + " empmst.dept=[deptdesc].[dept] AND " & strFile2 & ".Date >= #" & Format(txtFrom.Text, "dd/MMM/yyyy") & "# AND " & strFile2 & ".Date <= #" & Format(txtTo.Text, "dd/MMM/yyyy") & "# "
        End If
      If cboDept.Text <> "ALL" Then
         strSql = strSql + "  AND deptdesc.dept = " & cboDept.List(cboDept.ListIndex, 1)
      End If

      Dim rs As New ADODB.Recordset
      If rs.State = 1 Then rs.Close
        rs.Open strSql, ConMain, adOpenDynamic, adLockOptimistic
        TruncateTable ("ImportTbl")
       
    If rs.RecordCount > 0 Then
        For i = 0 To rs.RecordCount - 1
           ConMain.Execute " INSERT INTO ImportTbl (empcode, d1,d2,d3) values ('" & rs.Fields(1).Value & "', '" & Format(rs.Fields(0).Value, "dd/MMM/YYYY") & "','" & rs.Fields(2).Value & "', '" & rs.Fields(3) & "')"
           rs.MoveNext
        Next
    End If
    
    Dim ImRs As New ADODB.Recordset
    Dim Searchrs As New ADODB.Recordset
    
    strSql = Replace(strSql, "Trn", "TrnB")
    If rs.State = 1 Then rs.Close
    rs.Open strSql, ConMain, adOpenDynamic, adLockOptimistic
 
    For i = 0 To rs.RecordCount - 1
       If ImRs.State = 1 Then ImRs.Close
       sql = "Select * from Importtbl where d1 = '" & Format(rs.Fields("Date"), "dd/MMM/yyyy") & "' and  Empcode = '" & rs.Fields("Empcode") & "'"
       ImRs.Open sql, ConMain, adOpenDynamic, adLockOptimistic
       If ImRs.EOF Then
            ConMain.Execute " INSERT INTO ImportTbl (empcode, d1,d2,d4) values ('" & rs.Fields(1).Value & "', '" & Format(rs.Fields(0).Value, "dd/MMM/YYYY") & "','" & rs.Fields(2).Value & "', '" & rs.Fields(3) & "')"
        Else
            ConMain.Execute " Update ImportTbl Set D4 = '" & rs.Fields(3) & "' where empcode = '" & rs.Fields("Empcode") & "' and  d1 = '" & Format(rs.Fields("Date"), "dd/MMM/yyyy") & "'"
        End If
        rs.MoveNext
    Next
    
    If adrsForm.State = 1 Then adrsForm.Close
     adrsForm.Open "select * from ImportTbl", ConMain, adOpenDynamic, adLockOptimistic
      
     
   adrsForm.Sort = "Empcode, D1"
   strData = "Date" & vbTab & "Emp Code" & vbTab & "Emp Name" & vbTab & "T&A P/A" & vbTab & "Bio P/A"
   strData = strData & vbCrLf

     If adrsForm.RecordCount < 1 Then
       MsgBox "No Record found"
       GetBioAttendance = False
       Exit Function
    End If
    
    If bytRepMode = 1 Then
     
     
        For i = 0 To adrsForm.RecordCount - 1
             With adrsForm
             strLinedata = Format(adrsForm.Fields("d1"), "dd/MMM/yyyy") & vbTab & adrsForm.Fields("EmpCode") & vbTab & adrsForm.Fields("d2") & vbTab & adrsForm.Fields("d3") & vbTab & adrsForm.Fields("d4")
              strData = strData & strLinedata & vbCrLf
             End With
             adrsForm.MoveNext
        Next
    Else
    

         For i = 0 To adrsForm.RecordCount - 1
             With adrsForm
             strLinedata = Format(adrsForm.Fields("d1"), "dd/MMM/yyyy") & vbTab & adrsForm.Fields("EmpCode") & vbTab & adrsForm.Fields("d2") & vbTab & adrsForm.Fields("d3") & vbTab & adrsForm.Fields("d4")
              strData = strData & strLinedata & vbCrLf
             End With
             adrsForm.MoveNext
        Next
    End If
    

   GetBioAttendance = strData
   Exit Function
ERR_P:
      ShowError ("GetBioAttendance" & Me.Caption & "Erl:" & Erl)
Resume Next
End Function

Private Function GetExtended()

Dim EmpPunch As EmpPunch
Dim adrsEmp As New ADODB.Recordset

If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "Select EmpCode, name, Card from Empmst", ConMain
sql = "Select * from DailyPro order by Empcode, Dte, t_Punch, shift "

    If rs.State = 1 Then rs.Close
    rs.Open sql, ConMain
    Ein = False
    grd.Rows = 2
    i = 1
    For j = 0 To rs.RecordCount - 1
         If rs.EOF Then GoTo EOF
        EmpPunch.Empcode = rs.Fields("EmpCode").Value
        EmpPunch.PunchTime = rs.Fields("t_punch").Value
        EmpPunch.InOut = IIf(IsNull(rs.Fields("shift").Value), "", rs.Fields("shift").Value)
        EmpPunch.PDate = rs.Fields("Dte").Value
CONTINUE:
        rs.MoveNext
        If rs.EOF Then GoTo EOF
        
        If EmpPunch.Empcode = rs.Fields("EmpCode").Value And EmpPunch.PDate = rs.Fields("Dte").Value Then
             If EmpPunch.InOut = "O" Then 'And rs.Fields("shift").Value = "I" Then
                    If rs.Fields("shift").Value = "I" Then
                        adrsEmp.MoveFirst
                        adrsEmp.Find "Empcode =  '" & EmpPunch.Empcode & "'"
                        grd.TextMatrix(i, 2) = IIf(IsNull(adrsEmp("Name")), "", adrsEmp("Name"))
                        grd.TextMatrix(i, 1) = IIf(IsNull(adrsEmp("Card")), "", adrsEmp("Card"))
                        grd.TextMatrix(i, 0) = EmpPunch.Empcode
                        grd.TextMatrix(i, 3) = Format(EmpPunch.PDate, "dd/MMM/yyyy")
                        grd.TextMatrix(i, 4) = EmpPunch.PunchTime
                        grd.TextMatrix(i, 5) = rs.Fields("t_punch").Value
                        HoldSeconds = DateDiff("s", EmpPunch.PunchTime, rs.Fields("t_punch").Value)
                        rngDiff = Format(DateAdd("s", HoldSeconds, "00:00:00"), "hh:mm:ss")
                        rngOut = DateAdd("s", HoldSeconds, rngOut)
                        grd.TextMatrix(i, 6) = rngDiff ' DateAdd("s", HoldSeconds, rngOut)
                        grd.TextMatrix(i, 7) = rngDiff
                        grd.Rows = grd.Rows + 1
                        i = i + 1
                        grd.TextMatrix(i, 5) = "Total :"
                        grd.TextMatrix(i, 6) = rngOut
                        grd.TextMatrix(i, 7) = rngOut
                        
                        rs.MoveNext
                        If rs.EOF Then GoTo EOF
'                        rs.MoveNext
'                        If rs.EOF Then GoTo EOF
                        j = j + 2
                    Else
                        GoTo CONTINUE
                    End If
             End If
        Else
        
            grd.Rows = grd.Rows + 1
            i = i + 1
            rngOut = 0
            rs.MoveNext
EOF:
                        
        End If
        
    Next
    
    Dim strLinedata As String
    Dim strData As String
    
    strData = "Employee Card" & vbTab & "Employee Code" & vbTab & "Employee Name" & vbTab & "Date" & vbTab & "From Out Time" & vbTab & "To In Time" & vbTab & "Total Break" & vbTab & "Extended Break" & vbCrLf

   For i = 0 To grd.Rows - 1
        With grd
        strLinedata = .TextMatrix(i, 0) & vbTab & .TextMatrix(i, 1) & vbTab & .TextMatrix(i, 2) & vbTab & Format(.TextMatrix(i, 3), "dd/MMM/yyyy") & vbTab & Format(.TextMatrix(i, 4), "hh:nn:ss") & vbTab & Format(.TextMatrix(i, 5), "hh:nn:ss") & vbTab & Format(.TextMatrix(i, 6), "hh:nn:ss") & vbTab & Format(.TextMatrix(i, 7), "hh:nn:ss")
        If Len(Trim(strLinedata)) > 10 Then
        Debug.Print Len(Trim(strLinedata)) & "  " & strLinedata
            strData = strData & strLinedata & vbCrLf
         End If
        End With
   Next
   
      GetExtended = strData
      
End Function

Private Function GetINOUT()
    typDT.dtFrom = txtFrom.Text
    typDT.dtTo = txtTo.Text
    FillInstalltypes
    If Not AppendDataFile(frmExpTimeCard) Then Exit Function
    TruncateTable ("Dailypro")
    GetDataPunches ("(Select Empcode From Empmst)")
    
    If UCase(ExpType) = "INOUT" Then
        GetINOUT = CalculateInout
    Else
        GetINOUT = GetExtended
    End If
End Function


Private Function CalculateInout()

Dim EmpPunch As EmpPunch
Dim adsrsEmp As ADODB.Recordset

If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "Select EmpCode, name from Empmst", ConMain
sql = "Select * from DailyPro order by Empcode, Dte, t_Punch, shift "

    If rs.State = 1 Then rs.Close
    rs.Open sql, ConMain
    Ein = False
    grd.Rows = 2
i = 1
    For j = 0 To rs.RecordCount - 1
       
        EmpPunch.Empcode = rs.Fields("EmpCode").Value
        EmpPunch.PunchTime = rs.Fields("t_punch").Value
        EmpPunch.InOut = IIf(IsNull(rs.Fields("shift").Value), "", rs.Fields("shift").Value)
        EmpPunch.PDate = rs.Fields("Dte").Value
      
        rs.MoveNext
        If rs.EOF Then GoTo EOF
            
        If EmpPunch.Empcode = rs.Fields("EmpCode").Value And EmpPunch.PDate = rs.Fields("Dte").Value Then
             If EmpPunch.InOut = "I" Then
                    If CurrentInout = "I" Then InwO = InwO + 1
                    If Ein = False Then
                        grd.TextMatrix(i, 3) = EmpPunch.PunchTime
                        Ein = True
                    End If
                    CurrentInout = "I"
                ElseIf EmpPunch.InOut = "O" Then
                    If CurrentInout = "O" Then OutwI = OutwI + 1
                    grd.TextMatrix(i, 4) = EmpPunch.PunchTime
                    CurrentInout = "O"
             End If
                
             If EmpPunch.InOut = "I" And rs.Fields("shift").Value = "O" Then
                HoldSeconds = DateDiff("s", EmpPunch.PunchTime, rs.Fields("t_punch").Value)
                rngDiff = Format(DateAdd("s", HoldSeconds, "00:00:00"), "h:m:s")
                rngIn = DateAdd("s", HoldSeconds, rngIn)
'                rngDiff = TimDiff(Val(Replace(rs.Fields("t_punch").Value, ":", ".")), Val(Replace(EmpPunch.PunchTime, ":", ".")))
'                rngIn = TimAdd(rngIn, rngDiff)
              
              End If
            If EmpPunch.InOut = "O" And rs.Fields("shift").Value = "I" Then
                HoldSeconds = DateDiff("s", EmpPunch.PunchTime, rs.Fields("t_punch").Value)
                rngDiff = Format(DateAdd("s", HoldSeconds, "00:00:00"), "h:m:s")
                rngOut = DateAdd("s", HoldSeconds, rngOut)
                
'                rngDiff = TimDiff(Val(Replace(rs.Fields("t_punch").Value, ":", ".")), Val(Replace(EmpPunch.PunchTime, ":", ".")))
'                rngOut = TimAdd(rngOut, rngDiff)
              
              End If
           
        Else
EOF:
                If EmpPunch.InOut = "I" Then
                    If CurrentInout = "I" Then InwO = InwO + 1
                    If Ein = False Then
                        grd.TextMatrix(i, 3) = EmpPunch.PunchTime
                        Ein = True
                    End If
                    CurrentInout = "I"
                ElseIf EmpPunch.InOut = "O" Then
                    If CurrentInout = "O" Then OutwI = OutwI + 1
                    grd.TextMatrix(i, 4) = EmpPunch.PunchTime
                    CurrentInout = "O"
                End If

                If grd.TextMatrix(i, 4) <> "" Then
                    If grd.TextMatrix(i, 3) <> "" Then
'                        grd.TextMatrix(i, 5) = TimDiff(Val(Replace(grd.TextMatrix(i, 4), ":", ".")), Val(Replace(grd.TextMatrix(i, 3), ":", ".")))
                        HoldSeconds = DateDiff("s", CDate(grd.TextMatrix(i, 3)), CDate(grd.TextMatrix(i, 4)))
                        grd.TextMatrix(i, 5) = Format(DateAdd("s", HoldSeconds, "00:00:00"), "h:m:s")
                        
                    End If
                End If

                grd.TextMatrix(i, 0) = EmpPunch.Empcode
                
                adrsEmp.MoveFirst
                adrsEmp.Find "Empcode =  '" & EmpPunch.Empcode & "'"
                grd.TextMatrix(i, 1) = IIf(IsNull(adrsEmp("Name")), "", adrsEmp("Name"))
                 grd.TextMatrix(i, 2) = Format(EmpPunch.PDate, "dd/MMM/yyyy")
                 
                grd.TextMatrix(i, 6) = rngIn
                grd.TextMatrix(i, 7) = rngOut
                 
                grd.TextMatrix(i, 8) = InwO
                grd.TextMatrix(i, 9) = OutwI
                
                grd.Rows = grd.Rows + 1
                i = grd.Rows - 1
                Call Resett
CONTINUE:
                
        End If
        
    Next
    
    Dim strLinedata As String
    Dim strData As String
    
    strData = "Employee Code" & vbTab & "Employee Name" & vbTab & "Date" & vbTab & "First IN Time" & vbTab & "Last Out Time" & vbTab & "Total Day Time" & vbTab & "Total In" & vbTab & "Total Out" & vbTab & "In Without out" & vbTab & "Out Without In"

   For i = 0 To grd.Rows - 1
        With grd
        strLinedata = .TextMatrix(i, 0) & vbTab & .TextMatrix(i, 1) & vbTab & Format(.TextMatrix(i, 2), "dd/MMM/yyyy") & vbTab & Format(.TextMatrix(i, 3), "hh:nn") & vbTab & Format(.TextMatrix(i, 4), "hh:nn") & vbTab & Format(.TextMatrix(i, 5), "hh:nn") & vbTab & Format(.TextMatrix(i, 6), "hh:nn") & vbTab & Format(.TextMatrix(i, 7), "hh:nn") & vbTab & .TextMatrix(i, 8) & vbTab & .TextMatrix(i, 9)
        strData = strData & strLinedata & vbCrLf
        End With
   Next


   CalculateInout = strData
   
    
End Function

Private Sub Resett()
    InwO = 0
    OutwI = 0
    Ein = False
    CurrentInout = ""
    rngIn = 0
    rngOut = 0
End Sub

Private Function DataFound() As Boolean
On Error GoTo ERR_P
Dim bytTmp As Byte, bytCnt As Byte, strTmp As String, strlstdt As String
Dim dttodate As Date, dtfromdate As Date, cnt As Integer

strfile1 = MakeName(MonthName(Month(txtFrom.Text)), Year(txtFrom.Text), "trn")
strFile2 = MakeName(MonthName(Month(txtTo.Text)), Year(txtTo.Text), "trn")

If adrsForm.State = 1 Then adrsForm.Close
'    adrsForm.Open "Select name,div,Dept,location," & strfile1 & ".* from empmst," & strfile1 & " Where " & strfile1 & ".empcode=empmst.empcode and " & strfile1 & ".Empcode in " & strSelEmp & _
    " Order by " & strExpFileName & ".empcode, " & strKDate

If adrsForm.EOF Then
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
    Exit Function
End If

DataFound = True
Exit Function
ERR_P:
    ShowError ("DataFound::" & Me.Caption)
    Resume Next
End Function

Private Function FindDataFile() As Boolean
On Error GoTo ERR_P
'If Not FindTable() Then
'    If bytMode = 1 Then
'        MsgBox NewCaptionTxt("67024", adrsC) & cboMonth.Text
'        cboMonth.SetFocus
'    Else
'        MsgBox NewCaptionTxt("67025", adrsC) & cboYear.Text
'        cboYear.SetFocus
'    End If
'    Exit Function
'End If
FindDataFile = True
Exit Function
ERR_P:
    ShowError ("FindDataFile::" & Me.Caption)
End Function

Private Sub cmdSA_Click()
Call SelUnselAll(SELECTED_COLOR, MSF1)
End Sub

Private Sub cmdSR_Click()
Call SelUnsel(SELECTED_COLOR, MSF1, cboFrom, cboTo)
End Sub

Private Sub cmdUA_Click()
Call SelUnselAll(UNSELECTED_COLOR, MSF1)
End Sub

Private Sub cmdUR_Click()
Call SelUnsel(UNSELECTED_COLOR, MSF1, cboFrom, cboTo)
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P

txtFrom.Text = DateDisp(Date)
txtTo.Text = DateDisp(Date)

ExpSubMenu.mnuCap = ""

Call LoadSpecifics

If strCurrentUserType <> HOD Then cboDept.Text = "ALL"

CapGrid

Exit Sub
ERR_P:
    ShowError ("Load::" & Me.Caption)
End Sub

Private Sub CapGrid()
'' Sizing
MSF1.ColWidth(1) = MSF1.ColWidth(1) * 2.65
'' Aligning
MSF1.ColAlignment(0) = flexAlignLeftTop
End Sub

Private Sub cboDept_Click()
On Error GoTo ERR_P
If cboDept.ListIndex < 0 Then Exit Sub               '' If No Department
Call FillComboGrid
Call SelUnselAll(UNSELECTED_COLOR, MSF1)
Exit Sub
ERR_P:
    ShowError ("Department::" & Me.Caption)
End Sub

Private Sub FillComboGrid()
On Error GoTo ERR_P
Dim adrsTmp As New ADODB.Recordset, intEmpCnt As Integer, intTmpCnt As Integer
Dim strDeptTmp As String, strTempforCF As String
Call ComboFill(cboFrom, 16, 2, cboDept.List(cboDept.ListIndex, 0))
Call ComboFill(cboTo, 16, 2, cboDept.List(cboDept.ListIndex, 0))
If cboFrom.ListCount > 0 Then cboFrom.ListIndex = 0
If cboTo.ListCount > 0 Then cboTo.ListIndex = cboTo.ListCount - 1
If cboDept.Text = "ALL" Then
    strDeptTmp = "ALL"
Else
    strDeptTmp = cboDept.List(cboDept.ListIndex, 1)
    strDeptTmp = strDeptTmp
End If

    Select Case UCase(Trim(strDeptTmp))
        Case "", "ALL"
            strTempforCF = "select Empcode,name from empmst order by Empcode"               'Empcode,name
        Case Else
             If strCurrentUserType = HOD Then    ' 04-06
                strTempforCF = "select Empcode,name from empmst " & Replace(strCurrData, "'", "") & " and Empmst." & SELCRIT & "=" & _
                    strDeptTmp & " order by Empcode"    'Empcode,name
            Else
                'this If condition add by  for datatype
                If blnFlagForDept = True Then
                    strTempforCF = "select Empcode,name from empmst where Empmst." & SELCRIT & "=" & _
                    strDeptTmp & " order by Empcode"    'Empcode,name
                Else
                    strTempforCF = "select Empcode,name from empmst where Empmst." & SELCRIT & "=" & _
                    strDeptTmp & " order by Empcode"
                End If
            End If
    End Select

If GetFlagStatus("LocationRights") And strCurrentUserType <> ADMIN Then ' 28-01-09
    If UCase(Trim(strDeptTmp)) = "ALL" Then
        strTempforCF = "select Empcode,name from empmst " & strCurrData & " order by Empcode"
    Else
        strTempforCF = "select Empcode,name from empmst where Empmst.Dept = " & cboDept.Text & " And Empmst.Location in (" & UserLocations & ") order by Empcode"
    End If
End If
If adrsTmp.State = 1 Then adrsTmp.Close
adrsTmp.Open strTempforCF, ConMain, adOpenStatic, adLockReadOnly
If (adrsTmp.EOF And adrsTmp.BOF) Then
    MSF1.Rows = 1
    Exit Sub
End If
intEmpCnt = adrsTmp.RecordCount
intTmpCnt = intEmpCnt
MSF1.Rows = intEmpCnt + 1
For intEmpCnt = 0 To intTmpCnt - 1
    MSF1.TextMatrix(intEmpCnt + 1, 0) = adrsTmp(0)
    MSF1.TextMatrix(intEmpCnt + 1, 1) = adrsTmp(1)
    adrsTmp.MoveNext
Next
Exit Sub
ERR_P:
    ShowError ("Fill Employee Combos :: " & Me.Caption)
End Sub

Private Sub txtFrom_Click()
varCalDt = ""
varCalDt = Trim(txtFrom.Text)
txtFrom.Text = ""
Call ShowCalendar
End Sub

Private Sub txtTo_Click()
    varCalDt = ""
    varCalDt = Trim(txtTo.Text)
    txtTo.Text = ""
    Call ShowCalendar
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SaveExportData
' Author    : IVS
' Date      : 08/08/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function SaveExportData(ByVal strmain As String) As Boolean
On Error GoTo Err
    If Dir(CD1.FileName) <> "" Then Kill CD1.FileName
    Dim TmpFile As String
    TmpFile = IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\Exp.xls"
    Open TmpFile For Append As #2
    Print #2, strmain
    Close #2
    
    Dim xlApp As Excel.Application
    Set xlApp = New Excel.Application
    xlApp.Workbooks.Open TmpFile
    Select Case UCase(ExpType)
    Case "TIMECARD"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("F").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("G").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("H").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("I").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("J").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("K").NumberFormat = "0.00"
    Case "INOUT", "EXTENDED"
'        xlApp.ActiveWorkbook.ActiveSheet.Columns("D").NumberFormat = "0.00"
'        xlApp.ActiveWorkbook.ActiveSheet.Columns("E").NumberFormat = "0.00"
'        xlApp.ActiveWorkbook.ActiveSheet.Columns("F").NumberFormat = "0.00"
  '      xlApp.ActiveWorkbook.ActiveSheet.Columns("G").NumberFormat = "0.00"
'        xlApp.ActiveWorkbook.ActiveSheet.Columns("H").NumberFormat = "0.00"
 '       xlApp.ActiveWorkbook.ActiveSheet.Columns("I").NumberFormat = "0.00"
    Case "BIOINOUT"
    
    Case "ABSENTEE"
    
    End Select
    
    xlApp.Workbooks(1).SaveAs CD1.FileName, xlOpenXMLWorkbook, , , , False
     xlApp.Quit
    Set xlApp = Nothing

    Kill TmpFile
    Exit Function
Err:
    ShowError ("SaveExportData :: " & Me.Caption)
If Err.Number = 0 Then SaveExportData = True
End Function


