VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExp 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Data"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   435
      Left            =   5160
      TabIndex        =   13
      Top             =   6240
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6615
      Top             =   4710
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Next"
      Height          =   435
      Left            =   3480
      TabIndex        =   12
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Frame frOptions 
      Caption         =   "Options"
      Height          =   6075
      Index           =   0
      Left            =   240
      TabIndex        =   31
      Top             =   120
      Width           =   7125
      Begin VB.ListBox lstFields 
         Height          =   4350
         Index           =   0
         ItemData        =   "frmExp.frx":0000
         Left            =   120
         List            =   "frmExp.frx":0002
         TabIndex        =   0
         Top             =   540
         Width           =   2205
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "Add"
         Height          =   525
         Index           =   0
         Left            =   2370
         TabIndex        =   1
         Top             =   750
         Width           =   1395
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "Remove"
         Height          =   525
         Index           =   1
         Left            =   2370
         TabIndex        =   2
         Top             =   1305
         Width           =   1395
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "Add All"
         Height          =   525
         Index           =   2
         Left            =   2370
         TabIndex        =   3
         Top             =   1890
         Width           =   1395
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "Remove All"
         Height          =   525
         Index           =   3
         Left            =   2370
         TabIndex        =   4
         Top             =   2460
         Width           =   1395
      End
      Begin VB.ListBox lstFields 
         Height          =   4350
         Index           =   1
         ItemData        =   "frmExp.frx":0004
         Left            =   3840
         List            =   "frmExp.frx":0006
         TabIndex        =   6
         Top             =   525
         Width           =   2205
      End
      Begin VB.CommandButton cmdUPD 
         Caption         =   "Up"
         Height          =   525
         Index           =   0
         Left            =   6090
         TabIndex        =   7
         Top             =   780
         Width           =   945
      End
      Begin VB.CommandButton cmdUPD 
         Caption         =   "Down"
         Height          =   525
         Index           =   1
         Left            =   6090
         TabIndex        =   8
         Top             =   1440
         Width           =   945
      End
      Begin VB.Frame frOptions 
         Caption         =   "Options"
         Height          =   645
         Index           =   1
         Left            =   90
         TabIndex        =   34
         Top             =   5130
         Width           =   2235
         Begin VB.ComboBox cboOper 
            Height          =   315
            Index           =   0
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblOper 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delimiter"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   300
            Width           =   600
         End
      End
      Begin VB.CommandButton cmdExp 
         Caption         =   "Export"
         Height          =   435
         Left            =   5820
         TabIndex        =   11
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton cmdExportToExcell 
         Caption         =   "ExportToExcell"
         Height          =   435
         Left            =   4485
         TabIndex        =   10
         Top             =   5280
         Width           =   1215
      End
      Begin VB.ListBox lstExtraFields 
         Height          =   255
         Left            =   3810
         TabIndex        =   9
         Top             =   4920
         Width           =   2220
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   1740
         TabIndex        =   32
         Top             =   1875
         Visible         =   0   'False
         Width           =   2760
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Please Wait...."
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   -15
            TabIndex        =   33
            Top             =   105
            Width           =   2745
         End
      End
      Begin VB.Label lblOper 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fields Available"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label lblOper 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fields Available"
         Height          =   195
         Index           =   1
         Left            =   4050
         TabIndex        =   37
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label lblOper 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extra Fields Available"
         Height          =   195
         Index           =   3
         Left            =   2190
         TabIndex        =   36
         Top             =   4950
         Width           =   1500
      End
   End
   Begin VB.Frame frEmp 
      Caption         =   "Employee Selection"
      Height          =   5025
      Left            =   0
      TabIndex        =   19
      Top             =   1200
      Width           =   7125
      Begin VB.CommandButton cmdSR 
         Caption         =   "Select &Range"
         Height          =   435
         Left            =   4110
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdUR 
         Caption         =   "&Unselect Range"
         Height          =   465
         Left            =   4110
         TabIndex        =   22
         Top             =   1500
         Width           =   1335
      End
      Begin VB.CommandButton cmdSA 
         Caption         =   "&Select All"
         Height          =   435
         Left            =   4110
         TabIndex        =   21
         Top             =   2100
         Width           =   1335
      End
      Begin VB.CommandButton cmdUA 
         Caption         =   "U&nselect All"
         Height          =   435
         Left            =   4110
         TabIndex        =   20
         Top             =   2520
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   4065
         Left            =   510
         TabIndex        =   24
         ToolTipText     =   "Click to Toggle Selection"
         Top             =   900
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   7170
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSForms.ComboBox CboLocation 
         Height          =   315
         Left            =   5280
         TabIndex        =   47
         Top             =   570
         Width           =   1515
         VariousPropertyBits=   612390939
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2672;556"
         TextColumn      =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   225
         Left            =   5520
         TabIndex        =   46
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fro&m"
         Height          =   195
         Left            =   1020
         TabIndex        =   30
         Top             =   630
         Width           =   345
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T&o"
         Height          =   195
         Left            =   3270
         TabIndex        =   29
         Top             =   630
         Width           =   195
      End
      Begin MSForms.ComboBox cboFrom 
         Height          =   315
         Left            =   1830
         TabIndex        =   28
         Top             =   570
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
      Begin MSForms.ComboBox cboTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   27
         Top             =   570
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
      Begin MSForms.ComboBox cboDept 
         Height          =   315
         Left            =   1830
         TabIndex        =   26
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
      Begin VB.Label lblDeptCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         Height          =   195
         Left            =   540
         TabIndex        =   25
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.Frame frSel 
      Height          =   1035
      Left            =   0
      TabIndex        =   39
      Top             =   120
      Width           =   7125
      Begin VB.OptionButton optExp 
         Caption         =   "Daily Data"
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   43
         Top             =   270
         Width           =   1425
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   240
         Width           =   1515
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   600
         Width           =   1515
      End
      Begin VB.OptionButton optExp 
         Caption         =   "Monthly Data"
         Height          =   195
         Index           =   1
         Left            =   630
         TabIndex        =   40
         Top             =   690
         Width           =   1425
      End
      Begin VB.Label lblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         Height          =   195
         Left            =   3090
         TabIndex        =   45
         Top             =   300
         Width           =   450
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         Height          =   195
         Left            =   3090
         TabIndex        =   44
         Top             =   660
         Width           =   330
      End
   End
   Begin VB.Frame frSap 
      Height          =   1035
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   7125
      Begin VB.TextBox txtToPeri 
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
         Left            =   4710
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "D"
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtFrPeri 
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
         Left            =   2940
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "D"
         Top             =   240
         Width           =   1125
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
         Left            =   4230
         TabIndex        =   18
         Top             =   270
         Width           =   300
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
         TabIndex        =   17
         Top             =   270
         Width           =   2460
      End
   End
End
Attribute VB_Name = "frmExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset
Dim adrsForm As New ADODB.Recordset
Private sngArrTotl
Private statusArray1
'' Other Variables
'' Current Frame
Dim bytCurrFrame As Byte
'' For Selected Employee
Dim strSelEmp As String
'' For Delimeter
Dim strDel As String
'' For Constants
Private Const TRN_FIELDS = 22, FIXED_LVTRN_FIELDS = 11  'changes done by  TRN_FIELDS = 16, FIXED_LVTRN_FIELDS = 8
'Private Const TRN_FIELDS = 17, FIXED_LVTRN_FIELDS1 = 14
'' For Arrays
Dim strTrnData(1 To TRN_FIELDS) As String, strTrnFields(1 To TRN_FIELDS) As String
Dim strLvTrnData() As String, strLvTrnFields() As String

'' File Names
Dim strExpFileName As String, strPrvExpFileName As String, strtrn As String
Dim dtLst_Date As Date
Public ExpPath As String, Flt As String
Dim blnMnth As Boolean, blnLvTrn As Boolean, blnPer As Boolean
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet




Private Sub CboLocation_Change()
On Error GoTo ERR_P
'cbodept.Text = "ALL"
If cboDept.ListIndex < 0 Then Exit Sub               '' If No Department
Call FillLocaGrid
Call SelUnselAll(vbWhite, MSF1)
Exit Sub
ERR_P:
    ShowError ("Department::" & Me.Caption)
End Sub

Private Sub FillLocaGrid()     '' Fills Employee Combo and Grid
On Error GoTo ERR_P
Dim intEmpCnt As Integer, intTmpCnt As Integer
Dim strArrEmp() As String
Dim strDeptTmp As String, strTempforCF As String

intEmpCnt = 0
If CboLocation.Text = "ALL" Then
    strDeptTmp = "ALL"
Else
    strDeptTmp = CboLocation.List(CboLocation.ListIndex, 1)
End If


Select Case UCase(Trim(strDeptTmp))
    Case "", "ALL"
        strTempforCF = "Select Empcode,Name from empmst where (joindate is not null and joindate<=" & strDTEnc & _
        DateCompStr(Date) & strDTEnc & ") Order by EmpCode"
       'strTempforCF = "select Empcode,name from empmst where location=" & (cboLoc.Text) & " and company=" & (cboComp.Text) & " order by Empcode"
    Case Else
        If strCurrentUserType = HOD Then

            strTempforCF = "Select Empcode,Name from empmst " & strCurrData & " and (joindate is not null and joindate<=" & strDTEnc & DateCompStr(Date) & _
              strDTEnc & ") and Empmst." & SELCRIT1 & " = " & strDeptTmp & " Order by EmpCode"
         
        Else
            'this If condition add by  for datatype
'            If blnFlagForDept = True Then
                strTempforCF = "select Empcode,name from empmst where Empmst." & SELCRIT1 & "=" & _
                strDeptTmp & " order by Empcode"    'Empcode,name
'            End If
        End If
End Select

If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open strTempforCF, ConMain, adOpenStatic, adLockReadOnly
If (adrsEmp.EOF And adrsEmp.BOF) Then
    cboFrom.clear
    cboTo.clear
    MSF1.Rows = 1
    Exit Sub
End If
intEmpCnt = adrsEmp.RecordCount
intTmpCnt = intEmpCnt
MSF1.Rows = intEmpCnt + 1
ReDim strArrEmp(intTmpCnt - 1, 1)
For intEmpCnt = 0 To intTmpCnt - 1
    strArrEmp(intEmpCnt, 0) = adrsEmp(0)
    strArrEmp(intEmpCnt, 1) = adrsEmp(1)
    MSF1.TextMatrix(intEmpCnt + 1, 0) = adrsEmp(0)
    MSF1.TextMatrix(intEmpCnt + 1, 1) = adrsEmp(1)
    adrsEmp.MoveNext
Next
cboFrom.List = strArrEmp
cboTo.List = strArrEmp
cboFrom.ListIndex = 0
cboTo.ListIndex = cboTo.ListCount - 1
Erase strArrEmp
Exit Sub
ERR_P:
    ShowError ("Fill Combo Grid :: " & Me.Caption)
'    Resume Next
End Sub

Private Sub cmdExp_Click()
On Error GoTo ERR_P
If lstFields(1).ListCount = 0 Then
    MsgBox NewCaptionTxt("67021", adrsC), vbInformation
    cmdOper(2).SetFocus
    Exit Sub
End If

'' Get the File Name
With CD1
    .FileName = ""
    If cboOper(0).Text = "COMMA" Then    ' 24-03
        .Filter = "CSV (Comma Separated Value)|*.txt"
    ElseIf (cboOper(0).Text = "TAB") Then
        .Filter = "Text Files|*.Txt"
    Else
        Dim xlApp As Excel.Application
        Dim xlBook As Excel.Workbook
        Dim fFmt As Long
        Set xlApp = New Excel.Application
        Set xlBook = xlApp.Workbooks.Add
        fFmt = xlBook.FileFormat
        Select Case fFmt
            Case xlSYLK:
                .Filter = "*.slk|*.slk"
            Case xlWKS:
                .Filter = "*.wks|*.wks"
            Case xlWK1, xlWK1ALL, xlWK1FMT:
                .Filter = "*.wk1|*.wk1"
            Case xlCSV, xlCSVMac, xlCSVMSDOS, xlCSVWindows:
                .Filter = "*.csv|*.csv"
            Case xlDBF2, xlDBF3, xlDBF4:
                .Filter = "*.dbf|*.dbf"
            Case xlWorkbookNormal, xlExcel2FarEast, xlExcel3, xlExcel4, xlExcel4Workbook, xlExcel5, xlExcel7, xlExcel9795:
                .Filter = "Excel|*.xls"
            Case xlHtml:
                .Filter = "*.htm|*.htm"
            Case xlTextMac, xlTextWindows, xlTextMSDOS, xlUnicodeText, xlCurrentPlatformText:
                .Filter = "*.txt|*.txt"
            Case xlTextPrinter:
                .Filter = "*.prn|*.prn"
            Case 50:
                .Filter = "*.xlsb|*.xlsb"
            Case 51:
                .Filter = "*.xls|*.xls"
            Case 52:
                .Filter = "*.xlsm|*.xlsm"
            Case 56:
                .Filter = "*.xls|*.xls"
            Case Else:
                MsgBox "Do not find exact Excel File Format:: " & fFmt
                Exit Sub
        End Select
    End If
        CD1.ShowSave
        If Trim(.FileName) = "" Then Exit Sub
End With
Call ExportData
Exit Sub
ERR_P:
    ShowError ("Export::" & Me.Caption)
    'Resume Next
End Sub




Private Sub cmdExportToExcell_Click()
10    On Error GoTo Err
          Dim str_Input_Query As String
          Dim str_Column As String
          Dim str_Excell_File_Name As String
          Dim strTmp  As String
          '' Temporary Current Arrays
          Dim strTmpDataArr() As String, strTmpFieldsArr() As String
          '' Temporary Variables
          Dim bytTmp As Byte
          Dim strLvBal As String
20        Frame1.Visible = True
30        strTmpDataArr = GetCurrentArray(1)
40        strTmpFieldsArr = GetCurrentArray(2)
50        strLvBal = "Lvbal" & Right(strExpFileName, 2)
60        adrsForm.MoveFirst
70        For bytTmp = 0 To lstFields(1).ListCount - 1
80            strTmp = lstFields(1).List(bytTmp)
90            strTmp = GetFieldName(strTmpDataArr, strTmpFieldsArr, strTmp)
100           str_Column = str_Column & strExpFileName & "." & strTmp & ","
110       Next
          'for query
120       For bytTmp = 0 To lstExtraFields.ListCount - 1
130           str_Column = str_Column & "'" & GetINIString(lstExtraFields.List(bytTmp), _
              "DefaulValue", App.path & "\" & App.EXEName & ".ini") & "'" & " AS" & _
              " """ & lstExtraFields.List(bytTmp) & """" & ","
140       Next
          'for remove comma
          'str_Column = Left(str_Column, Len(str_Column) - 1)
150       str_Column = str_Column & strLvBal & ".PL AS PL_Balance," _
            & strLvBal & ".CL AS CL_Balance," _
            & strLvBal & ".CO AS CO_Balance," _
            & strLvBal & ".SL AS SL_Balance"
          
          
160       str_Input_Query = "SELECT " & str_Column & _
          " FROM " & strLvBal & "," & strExpFileName & " WHERE " & strLvBal & _
          ".empcode=" & strExpFileName & ".empcode AND " & strExpFileName & "." & _
          Mid(UCase(adrsForm.Source), InStr(1, UCase(adrsForm.Source), "WHERE") + Len("WHERE") + 1, _
          Len(UCase(adrsForm.Source))) & ""
          
          'for Excell File Name
170       str_Excell_File_Name = Left(cboMonth.Text, 3) & Right(cboYear.Text, 2) & "ExportToExcell"
          'for export
180       MsgBox "Exported following file " & ExportIntoFile(str_Excell_File_Name, str_Input_Query), vbInformation
190       Frame1.Visible = False
200   Exit Sub
Err:
          Frame1.Visible = False
210       Call ShowError("cmdExportToExcell_Click Line : " & Erl)
End Sub


Private Sub cmdOper_Click(Index As Integer)
Dim bytTmp As Byte
Select Case Index
    Case 0  '' Add
        If lstFields(0).ListIndex < 0 Then Exit Sub
        'this ValidColumn add by  for duplicate column MIS2007DF023
        If ValidColumn(lstFields(0)) Then
            lstFields(1).AddItem lstFields(0).List(lstFields(0).ListIndex)
        End If
        
        If lstFields(0).ListIndex < lstFields(0).ListCount - 1 Then
            lstFields(0).ListIndex = lstFields(0).ListIndex + 1
        End If
        lstFields(1).ListIndex = lstFields(1).ListCount - 1
        
'
'        If ValidColumn(lstExtraFields) Then
'            lstFields(1).AddItem lstExtraFields.List(lstExtraFields.ListIndex)
'        End If
        
    Case 1  '' Remove
        If lstFields(1).ListCount = 0 Then Exit Sub
        If lstFields(1).ListIndex < 0 Then Exit Sub
        lstFields(1).RemoveItem (lstFields(1).ListIndex)
        If lstFields(1).ListIndex < lstFields(1).ListCount - 1 Then
            lstFields(1).ListIndex = lstFields(1).ListIndex + 1
        End If
    Case 2  '' Add All
        lstFields(1).clear
        For bytTmp = 0 To lstFields(0).ListCount - 1
            lstFields(1).AddItem lstFields(0).List(bytTmp)
        Next
        lstFields(1).ListIndex = 0
'
'        For bytTmp = 0 To lstExtraFields.ListCount - 1
'            lstFields(1).AddItem lstExtraFields.List(bytTmp)
'        Next
        
    Case 3  '' Remove All
        lstFields(1).clear
End Select
End Sub
'this function add by  for remove duplicate column
Private Function ValidColumn(lstI As ListBox) As Boolean
Dim i As Integer
If lstI.ListIndex < 0 Then Exit Function
For i = 0 To lstFields(1).ListCount
    If lstI.Text = lstFields(1).List(i) Then
        ValidColumn = False
        Exit Function
    Else
        ValidColumn = True
    End If
Next
End Function
Private Sub cmdBack_Click()

Select Case bytCurrFrame
    Case 1
        '' Make File Name
        Select Case bytMode
            Case 1
  
                strExpFileName = Left(cboMonth.Text, 3) & Right(cboYear.Text, 2) & "Trn"
          
                lstExtraFields.Visible = False
                lblOper(3).Visible = False
                cmdExportToExcell.Visible = False
            Case 2
                If CByte(pVStar.Yearstart) > MonthNumber(cboMonth.Text) Then
                    strExpFileName = "lvtrn" & Right(CStr(CInt(cboYear.Text) - 1), 2)
                Else
                    strExpFileName = "LvTrn" & Right(cboYear.Text, 2)
                End If
                dtLst_Date = CDate("01-" & cboMonth.Text & "-" & cboYear.Text)
                dtLst_Date = FdtLdt(Month(dtLst_Date), cboYear.Text, "L")
      
        End Select
            If Not FindDataFile Then Exit Sub
            If Not CheckEmployees Then Exit Sub
            If Not DataFound Then Exit Sub
        Call FillListBoxes
        bytCurrFrame = bytCurrFrame + 1
    Case 2
        bytCurrFrame = bytCurrFrame - 1
End Select
Call ToggleState
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub cmdUPD_Click(Index As Integer)
On Error GoTo ERR_P
Dim bytTmp As Byte, strTmp As String
With lstFields(1)
    Select Case Index
        Case 0      '' Up
            If .ListCount = 0 Then Exit Sub
            If .ListIndex < 1 Then Exit Sub
            bytTmp = .ListIndex
            strTmp = .List(bytTmp - 1)
            .RemoveItem bytTmp - 1
            .AddItem strTmp, bytTmp
        Case 1      '' Down
            If .ListCount = 0 Then Exit Sub
            If .ListIndex = .ListCount - 1 Then Exit Sub
            bytTmp = .ListIndex
            strTmp = .List(bytTmp + 1)
            .RemoveItem bytTmp + 1
            .AddItem strTmp, bytTmp
    End Select
End With
Exit Sub
ERR_P:
    ShowError ("UPD:: " & Me.Caption)
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P

txtFrPeri.Text = DateDisp(Date)
txtToPeri.Text = DateDisp(Date)

ExpSubMenu.mnuCap = ""
Call SetFormIcon(Me)
Call RetCaptions
Call GetRights
Call LoadSpecifics
Call SetTagOption
optExp(0).Value = True
If strCurrentUserType <> HOD Then cboDept.Text = "ALL"
lblLocation.Visible = GetFlagStatus("pratham")
CboLocation.Visible = GetFlagStatus("pratham")
Exit Sub
ERR_P:
    ShowError ("Load::" & Me.Caption)
End Sub

Private Function SetTagOption()

        cmdExportToExcell.Visible = False
        lblOper(3).Visible = False
        lstExtraFields.Visible = False
  
End Function
Private Sub GetRights()
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(4, 19, 7, 1)
If strTmp = "1" Then
    cmdBack.Enabled = True
Else
    MsgBox NewCaptionTxt("00001", adrsMod), vbExclamation
    cmdBack.Enabled = False
End If
Exit Sub
ERR_P:
    ShowError ("GetRights::" & Me.Caption)
    cmdBack.Enabled = False
End Sub

Private Sub ExportData()
On Error GoTo ERR_P
'' Get the Delimiter
Dim strmain As String
strDel = IIf(cboOper(0).ListIndex = 0, ",", vbTab)
strmain = GetExportData
If SaveExportData(strmain) Then
    MsgBox NewCaptionTxt("67022", adrsC), vbInformation
Else
    MsgBox NewCaptionTxt("67023", adrsC), vbInformation
End If
Exit Sub
ERR_P:
    ShowError ("ExportData::" & Me.Caption)
End Sub

Private Function SaveExportData(ByVal strmain As String) As Boolean
'' On Error Resume Next
If Dir(CD1.FileName) <> "" Then Kill CD1.FileName

    Open CD1.FileName For Append As #2
    Print #2, strmain
    Close #2

If Err.Number = 0 Then SaveExportData = True
End Function


Private Function GetExportData() As String

10    On Error GoTo ERR_P
      '' Data String
      Dim strLinedata As String, strData As String
      '' Temporary Current Arrays
      Dim strTmpDataArr() As String, strTmpFieldsArr() As String
      '' Temporary Variables
      Dim bytTmp As Byte, strTmp As String
20    strTmpDataArr = GetCurrentArray(1)
30    strTmpFieldsArr = GetCurrentArray(2)
40    adrsForm.MoveFirst
50    For bytTmp = 0 To lstFields(1).ListCount - 1
60        strTmp = lstFields(1).List(bytTmp)
70        strTmp = GetFieldName(strTmpDataArr, strTmpFieldsArr, strTmp)
80        If strTmp <> "" Then
90                strData = strData & strTmp & strDel
100       Else
 
                strData = strData & strDel
         
120       End If
130   Next
140   strData = strData & vbCrLf

150   Do While Not adrsForm.EOF
160       strLinedata = ""
170       For bytTmp = 0 To lstFields(1).ListCount - 1
180           strTmp = lstFields(1).List(bytTmp)
190           strTmp = GetFieldName(strTmpDataArr, strTmpFieldsArr, strTmp)
200           If strTmp <> "" Then
210               If Not IsNull(adrsForm(strTmp)) Then
                      If UCase(strTmp) = "DATE" Then
                        strLinedata = strLinedata & Format(adrsForm(strTmp), "dd/mmm/yy") & strDel
                      Else
220                     strLinedata = strLinedata & adrsForm(strTmp) & strDel
                      End If
230               Else
240                   strLinedata = strLinedata & strDel
250               End If
260           End If
270       Next
280       strLinedata = Left(strLinedata, Len(strLinedata) - 1)
290       strData = strData & strLinedata & vbCrLf
300       adrsForm.MoveNext
310   Loop
320   GetExportData = strData
330   Exit Function
ERR_P:
      ShowError ("GetExportData::" & Me.Caption & "Erl:" & Erl)
Resume Next
End Function

Private Function GetFieldName(ByRef strTmpDataArr() As String, _
ByRef strTmpFieldsArr() As String, ByVal strTmp As String) As String
'' On Error Resume Next
Dim bytTmp As Byte
For bytTmp = 1 To UBound(strTmpDataArr)
    If Trim(UCase(strTmp)) = Trim(UCase(strTmpDataArr(bytTmp))) Then
        strTmp = strTmpFieldsArr(bytTmp)
'        If strTmp = "Empcode" Then
'        strTmp = "Employee Number"
'        End If
        Exit For
    End If
Next
GetFieldName = strTmp
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

Private Function DataFound() As Boolean
On Error GoTo ERR_P
Dim bytTmp As Byte, bytCnt As Byte, strTmp As String, strlstdt As String
Dim dttodate As Date, dtfromdate As Date, cnt As Integer
If adrsForm.State = 1 Then adrsForm.Close
If InStr(1, UCase(strExpFileName), "LV") = 0 Then
    
    adrsForm.Open "Select name,div,Dept,location," & strExpFileName & ".* from empmst," & strExpFileName & " Where " & strExpFileName & ".empcode=empmst.empcode and " & strExpFileName & ".Empcode in " & strSelEmp & _
    " Order by " & strExpFileName & ".empcode, " & strKDate 'Changes done by  21-08-09

Else
    If optExp(0).Value = True Then
        adrsForm.Open "Select * from " & strExpFileName & " Where Empcode in " & strSelEmp & _
        "Order by Empcode"
    Else
        Select Case bytBackEnd
            Case 1, 2
                dtfromdate = DateCompDate(FdtLdt(MonthNumber(cboMonth.Text), cboYear.Text, "f"))
                dttodate = DateCompDate(FdtLdt(MonthNumber(cboMonth.Text), cboYear.Text, "l"))
                strlstdt = Day(dttodate)
                For bytCnt = 1 To strlstdt
                  If UCase(WeekdayName(WeekDay(dtfromdate + bytCnt), True)) = "SUN" Then
                  cnt = cnt + 1
                  End If
                Next
                adrsForm.Open "Select name," & strExpFileName & ".*," & cnt & " as sunday from empmst," & strExpFileName & " Where " & strExpFileName & ".empcode=empmst.empcode and " & strExpFileName & ".Empcode in " & strSelEmp & _
                "  and  month(lst_date)=" & MonthNumber(cboMonth.Text) & " Order by " & strExpFileName & ".Empcode" 'Changes done by  21-08-09
            Case 3  ' 25-05
                adrsForm.Open "Select name," & strExpFileName & ".* from empmst," & strExpFileName & " Where " & strExpFileName & ".empcode=empmst.empcode and " & strExpFileName & ".Empcode in " & strSelEmp & _
                "  and  TO_CHAR(lst_date,'MM')=" & MonthNumber(cboMonth.Text) & " Order by " & strExpFileName & ".Empcode"  'Changes done by  21-08-09
        End Select
    End If
End If
If adrsForm.EOF Then
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
    Exit Function
End If
If bytMode = 2 Then
    '' Set Variables
    bytTmp = 0: bytCnt = 0


        ReDim strLvTrnData(1 To FIXED_LVTRN_FIELDS)
        ReDim strLvTrnFields(1 To FIXED_LVTRN_FIELDS)
        
        For bytTmp = 1 To FIXED_LVTRN_FIELDS
            '' Display Names
            strLvTrnData(bytTmp) = Choose(bytTmp, "Employee", "Employee Name", "Paid Days", "Absent", "Present", _
            "WeekOff", "Holidays", "Night Shifts", "Overtime Paid", "WorkHrs", "Sunday") 'Changes done by  21-08-09
            '' Actual Field Names
            strLvTrnFields(bytTmp) = Choose(bytTmp, "Empcode", "Name", "PaidDays", (pVStar.AbsCode), _
            (pVStar.PrsCode), Trim(pVStar.WosCode), Trim(pVStar.HlsCode), "Night", "Otpd_Hrs", "wrk_hrs", "sunday") 'Changes done by  21-08-09
         Next
End If
    For bytTmp = 0 To adrsForm.Fields.Count - 1
        strTmp = adrsForm.Fields(CInt(bytTmp)).name
        If Len(strTmp) = 2 Then
            Select Case UCase(strTmp)
                Case UCase(Trim(pVStar.AbsCode)), UCase(Trim(pVStar.PrsCode)), UCase((pVStar.AbsCode)), UCase((pVStar.PrsCode)) 'last two case add by
                Case UCase(Trim(pVStar.WosCode)), UCase(Trim(pVStar.HlsCode))
                Case Else
                    ReDim Preserve strLvTrnData(1 To (FIXED_LVTRN_FIELDS + 1 + bytCnt))
                    ReDim Preserve strLvTrnFields(1 To (FIXED_LVTRN_FIELDS + 1 + bytCnt))
                    strLvTrnData(UBound(strLvTrnData)) = strTmp
                    strLvTrnFields(UBound(strLvTrnFields)) = strTmp
                    bytCnt = bytCnt + 1
            End Select
        End If
     Next

DataFound = True
Exit Function
ERR_P:
    ShowError ("DataFound::" & Me.Caption)
    Resume Next
End Function

Private Function FindDataFile() As Boolean
On Error GoTo ERR_P
If Not FindTable(strExpFileName) Then
    If bytMode = 1 Then
        MsgBox NewCaptionTxt("67024", adrsC) & cboMonth.Text
        cboMonth.SetFocus
    Else
        MsgBox NewCaptionTxt("67025", adrsC) & cboYear.Text
        cboYear.SetFocus
    End If
    Exit Function
End If
FindDataFile = True
Exit Function
ERR_P:
    ShowError ("FindDataFile::" & Me.Caption)
End Function

Private Sub FillListBoxes()
On Error GoTo ERR_P
Dim bytTmp As Byte
Select Case bytMode
    Case 1      '' Daily Data
        With lstFields(0)
            .clear
            For bytTmp = 1 To TRN_FIELDS
                .AddItem strTrnData(bytTmp)
            Next
        End With
    Case 2      '' Monthly data
        With lstFields(0)
            .clear
            For bytTmp = 1 To UBound(strLvTrnData)
                .AddItem strLvTrnData(bytTmp)
            Next
        End With
End Select
If lstFields(0).ListCount > 0 Then lstFields(0).ListIndex = 0
Call cmdOper_Click(2)
Exit Sub
ERR_P:
    ShowError ("FillListBoxes:" & Me.Caption)
End Sub

Private Function CheckEmployees() As Boolean
On Error GoTo ERR_P
Dim intTmp As Integer
Call TruncateTable("ECode") ' 04-06
strSelEmp = ""
With MSF1
    For intTmp = 1 To .Rows - 1
        .row = intTmp
        If .CellBackColor = SELECTED_COLOR Then
            strSelEmp = strSelEmp & "'" & .TextMatrix(intTmp, 0) & "',"
            ConMain.Execute "insert into ECode values('" & .TextMatrix(intTmp, 0) & "')"   ' 04-06
        End If
    Next
End With
If strSelEmp = "" Then
    MsgBox NewCaptionTxt("00049", adrsMod), vbInformation
    cmdSA.SetFocus
    Exit Function
End If
strSelEmp = Left(strSelEmp, Len(strSelEmp) - 1)
strSelEmp = "(" & strSelEmp & ")"
CheckEmployees = True
Exit Function
ERR_P:
    ShowError ("CheckEmployees::" & Me.Caption)
End Function

Private Sub FillFieldsArrays()
On Error Resume Next
Dim bytTmp As Byte
For bytTmp = 1 To TRN_FIELDS
    '' Display Names
    strTrnData(bytTmp) = Choose(bytTmp, "Employee", "Employee Name", "Date", "Entry", "Shift", "Status", _
    "Arrival", "Break Out Punch", "Break In Punch", "Departure", "Work Hours", "Present", "Late Arrival", _
    "Early Departure", "Comp Off", "OT", "OT Authorization", "OT Remarks", "Division", "Location", "Department") 'Changes done by  21-08-09
    '' Actual Field Names
    strTrnFields(bytTmp) = Choose(bytTmp, "Empcode", "Name", "Date", "Entry", "Shift", "Presabs", _
    "Arrtim", "Actrt_O", "Actrt_I", "Deptim", "WrkHrs", "present", "LateHrs", "EarlHrs", "Cof", _
    "Ovtim", "OTConf", "OTConf", "div", "Location", "dept")
Next
End Sub

Private Sub RetCaptions()
'' On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '67%'", ConMain, adOpenStatic
Call SetGridDetails(Me, frEmp, MSF1, lblFrom, lblTo)
End Sub

Private Sub FillCombos()
On Error GoTo ERR_P
Dim intTmp As Integer
'Call SetCritCombos(cboDept)

    Call SetCritCombos(cboDept)
    Call ComboFill(CboLocation, 11, 2)
'cbodept.ListIndex = cbodept.ListCount - 1
'' Fill Month Year Combos
For intTmp = 1 To 12
    cboMonth.AddItem MonthName(intTmp)
Next
For intTmp = 2002 To 2029
    cboYear.AddItem CStr(intTmp)
Next
cboMonth.Text = MonthName(Month(Date))
cboYear.Text = CStr(Year(Date))
'' Fill Delimeter and Options Combo
With cboOper
    '' Delimeter
    If Not GetFlagStatus("EXCELPASSWORD") Then
    .Item(0).AddItem "COMMA"
    .Item(0).AddItem "TAB"
    End If
    .Item(0).AddItem "EXCEL"    ' 25-05
    .Item(0).ListIndex = 0
End With
Exit Sub
ERR_P:
    ShowError ("FillCombos::" & Me.Caption)
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
    '        ''For Mauritius 09-08-2003
    '        ''Original ->strTempforCF = "select Empcode,name from empmst Where " & SELCRIT & "=" & _
    '        strDeptTmp & " order by Empcode"                               'Empcode,name
    '        strTempforCF = "select Empcode,name from empmst " & strCurrData & " order by Empcode"    'Empcode,name
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

Private Sub LoadSpecifics()
On Error GoTo ERR_P
frOptions(0).Visible = False
Call CapGrid
Call FillCombos
bytCurrFrame = 1
bytMode = 1
Call ToggleState
'optExp(0).Value = True
Call FillFieldsArrays
'' Set Recordset
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

Private Sub ToggleState()
'' On Error Resume Next
Select Case bytCurrFrame
    Case 1      '' Emp Selection Frame
        cmdBack.Caption = "&Next"
        frOptions(0).Visible = False
    Case 2      '' Export Frame
        cmdBack.Caption = "&Back"
        frOptions(0).Visible = True
        frOptions(0).Top = frSel.Top
        frOptions(0).Left = frSel.Left
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call FreeRes
End Sub

Private Sub FreeRes()
'' On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
If adrsForm.State = 1 Then adrsForm.Close
Set adrsC = Nothing
Set adrsForm = Nothing
'' Erase Arrays
Erase strTrnData: Erase strTrnFields
Erase strLvTrnData: Erase strLvTrnFields
End Sub

Private Sub MSF1_Click()
If MSF1.Rows = 1 Then Exit Sub
If MSF1.CellBackColor = SELECTED_COLOR Then
    With MSF1
        .Col = 0
        .CellBackColor = UNSELECTED_COLOR
        .Col = 1
        .CellBackColor = UNSELECTED_COLOR
    End With
Else
    With MSF1
        .Col = 0
        .CellBackColor = SELECTED_COLOR
        .Col = 1
        .CellBackColor = SELECTED_COLOR
    End With
End If
End Sub

Private Sub MSF1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then Call MSF1_Click
End Sub

Private Sub CapGrid()
'' Sizing
MSF1.ColWidth(1) = MSF1.ColWidth(1) * 2.65
'' Aligning
MSF1.ColAlignment(0) = flexAlignLeftTop
End Sub

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

Private Sub optExp_Click(Index As Integer)
Select Case Index
    Case 0  '' Daily Data
        bytMode = 1
        lblMonth.Visible = True
        cboMonth.Visible = True
    Case 1  '' Monthly Data
        bytMode = 2
        lblMonth.Visible = True
        cboMonth.Visible = True
End Select
End Sub

Private Sub txtFrPeri_Click()
varCalDt = ""
varCalDt = Trim(txtFrPeri.Text)
txtFrPeri.Text = ""
Call ShowCalendar
End Sub

Private Sub txtFrPeri_GotFocus()
    Call GF(txtFrPeri)
End Sub

Private Sub txtFrPeri_KeyPress(KeyAscii As Integer)
    Call CDK(txtFrPeri, KeyAscii)
End Sub

Private Sub txtFrPeri_Validate(Cancel As Boolean)
If Not ValidDate(txtFrPeri) Then
    txtFrPeri.SetFocus
    Cancel = True
End If
End Sub

Private Sub txtToPeri_Click()
varCalDt = ""
varCalDt = Trim(txtToPeri.Text)
txtToPeri.Text = ""
Call ShowCalendar
End Sub

Private Sub txtToPeri_GotFocus()
    Call GF(txtToPeri)
End Sub

Private Sub txtToPeri_KeyPress(KeyAscii As Integer)
    Call CDK(txtToPeri, KeyAscii)
End Sub

Private Sub txtToPeri_Validate(Cancel As Boolean)
If Not ValidDate(txtToPeri) Then
    txtToPeri.SetFocus
    Cancel = True
End If
End Sub

Public Function SelPath() As Boolean 'Added by  30-03
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    SelPath = True
    With CD1
        .FileName = ""
        Flt = ExpExtension(xlBook.FileFormat)
        If Flt = "FALSE" Then
            SelPath = False
            Exit Function
        Else
            .Filter = Flt
        End If
        .Flags = cdlOFNFileMustExist
        .ShowSave
        ExpPath = .FileName
    End With
    If ExpPath = "" Then
        SelPath = False
        Exit Function
    End If
    If Dir(ExpPath, vbDirectory) <> "" Then
        If MsgBox(ExpPath & " already exist.Overwrite?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Set xlApp = Nothing: Set xlBook = Nothing: Set xlSheet = Nothing: SelPath = False: Exit Function
        End If
    End If
End Function


Public Sub SaveExlExport(ByVal strqry As String)    ' 07-05
    Dim rsExp As New ADODB.Recordset
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)

    On Error GoTo ERR_P
    If rsExp.State = 1 Then rsExp.Close
    rsExp.Open strqry, ConMain, adOpenStatic, adLockReadOnly
    If rsExp.BOF And rsExp.EOF Then
        MsgBox "Data Not Found"
        Exit Sub
    End If
    With CD1
        .FileName = ""
        Flt = ExpExtension(xlBook.FileFormat)
        If Flt = "FALSE" Then
            Exit Sub
        Else
            .Filter = Flt
        End If
        .Flags = cdlOFNFileMustExist
        .ShowSave
        If Trim(.FileName) = "" Then Exit Sub
    End With
    ExpPath = CD1.FileName
    If Dir(ExpPath, vbDirectory) <> "" Then
        If MsgBox(ExpPath & " already exist.Overwrite?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Set xlApp = Nothing: Set xlBook = Nothing: Set xlSheet = Nothing
            Exit Sub
        End If
    End If
    xlBook.Close
    xlApp.Quit
    Set xlApp = Nothing: Set xlBook = Nothing: Set xlSheet = Nothing
    Exit Sub
ERR_P:
    ShowError ("Save Excel Export:: " & Me.Caption)
    'Resume Next
End Sub

