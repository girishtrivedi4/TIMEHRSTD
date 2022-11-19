VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmShfCh 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Shedule For All"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optDate 
      Caption         =   "Date-Wise"
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
      Left            =   3720
      TabIndex        =   3
      Top             =   345
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.OptionButton optMnth 
      Caption         =   "Month-Wise"
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
      Left            =   3720
      TabIndex        =   2
      Top             =   75
      Visible         =   0   'False
      Width           =   1365
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   10755
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   435
      Left            =   8280
      TabIndex        =   15
      Top             =   7470
      Width           =   1395
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   435
      Left            =   8280
      TabIndex        =   14
      Top             =   6960
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   6960
      Width           =   5175
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
         Height          =   435
         Left            =   3240
         TabIndex        =   11
         Top             =   200
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Rotation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   300
         Width           =   1230
      End
      Begin MSForms.ComboBox cboRot 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   975
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1720;661"
         ColumnCount     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CommandButton cmdSA 
      Caption         =   "Select All"
      Height          =   435
      Left            =   6720
      TabIndex        =   12
      Top             =   6960
      Width           =   1395
   End
   Begin VB.CommandButton cmdUA 
      Caption         =   "Unselect All"
      Height          =   435
      Left            =   6720
      TabIndex        =   13
      Top             =   7440
      Width           =   1395
   End
   Begin VB.ComboBox cboMonth 
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   60
      Width           =   1275
   End
   Begin VB.ComboBox cboYear 
      Height          =   315
      Left            =   7740
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   30
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   6255
      Left            =   0
      TabIndex        =   19
      Top             =   585
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11033
      _Version        =   393216
      Cols            =   33
      FixedCols       =   2
      HighLight       =   2
      Appearance      =   0
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   435
      Left            =   9840
      TabIndex        =   16
      Top             =   6960
      Width           =   1395
   End
   Begin VB.Label lblMonth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5340
      TabIndex        =   4
      Top             =   120
      Width           =   525
   End
   Begin VB.Label lblYear 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7320
      TabIndex        =   6
      Top             =   60
      Width           =   375
   End
   Begin MSForms.ComboBox cboDept 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   30
      Width           =   2595
      VariousPropertyBits=   612390939
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4577;556"
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   17
      Top             =   105
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Indicates SUNDAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   18
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmShfCh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFName As String
Dim strSelEmp As String
Dim bytTmp As Byte
Dim ColArr(32) As Integer   ' 21-05
Private Sub cboDept_Change() ' 11-07
Dim i As Integer
For i = 1 To 32
    ColArr(i) = 0
Next
End Sub

Private Sub cmdExport_Click()
     Dim i As Integer
    If optMnth.Value = True Then
        STRECODE = ""
        For i = 1 To MSF1.Rows - 1
        MSF1.Row = i
        If MSF1.CellBackColor = SELECTED_COLOR Then
            STRECODE = STRECODE & "'" & MSF1.TextMatrix(i, 1) & "',"
        End If
        
    Next
    
    STRECODE = Mid(STRECODE, 1, Len(STRECODE) - 1)
    
      Export ("Select * from " & strFName & " Where Empcode in (" & STRECODE & ")")
    Else       ' 21-05
        Dim strDate As String, strDept As String
       
        For i = 1 To 31
            If ColArr(i) = 1 Then
                strDate = strDate & "d" & i & ","
            End If
        Next
        If strDate = "" Then
            MsgBox "Please select the date."
            Exit Sub
        Else
            strDate = Left(strDate, Len(strDate) - 1)
            strDept = cboDept.List(cboDept.ListIndex, 0)
            strDept = EncloseQuotes(strDept)
            Select Case UCase(Trim(strDept))
                Case "", "ALL"
                    strDate = "Select empcode," & strDate & " from " & strFName
                Case Else
                    strDate = "Select " & strFName & ".empcode," & strDate & " from " & strFName & ",Empmst where Empmst.Empcode = " & _
                        strFName & ".Empcode And Empmst.Dept = " & cboDept.Text
            End Select
            Select Case bytBackEnd
                Case 1, 2 ''SQLServer,MSAccess
                    strDate = strDate & " Order by " & strFName & ".Empcode"
                Case 3  ''Oracle
                    strDate = strDate & " Order by Empcode"
            End Select
            Export (strDate)
        End If
    End If
End Sub
'added by
Private Sub cmdImport_Click()
CD1.Flags = cdlOFNHideReadOnly
CD1.Filter = "Excel Files|*.xls;*.xlsx"
CD1.FilterIndex = 1
CD1.FileName = ""
CD1.ShowOpen
Call ImportExcel(CD1.FileName)
End Sub
Private Function ImportExcel(FilePath As String)
On Error GoTo ERR_P
Dim strSql As String
Dim lng As Long
Dim bytTmp As Integer, i As Integer, Flag As Integer
Dim rsImport As New ADODB.Recordset
Dim rsSftCode As New ADODB.Recordset
Dim rsLvCode As New ADODB.Recordset
Dim xlSheet As Excel.Worksheet
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim Empcode As String, D1 As String, D2 As String, D3 As String, D4 As String, D5 As String, D6 As String, D7 As String, D8 As String
Dim D16 As String, D9 As String, D10 As String, D11 As String, D12 As String, D13 As String, D14 As String, D15 As String, D17 As String
Dim D18 As String, D19 As String, D20 As String, D21 As String, D22 As String, D23 As String, D24 As String, D25 As String, D26 As String
Dim D27 As String, D28 As String, D29 As String, D30 As String, D31 As String
    If CD1.FileName = "" Then Exit Function
    Screen.MousePointer = vbHourglass
    Set xlApp = New Excel.Application
    xlApp.Workbooks.Open CD1.FileName
    xlApp.Workbooks(1).Activate
    Set xlSheet = xlApp.ActiveWorkbook.Sheets(1)
If FilePath = "" Then
    MsgBox "No file has been selected.", vbInformation
    Exit Function
Else
    If UCase(Right(FilePath, 5)) = ".XLS" Or UCase(Right(FilePath, 5)) = ".XLSX" Then
        If FindTable("ImportTbl") Then ConMain.Execute "drop table ImportTbl"
        Call CreateTableIntoAs("*", "shfinfo", "ImportTbl", " where 1=2")
'       strSql = "select * into ImportTbl from openrowset('microsoft.jet.oledb.4.0','excel 8.0; database=" & FilePath & "','select * from [sheet1$]')"
         Select Case bytBackEnd
'         Case 2
'         strSql = "select * from openrowset('microsoft.jet.oledb.4.0','excel 8.0; database=" & FilePath & "','select * from [sheet1$]')"
'         conmain.Execute "insert into ImportTbl " & strSql, lng, adExecuteNoRecords
        Case 2
            Dim cn As New ADODB.Connection
            Set cn = New ADODB.Connection
            If cn.State = 1 Then cn.Close
            cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"
            cn.Open
            'conmain.Execute "delete from ImportTbl"
            ConMain.Execute "drop table ImportTbl"
            strSql = "select * into ImportTbl FROM [Excel 8.0;DATABASE=" & FilePath & ";HDR=Yes].[sheet1$]"
            ConMain.Execute strSql
        Case 1, 3
            'Dim cn As New ADODB.Connection
            'Set cn = New ADODB.Connection
            Dim rsTemp As New ADODB.Recordset
            Dim rsAccess As New ADODB.Recordset
            If cn.State = 1 Then cn.Close
            cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"
            cn.Open
            ConMain.Execute "delete from ImportTbl"
            'conmain.Execute "drop table ImportTbl"
            strSql = "select * FROM [Excel 8.0;DATABASE=" & FilePath & ";HDR=Yes].[sheet1$]"
            If rsTemp.State = 1 Then rsTemp.Close
            rsTemp.Open strSql, cn, adOpenStatic
         
         If rsAccess.State = 1 Then rsAccess.Close
        rsAccess.Open "select * from ImportTbl", ConMain, adOpenDynamic, adLockOptimistic
        For i = 1 To rsTemp.RecordCount 'looping through all records of the excel file
            Empcode = Trim(xlSheet.Cells(i + 1, 1).Value)
            'If Not (ECode = "") Then
                'If Len(ECode) <> ECodeSize Then ECode = ZeroPadding(ECodeSize - Len(ECode), ECode)
                D1 = xlSheet.Cells(i + 1, 2).Value
                D2 = xlSheet.Cells(i + 1, 3).Value
                D3 = xlSheet.Cells(i + 1, 4).Value
                D4 = xlSheet.Cells(i + 1, 5).Value
                D5 = xlSheet.Cells(i + 1, 6).Value
                D6 = xlSheet.Cells(i + 1, 7).Value
                D7 = xlSheet.Cells(i + 1, 8).Value
                D8 = xlSheet.Cells(i + 1, 9).Value
                D9 = xlSheet.Cells(i + 1, 10).Value
                D10 = xlSheet.Cells(i + 1, 11).Value
                D11 = xlSheet.Cells(i + 1, 12).Value
                D12 = xlSheet.Cells(i + 1, 13).Value
                D13 = xlSheet.Cells(i + 1, 14).Value
                D14 = xlSheet.Cells(i + 1, 15).Value
                D15 = xlSheet.Cells(i + 1, 16).Value
                D16 = xlSheet.Cells(i + 1, 17).Value
                D17 = xlSheet.Cells(i + 1, 18).Value
                D18 = xlSheet.Cells(i + 1, 19).Value
                D19 = xlSheet.Cells(i + 1, 20).Value
                D20 = xlSheet.Cells(i + 1, 21).Value
                D21 = xlSheet.Cells(i + 1, 22).Value
                D22 = xlSheet.Cells(i + 1, 23).Value
                D23 = xlSheet.Cells(i + 1, 24).Value
                D24 = xlSheet.Cells(i + 1, 25).Value
                D25 = xlSheet.Cells(i + 1, 26).Value
                D26 = xlSheet.Cells(i + 1, 27).Value
                D27 = xlSheet.Cells(i + 1, 28).Value
                D28 = xlSheet.Cells(i + 1, 29).Value
                D29 = xlSheet.Cells(i + 1, 30).Value
                D30 = xlSheet.Cells(i + 1, 31).Value

                rsAccess.AddNew
                rsAccess.Fields("Empcode") = Empcode
                rsAccess.Fields("D1") = D1
                rsAccess.Fields("D2") = D2
                rsAccess.Fields("D3") = D3
                rsAccess.Fields("D4") = D4
                rsAccess.Fields("D5") = D5
                rsAccess.Fields("D6") = D6
                rsAccess.Fields("D7") = D7
                rsAccess.Fields("D8") = D8
                rsAccess.Fields("D9") = D9
                rsAccess.Fields("D10") = D10
                rsAccess.Fields("D11") = D11
                rsAccess.Fields("D12") = D12
                rsAccess.Fields("D13") = D13
                rsAccess.Fields("D14") = D14
                rsAccess.Fields("D15") = D15
                rsAccess.Fields("D16") = D16
                rsAccess.Fields("D17") = D17
                rsAccess.Fields("D18") = D18
                rsAccess.Fields("D19") = D19
                rsAccess.Fields("D20") = D20
                rsAccess.Fields("D21") = D21
                rsAccess.Fields("D22") = D22
                rsAccess.Fields("D23") = D23
                rsAccess.Fields("D24") = D24
                rsAccess.Fields("D25") = D25
                rsAccess.Fields("D26") = D26
                rsAccess.Fields("D27") = D27
                rsAccess.Fields("D28") = D28
                rsAccess.Fields("D29") = D29
                rsAccess.Fields("D30") = D30
                rsAccess.Fields("D31") = D31
                rsAccess.Update
            'End If
            Empcode = "": D1 = "": D2 = "": D3 = "": D4 = "": D5 = "": D6 = "": D7 = "": D8 = "": D9 = "": D10 = ""
            D11 = "": D12 = "": D13 = "": D14 = "": D15 = "": D16 = "": D17 = "": D18 = "": D19 = "": D20 = ""
            D21 = "": D22 = "": D23 = "": D24 = "": D25 = "": D26 = "": D27 = "": D28 = "": D29 = "": D30 = "": D31 = ""
            rsTemp.MoveNext
        Next i
        End Select
        
        If rsSftCode.State = 1 Then rsSftCode.Close
        rsSftCode.Open "select distinct shift from instshft", ConMain, adOpenStatic
        
        If rsLvCode.State = 1 Then rsLvCode.Close
        rsLvCode.Open "select distinct lvcode from Leavdesc", ConMain, adOpenStatic
        
        If rsImport.State = 1 Then rsImport.Close
        rsImport.Open "select * from ImportTbl", ConMain, adOpenStatic
        

        
        If Not (rsImport.EOF And rsImport.BOF) Then
            Screen.MousePointer = vbHourglass
            rsImport.MoveFirst
            While Not (rsImport.EOF)
                ConMain.Execute "update " & strFName & " set d1='" & UCase(rsImport.Fields("d1")) & "',d2='" & UCase(rsImport.Fields("d2")) & _
                "',d3='" & UCase(rsImport.Fields("d3")) & "',d4='" & UCase(rsImport.Fields("d4")) & "',d5='" & UCase(rsImport.Fields("d5")) & "',d6='" & UCase(rsImport.Fields("d6")) & _
                "',d7='" & UCase(rsImport.Fields("d7")) & "',d8='" & UCase(rsImport.Fields("d8")) & "',d9='" & UCase(rsImport.Fields("d9")) & "',d10='" & UCase(rsImport.Fields("d10")) & _
                "',D11='" & UCase(rsImport.Fields("d11")) & "',D12='" & UCase(rsImport.Fields("d12")) & "',d13='" & UCase(rsImport.Fields("d13")) & "',d14='" & UCase(rsImport.Fields("d14")) & _
                "',d15='" & UCase(rsImport.Fields("d15")) & "',d16='" & UCase(rsImport.Fields("d16")) & "',d17='" & UCase(rsImport.Fields("d17")) & "',d18='" & UCase(rsImport.Fields("d18")) & _
                "',d19='" & UCase(rsImport.Fields("d19")) & "',d20='" & UCase(rsImport.Fields("d20")) & "',d21='" & UCase(rsImport.Fields("d21")) & "',d22='" & UCase(rsImport.Fields("d22")) & _
                "',d23='" & UCase(rsImport.Fields("d23")) & "',d24='" & UCase(rsImport.Fields("d24")) & "',d25='" & UCase(rsImport.Fields("d25")) & "',d26='" & UCase(rsImport.Fields("d26")) & _
                "',d27='" & UCase(rsImport.Fields("d27")) & "',d28='" & UCase(rsImport.Fields("d28")) & "',d29='" & UCase(rsImport.Fields("d29")) & "',d30='" & UCase(rsImport.Fields("d30")) & _
                "',d31='" & UCase(rsImport.Fields("d31")) & "' where empcode='" & rsImport.Fields("empcode") & "'"
                rsImport.MoveNext
            Wend
            Call FillComboGrid
            Screen.MousePointer = vbNormal
            MsgBox "Import sucessfully."
         End If
    Else
         MsgBox "Selected File is not excel file. Select excel file and Try Again.", vbOKOnly + vbInformation, "Imoport Excel File"
         Screen.MousePointer = vbNormal
          Exit Function
    End If
End If
Exit Function
ERR_P:
'MsgBox Err.Number & " " & Err.Description
If Err.Number = -2147217900 Then
'Resume Next
    MsgBox "Excel file is in edit mode. First close the excel file and then import.", vbOKOnly + vbInformation, "Imoport Excel File"
'    MsgBox Err.Description
Else
    ShowError ("Import Data :: " & Me.Caption)
End If
Resume Next
End Function

Private Sub Form_Load()
On Error GoTo ERR_P
bytTmp = 1
'********
Select Case bytBackEnd
    Case 1, 2
        cmdImport.Visible = True
    Case 3
        cmdImport.Visible = True
End Select
Dim i As Integer    ' 21-05
optMnth.Value = True

For i = 0 To 31
    ColArr(i) = 0
Next
'********
Call SetFormIcon(Me)
Call GetRights
Call SetGridWidth
Call FillRotCombo
Call FillCombos
MSF1.ToolTipText = "Click on Empcode Column to SELECT Employee OR Click on any other Column to edit data"
bytTmp = 2
Exit Sub
ERR_P:
    ShowError ("Load :: " & Me.Caption)
End Sub

Private Sub GetRights()
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(4, 5, 2, 1)
If strTmp = "1" Then
    cmdChange.Enabled = True
    MSF1.Enabled = True
Else
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    cmdChange.Enabled = False
    MSF1.Enabled = False
End If
Exit Sub
ERR_P:
    ShowError ("GetRights :: " & Me.Caption)
    cmdChange.Enabled = False
    MSF1.Enabled = False
End Sub

Private Sub SetGridWidth()
Dim bytCnt As Byte
With MSF1                       '' Sizing
    .ColWidth(0) = 2500
    .ColWidth(1) = 1000
    .ColAlignment(1) = flexAlignCenterCenter
    For bytCnt = 2 To 32
        .ColWidth(bytCnt) = .ColWidth(bytCnt) * 0.45
        .ColAlignment(bytCnt) = flexAlignLeftCenter
    Next
End With
End Sub

Private Sub FillRotCombo()      '' Fills Rotation ComboBox
On Error GoTo ERR_P
Dim strArrTmp() As String, bytTmp As Byte
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select SCode,Mon_Oth,Skp,Pattern from Ro_Shift where scode <> '100' order by SCode", _
ConMain, adOpenStatic
If Not (adrsDept1.BOF And adrsDept1.EOF) Then
    ReDim strArrTmp(adrsDept1.RecordCount - 1, 3)
    cboRot.ColumnCount = 4
    cboRot.ListWidth = "6 cm"
    cboRot.ColumnWidths = "1cm;1 cm;2 cm; 2 cm"
    For bytTmp = 0 To adrsDept1.RecordCount - 1
        strArrTmp(bytTmp, 0) = adrsDept1("SCode")           '' Shift Code
        strArrTmp(bytTmp, 1) = adrsDept1("Mon_Oth")         '' Type
        strArrTmp(bytTmp, 2) = adrsDept1("Skp")             '' Skip type
        strArrTmp(bytTmp, 3) = adrsDept1("Pattern")         '' Shift Pattern
        adrsDept1.MoveNext
    Next
    cboRot.List = strArrTmp
    Erase strArrTmp
End If
Exit Sub
ERR_P:
    ShowError ("FillRotCombo :: " & Me.Caption)
End Sub

Private Sub FillCombos()
On Error GoTo ERR_P
Dim intTmp As Integer
cboMonth.clear
For intTmp = 1 To 12
    cboMonth.AddItem Choose(intTmp, "January", "February", "March", "April", "May", "June", _
    "July", "August", "September", "October", "November", "December")
Next
cboYear.clear
For intTmp = 1996 To 2097
    cboYear.AddItem CStr(intTmp)
Next
cboMonth.Text = MonthName(Month(Date))
cboYear.Text = pVStar.YearSel

Call SetCritCombos(cboDept)
If strCurrentUserType <> HOD Then cboDept.Text = "ALL"
Exit Sub
ERR_P:
    ShowError ("FillCombos::" & Me.Caption)
End Sub

Private Sub cboDept_Click()
On Error GoTo ERR_P
If cboDept.ListIndex < 0 Then Exit Sub               '' If No Department
Call ShowDays
Exit Sub
ERR_P:
    ShowError ("Department::" & Me.Caption)
End Sub

Private Sub cboMonth_Click()
    Call ShowDays
End Sub

Private Sub cboYear_Click()
    Call ShowDays
End Sub

Private Sub ShowDays()
If Trim(cboMonth.Text) = "" Then Exit Sub
If Trim(cboYear.Text) = "" Then Exit Sub
'' Check if File Exists
strFName = Left(cboMonth.Text, 3) & Right(cboYear.Text, 2) & "Shf"
Call SetSunday
If Not FindTable(strFName) Then
    If bytTmp = 1 Then Exit Sub
    MsgBox "Monthly shift file not found for " & cboMonth.Text & " " & cboYear.Text, vbExclamation
    MSF1.Rows = 1
    cmdExport.Enabled = False   ' 21-05
    Exit Sub
Else
    cmdExport.Enabled = True    ' 21-05
End If
Call FillComboGrid

End Sub

Public Function DaysInMonth(ByVal dDate As Date) As Integer
    DaysInMonth = Day(DateAdd("m", 1, dDate - Day(dDate) + 1) - 1)
End Function

Private Sub SetSunday()
Dim dttmp As Date, strTmp As String
Dim bytCnt As Byte, bytMon As Byte
Dim TotalMDays As Byte
TotalMDays = DaysInMonth(FdtLdt(MonthNumber(cboMonth.Text), cboYear.Text))
For bytCnt = 0 To TotalMDays + 1 ' 31
'For bytCnt = 1 To 31   'dont delete commented by  21-05
    MSF1.Col = bytCnt
    MSF1.Row = 0
    'MSF1.CellBackColor = vbNormal
    MSF1.CellBackColor = &H8000000F
Next
strTmp = FdtLdt(MonthNumber(cboMonth.Text), cboYear.Text)
dttmp = strTmp
For bytCnt = 0 To 6
    If UCase(Format(dttmp + bytCnt, "ddd")) = "SUN" Then Exit For
'     If UCase(Format(WeekDay(dttmp) + bytCnt, "ddd")) = "FRI" Then Exit For
Next
dttmp = dttmp + bytCnt
bytCnt = bytCnt + 1
bytMon = Month(dttmp)
Do While bytCnt <= TotalMDays And bytMon = Month(dttmp)
    MSF1.Col = bytCnt + 1
    MSF1.Row = 0
    MSF1.CellBackColor = vbRed
    dttmp = dttmp + 7
    bytCnt = Day(dttmp)
Loop
End Sub

Private Sub FillComboGrid()     '' Fills Employee Combo and Grid
On Error GoTo ERR_P
If Trim(strFName) = "" Then Exit Sub
If Trim(cboDept.Text) = "" Then Exit Sub
Dim strDeptTmp As String, strTempforCF As String
strDeptTmp = cboDept.List(cboDept.ListIndex, 0)
strDeptTmp = EncloseQuotes(strDeptTmp)
Select Case UCase(Trim(strDeptTmp))
    Case "", "ALL"
            strTempforCF = "Select Empmst.name, " & strFName & ".*  from " & strFName & ",Empmst where Empmst.Empcode = " & strFName & ".Empcode"
    Case Else

            strTempforCF = "Select Empmst.name, " & strFName & ".*  from " & strFName & ",Empmst where Empmst.Empcode = " & _
            strFName & ".Empcode And Empmst.Dept = " & cboDept.List(cboDept.ListIndex, 1)
End Select


Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MSAccess
        strTempforCF = strTempforCF & " Order by " & strFName & ".Empcode"
    Case 3  ''Oracle
        strTempforCF = strTempforCF & " Order by Empcode"
End Select
Call FillGrid(strTempforCF)
Exit Sub
ERR_P:
    ShowError ("Fill Combo Grid :: " & Me.Caption)
End Sub

Private Sub FillGrid(ByVal strTName As String)
On Error GoTo ERR_P
Dim intCnt As Integer
Dim DaysInM As Byte
DaysInM = DaysInMonth(FdtLdt(MonthNumber(cboMonth.Text), cboYear.Text))

If DaysInM = 28 Then
    MSF1.ColWidth(30) = 0
    MSF1.ColWidth(31) = 0
    MSF1.ColWidth(32) = 0
ElseIf DaysInM = 29 Then
    MSF1.ColWidth(31) = 0
    MSF1.ColWidth(32) = 0
ElseIf DaysInM = 30 Then
    MSF1.ColWidth(32) = 0
Else
    MSF1.ColWidth(30) = 450
    MSF1.ColWidth(31) = 450
    MSF1.ColWidth(32) = 450

End If
If adrsLeave.State = 1 Then adrsLeave.Close
adrsLeave.Open strTName, ConMain, adOpenStatic
If Not (adrsLeave.EOF And adrsLeave.BOF) Then
    For intCnt = 0 To DaysInM + 1 'adrsLeave.Fields.Count - 1
        MSF1.TextMatrix(0, intCnt) = UCase(adrsLeave.Fields(intCnt).name)
    Next
    MSF1.Rows = 1
    Do While Not adrsLeave.EOF
        If Not adrsLeave.EOF Then
            MSF1.Rows = MSF1.Rows + 1
        Else
            Exit Do
        End If
        With MSF1
            For intCnt = 0 To DaysInM + 1 'adrsLeave.Fields.Count - 1
                .TextMatrix(MSF1.Rows - 1, intCnt) = IIf(IsNull(adrsLeave(intCnt)), "", adrsLeave(intCnt))
            Next
        End With
        adrsLeave.MoveNext
    Loop
Else
    MSF1.Rows = 1
End If
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
    ''Resume Next
End Sub

Private Sub MSF1_Click()
Dim bytCnt As Byte
With MSF1
    If optDate.Value = True Then    ' 21-05
        Call DateWiseSel
    End If
    If .MouseRow = 0 Then Exit Sub   '*************
    If .MouseCol <> 0 And .MouseCol <> 1 Then
        bytShfMode = 6
        
        .CellBackColor = vbBlue
        frmSingleS.Show vbModal
        If bytShfMode = 9 Then
            If Not UpdateShift Then GoTo Norm
            .TextMatrix(.Row, .Col) = strDjFileN
        End If
        bytShfMode = 1
Norm:
        If (optDate.Value = True) And (ColArr(.Col) = 1) Then   ' 21-05
            .CellBackColor = &HC0FFFF
        Else
            .CellBackColor = vbNormal
        End If
    Else
        If .Rows = 1 Then Exit Sub
            If .CellBackColor = &HC0FFFF Then
                For bytCnt = 0 To 32
                    .Col = bytCnt
                    .CellBackColor = vbWhite
                Next
                For bytCnt = 1 To 32    ' 21-05
                    If ColArr(bytCnt) = 1 Then
                        .Col = bytCnt
                        .CellBackColor = &HC0FFFF
                    End If
                Next
                .Row = 0
            Else
                For bytCnt = 0 To 32
                    .Col = bytCnt
                    .CellBackColor = &HC0FFFF
                Next
            End If
    End If
End With
End Sub
Private Sub DateWiseSel()      ' 21-05
Dim bytCnt As Long
With MSF1
    If .Col <> 0 And .Row = 0 Then
        If ColArr(.Col) = 0 Then
            Screen.MousePointer = vbHourglass
            For bytCnt = 1 To .Rows - 1
                .Row = bytCnt
                .CellBackColor = &HC0FFFF
            Next
            Screen.MousePointer = vbNormal
            ColArr(.Col) = 1
        ElseIf ColArr(.Col) = 1 Then
            Screen.MousePointer = vbHourglass
            For bytCnt = 1 To .Rows - 1
                .Row = bytCnt
                .CellBackColor = vbWhite
            Next
            Screen.MousePointer = vbNormal
            ColArr(.Col) = 0
        End If
        .Row = 0
    End If
End With
End Sub

Private Function UpdateShift() As Boolean
On Error GoTo ERR_P
ConMain.Execute "Update " & strFName & " set " & MSF1.TextMatrix(0, MSF1.Col) & " = '" & _
strDjFileN & "' where Empcode = '" & MSF1.TextMatrix(MSF1.Row, 1) & "'"
UpdateShift = True
Exit Function
ERR_P:
    ShowError ("UpdateShift :: " & Me.Caption)
End Function

Private Sub cmdSA_Click()
    Call SelUnselAll(&HC0FFFF, MSF1)
End Sub

Private Sub cmdUA_Click()
    Call SelUnselAll(vbWhite, MSF1)
End Sub

Private Sub cmdChange_Click()
If Not CheckEmployee Then Exit Sub
If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "Select EmpCode,Name,location,Cat,STyp,F_Shf,SCode," & strKOff & ",Off2,WO_1_3,WO_2_4,Shf_Date," & _
"JoinDate,LeavDate From EmpMst Where EmpCode in " & strSelEmp, ConMain
If Not (adrsEmp.EOF And adrsEmp.BOF) Then
    Screen.MousePointer = vbHourglass
    Do While Not adrsEmp.EOF
        Call FillSchShift
        adrsEmp.MoveNext
    Loop
    Screen.MousePointer = vbNormal
End If
Call FillComboGrid
End Sub

Private Function CheckEmployee() As Boolean     '' Function to Check if Employees are
strSelEmp = ""                                  '' Selected or not

If cboRot.Text = "" Then
    MsgBox "Please select Rotation to change", vbInformation
    cboRot.SetFocus
    Exit Function
End If
If MSF1.Rows = 1 Then
    MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation, App.EXEName
    Exit Function
End If
MSF1.Col = 0
For i = 1 To MSF1.Rows - 1
    MSF1.Row = i
    If MSF1.CellBackColor = SELECTED_COLOR Then
        strSelEmp = strSelEmp & "'" & MSF1.Text & "',"
    End If
Next
If strSelEmp = "" Then
    MsgBox NewCaptionTxt("00049", adrsMod), vbExclamation
    Exit Function
Else
    strSelEmp = Left(strSelEmp, Len(strSelEmp) - 1)
    strSelEmp = "(" & strSelEmp & ")"
End If
CheckEmployee = True
End Function

Private Sub FillSchShift()          '' Fills Employee Shift For the Newly Added Employee
On Error GoTo ERR_P                 '' for the Current Month
'' Start Date Checks
Dim dttmp As Date
'' If the Employee has Already Left
If Not IsNull(adrsEmp("LeavDate")) Then
    dttmp = FdtLdt(cboMonth.ListIndex + 1, cboYear.Text, "F")
    If DateCompDate(adrsEmp("LeavDate")) <= dttmp Then Exit Sub
End If
'' Get Current Months Last Process Date
dttmp = FdtLdt(cboMonth.ListIndex + 1, cboYear.Text, "L")
'' Check on JoinDate
If DateCompDate(adrsEmp("JoinDate")) > dttmp Then Exit Sub
'' Check on Shift Date
If DateCompDate(Shft.startdate) > dttmp Then Exit Sub
'' End Date Checks
Call GetSENums(cboMonth.Text, cboYear.Text)
Call FillEmployeeDetails(adrsEmp("Empcode"))
''Heart of the Logic
typEmpRot.strShifttype = "R"     '' Shift Type
typEmpRot.strShiftCode = cboRot.Text
''
If UCase(MonthName(Month(Shft.startdate))) = UCase(cboMonth.Text) And Year(Shft.startdate) = CInt(cboYear.Text) Then Call AdjustSENums(DateCompDate(Shft.startdate))
'' For Rotation Shifts
'' Fill Other Skip Pattern and Shift Pattern Array
Call FillArrays
Select Case strCapSND
    Case "O"        '' After Specific Number of Days
        Call SpecificDaysShifts(adrsEmp("Empcode"), cboMonth.Text, cboYear.Text)
    Case "D"        '' Only on Fixed Days
        Call FixedDaysShifts(adrsEmp("Empcode"), cboMonth.Text, cboYear.Text)
    Case "W"        '' Only On Fixed Week days
        Call WeekDaysShifts(adrsEmp("Empcode"), cboMonth.Text, cboYear.Text)
End Select
'' Add that Record to the Shift File
Call UpdateAfterShiftDate(cboMonth.Text, cboYear.Text, adrsEmp("Empcode"))
Exit Sub
ERR_P:
    ShowError ("FillSchShift :: " & Me.Caption)
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub optDate_Click()    ' 21-05
    cmdImport.Enabled = False
    Call SelUnselAll(vbWhite, MSF1)
End Sub

Private Sub optMnth_Click()    ' 21-05
    Dim i As Integer
    Dim k As Long
    cmdImport.Enabled = True
    Call SelUnselAll(vbWhite, MSF1)
    With MSF1
    Screen.MousePointer = vbHourglass
    For i = 1 To 31
        If ColArr(i) = 1 Then
            For k = 1 To .Rows - 1
                .Col = i
                .Row = k
                .CellBackColor = vbWhite
                ColArr(i) = 0
            Next
        End If
    Next
    Screen.MousePointer = vbNormal
    .Row = 0
    End With
End Sub


