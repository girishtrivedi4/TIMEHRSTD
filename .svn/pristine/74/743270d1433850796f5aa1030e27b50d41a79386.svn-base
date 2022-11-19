VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExportReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frSap 
      Height          =   795
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   6885
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
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "D"
         Top             =   240
         Width           =   1155
      End
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
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   3
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
         Left            =   4590
         TabIndex        =   6
         Top             =   300
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
         Left            =   240
         TabIndex        =   5
         Top             =   300
         Width           =   2460
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   435
      Left            =   3960
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   435
      Left            =   2280
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5880
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.ComboBox cboDept 
      Height          =   315
      Index           =   0
      Left            =   3450
      TabIndex        =   8
      Top             =   1020
      Width           =   3315
      VariousPropertyBits=   612915227
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
      Index           =   0
      Left            =   2160
      TabIndex        =   7
      Top             =   1080
      Width           =   825
   End
End
Attribute VB_Name = "frmExportReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ReportModComplete As Boolean

Private Sub cmdExit_Click()
'    ReportType = "Periodic"
'    Load frmReports
'    frmReports.txtFrPeri.Text = txtFrom.Text
'    frmReports.txtToPeri.Text = txtTo.Text
'    frmReports.optGrp(2).Value = True
'    frmReports.optGrp_Click (2)
'    frmReports.Show vbModal
    Unload Me
End Sub

Private Sub cmdExport_Click()
   Dim path As String
   If strRepName = "Performance" Then
        CD.FileName = ""
        CD.Filter = "*.xls|*.xls"
        CD.Flags = cdlOFNFileMustExist
        CD.ShowSave
        If Trim(CD.FileName) = "" Then Exit Sub
        If CD.FileName = "" Then
            MsgBox "Exprt File is not given", vbCritical
            Exit Sub
        End If
        path = CD.FileName
    
        ReportType = "Periodic"
        Load frmReports
        frmReports.txtFrPeri.Text = txtFrom.Text
        frmReports.txtToPeri.Text = txtTo.Text
        frmReports.optGrp(2).Value = True
        If cboDept(0).Value <> "ALL" Then
            frmReports.cmbFrDepSel.Text = cboDept(0).Column(1)
            frmReports.cmbToDepSel.Text = cboDept(0).Column(1)
        End If
        
            typOptIdx.bytPer = 17
            frmReports.Visible = False
            Call frmReports.ReportsMod
            If ReportModComplete = False Then Exit Sub
        Call ExportExcel(path)
        Unload frmCRV
        Unload frmReports
    Else
        typDT.dtFrom = txtFrom.Text
        typDT.dtTo = txtTo.Text
        
        If (CDate(txtFrom.Text) + 31) < CDate(txtTo.Text) Then
            MsgBox "Date Range Should Not Greater Then 31 Days", vbInformation
            Exit Sub
        End If
        
        If Not AppendDataFile(Me) Then Exit Sub
        Call MonthEntries
        strSql = "SELECT Empmst.name AS EmpName, ImportTbl.* FROM Empmst INNER JOIN ImportTbl ON VAL(Empmst.empcode) = VAL(ImportTbl.empcode) "
        If cboDept(0).Value <> "ALL" Then strSql = strSql + " WHERE Empmst.dept = " & cboDept(0).Column(1)
        If Not Export(strSql) Then Exit Sub
    End If
End Sub


Public Sub ExportExcel(FilePath As String)
On Error GoTo Err
    Report.ExportOptions.FormatType = crEFTExcelDataOnly
    Report.ExportOptions.DiskFileName = FilePath
    Report.ExportOptions.DestinationType = crEDTDiskFile
    Report.ExportOptions.ExcelAreaType = crDetail
    Report.Export False
    MsgBox "Performance Export Data Completed", vbInformation, "Performance Data"
Err:
End Sub

Private Function SaveExportData(ByVal strmain As String) As Boolean
On Error GoTo Err
    If Dir(CD.FileName) <> "" Then Kill CD.FileName
    Dim TmpFile As String
    TmpFile = IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\Exp.xls"
    Open TmpFile For Append As #2
    Print #2, strmain
    Close #2
    
    Dim xlApp As Excel.Application
    Set xlApp = New Excel.Application
    xlApp.Workbooks.Open TmpFile
        xlApp.ActiveWorkbook.ActiveSheet.Columns("F").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("G").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("H").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("I").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("J").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("K").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("N").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("O").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("P").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("L").NumberFormat = "0.00"
        xlApp.ActiveWorkbook.ActiveSheet.Columns("M").NumberFormat = "0.00"
    xlApp.Workbooks(1).SaveAs CD1.FileName, xlOpenXMLWorkbook, , , , False
     xlApp.Quit
    Set xlApp = Nothing

    Kill TmpFile
    Exit Function
Err:
    ShowError ("SaveExportData :: " & Me.Caption)
If Err.Number = 0 Then SaveExportData = True
End Function

Private Sub Form_Load()
    txtFrom.Text = DateDisp(Date)
    txtTo.Text = DateDisp(Date)
    Call SetCritCombos(cboDept(0))
    cboDept(0).Text = "ALL"
    If strRepName = "Performance" Then
        Me.Caption = "Performance Export"
    Else
        Me.Caption = "Monthly Entries Export"
    End If
End Sub

Private Sub txtFrom_Click()
    Call ShowCalendar
End Sub

Private Sub txtTo_Click()
    Call ShowCalendar
End Sub
