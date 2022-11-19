VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExportCustom 
   Caption         =   "Export"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   6750
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboMonth 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmExportCustom.frx":0000
      Left            =   3570
      List            =   "frmExportCustom.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   1275
   End
   Begin VB.ComboBox cboYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4890
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   825
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   435
      Left            =   4080
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Export"
      Height          =   435
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblMonYea 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Month For Export"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmExportCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim strFile As String
    Dim lvtrnFile As String
    Dim Empcode As String
    Dim strLinedata As String, strData As String
    
    strFile = MakeName(cboMonth.Text, cboYear.Text, "trn")
    If Not (FindTable(strFile)) Then
        MsgBox "Monthly Transaction file not present for selected month", vbInformation
        Exit Sub
    End If
    lvtrnFile = "LvTrn" & Right(cboYear, 2)
    If Not FindTable(lvtrnFile) Then
        MsgBox "Yearly Transaction file not present for selected month", vbInformation
        Exit Sub
    
    End If
    
        CD.Filter = "*.xls|*.xls"
        CD.ShowSave
        If Trim(CD.FileName) = "" Then Exit Sub
       Dim rsTrn As New ADODB.Recordset
       rsTrn.Open "Select * from " & lvtrnFile & " Where MONTH(lst_date)= " & cboMonth.ListIndex + 1, ConMain, adOpenKeyset, adLockOptimistic
       
'      strData = vbTab & "Date" & vbTab & "Time In" & vbTab & "Time Out" & vbTab & "Abs/Pres" & vbTab & "Total Working Hours" & vbTab & "Overtime" & vbCrLf
      
     strSql = "SELECT " & strFile & ".*, empmst.name FROM " & strFile & " , empmst"
     strSql = strSql + " WHERE " & strFile & ".Empcode=[empmst].[Empcode]  "
     strSql = strSql + " Order By " & strFile & ".Empcode, " & strFile & ".Date"
    
    Dim adrsForm As New ADODB.Recordset
    If adrsForm.State = 1 Then adrsForm.Close
     adrsForm.Open strSql, ConMain, adOpenKeyset, adLockOptimistic
     
    Dim wrkHrs As Single, Twrkrs As Single
    Dim StrHead As String
    
   Do While Not adrsForm.EOF
    wrkHrs = IIf(adrsForm.Fields("wrkhrs") > 8, 8, adrsForm.Fields("wrkhrs"))
    Twrkrs = TimAdd(Twrkrs, wrkHrs)
'     strLinedata = vbTab & IIf(InStr(1, strData, adrsForm.Fields("name")), "", adrsForm.Fields("name"))
     
    If InStr(1, strData, adrsForm.Fields("name")) < 1 Then
        StrHead = vbTab & "Employee name : " & adrsForm.Fields("name") & vbCrLf
        StrHead = StrHead & vbTab & "Date" & vbTab & "Time In" & vbTab & "Time Out" & vbTab & "Abs/Pres" & vbTab & "Total Working Hours" & vbTab & "Overtime" & vbCrLf
    Else
        StrHead = ""
    End If
     strLinedata = ""
     strLinedata = StrHead & strLinedata & vbTab & Format(adrsForm.Fields("date"), "dd/MMM/yyyy") & vbTab & Format(adrsForm.Fields("arrtim"), "00.00") & vbTab & Format(adrsForm.Fields("deptim"), "00.00") & vbTab & adrsForm.Fields("presabs") & vbTab & wrkHrs & vbTab & Format(adrsForm.Fields("ovtim"), "00.00")
     strData = strData & strLinedata & vbCrLf
      
      Empcode = adrsForm.Fields("Empcode")
      adrsForm.MoveNext
        
      If Not adrsForm.EOF Then
        If Empcode <> adrsForm.Fields("Empcode") Then
            rsTrn.Filter = "Empcode= '" & Empcode & "'"
            If Not rsTrn.EOF Then
                strData = strData & vbTab & "Total Presesnt Days = " & rsTrn.Fields("paiddays") & " Days (Including Holiday)" & vbCrLf
                strData = strData & vbTab & "Total Workng Hours = " & Format(Twrkrs, "0.00") & vbCrLf
                strData = strData & vbTab & "Total Overtime in Hours /Days = " & Format(rsTrn.Fields("ot_hrs"), "0.00") & vbCrLf & vbCrLf & vbCrLf
            Else
                strData = strData & vbTab & "Total Presesnt Days = 0 Days (Including Holiday)" & vbCrLf
                strData = strData & vbTab & "Total Working Hours = 0" & vbCrLf
                strData = strData & vbTab & "Total Overtime in Hours /Days = 0 " & vbCrLf & vbCrLf & vbCrLf
            End If
            rsTrn.Filter = ""
            Twrkrs = 0
        End If
      Else
            rsTrn.Filter = "Empcode= '" & Empcode & "'"
            
            If Not rsTrn.EOF Then
              strData = strData & vbTab & "Total Presesnt Days = " & rsTrn.Fields("paiddays") & " (Including Holiday)" & vbCrLf
                strData = strData & vbTab & "Total Working Hours = " & Format(Twrkrs, "0.00") & vbCrLf
                strData = strData & vbTab & "Total Overtime in Hours /Days = " & Format(rsTrn.Fields("ot_hrs"), "0.00") & vbCrLf & vbCrLf & vbCrLf
            Else
            
            End If
            Twrkrs = 0
        
      End If
      
        
   Loop
    
        Dim TmpFile As String
    
    TmpFile = IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\Exp.xls"
    Open TmpFile For Append As #2
    Print #2, strData
    Close #2

If Dir(CD.FileName) <> "" Then Kill CD.FileName
    Set xlApp = New Excel.Application
    xlApp.Workbooks.Open TmpFile
    xlApp.Workbooks(1).SaveAs CD.FileName, xlOpenXMLWorkbook, , , , False
    
    xlApp.Visible = True
  xlApp.ActiveWorkbook.ActiveSheet.Columns("B:H").EntireColumn.AutoFit
  xlApp.ActiveWorkbook.ActiveSheet.Range("B:H").HorizontalAlignment = Excel.xlRight
  xlApp.ActiveWorkbook.ActiveSheet.Columns("D:H").NumberFormat = "0.00"
  
   
    Set xlApp = Nothing
'    xlApp.Quit
    Kill TmpFile
    
    
    MsgBox "Excel Export Completed Successfully", vbInformation, "Excel Export"
End Sub

Private Sub Form_Load()
    For intTmp = 1 To 12
    cboMonth.AddItem MonthName(intTmp)
    Next
    For intTmp = 2002 To 2029
    cboYear.AddItem CStr(intTmp)
    Next
    cboMonth.Text = MonthName(Month(Date))
    cboYear.Text = CStr(Year(Date))
End Sub
