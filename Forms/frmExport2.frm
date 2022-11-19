VERSION 5.00
Begin VB.Form frmExport2 
   Caption         =   "Export Utility"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboYear 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox PathTxt 
      Height          =   375
      Left            =   120
      ScrollBars      =   1  'Horizontal
      TabIndex        =   5
      Text            =   " "
      Top             =   1200
      Width           =   3195
   End
   Begin VB.CommandButton PathCmd 
      Caption         =   "..."
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      ToolTipText     =   "Click to Browse the .DAT File Path"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "E&XPORT"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cboMonth 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Select The Path for the Exported Dat File"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Year"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Month"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   270
      Width           =   975
   End
End
Attribute VB_Name = "frmExport2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strMonth As String, strMonth1 As String, strYear As String, strYear2 As String
Dim strFile1 As String, strFile2 As String, strWkDay1 As String, strWkday2 As String
Dim dtStart As Date, dtLast As Date, dtLast1 As Date
Dim strMonth1end As Integer, strMonthend As Integer
Dim adrsMon As New ADODB.Recordset
Dim adrsC As New ADODB.Recordset
Public strFilePath, FileName As String

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdExport_Click()
If strFilePath = "" Then
PathCmd.SetFocus
MsgBox "First Select The Path For The Dat File"
Exit Sub
End If
strMonth = cboMonth.Text
strYear = cboYear.Text
If strMonth = "January" Then
    strMonth1 = "december"
    strYear2 = strYear - 1
Else
    strMonth1 = MonthName(MonthNumber(strMonth) - 1)
    strYear2 = strYear
End If
dtStart = GetDateOfDay(1, strMonth, strYear)
'' Make the Shift Start Date of the Current Month
dtLast = FdtLdt(MonthNumber(strMonth), strYear, "L")
dtLast1 = FdtLdt(MonthNumber(strMonth1), strYear2, "L")
strFile1 = MakeName(strMonth, strYear, "trn")
strFile2 = MakeName(strMonth1, strYear2, "trn")
Call GetSENums(strMonth, strYear)
strMonthend = typSENum.bytEnd
Call GetSENums(strMonth1, strYear2)
strMonth1end = typSENum.bytEnd
Call MakeFileName
Call Export
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)            '' Sets the Forms Icon
frmExport2.Caption = "EXPORT UTILITY"
'Call RetCaption
Call FillCombos
'Call GetRights
cboMonth.Text = MonthName(Month(Date))
    cboYear.Text = pVStar.YearSel
    
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions", VstarDataEnv.cnDJConn, adOpenStatic
'Me.Caption = NewCaptionTxt("36001", adrsC)
End Sub

Private Sub FillCombos()
Dim intTmp As Integer
With cboMonth           '' Month
    .AddItem "January"
    .AddItem "February"
    .AddItem "March"
    .AddItem "April"
    .AddItem "May"
    .AddItem "June"
    .AddItem "July"
    .AddItem "August"
    .AddItem "September"
    .AddItem "October"
    .AddItem "November"
    .AddItem "December"
End With
With cboYear            '' Year
    For intTmp = 1997 To 2096
        .AddItem CStr(intTmp)
    Next
End With
End Sub

Public Sub OpenMasters()
On Error GoTo Err_particular
'' Open Category Master
If AdrsCat.State = 1 Then AdrsCat.Close
AdrsCat.Open "Select Cat," & strKDesc & ",HalfCutER,HalfCutLT,Erl_Allow,Erl_Ignore," & _
"Lt_Allow,Lt_Ignore from CatDesc where cat <> '100'", VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
'' Open Employee Master
If strSelEmp <> "" Then
    strSelEmp = " where Empcode in " & strSelEmp
End If
If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "Select Empcode,Name,Shf_Chg,Entry,Cat,Card,Conv," & strKOff & ",Off2,Wo_1_3," & _
"Wo_2_4,Styp,Joindate,Leavdate,SCode,F_Shf,OTCode,COCode,WOHLAction,Action3Shift," & _
"AutoForPunch,ActionBlank from Empmst", VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
'' Open Shift Master
If adRsintshft.State = 1 Then adRsintshft.Close
adRsintshft.Open "Select * from instshft where shift <> '100'", VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
'' Unuseable Recordsets
Exit Sub
Err_particular:
    Call ShowError("OpenMasters")
End Sub
Private Function GetDateOfDay(ByVal bytDay As Byte, ByVal strMonth As String, _
strYear As String) As String        '' Function to make Date
On Error GoTo ERR_P
Select Case bytDateF
    Case 1      '' American (MM/DD/YY)
        GetDateOfDay = Format(MonthNumber(strMonth), "00") & "/" & Format(bytDay, "00") & _
        "/" & strYear
    Case 2      '' British  (DD/MM/YY)
        GetDateOfDay = Format(bytDay, "00") & "/" & Format(MonthNumber(strMonth), "00") & _
        "/" & strYear
End Select
Exit Function
ERR_P:
    ShowError ("Get Date of day :: Rotation Module")
End Function

Public Sub Export()
Dim statusArray, ExpArray As Variant
Dim sngstatus As String
Dim i As Integer
If Not FindTable(strFile1) Then
    MsgBox NewCaptionTxt("36019", adrsC) & cboMonth.Text & " " & _
        cboYear.Text & NewCaptionTxt("00055", adrsMod), vbExclamation
    Exit Sub
End If
If Dir(strFilePath & "\" & FileName) <> "" Then Kill strFilePath & "\" & FileName
 statusArray = Array("A A ", "ADAD", "COCO", "CLCL", "ELEL", "ESES", "LWLW", "WOWO", "RHRH", "SLSL", "P P ", "HLHL", "OTHER")
    ExpArray = Array("ABS", "ADJ", "COF", "CSL", "ERL", "ESI", "LWP", "OFF", "RHO", "SKL", "WKD", "PDH", "WKD", "OTH")
If adrsMon.State = 1 Then adrsMon.Close
adrsMon.Open "select * from " & strFile1 & " Order by Empcode, date", VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
If adrsMon.EOF = False Then
adrsMon.MoveFirst
End If
Do While adrsMon.EOF = False
    sngstatus = "1.0"
    strprint = adrsMon!Empcode
    strprint = strprint & Space(7 - Len(strprint))
    strprint = strprint & ","
    strprint = strprint & Format(adrsMon!Date, "YYYY") & Format(adrsMon!Date, "MM") & Format(adrsMon!Date, "DD")
    strprint = strprint & ","
    For i = 0 To UBound(statusArray) - 1
        'If adrsMon!ovtim  Then
            strprint = strprint & ExpArray(UBound(ExpArray))
            strprint = strprint & "," & sngstatus
            Exit For
        'End If
       If i < UBound(statusArray) - 1 Then
         If adrsMon!presabs = statusArray(i) Then
            strprint = strprint & ExpArray(i)
            strprint = strprint & "," & sngstatus
            Exit For
         End If
       ElseIf i = UBound(statusArray) - 1 Then
        If adrsMon!presabs = statusArray(i) Then
         If adrsMon!arrtim > 0 Then
            strprint = strprint & ExpArray(i)
            strprint = strprint & "," & sngstatus
            Exit For
         Else
            strprint = strprint & ExpArray(i + 1)
            strprint = strprint & "," & sngstatus
            Exit For
         End If
        End If
       End If
    Next
    If i = UBound(statusArray) Then
        sngstatus = "0.5"
        strprint = strprint & Trim(Left(adrsMon!presabs, 2)) & Trim(Right(adrsMon!presabs, 2))
        strprint = strprint & "," & sngstatus
    End If
    Call WriteToFile
    adrsMon.MoveNext
    If adrsMon.EOF = True Then
    Exit Do
    End If
    
Loop
MsgBox "Data is Exported Successfully."
End Sub


Public Sub WriteToFile()
Open strFilePath & "\" & FileName For Append As #1
Print #1, strprint
Close #1
End Sub

Private Sub PathCmd_Click()
DirPathfrm.Show vbModal
End Sub

Public Sub MakeFileName()
FileName = "PYAT" & Right(cboYear.Text, 2) & Format(MonthNumber(cboMonth.Text), "00") & ".DAT"
End Sub
