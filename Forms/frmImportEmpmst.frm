VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B6A4AB44-1E6F-4C19-99CC-D9A1E4629B80}#4.0#0"; "PEDataGrid.ocx"
Begin VB.Form frmImportEmpmst 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Import"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   2760
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   2760
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Please Wait...."
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -15
         TabIndex        =   23
         Top             =   105
         Width           =   2745
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5415
      Left            =   7920
      TabIndex        =   14
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picControls 
      Height          =   10095
      Left            =   120
      ScaleHeight     =   10035
      ScaleWidth      =   7755
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.ListBox lstStatus 
         Height          =   1260
         Left            =   120
         TabIndex        =   21
         Top             =   8280
         Width           =   7575
      End
      Begin MSComDlg.CommonDialog CDialogExcellOpen 
         Left            =   6480
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   480
         Picture         =   "frmImportEmpmst.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   15
         Top             =   120
         Width           =   480
      End
      Begin VB.Frame Frame1 
         Caption         =   "Available Sheet"
         Height          =   1215
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   7575
         Begin VB.ComboBox cmbAvailableSheet 
            Height          =   360
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   360
            Width           =   2895
         End
         Begin LVbuttons.LaVolpeButton cmdShowData 
            Height          =   330
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            BTYPE           =   3
            TX              =   "&Show"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            BCOL            =   12648447
            FCOL            =   4210752
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "frmImportEmpmst.frx":030A
            ALIGN           =   1
            IMGLST          =   "(None)"
            IMGICON         =   "(None)"
            ICONAlign       =   0
            ORIENT          =   0
            STYLE           =   1
            IconSize        =   2
            SHOWF           =   -1  'True
            BSTYLE          =   0
         End
         Begin VB.Label Label4 
            Caption         =   "Choose Sheet of Excell"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Selection Input"
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   7575
         Begin VB.ComboBox cmbBackEnd 
            Height          =   360
            Left            =   4560
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   360
            Width           =   2895
         End
         Begin VB.ComboBox ComMonth_I 
            Height          =   360
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
         Begin VB.ComboBox ComYear_I 
            Height          =   360
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   720
            Width           =   1455
         End
         Begin VB.ListBox lstFileName 
            Height          =   540
            Left            =   1800
            TabIndex        =   4
            Top             =   1200
            Width           =   5655
         End
         Begin LVbuttons.LaVolpeButton cmdOpen 
            Height          =   450
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   794
            BTYPE           =   3
            TX              =   "&Open"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            BCOL            =   12648447
            FCOL            =   4210752
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "frmImportEmpmst.frx":0326
            ALIGN           =   1
            IMGLST          =   "(None)"
            IMGICON         =   "(None)"
            ICONAlign       =   0
            ORIENT          =   0
            STYLE           =   1
            IconSize        =   2
            SHOWF           =   -1  'True
            BSTYLE          =   0
         End
         Begin VB.Label Label5 
            Caption         =   "Select &BackEnd"
            Height          =   615
            Left            =   3360
            TabIndex        =   17
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Select &Month"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Select &Year"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1455
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   4560
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6376
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Available Data"
         TabPicture(0)   =   "frmImportEmpmst.frx":0342
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "PEDataGrid1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Message"
         TabPicture(1)   =   "frmImportEmpmst.frx":035E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstMessage"
         Tab(1).ControlCount=   1
         Begin PEDataGridControl.PEDataGrid PEDataGrid1 
            Height          =   3015
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   5318
         End
         Begin MSComctlLib.ListView lstMessage 
            Height          =   3015
            Left            =   -74880
            TabIndex        =   2
            Top             =   480
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   5318
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   1680
         TabIndex        =   19
         Top             =   9600
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin LVbuttons.LaVolpeButton cmdImport 
         Height          =   330
         Left            =   120
         TabIndex        =   20
         Top             =   9600
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "&Import"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   12648447
         FCOL            =   4210752
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmImportEmpmst.frx":037A
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.Label Label6 
         Caption         =   "This tab use only for viewing record and error message"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   4080
         Width           =   7575
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   7800
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label3 
         Caption         =   "This option use for import employee from outside source to MIS"
         Height          =   495
         Left            =   1440
         TabIndex        =   16
         Top             =   120
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmImportEmpmst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmImportEmpmst
' DateTime  : 13/10/07 12:31
' Author    : jagdish
' Purpose   :
'---------------------------------------------------------------------------------------

Dim cn As New ADODB.Connection
Dim fso As New FileSystemObject
Dim iRows As Long
Dim jColumns As Long
Dim strEmpCode As String
Dim rs As New ADODB.Recordset

Private Sub cmdImport_Click()
On Error GoTo Err
    rs.Requery
    lstStatus.Clear
    lstMessage.ListItems.Clear
    If ImportData Then
        Call AddRecordToListBox(lstStatus, "Import Completed")
    End If
Exit Sub
Err:
    Call ShowError("cmdImport_Click")
End Sub

Private Sub cmdOpen_Click()
On Error GoTo Err
    Me.Refresh
    Frame3.Visible = True
    lstFileName.Clear
    CDialogExcellOpen.Filter = "*.xls|*.xls"
    CDialogExcellOpen.ShowOpen
    If CDialogExcellOpen.FileName = "" Then Frame3.Visible = False: Exit Sub
    lstFileName.AddItem CDialogExcellOpen.FileName
    Call FillSheetCombo(cmbAvailableSheet, CDialogExcellOpen.FileName)
    Frame3.Visible = False
Exit Sub
Err:
    Call ShowError("cmdOpen_Click")
End Sub
Private Function FillSheetCombo(cmbI As ComboBox, XLsfile As String)
On Error GoTo ErrFillSheetCombo
    Dim rs As New ADODB.Recordset
    cn.CursorLocation = adUseClient
    If cn.State = 1 Then cn.Close
    cn.Open "Provider = " & _
    " Microsoft.Jet.OLEDB.4.0;" & " Data Source=" & XLsfile & _
    ";Extended Properties=Excel 8.0;"
    If rs.State = 1 Then rs.Close
    Set rs = cn.OpenSchema(adSchemaTables)
    cmbI.Clear
    Do While Not rs.EOF
        cmbI.AddItem rs.Fields("TABLE_NAME")
        rs.MoveNext
    Loop
    
    If cmbI.ListCount > -1 Then
        cmbI.Text = cmbI.List(0)
    End If
    
    Set rs = Nothing
    
Exit Function
ErrFillSheetCombo:
    Set cn = Nothing
    Set rs = Nothing
    Call ShowError("Error in FillSheetCombo")
End Function

Private Sub cmdShowData_Click()
On Error GoTo Err:
    If cmbAvailableSheet.Text <> "" Then
         If rs.State = 1 Then rs.Close
         rs.CursorLocation = adUseClient
         Call SetGridProperty(PEDGImport)
         rs.Open "SELECT * FROM [" & cmbAvailableSheet.Text & "]", cn
         Set PEDGImport.DataSource = rs
         cmdImport.Visible = True
    End If
Exit Sub
Err:
    Call ShowError("cmdShowData_Click")
End Sub
Private Function SetGridProperty(PEDGI As PEDataGrid)
    'PEDGI.CheckBoxes = True
End Function
Private Sub Form_Load()
On Error GoTo Err
    Call SetFormIcon(frmImportEmpmst)
    If ComboValue = False Then
        MsgBox "Initial Value Setting Problem in combo month and year", vbInformation
        End
    End If
    If fso.FolderExists(App.Path & "\ExpImp") = False Then
        fso.CreateFolder (App.Path & "\ExpImp")
    End If
    ComYear_I.Text = Year(Date)
    ComMonth_I.Text = MonthName(Month(Date), False)
    With VScroll1
        .Max = picControls.Height - 5200
        .SmallChange = .Max / 10
        .LargeChange = .Max / 4
    End With
    Call FillBackCombo
    Call SetColumn(lstMessage)
Exit Sub
Err:
    Call ShowError("Form_load" & Me.Caption)
End Sub
Private Function SetDataToList(lstI As ListView, _
strEmpCode As String, strDesc As String)
On Error GoTo Err
    Dim MiTEM As ListItem
    Set MiTEM = lstI.ListItems.Add()
    MiTEM.SubItems(1) = CStr(strEmpCode)
    MiTEM.SubItems(2) = CStr(strDesc)
Exit Function
Err:
    Call ShowError("SetDataToList")
End Function
Private Function SetColumn(lstI As ListView)
On Error GoTo Err
    lstI.ColumnHeaders.Add 1
    lstI.ColumnHeaders(1).Width = 0
    lstI.ColumnHeaders.Add 2, "Empcode", "Empcode"
    lstI.ColumnHeaders.Add 3, "Description", "Description"
Exit Function
Err:
    Call ShowError("SetColumn")
End Function

Private Function FillBackCombo()
On Error GoTo Err
    cmbBackEnd.AddItem "SQL"
    cmbBackEnd.AddItem "Access"
    cmbBackEnd.AddItem "Oracle"
    cmbBackEnd.Text = cmbBackEnd.List(2)
Exit Function
Err:
    Call ShowError("FillBackCombo")
End Function
Private Function ComboValue() As Boolean
On Error GoTo errorHandler
    For cnt = 1 To 10
        ComYear_I.AddItem (2004 + cnt)
    Next
    For cnt = 1 To 12
        ComMonth_I.AddItem MonthName(cnt)
    Next
    ComboValue = True
Exit Function
errorHandler:
    ComboValue = False
End Function

Private Function ImportData() As Boolean
On Error GoTo errorHandler

If Not FindTable("LVBAL" & Right(ComYear_I.Text, 2)) Then
    MsgBox "LVBAL" & Right(ComYear_I.Text, 2) & "TABLE NOT FOUND", vbInformation
    Exit Function
End If
'' Check for Demo Version
If InVar.blnVerType = "1" Then
    If CInt(InVar.lngEmp) <= adrsEmp.RecordCount Then
        MsgBox "Error", vbInformation
        Exit Function
    End If
End If

If Not (rs.EOF And rs.BOF) Then rs.MoveFirst
iRows = 0
ProgressBar1.Max = rs.RecordCount - 1
        
Do While Not rs.EOF
    Call AddRecordToListBox(lstStatus, "Import " & strEmpCode & " Employee")
    strEmpCode = FilterNull(rs.Fields("empcode"))
    Call AddRecordToListBox(lstStatus, "Start primary keys checking on " & strEmpCode & " Employee")
    If PrimaryForKeyCheck Then
        Call AddRecordToListBox(lstStatus, "Primary Key checking successfully on " & strEmpCode & " Employee")
        Call AddRecordToListBox(lstStatus, "Start Validation on following employee " & strEmpCode & " Employee")
        If ValidateModMaster Then
            Call AddRecordToListBox(lstStatus, "Validation completed successfully " & strEmpCode & " Employee")
            Call AddRecordToListBox(lstStatus, "Start validation add function on following " & strEmpCode & " Employee")
            If ValidateAddmaster Then
                Call AddRecordToListBox(lstStatus, "Completed validation add function for following " & strEmpCode & " Employee")
                Call AddRecordToListBox(lstStatus, "Start saving record following employee " & strEmpCode & " Employee")
                If SaveAddMaster Then
                     Call AddRecordToListBox(lstStatus, "Save Completed for following " & strEmpCode & " Employee")
                     Call SaveAddLog
                     Call AddRecordToListBox(lstStatus, "Open leave master for following employee " & strEmpCode & " Employee")
                     Call OpenLeaveMaster
                     Call AddRecordToListBox(lstStatus, "Start adding new employee in leave master " & strEmpCode & " Employee")
                     Call UpdateNewEmpLeave(FilterNull(rs.Fields("empcode")), FilterNull(rs.Fields("joindate")), _
                     rs!cat, Val(pVStar.YearSel))
                     Call AddRecordToListBox(lstStatus, "Completed leave master updation for following employee " & strEmpCode & " Employee")
                     Call AddRecordToListBox(lstStatus, "Fill employee details in shift master " & strEmpCode & " Employee")
                     Call FillEmpShift
                     Call AddRecordToListBox(lstStatus, "Completed shift filled data for following employee " & strEmpCode & " Employee")
                End If
            End If
        End If
    End If
    rs.MoveNext
    Call AddRecordToListBox(lstStatus, "------------------------------" & "Start Next employee")
    ProgressBar1.Value = iRows
    iRows = iRows + 1
    Me.Refresh
Loop
ImportData = True
Exit Function
errorHandler:
    Call ShowError("ImportData")
End Function
Private Function AddRecordToListBox(lstListbox As ListBox, strMessage As String)
On Error GoTo Err
    lstListbox.AddItem strMessage
    lstListbox.Refresh
    If lstListbox.ListCount > 0 Then
        lstListbox.ListIndex = lstListbox.ListCount - 1
    End If
Exit Function
Err:
    Call ShowError("AddRecordToListBox")
    Resume Next
End Function
Private Function ValidateAddmaster() As Boolean  '' Validates New record before Saving
Dim strEmpCode As String
On Error GoTo Err_P
ValidateAddmaster = True
If FilterNull(rs.Fields("Empcode")) <> "" Then
    If Len(FilterNull(rs.Fields("Empcode"))) <> pVStar.CodeSize Then
        Call SetDataToList(lstMessage, strEmpCode, "Employee code lenght is invalid")
        ValidateAddmaster = False
        Exit Function
    End If
Else
    Call SetDataToList(lstMessage, strEmpCode, "Employee code not allow blank")
    ValidateAddmaster = False
    Exit Function
End If

If CheckPrimaryKey("Empmst", "Empcode", FilterNull(rs.Fields("Empcode"))) Then
    Call SetDataToList(lstMessage, FilterNull(rs.Fields("Empcode")), "Empcode already resent")
    ValidateAddmaster = False
    Exit Function
End If

If FilterNull(rs.Fields("card")) <> "" Then
    If Len(FilterNull(rs.Fields("card"))) <> pVStar.CardSize Then
        Call SetDataToList(lstMessage, strEmpCode, "Employee card lenght is invalid")
        ValidateAddmaster = False
        Exit Function
    End If
Else
    Call SetDataToList(lstMessage, strEmpCode, "Employee card not allow blank")
    ValidateAddmaster = False
    Exit Function
End If

If CheckPrimaryKey("Empmst", "card", FilterNull(rs.Fields("card"))) Then
    Call SetDataToList(lstMessage, strEmpCode, "emp card already resent")
    ValidateAddmaster = False
    Exit Function
End If


If FilterNull(rs.Fields("name")) = "" Then
    Call SetDataToList(lstMessage, strEmpCode, "Employee name not allow blank")
    ValidateAddmaster = False
    Exit Function
End If

'' Check for Empty OT Rule
If FilterNull(rs.Fields("OTCODE")) = "" Then
    Call SetDataToList(lstMessage, strEmpCode, "OT Rule not allow blank")
    ValidateAddmaster = False
    Exit Function
End If

If FilterNull(rs.Fields("COCODE")) = "" Then
    Call SetDataToList(lstMessage, strEmpCode, "Co rule not allow blank")
    ValidateAddmaster = False
    Exit Function
End If

'' Check for Empty JoinDate
If Trim(FilterNull(rs.Fields("JOINDATE"))) = "" Then
    Call SetDataToList(lstMessage, strEmpCode, "Empty join date")
    ValidateAddmaster = False
    Exit Function
End If
'' Check for JoinDate Greater then Shift Date
If CDate(FilterNull(rs.Fields("JOINDATE"))) > DateCompDate(Shft.StartDate) Then
        Call SetDataToList(lstMessage, strEmpCode, "Invalid join date Its greater than shift date")
        ValidateAddmaster = False
        Exit Function
End If

'' Check for Empty Shift Code
If Shft.ShiftCode = "" Then
    Call SetDataToList(lstMessage, strEmpCode, "Shiftcode not allow blank")
    ValidateAddmaster = False
    Exit Function
End If
Exit Function
Err_P:
    Call SetDataToList(lstMessage, strEmpCode, Err.Description)
    ValidateAddmaster = False
    Resume Next
End Function

Private Function ValidateModMaster() As Boolean     '' Checks for the Validations
On Error GoTo Err_P
ValidateModMaster = True
If FilterNull(rs.Fields("shf_date")) = "" Then
    Call SetDataToList(lstMessage, strEmpCode, _
     "Shift Date Invalid")
    ValidateModMaster = False
    Exit Function
Else
    If bytDateF = 1 Then    '' American
        Shft.StartDate = DateDisp(Format(FilterNull(rs.Fields("shf_date")), "MM/DD/YYYY"))
    Else                    '' British
        Shft.StartDate = DateDisp(Format(FilterNull(rs.Fields("shf_date")), "DD/MM/YYYY"))
    End If
End If

If FilterNull(rs.Fields("joindate")) = "" Then
    Call SetDataToList(lstMessage, strEmpCode, "Join Date Invalid")
    ValidateModMaster = False
    Exit Function
End If

If Not CDate(Shft.StartDate) >= CDate(FilterNull(rs.Fields("joindate"))) Then
    Call SetDataToList(lstMessage, strEmpCode, "Join Date Invalid")
    ValidateModMaster = False
    Exit Function
End If

Select Case FilterNull(rs.Fields("styp"))
    Case ""
        Call SetDataToList(lstMessage, strEmpCode, _
            "Invalid Shift")
        ValidateModMaster = False
        Exit Function
    Case "F"
        If FilterNull(rs.Fields("f_shf")) = 100 Then
            Call SetDataToList(lstMessage, strEmpCode, _
            "Invalid Shift")
            ValidateModMaster = False
            Exit Function
        Else
            Shft.ShiftType = FilterNull(rs.Fields("styp"))
            Shft.ShiftCode = FilterNull(rs.Fields("f_shf"))
        End If
    Case "R"
        If FilterNull(rs.Fields("scode")) = "" Then
            Call SetDataToList(lstMessage, strEmpCode, _
            "Invalid Shift")
            ValidateModMaster = False
            Exit Function
        Else
            Shft.ShiftType = FilterNull(rs.Fields("styp"))
            Shft.ShiftCode = FilterNull(rs.Fields("scode"))
        End If
End Select

Shft.WO = Left(FilterNull(rs.Fields("off")), 2)        '' Week Off
Shft.WO1 = Left(FilterNull(rs.Fields("off2")), 2)
Shft.WO2 = Left(FilterNull(rs.Fields("wo_1_3")), 2)
Shft.WO3 = Left(FilterNull(rs.Fields("wo_2_4")), 2)
Shft.Action3Shift = FilterNull(rs.Fields("Action3Shift"))
Shft.WOHLAction = FilterNull(rs.Fields("WOHLAction"))
Shft.AutoOnPunch = IIf(FilterNull(rs.Fields("AutoForPunch")) = "1", True, False)
Shft.ActionBlank = FilterNull(rs.Fields("ActionBlank"))
Exit Function
Err_P:
    Call SetDataToList(lstMessage, strEmpCode, Err.Description)
    ValidateModMaster = False
    Resume Next
End Function

Private Function SaveAddMaster() As Boolean     '' Saves New Added Record
On Error GoTo Err_P
Dim strTmp(4) As String
Dim temp As String
Dim empname
If Shft.ShiftType = "F" Then
    strTmp(0) = "'F','" & Shft.ShiftCode & "','100'," & strDTEnc & DateCompStr(Shft.StartDate) & _
    strDTEnc & ",'" & Shft.WO & "','" & Shft.WO1 & "','" & Shft.WO2 & "','" & Shft.WO3 & "'"
Else
    strTmp(0) = "'R','100','" & Shft.ShiftCode & "'," & strDTEnc & DateCompStr(Shft.StartDate) & _
    strDTEnc & ",'" & Shft.WO & "','" & Shft.WO1 & "','" & Shft.WO2 & "','" & Shft.WO3 & "'"

End If
''For Mauritius 05-08-2003
''Start
If Trim(FilterNull(rs!LeavDate)) = "" Then
    strTmp(1) = "NULL"
Else
    strTmp(1) = GetTO_Char(CStr(rs!LeavDate))
End If
If Trim(FilterNull(rs!Birth_Dt)) = "" Then
    strTmp(2) = "NULL"
Else
    strTmp(2) = GetTO_Char(CStr(rs!Birth_Dt))
End If
If Trim(FilterNull(rs!confmdt)) = "" Then
    strTmp(3) = "NULL"
Else
    strTmp(3) = GetTO_Char(CStr(rs!confmdt))
End If
strTmp(4) = GetTO_Char(CStr(rs!joindate))
''End
'If Trim(strnullcheck(rs!salary)) = "" Then txtSal.Text = "0"
'' Insert Query

empname = Replace(CStr(rs!Name), "'", "''")
'MsgBox empname
temp = "insert into EmpMst(EmpCode,Card,Name,Designatn,Entry,OTCode," & _
"COCode,Shf_Chg,Cat,Dept," & strKGroup & " ,Location,Company,Div,Conv,STyp,F_Shf,SCode,Shf_Date," & _
strKOff & " ,Off2,WO_1_3,WO_2_4,LeavDate,Birth_Dt,BG,JoinDate,ConfmDt,Sex,Qualf,Email_Id,Salary," & _
"Reference,ResAdd1,ResAdd2,City,Pin,Phone,UDf1,UDf2,UDf3,UDf4,UDf5,UDf6,UDf7,UDf8,UDf9,UDf10," & _
"Name2,WOHLAction,Action3Shift,AutoForPunch,ActionBlank) Values('" & rs!EmpCode & "','" & _
rs!card & "','" & empname & "','" & FilterNull(rs!Designatn) & "','" & FilterNull(rs!entry) & _
"'," & FilterNull(rs!OTCode) & "," & _
FilterNull(rs!COCOde) & "," & IIf(FilterNull(rs!shf_chg) = "0", 0, 1) & _
",'" & FilterNull(rs!cat) & "'," & FilterNull(rs!dept) & "," & FilterNull(rs.Fields(20)) & "," & FilterNull(rs!Location) & _
"," & rs!Company & "," & rs!div & ",'" & IIf(FilterNull(rs!conv) = "Bus", "B", "O") & _
"'," & strTmp(0) & "," & strTmp(1) & "," & strTmp(2) & ",'" & FilterNull(rs!bg) & "'," & strTmp(4) & _
"," & strTmp(3) & ",'" & FilterNull(rs!sex) & "','" & FilterNull(rs!qualf) & "','" & FilterNull(rs!email_id) & "'," & _
IIf(FilterNull(rs!salary) = "", 0, rs!salary) & ",'" & FilterNull(rs!reference) & "','" & FilterNull(rs!resadd1) & "','" & FilterNull(rs!resadd2) & "','" & FilterNull(rs!city) & _
"','" & FilterNull(rs!pin) & "','" & FilterNull(rs!Phone) & "','" & FilterNull(rs!udf1) & "','" & FilterNull(rs!udf2) & "','" & _
FilterNull(rs!udf3) & "','" & FilterNull(rs!udf4) & "','" & FilterNull(rs!udf5) & "','" & FilterNull(rs!udf6) & "','" & _
FilterNull(rs!UDf7) & "','" & FilterNull(rs!udf8) & "','" & FilterNull(rs!udf9) & "','" & FilterNull(rs!udf10) & _
"','" & Trim(FilterNull(rs!Name2)) & "'," & Shft.WOHLAction & ",'" & Shft.Action3Shift & "'," & _
IIf(Shft.AutoOnPunch, 1, 0) & ",'" & Shft.ActionBlank & "' )"

VstarDataEnv.cnDJConn.Execute temp
SaveAddMaster = True
Exit Function
Err_P:
    Call SetDataToList(lstMessage, strEmpCode, Err.Description)
    SaveAddMaster = False
    Resume Next
End Function
Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo Err_P
Call AddActivityLog(lgADD_MODE, 1, 14)     '' Add Activity
Exit Sub
Err_P:
    Call SetDataToList(lstMessage, strEmpCode, Err.Description)
    Resume Next
End Sub
Private Sub FillEmpShift()          '' Fills Employee Shift For the Newly Added Employee
On Error GoTo Err_P                 '' for the Current Month
'' Start Date Checks
Dim dtTmp As Date
'' If the Employee has Already Left
If Trim(FilterNull(rs!LeavDate)) <> "" Then
    dtTmp = FdtLdt(Month(Date), CStr(Year(Date)), "F")
    If DateCompDate(rs!LeavDate) <= dtTmp Then Exit Sub
End If
'' Get Current Months Last Process Date
dtTmp = FdtLdt(Month(Date), CStr(Year(Date)), "L")
'' Check on JoinDate
If CDate(FilterNull(rs!joindate)) > dtTmp Then Exit Sub
'' Check on Shift Date
If DateCompDate(Shft.StartDate) > dtTmp Then Exit Sub
'' End Date Checks
If Not FindTable(Left(MonthName(Month(Date)), 3) & Right(CStr(Year(Date)), 2) & "Shf") Then
    'VstarDataEnv.cnDJConn.Execute "Select * into " & Left(MonthName(Month(Date)), 3) & _
    Right(CStr(Year(Date)), 2) & "Shf" & " from shfinfo where " & "1=2"
    Call CreateTableIntoAs("*", "shfinfo", Left(MonthName(Month(Date)), 3) & _
    Right(CStr(Year(Date)), 2) & "Shf", " Where 1=2")
    Call CreateTableIndexAs("MONYYSHF", Left(MonthName(Month(Date)), 3), Right(CStr(Year(Date)), 2))
    Call GetSENums(MonthName(Month(Date)), CStr(Year(Date)))
End If



Call GetSENums(MonthName(Month(Date)), CStr(Year(Date)))
adrsEmp.Requery
Call FillEmployeeDetails(FilterNull(rs!EmpCode))
If Month(Date) = Month(Shft.StartDate) And Year(Shft.StartDate) = Year(Date) Then Call AdjustSENums(DateCompDate(Shft.StartDate))
If typEmpRot.strShifttype = "F" Then
        '' If Fixed Shifts
        Call FixedShifts(rs!EmpCode, MonthName(Month(Date)), CStr(Year(Date)))
    Else
        '' if Rotation Shifts
        '' Fill Other Skip Pattern and Shift Pattern Array
        Call FillArrays
        Select Case strCapSND
            Case "O"        '' After Specific Number of Days
                Call SpecificDaysShifts(rs!EmpCode, MonthName(Month(Date)), CStr(Year(Date)))
            Case "D"        '' Only on Fixed Days
                Call FixedDaysShifts(rs!EmpCode, MonthName(Month(Date)), CStr(Year(Date)))
            Case "W"        '' Only On Fixed Week days
                Call WeekDaysShifts(rs!EmpCode, MonthName(Month(Date)), CStr(Year(Date)))
        End Select
    End If
    '' Add that Record to the Shift File
    Call AddRecordsToShift(MonthName(Month(Date)), CStr(Year(Date)), rs!EmpCode)
Exit Sub
Err_P:
    Call SetDataToList(lstMessage, strEmpCode, Err.Description)
    Resume Next
End Sub
Private Function CheckPrimaryKey(strMasterTable As String, _
    strPrimary As String, strKeyValue As String) As Boolean
Dim adrsTemp As New ADODB.Recordset
CheckPrimaryKey = True
If strKeyValue <> "" Then
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "SELECT " & strPrimary & " FROM " & _
        strMasterTable & "", VstarDataEnv.cnDJConn, adOpenKeyset, adLockOptimistic
    
    If Not adrsTemp.BOF And adrsTemp.EOF Then adrsTemp.MoveFirst
    adrsTemp.Find "" & strPrimary & "='" & _
        CStr(strKeyValue) & "'"
    
    If adrsTemp.EOF Then
        Call SetDataToList(lstMessage, strEmpCode, _
            "Referance problem " & strKeyValue & _
            " Value not present in " & strPrimary)
        CheckPrimaryKey = False
        Exit Function
    End If
Else
    Call SetDataToList(lstMessage, strEmpCode, "Blank referance not allow in following column " & strPrimary)
    CheckPrimaryKey = False
    Exit Function
End If
End Function

Public Function PrimaryForKeyCheck() As Boolean
On Error GoTo err_a
PrimaryForKeyCheck = True

If Not CheckPrimaryKey("Division", _
    "div", FilterNull(rs.Fields("div"))) Then
    PrimaryForKeyCheck = False
    Exit Function
End If
If Not CheckPrimaryKey("deptdesc", _
    "dept", FilterNull(rs.Fields("dept"))) Then
    PrimaryForKeyCheck = False
    Exit Function
End If
'indian airline not have group
'If Not CheckPrimaryKey("GroupMst", _
'    "" & strKGroup & "", FilterNull(rs.Fields("Group"))) Then
'    PrimaryForKeyCheck = False
'    Exit Function
'End If
If Not CheckPrimaryKey("catdesc", _
    "cat", FilterNull(rs.Fields("cat"))) Then
    PrimaryForKeyCheck = False
    Exit Function
End If
If Not CheckPrimaryKey("Location", _
    "Location", FilterNull(rs.Fields("Location"))) Then
    PrimaryForKeyCheck = False
    Exit Function
End If
If Not CheckPrimaryKey("company", _
    "company", FilterNull(rs.Fields("company"))) Then
    PrimaryForKeyCheck = False
End If
If FilterNull(rs.Fields("styp")) <> "" Then
    'for rotation shift
    If FilterNull(rs.Fields("styp")) = "R" Then
        If Not CheckPrimaryKey("ro_shift", _
            "scode", FilterNull(rs.Fields("scode"))) Then
            PrimaryForKeyCheck = False
            Exit Function
        End If
    'for fixed code
    ElseIf FilterNull(rs.Fields("styp")) = "F" Then
        If Not CheckPrimaryKey("instshft", _
            "shift", FilterNull(rs.Fields("shift"))) Then
            PrimaryForKeyCheck = False
            Exit Function
        End If
    End If
Else
    PrimaryForKeyCheck = False
    Exit Function
End If
If FilterNull(rs.Fields("OTCode")) = 100 Then
    If Not CheckPrimaryKey("CORul", _
        "COCode", FilterNull(rs.Fields("COCode"))) Then
        PrimaryForKeyCheck = False
        Exit Function
    End If
End If
If FilterNull(rs.Fields("COCOde")) = 100 Then
    If Not CheckPrimaryKey("OTRul", _
        "OTCode", FilterNull(rs.Fields("OTCode"))) Then
        PrimaryForKeyCheck = False
        Exit Function
    End If
End If
If FilterNull(rs.Fields("confmdt")) <> "" Then
    Call SetDataToList(lstMessage, strEmpCode, _
        "Check Position of month and date OR Invalid confirm date")
    PrimaryForKeyCheck = False
    Exit Function
End If
PrimaryForKeyCheck = True
Exit Function
err_a:
    Call SetDataToList(lstMessage, strEmpCode, Err.Description)
    PrimaryForKeyCheck = False
    Resume Next
End Function

Private Sub Text1_GotFocus()
    ComMonth_I.SetFocus
End Sub

Private Sub Text2_GotFocus()
    ComYear_I.SetFocus
End Sub

Private Sub MSF1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cn = Nothing
End Sub

Private Sub lstFileName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lstFileName.ListCount > 0 Then
        lstFileName.ToolTipText = lstFileName.List(lstFileName.ListIndex)
    End If
End Sub

Private Sub VScroll1_Change()
    picControls.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub
