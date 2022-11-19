VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Employee Details"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   405
      Left            =   9570
      TabIndex        =   31
      Top             =   6480
      Width           =   1395
   End
   Begin VB.Frame frSet 
      Caption         =   "Select Details"
      Height          =   6435
      Left            =   7920
      TabIndex        =   10
      Top             =   0
      Width           =   4665
      Begin VB.CommandButton cmdComp 
         Caption         =   "Set Company"
         Height          =   375
         Left            =   30
         TabIndex        =   29
         Top             =   5520
         Width           =   1725
      End
      Begin VB.CommandButton cmdDiv 
         Caption         =   "Set Division"
         Height          =   375
         Left            =   30
         TabIndex        =   27
         Top             =   4980
         Width           =   1725
      End
      Begin VB.CommandButton cmdDesig 
         Caption         =   "Set Designation"
         Height          =   375
         Left            =   60
         TabIndex        =   25
         Top             =   4410
         Width           =   1725
      End
      Begin VB.CommandButton cmdEnt 
         Caption         =   "Set Entries"
         Height          =   375
         Left            =   60
         TabIndex        =   23
         Top             =   3810
         Width           =   1725
      End
      Begin VB.CommandButton cmdCO 
         Caption         =   "Set CO Rule"
         Height          =   375
         Left            =   60
         TabIndex        =   21
         Top             =   3210
         Width           =   1725
      End
      Begin VB.CommandButton cmdOT 
         Caption         =   "Set OT Rule"
         Height          =   375
         Left            =   60
         TabIndex        =   19
         Top             =   2610
         Width           =   1725
      End
      Begin VB.CommandButton cmdLoca 
         Caption         =   "Set Location"
         Height          =   375
         Left            =   60
         TabIndex        =   17
         Top             =   2010
         Width           =   1725
      End
      Begin VB.CommandButton cmdGroup 
         Caption         =   "Set Group"
         Height          =   375
         Left            =   60
         TabIndex        =   15
         Top             =   1410
         Width           =   1725
      End
      Begin VB.CommandButton cmdDept 
         Caption         =   "Set Department"
         Height          =   375
         Left            =   60
         TabIndex        =   13
         Top             =   810
         Width           =   1725
      End
      Begin VB.CommandButton cmdcat 
         Caption         =   "Set Category"
         Height          =   375
         Left            =   60
         TabIndex        =   11
         Top             =   210
         Width           =   1725
      End
      Begin MSForms.ComboBox cboDesig 
         Height          =   315
         Left            =   1920
         TabIndex        =   26
         Top             =   4440
         Width           =   2625
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4630;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1500;4500"
      End
      Begin MSForms.ComboBox cboComp 
         Height          =   315
         Left            =   1890
         TabIndex        =   30
         Top             =   5520
         Width           =   2625
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4630;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1500;4500"
      End
      Begin MSForms.ComboBox cboDiv 
         Height          =   315
         Left            =   1890
         TabIndex        =   28
         Top             =   4980
         Width           =   2625
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4630;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1500;4500"
      End
      Begin MSForms.ComboBox cboEnt 
         Height          =   315
         Left            =   1890
         TabIndex        =   24
         Top             =   3840
         Width           =   2625
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4630;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1500;4500"
      End
      Begin MSForms.ComboBox cboCO 
         Height          =   315
         Left            =   1890
         TabIndex        =   22
         Top             =   3210
         Width           =   2625
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4630;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1500;4500"
      End
      Begin MSForms.ComboBox cboOT 
         Height          =   315
         Left            =   1890
         TabIndex        =   20
         Top             =   2640
         Width           =   2625
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4630;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1500;4500"
      End
      Begin MSForms.ComboBox cboLoca 
         Height          =   315
         Left            =   1890
         TabIndex        =   18
         Top             =   2040
         Width           =   2625
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4630;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1500;4500"
      End
      Begin MSForms.ComboBox cboGroup 
         Height          =   315
         Left            =   1890
         TabIndex        =   16
         Top             =   1410
         Width           =   2625
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4630;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1500;4500"
      End
      Begin MSForms.ComboBox cboDept 
         Height          =   315
         Left            =   1890
         TabIndex        =   14
         Top             =   810
         Width           =   2625
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4630;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1500;4500"
      End
      Begin MSForms.ComboBox cboCat 
         Height          =   315
         Left            =   1890
         TabIndex        =   12
         Top             =   210
         Width           =   2625
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4630;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1500;4500"
      End
   End
   Begin VB.Frame frEmp 
      Caption         =   "Employee Selection"
      Height          =   6435
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   7905
      Begin VB.CommandButton cmdUA 
         Caption         =   "Unselect All"
         Height          =   435
         Left            =   6390
         TabIndex        =   9
         Top             =   2580
         Width           =   1425
      End
      Begin VB.CommandButton cmdSA 
         Caption         =   "Select All"
         Height          =   435
         Left            =   6390
         TabIndex        =   8
         Top             =   2160
         Width           =   1425
      End
      Begin VB.CommandButton cmdUR 
         Caption         =   "Unselect Range"
         Height          =   465
         Left            =   6390
         TabIndex        =   7
         Top             =   1560
         Width           =   1425
      End
      Begin VB.CommandButton cmdSR 
         Caption         =   "Select Range"
         Height          =   435
         Left            =   6390
         TabIndex        =   6
         Top             =   1140
         Width           =   1425
      End
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   5595
         Left            =   0
         TabIndex        =   5
         ToolTipText     =   "Click to Toggle Selection"
         Top             =   780
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   9869
         _Version        =   393216
         Rows            =   1
         Cols            =   13
         FixedCols       =   0
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSForms.ComboBox cboTo 
         Height          =   315
         Left            =   4800
         TabIndex        =   4
         Top             =   270
         Width           =   2925
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "5159;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1500;4500"
      End
      Begin MSForms.ComboBox cboFrom 
         Height          =   315
         Left            =   930
         TabIndex        =   2
         Top             =   270
         Width           =   3045
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "5371;556"
         ListWidth       =   6000
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
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
         Left            =   4440
         TabIndex        =   3
         Top             =   330
         Width           =   195
      End
      Begin VB.Label lblEmp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From "
         Height          =   195
         Left            =   510
         TabIndex        =   1
         Top             =   330
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset
'' Other Variables
Dim strEmp As String

Private Sub cboFrom_Click()
If cboFrom.ListIndex < 0 Then Exit Sub
cboTo.ListIndex = cboTo.ListCount - 1
End Sub

Private Sub cmdcat_Click()
If Trim(cboCat.Text) = "" Then Exit Sub
strEmp = CheckEmployees
If strEmp = "" Then Exit Sub
Call ReflectChanges("CAT", cboCat.List(cboCat.ListIndex, 1), True)
Call AuditInfo("UPDATE", Me.Caption, "Set Category:  " & cboCat.Text)
End Sub

Private Sub cmdCO_Click()
If Trim(cboCO.Text) = "" Then Exit Sub
strEmp = CheckEmployees
If strEmp = "" Then Exit Sub
Call ReflectChanges("COCode", IIf(UCase(Trim(cboCO.Text)) = "NONE", 100, cboCO.List(cboCO.ListIndex, 1)))
Call AuditInfo("UPDATE", Me.Caption, "Set CO Rules:  " & cboCO.Text)
End Sub

Private Sub cmdComp_Click()
If Trim(cboComp.Text) = "" Then Exit Sub
strEmp = CheckEmployees
If strEmp = "" Then Exit Sub
Call ReflectChanges("COMPANY", cboComp.List(cboComp.ListIndex, 1))
Call AuditInfo("UPDATE", Me.Caption, "Set Company:  " & cboComp.Text)
End Sub

Private Sub cmdDept_Click()
If Trim(cboDept.Text) = "" Then Exit Sub
strEmp = CheckEmployees
If strEmp = "" Then Exit Sub
Call ReflectChanges("DEPT", cboDept.List(cboDept.ListIndex, 1))
Call AuditInfo("UPDATE", Me.Caption, "Set Department:  " & cboDept.Text)
End Sub

Private Sub cmdDesig_Click()
    If Trim(cboDesig.Text) = "" Then Exit Sub
    strEmp = CheckEmployees
    If strEmp = "" Then Exit Sub
    Call ReflectChanges("DESIGNATN", cboDesig.List(cboDesig.ListIndex, 1))
    Call AuditInfo("UPDATE", Me.Caption, "Set Designation:  " & cboDesig.Text)
End Sub

Private Sub cmdDesig_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 3))))
End If
End Sub

Private Sub cmdDiv_Click()
If Trim(cboDiv.Text) = "" Then Exit Sub
strEmp = CheckEmployees
If strEmp = "" Then Exit Sub
Call ReflectChanges("DIV", cboDiv.List(cboDiv.ListIndex, 1))
Call AuditInfo("UPDATE", Me.Caption, "Set Division:  " & cboDiv.Text)
End Sub

Private Sub cmdEnt_Click()
If Trim(cboEnt.Text) = "" Then Exit Sub
strEmp = CheckEmployees
If strEmp = "" Then Exit Sub
Call ReflectChanges("ENTRY", cboEnt.Text)
Call AuditInfo("UPDATE", Me.Caption, "Set Entries:  " & cboEnt.Text)
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdGroup_Click()
If Trim(cboGroup.Text) = "" Then Exit Sub
strEmp = CheckEmployees
If strEmp = "" Then Exit Sub
Call ReflectChanges("" & strKGroup & "", cboGroup.List(cboGroup.ListIndex, 1))
Call AuditInfo("UPDATE", Me.Caption, "Set Group:  " & cboGroup.Text)
End Sub

Private Sub cmdLoca_Click()
If Trim(cboLoca.Text) = "" Then Exit Sub
strEmp = CheckEmployees
If strEmp = "" Then Exit Sub
Call ReflectChanges("LOCATION", cboLoca.List(cboLoca.ListIndex, 1))
Call AuditInfo("UPDATE", Me.Caption, "Set Location:  " & cboLoca.Text)
End Sub

Private Sub cmdOT_Click()
If Trim(cboOT.Text) = "" Then Exit Sub
strEmp = CheckEmployees
If strEmp = "" Then Exit Sub
Call ReflectChanges("OTCode", IIf(UCase(Trim(cboOT.Text)) = "NONE", 100, cboOT.List(cboOT.ListIndex, 1)))
Call AuditInfo("UPDATE", Me.Caption, "Set OT Rules:  " & cboOT.Text)
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Call SetFormIcon(Me, True)
Call GetRights
Call RetCaptions
Call FillCombos
Call FillGrid
Exit Sub
ERR_P:
    ShowError ("Load::" & Me.Caption)
End Sub

Private Sub FillGrid()
On Error GoTo ERR_P
Dim intTmp As Integer, strTmp As String
If strCurrentUserType = HOD Then
    ''Original ->strTmp = " Where Dept=" & intCurrDept & " "
    strTmp = strCurrData
End If
If GetFlagStatus("LocationRights") Then strTmp = strCurrData
If adrsEmp.State = 1 Then adrsEmp.Close

adrsEmp.Open "Select Empcode,Name,Empmst.Cat,Empmst.Dept,Empmst." & strKGroup & _
",Empmst.Location,OTCode,COCode,Entry,Designatn,empmst.Company,empmst.Div " & _
" from Empmst " & strTmp & " order by Empcode", ConMain, adOpenStatic, adLockReadOnly

intTmp = 1
If adrsEmp.EOF Then
    MSF3.Rows = 1
    Exit Sub
Else
    MSF3.Rows = adrsEmp.RecordCount + 1
End If
Do While Not adrsEmp.EOF
    With MSF3
        .TextMatrix(intTmp, 0) = adrsEmp("Empcode")
        .TextMatrix(intTmp, 1) = IIf(IsNull(adrsEmp("Name")), "", adrsEmp("Name"))
        .TextMatrix(intTmp, 2) = IIf(IsNull(adrsEmp("Cat")), "", adrsEmp("Cat"))
        .TextMatrix(intTmp, 3) = IIf(IsNull(adrsEmp("Dept")), "", adrsEmp("Dept"))
        .TextMatrix(intTmp, 4) = IIf(IsNull(adrsEmp("Group")), "", adrsEmp("Group"))
        .TextMatrix(intTmp, 5) = IIf(IsNull(adrsEmp("Location")), "", adrsEmp("Location"))
        .TextMatrix(intTmp, 6) = IIf(IsNull(adrsEmp("OTCode")) Or adrsEmp("OTCode") = 100, "None", adrsEmp("OTCode"))
        .TextMatrix(intTmp, 7) = IIf(IsNull(adrsEmp("COCode")) Or adrsEmp("COCode") = 100, "None", adrsEmp("COCode"))
        .TextMatrix(intTmp, 8) = IIf(IsNull(adrsEmp("Entry")), "", adrsEmp("Entry"))
        .TextMatrix(intTmp, 9) = IIf(IsNull(adrsEmp("Designatn")), "", adrsEmp("Designatn"))
        .TextMatrix(intTmp, 10) = IIf(IsNull(adrsEmp("Company")), "", adrsEmp("Company"))
        .TextMatrix(intTmp, 11) = IIf(IsNull(adrsEmp("Div")), "", adrsEmp("Div"))

        intTmp = intTmp + 1
    End With
    adrsEmp.MoveNext
Loop
Exit Sub
ERR_P:
    ShowError ("FillGrid::" & Me.Caption)
   ' Resume Next
End Sub

Private Sub FillCombos()
On Error GoTo ERR_P
Dim bytTmp As Byte
'' Employee
Call ComboFill(cboFrom, 1, 2)
If cboFrom.ListCount > 0 Then cboFrom.ListIndex = 0
Call ComboFill(cboTo, 1, 2)
If cboTo.ListCount > 0 Then cboTo.ListIndex = cboTo.ListCount - 1
'' Category
Call ComboFill(cboCat, 3, 2)
If cboCat.ListCount > 0 Then cboCat.ListIndex = 0
If cboCat.ListCount = 0 Then cmdcat.Enabled = False
'' Department
Call ComboFill(cboDept, 2, 2)
If cboDept.ListCount > 0 Then cboDept.ListIndex = 0
If cboDept.ListCount = 0 Then cmdDept.Enabled = False
'' Group
Call ComboFill(cboGroup, 8, 2)
If cboGroup.ListCount > 0 Then cboGroup.ListIndex = 0
If cboGroup.ListCount = 0 Then cmdGroup.Enabled = False
'' Location
Call ComboFill(cboLoca, 11, 2)
If cboLoca.ListCount > 0 Then cboLoca.ListIndex = 0
If cboLoca.ListCount = 0 Then cmdLoca.Enabled = False
'' OT Rule
Call ComboFill(cboOT, 9, 2)
cboOT.AddItem "None"
If cboOT.ListCount > 0 Then cboOT.ListIndex = 0
If cboOT.ListCount = 0 Then cmdOT.Enabled = False
'' CO Rule
Call ComboFill(cboCO, 10, 2)
cboCO.AddItem "None"
If cboCO.ListCount > 0 Then cboCO.ListIndex = 0
If cboCO.ListCount = 0 Then cmdCO.Enabled = False
'' Entries
For bytTmp = 1 To 6
    cboEnt.AddItem Choose(bytTmp, "0", "1", "2", "4", "6", "8")
Next bytTmp
cboEnt.ListIndex = 2
'' Division
Call ComboFill(cboDiv, 13, 2)
If cboDiv.ListCount > 0 Then cboDiv.ListIndex = 0
If cboDiv.ListCount = 0 Then cmdDiv.Enabled = False
'' Company
Call ComboFill(cboComp, 5, 2)
If cboComp.ListCount > 0 Then cboComp.ListIndex = 0
If cboComp.ListCount = 0 Then cmdComp.Enabled = False
'' Designation
    Call ComboFill(cboDesig, 20, 2)
    
    If cboDesig.ListCount > 0 Then cboDesig.ListIndex = 0
    If cboDesig.ListCount = 0 Then cboDesig.Enabled = False
 
Exit Sub
ERR_P:
    ShowError ("FillCombos::" & Me.Caption)
End Sub

Private Sub RetCaptions()                   '' Gets and Sets the Form Captions
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '64%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("64001", adrsC)              '' Form Caption
'' Button Captions
frSet.Caption = NewCaptionTxt("64002", adrsC)
'cmdcat.Caption = NewCaptionTxt("64003", adrsC)
'cmdDept.Caption = NewCaptionTxt("64004", adrsC)
'cmdGroup.Caption = NewCaptionTxt("64005", adrsC)
'cmdLoca.Caption = NewCaptionTxt("64006", adrsC)
'cmdOT.Caption = NewCaptionTxt("64007", adrsC)
'cmdCO.Caption = NewCaptionTxt("64008", adrsC)
'cmdEnt.Caption = NewCaptionTxt("64009", adrsC)
'cmdDesig.Caption = NewCaptionTxt("64010", adrsC)
'cmdDiv.Caption = NewCaptionTxt("64012", adrsC)
'cmdExit.Caption = NewCaptionTxt("00008", adrsMod)
Call SetGridDetails(Me, frEmp, MSF3, lblEmp, lblTo)
Call CapGrid
End Sub

Private Sub CapGrid()
With MSF3
    .TextMatrix(0, 2) = NewCaptionTxt("00051", adrsMod)    '' Category
    .TextMatrix(0, 3) = NewCaptionTxt("00058", adrsMod)    '' Department
    .TextMatrix(0, 4) = NewCaptionTxt("00059", adrsMod)    '' Group
    .TextMatrix(0, 5) = NewCaptionTxt("00110", adrsMod)    '' Location
    .TextMatrix(0, 6) = NewCaptionTxt("00090", adrsMod)    '' OT Rule
    .TextMatrix(0, 7) = NewCaptionTxt("00091", adrsMod)    '' CO Rule
    .TextMatrix(0, 8) = NewCaptionTxt("00121", adrsMod)    '' Entry
    .TextMatrix(0, 9) = NewCaptionTxt("00123", adrsMod)    '' Designation
    .TextMatrix(0, 10) = "Comp."    '' Designation
    .TextMatrix(0, 11) = "Div."     '' Designation
    .ColWidth(12) = 0  ''Machinecodes                'by
    '' Resizing
    .ColWidth(3) = .ColWidth(3) * 1.1
    .ColWidth(9) = .ColWidth(9) * 1.15
    '' Aligning
    .ColAlignment(0) = flexAlignLeftTop
    .ColAlignment(1) = flexAlignLeftTop
    .ColAlignment(2) = flexAlignLeftTop
    .ColAlignment(3) = flexAlignLeftTop
    .ColAlignment(4) = flexAlignLeftTop
    .ColAlignment(5) = flexAlignLeftTop
    .ColAlignment(6) = flexAlignLeftTop
    .ColAlignment(7) = flexAlignLeftTop
    .ColAlignment(8) = flexAlignLeftTop
    .ColAlignment(9) = flexAlignLeftTop
End With
End Sub

Private Sub GetRights()
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 11, 1)
If Mid(strTmp, 2, 1) = "1" Then
    frSet.Enabled = True
Else
    MsgBox NewCaptionTxt("00001", adrsMod), vbExclamation
    frSet.Enabled = False
End If
Exit Sub
ERR_P:
    ShowError ("GetRights::" & Me.Caption)
    frSet.Enabled = False
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


Private Sub MSF3_Click()
Dim bytTmp As Byte
If MSF3.Rows = 1 Then Exit Sub
If MSF3.CellBackColor = SELECTED_COLOR Then
    With MSF3
        For bytTmp = 0 To .Cols - 1
            .Col = bytTmp
            .CellBackColor = UNSELECTED_COLOR
        Next
        .Col = 0
    End With
Else
    With MSF3
        For bytTmp = 0 To .Cols - 1
            .Col = bytTmp
            .CellBackColor = SELECTED_COLOR
        Next
        .Col = 0
    End With
End If
End Sub

Private Sub MSF3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then Call MSF3_Click
End Sub

Private Function CheckEmployees() As String
On Error GoTo ERR_P
Dim strTmp As String, intTmp As Integer
strTmp = ""
With MSF3
    .Col = 0
    For intTmp = 1 To .Rows - 1
        .row = intTmp
        If .CellBackColor = SELECTED_COLOR Then
            strTmp = strTmp & "'" & .Text & "',"
        End If
    Next
End With
If strTmp <> "" Then
    strTmp = Left(strTmp, Len(strTmp) - 1)
Else
    MsgBox NewCaptionTxt("00049", adrsMod)
End If
CheckEmployees = strTmp
Exit Function
ERR_P:
    ShowError ("CheckEmployees" & Me.Caption)
End Function

Private Sub ReflectChanges(ByVal strFld As String, ByVal strVal As String, Optional blnChar As Boolean = False)
On Error GoTo ERR_P
If blnChar Then strVal = "'" & strVal & "'"
ConMain.Execute "Update empmst Set " & strFld & "=" & strVal & " Where " & _
"Empcode in (" & strEmp & ")"
Call FillGrid
Exit Sub
ERR_P:
    ShowError ("ReflectChanges::" & Me.Caption)
End Sub


