VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmDecH 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   5610
      TabIndex        =   10
      Top             =   4440
      Width           =   1875
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3750
      TabIndex        =   9
      Top             =   4440
      Width           =   1875
   End
   Begin VB.CommandButton cmdEditCan 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1890
      TabIndex        =   8
      Top             =   4440
      Width           =   1875
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Command1"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4440
      Width           =   1905
   End
   Begin TabDlg.SSTab TB1 
      Height          =   4425
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   7805
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "List"
      TabPicture(0)   =   "frmDecH.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSF1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Details"
      TabPicture(1)   =   "frmDecH.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DetailsFrame"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame DetailsFrame 
         Height          =   4035
         Left            =   -74970
         TabIndex        =   13
         Top             =   330
         Width           =   7395
         Begin VB.CheckBox chkCat 
            Caption         =   "Add this Holiday / WeekOff for all categories"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   150
            TabIndex        =   0
            Top             =   345
            Width           =   4950
         End
         Begin VB.Frame frAs 
            Height          =   555
            Left            =   1755
            TabIndex        =   19
            Top             =   2535
            Width           =   3660
            Begin VB.OptionButton optWO 
               Caption         =   "WeekOff"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   150
               TabIndex        =   5
               Top             =   255
               Width           =   1305
            End
            Begin VB.OptionButton optHL 
               Caption         =   "Holiday"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1905
               TabIndex        =   6
               Top             =   210
               Width           =   1305
            End
         End
         Begin VB.TextBox txtDesc 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   1680
            MaxLength       =   49
            TabIndex        =   4
            Text            =   " "
            Top             =   2070
            Width           =   4215
         End
         Begin VB.TextBox txtDate 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1680
            TabIndex        =   2
            Tag             =   "D"
            Text            =   " "
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox txtComp 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5460
            TabIndex        =   3
            Tag             =   "D"
            Text            =   " "
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label lblCat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category      "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   14
            Top             =   930
            Width           =   1140
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date            "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   150
            TabIndex        =   15
            Top             =   1530
            Width           =   1125
         End
         Begin VB.Label lblDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description  :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   135
            TabIndex        =   17
            Top             =   2115
            Width           =   1155
         End
         Begin VB.Label lblAs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Declare as   :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   150
            TabIndex        =   18
            Top             =   2715
            Width           =   1170
         End
         Begin VB.Label lblComp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compensate Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3360
            TabIndex        =   16
            Top             =   1530
            Width           =   1620
         End
         Begin MSForms.ComboBox cboCat 
            Height          =   375
            Left            =   1680
            TabIndex        =   1
            Top             =   840
            Width           =   1575
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "2778;661"
            MatchEntry      =   0
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   4095
         Left            =   30
         TabIndex        =   12
         Top             =   330
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   7223
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   12632256
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmDecH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset

Private Sub cboCat_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then Chr (9)
End Sub

Private Sub chkCat_Click()
If chkCat.Value = 1 Then
    cboCat.Value = ""
    cboCat.Enabled = False
Else
    cboCat.Enabled = True
End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)    '' Set the Form Icon
Call SetToolTipText(Me) '' Set the ToolTipText
Call RetCaptions        '' Get the Appropriate Captions
Call OpenMasterTable    '' Open Master Table
Call FillGrid           '' Fill Grid
Call FillComboCat       '' Fill Category Combo
TB1.Tab = 0             '' Set the Tab to List
Call GetRights          '' Gets Rights for the Operations
bytMode = 1
Call ChangeMode         '' Take Action on the Appropriate Mode
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '19%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("19001", adrsC)              '' Form caption
TB1.TabCaption(0) = NewCaptionTxt("00013", adrsMod)       '' List
TB1.TabCaption(1) = NewCaptionTxt("00014", adrsMod)       '' Details
lblAs.Caption = NewCaptionTxt("19006", adrsC)           '' Declare As

    chkCat.Caption = NewCaptionTxt("19004", adrsC)          '' Category Check Box
    lblCat.Caption = NewCaptionTxt("00051", adrsMod)          '' Category

lblDate.Caption = NewCaptionTxt("00030", adrsMod)         '' Date
lblComp.Caption = NewCaptionTxt("19005", adrsC)         '' Compensate Date
lblDesc.Caption = NewCaptionTxt("00052", adrsMod)         '' Description
optWO.Caption = NewCaptionTxt("19007", adrsC)           '' Week Off
optHL.Caption = NewCaptionTxt("19008", adrsC)           '' Holiday
Call SetGButtonCap(Me)                           '' Sets Appropriate Caption for the Buttons
Call CapGrid                                '' Captions For Grid
End Sub

Private Sub CapGrid()           '' Gives the Captions to the Grid
With MSF1
    '' Sets the Column Widhts
    .ColWidth(0) = .ColWidth(0) * 0.95
    .ColWidth(1) = .ColWidth(1) * 1.25
    .ColWidth(2) = .ColWidth(2) * 2.97
    .ColWidth(3) = .ColWidth(3) * 1.35
    .ColWidth(4) = .ColWidth(4) * 0.4
    '' Sets the Column Alignment
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    .ColAlignment(4) = flexAlignLeftCenter
    '' Sets the Appropriate Captions

        .TextMatrix(0, 0) = NewCaptionTxt("00051", adrsMod)   '' Category

    .TextMatrix(0, 1) = NewCaptionTxt("00030", adrsMod)   '' Date
    .TextMatrix(0, 2) = NewCaptionTxt("00052", adrsMod)   '' Description
    .TextMatrix(0, 3) = NewCaptionTxt("19002", adrsC)   '' Compensate On
    .TextMatrix(0, 4) = NewCaptionTxt("19003", adrsC)   '' As
End With
End Sub

Private Sub FillComboCat()      '' Fills Category Combo
On Error GoTo ERR_P
cboCat.clear
If AdrsCat.State = 1 Then AdrsCat.Close

    AdrsCat.Open "Select " & strKDesc & "  from catdesc where cat <> '100'", ConMain

If Not (AdrsCat.BOF And AdrsCat.EOF) Then
    Do While Not AdrsCat.EOF
        cboCat.AddItem AdrsCat(0)
        AdrsCat.MoveNext
    Loop
End If
AdrsCat.Close
Exit Sub
ERR_P:
    ShowError ("FillComboCat :: " & Me.Caption)
End Sub

Private Sub GetRights()             '' Gets and Sets the Appropriate Rights
On Error GoTo ERR_P
Dim strTmp As String
strTmp = RetRights(1, 13)
If Mid(strTmp, 1, 1) = "1" Then AddRights = True
If Mid(strTmp, 2, 1) = "1" Then EditRights = True
If Mid(strTmp, 3, 1) = "1" Then DeleteRights = True
Exit Sub
ERR_P:
    ShowError ("GetRights :: " & Me.Caption)
    AddRights = False
    EditRights = False
    DeleteRights = False
End Sub

Private Sub ChangeMode()
Select Case bytMode
    Case 1  '' View
        Call ViewAction
    Case 2  '' Add
        Call AddAction
    Case 3  '' Modify
        Call EditAction
End Select
End Sub

Private Sub ViewAction()    '' Procedure for Viewing Mode
'' Enable Needed Controls
cmdAddSave.Enabled = True   '' Enable ADD/SAVE Button
cmdEditCan.Enabled = True   '' Enable Edit/Cancel Button
cmdDel.Enabled = True       '' Enable Disable Button
'' Disable Needed Controls
chkCat.Enabled = False      '' Disable Category CheckBox
chkCat.Value = 0            '' Set Value to 0 in View Mode
cboCat.Enabled = False      '' Disable Category ComboBox
txtDate.Enabled = False     '' Disable Date TextBox
txtComp.Enabled = False     '' Disable Compensate Date TextBox
txtDesc.Enabled = False     '' Disable Description TextBox
frAs.Enabled = False        '' Disable As Frame
'' Give Captions to the Needed Controls
Call SetGButtonCap(Me)
TB1.Tab = 0     '' Set the Tab to the First Tab
End Sub

Private Sub AddAction()     '' Procedure for Addition Mode
If TB1.TabEnabled(1) = False Then TB1.TabEnabled(1) = True
TB1.Tab = 1
'' Enable Necessary Controls
chkCat.Enabled = True       '' Enable Category CheckBox
cboCat.Enabled = True       '' Enable Category ComboBox
txtDate.Enabled = True      '' Enable Date TextBox
txtComp.Enabled = True      '' Enable Compensate Date TextBox
txtDesc.Enabled = True      '' Enable Description TextBox
frAs.Enabled = True         '' Enable As Frame
'' Disable Necessary Controls
cmdDel.Enabled = False     '' Disable Delete Button
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Clear Necessary Controls
chkCat.Value = 0            '' Clear Category CheckBox
cboCat.Value = ""           '' Clear the Category ComboBox
txtDate.Text = ""           '' Clear Date TextBox
txtComp.Text = ""           '' Clear Compensate Date TextBox
txtDesc.Text = ""           '' Clear Description TextBox
optHL.Value = True          '' Keep Holiday as Default
chkCat.SetFocus             '' Set Focus on the Category CheckBox
End Sub

Private Sub EditAction()    '' Procedure for Edit Mode
'' Enable Necessary Controls
txtDesc.Enabled = True  '' Enable Description TextBox
'' Give Caption to the Needed Controls
Call SetGButtonCap(Me, 2)
'' Disable Necessary Controls
cmdDel.Enabled = False          '' Disable Delete Button
txtDesc.SetFocus                '' Set Focus on the Description TextBox
If TB1.Tab = 1 Then
    Exit Sub
Else
    If TB1.TabEnabled(1) Then TB1.Tab = 1
End If
End Sub

Private Sub OpenMasterTable()             '' Open the recordset for the Display purposes
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close
Dim strHolFrom As String, strHolTo As String
strHolFrom = CStr(CDate("01-" & Left(MonthName(pVStar.Yearstart), 3) & "-" & pVStar.YearSel))
strHolTo = CDate(strHolFrom) + IIf(Year(CDate(strHolFrom)) Mod 4, 364, 365)
adrsDept1.Open "Select Cat," & strKDate & " ," & strKDesc & " ,Hcode,Compensdt,Declas from DeclWohl where (" & strKDate & "  between " & _
strDTEnc & DateCompStr(strHolFrom) & strDTEnc & " and " & strDTEnc & DateCompStr(strHolTo) & _
strDTEnc & ") Order by Cat," & strKDate & " ", ConMain, adOpenStatic
Exit Sub
ERR_P:
    ShowError ("OpenMasterTable :: " & Me.Caption)
End Sub

Private Sub Display()       '' Displays the Given Master Records
On Error GoTo ERR_P
cboCat.Value = GetDesc(MSF1.TextMatrix(MSF1.row, 0))     '' Category Code
If cboCat.Value = "" Then

        MsgBox NewCaptionTxt("19009", adrsC), vbCritical
  
    bytMode = 1
    Call ChangeMode
    Exit Sub
End If
txtDate = DateDisp(MSF1.TextMatrix(MSF1.row, 1))        '' Date
txtDesc = MSF1.TextMatrix(MSF1.row, 2)                  '' Description
txtComp.Text = MSF1.TextMatrix(MSF1.row, 3)             '' Compensate Date
Select Case MSF1.TextMatrix(MSF1.row, 4)                '' Declare As
    Case pVStar.WosCode: optWO.Value = True     '' Week Off
    Case pVStar.HlsCode: optHL.Value = True     '' Holiday
    Case Else: optHL.Value = True               '' Holiday
End Select
Exit Sub
ERR_P:
    ShowError ("Display  :: " & Me.Caption)
End Sub

Private Function GetDesc(ByVal strCatName As String) As String
On Error GoTo ERR_P         '' Gets the Description for a Particular Code"
If adrsTemp.State = 1 Then adrsTemp.Close

    adrsTemp.Open "Select " & strKDesc & "  from Catdesc where Cat='" & strCatName & "'" _
    , ConMain

If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    GetDesc = adrsTemp.Fields(0)  'GetDesc = adrsTemp("Desc")
Else
    GetDesc = ""
End If
Exit Function
ERR_P:
    ShowError ("GetDesc  :: " & Me.Caption)
End Function

Private Sub FillGrid()          '' Fills the Grid
On Error GoTo ERR_P
Dim intCounter As Integer
adrsDept1.Requery               '' Requeries the Recordset for any Updated Values
'' Put Appropriate Rows in the Grid
If adrsDept1.EOF Then
    MSF1.Rows = 1
    TB1.TabEnabled(1) = False       '' Disables Tab 1 If No Records are Found
    Exit Sub
End If
MSF1.Rows = adrsDept1.RecordCount + 1   '' Sets Rows Appropriately
adrsDept1.MoveFirst
For intCounter = 1 To adrsDept1.RecordCount     '' Fills the Grid
    With MSF1
        .TextMatrix(intCounter, 0) = adrsDept1("Cat")                   '' Category
        .TextMatrix(intCounter, 1) = DateDisp(adrsDept1("Date"))        '' Date
        '' Description
        .TextMatrix(intCounter, 2) = IIf(IsNull(adrsDept1("Desc")), "", adrsDept1("Desc"))
        .TextMatrix(intCounter, 3) = DateDisp(adrsDept1("Compensdt"))   '' Compensate Date
        '' Declare As
        .TextMatrix(intCounter, 4) = IIf(IsNull(adrsDept1("DeclAs")), "", adrsDept1("DeclAs"))
    End With
    adrsDept1.MoveNext
Next
TB1.TabEnabled(1) = True        '' If Records Found Enables the TAB 1
TB1.Tab = 0
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Function ValidateAddmaster() As Boolean
On Error GoTo ERR_P
ValidateAddmaster = True
If chkCat.Value = False And cboCat.Text = "" Then

        MsgBox NewCaptionTxt("19010", adrsC), vbExclamation

    cboCat.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If chkCat.Value = 1 And cboCat.ListCount = 0 Then

        MsgBox NewCaptionTxt("19011", adrsC), vbExclamation

    chkCat.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If Trim(txtDate.Text) = "" Then
    MsgBox NewCaptionTxt("19012", adrsC), vbExclamation
    txtDate.SetFocus
    ValidateAddmaster = False
    Exit Function
Else
    If DateDiff("m", Year_Start, CDate(txtDate.Text)) > 11 Or DateDiff("m", Year_Start, CDate(txtDate.Text)) < 0 Then
        MsgBox NewCaptionTxt("19013", adrsC) & txtDate.Text & NewCaptionTxt("00021", adrsMod), vbExclamation
        txtDate.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If
If Trim(txtComp.Text) = "" Then
    MsgBox NewCaptionTxt("19012", adrsC), vbExclamation
    txtComp.SetFocus
    ValidateAddmaster = False
    Exit Function
Else
    If DateDiff("m", Year_Start, CDate(txtComp.Text)) > 11 Or DateDiff("m", Year_Start, CDate(txtComp.Text)) < 0 Then
        MsgBox NewCaptionTxt("19014", adrsC) & txtComp.Text & NewCaptionTxt("00021", adrsMod), vbExclamation
        txtComp.SetFocus
        ValidateAddmaster = False
        Exit Function
    End If
End If
If Trim(txtDate.Text) = Trim(txtComp.Text) Then
    MsgBox NewCaptionTxt("19015", adrsC), vbExclamation
    txtComp.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
If Trim(txtDesc.Text) = "" Then
    MsgBox NewCaptionTxt("19016", adrsC), vbExclamation
    txtDesc.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
'' if Week Off And Holiday Fall On Same Date as Any Week Off Then
'PramacyForDecHoliday this flag add by  for Pramacy
    Select Case CheckWOHLOnSameday
        Case -1
            MsgBox NewCaptionTxt("19019", adrsC), vbExclamation
            txtDate.SetFocus
            ValidateAddmaster = False
            Exit Function
        Case 0
            '' Do Nothing
        Case Is > 0
            MsgBox NewCaptionTxt("19017", adrsC), vbExclamation
            txtDate.SetFocus
            ValidateAddmaster = False
            Exit Function
    End Select

If CheckDuplicates > 0 Then
    MsgBox NewCaptionTxt("19018", adrsC), vbExclamation
    txtDate.SetFocus
    ValidateAddmaster = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster :: " & Me.Caption)
    ValidateAddmaster = False
End Function

Private Function CheckWOHLOnSameday()   '' Cheks if the Week Off or Hoilday Fall on
On Error GoTo ERR_P                     '' the Same Date as Week Off OR Not
Dim bytDH As Byte
bytDH = 0
If Not GetFlagStatus("LOCATIONWISEHL") Then Exit Function
If chkCat.Value = 0 Then
    If adrsTemp.State = 1 Then adrsTemp.Close

        adrsTemp.Open " select empmst.cat," & strKOff & " ,wo_1_3,wo_2_4," & strKDesc & "  from empmst,catdesc " & _
        "where catdesc.cat= empmst.cat and catdesc." & strKDesc & "  = '" & cboCat.Text & "'", ConMain, adOpenStatic
  
    If Not (adrsTemp.EOF And adrsTemp.BOF) Then
        adrsTemp.MoveFirst
        For i = 0 To adrsTemp.RecordCount - 1

                If UCase(cboCat.Text) = UCase(adrsTemp("desc")) And _
                (Left(WeekdayName(WeekDay(CDate(txtDate.Text), vbUseSystemDayOfWeek)), 2) = adrsTemp("off") Or _
                Left(WeekdayName(WeekDay(CDate(txtDate.Text), vbUseSystemDayOfWeek)), 2) = adrsTemp!wo_1_3 Or _
                Left(WeekdayName(WeekDay(CDate(txtDate.Text), vbUseSystemDayOfWeek)), 2) = adrsTemp!wo_2_4) Then
                    bytDH = bytDH + 1
                End If
     
            adrsTemp.MoveNext
        Next
    Else
        CheckWOHLOnSameday = -1
        Exit Function
    End If
Else
    If adrsRits.State = 1 Then adrsRits.Close

        adrsRits.Open "Select distinct(cat) from catdesc where cat <> '100'", ConMain
  
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open " select cat," & strKOff & " ,wo_1_3,wo_2_4 from empmst", ConMain, _
    adOpenStatic
    If Not adrsTemp.EOF Then
        Do While Not adrsRits.EOF
            adrsTemp.MoveFirst
            For i = 0 To adrsTemp.RecordCount - 1
                If adrsRits(0) = adrsTemp!cat And _
                (Left(WeekdayName(WeekDay(CDate(txtDate.Text), vbUseSystemDayOfWeek)), 2) = adrsTemp("off") Or _
                Left(WeekdayName(WeekDay(CDate(txtDate.Text), vbUseSystemDayOfWeek)), 2) = adrsTemp!wo_1_3 Or _
                Left(WeekdayName(WeekDay(CDate(txtDate.Text), vbUseSystemDayOfWeek)), 2) = adrsTemp!wo_2_4) Then
                    bytDH = bytDH + 1
                End If
                adrsTemp.MoveNext
            Next
            adrsRits.MoveNext
        Loop
    Else
        CheckWOHLOnSameday = -1
        Exit Function
    End If
End If
CheckWOHLOnSameday = bytDH
Exit Function
ERR_P:
    ShowError ("CheckWOHLOnSameday :: " & Me.Caption)
    CheckWOHLOnSameday = 1
    'Resume Next
End Function

Private Function CheckDuplicates() As Byte
On Error GoTo ERR_P
Dim bytDH As Byte
bytDH = 0
If chkCat.Value = 1 Then
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "select * from holiday where " & strKDate & " =" & strDTEnc & _
    DateCompStr(txtDate.Text) & strDTEnc, _
    ConMain, adOpenKeyset, adLockOptimistic
    If Not (adrsTemp.BOF And adrsTemp.EOF) Then bytDH = bytDH + 1
    
    adrsTemp.Close
    adrsTemp.Open "select * from Declwohl where " & strKDate & " =" & _
    strDTEnc & DateCompStr(txtDate.Text) & strDTEnc & " or " & strKDate & " =" & _
    strDTEnc & DateCompStr(txtComp.Text) & strDTEnc & _
    " or compensdt=" & strDTEnc & DateCompStr(txtDate.Text) & _
    strDTEnc & " or compensdt=" & strDTEnc & DateCompStr(txtComp.Text) & strDTEnc, _
    ConMain, adOpenKeyset, adLockOptimistic
    If Not (adrsTemp.BOF And adrsTemp.EOF) Then bytDH = bytDH + 1
    
Else
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "select * from Holiday where cat=" & _
    "'" & CatCode(cboCat.Text) & "'" & " and " & strKDate & " =" & strDTEnc & _
    DateCompDate(txtDate.Text) & strDTEnc, _
    ConMain, adOpenKeyset, adLockOptimistic
    If Not (adrsTemp.BOF And adrsTemp.EOF) Then bytDH = bytDH + 1
    adrsTemp.Close
    adrsTemp.Open "select * from declwohl where (cat=" & _
    "'" & CatCode(cboCat.Text) & "'" & " and " & strKDate & " =" & _
    strDTEnc & DateCompDate(txtDate.Text) & strDTEnc & ")" & _
    " or ( cat=" & "'" & CatCode(cboCat.Text) & "'" & " and compensdt=" & _
    strDTEnc & DateCompDate(txtDate.Text) & strDTEnc & ")" & _
    " or ( cat=" & "'" & CatCode(cboCat.Text) & "'" & " and " & strKDate & " =" & _
    strDTEnc & DateCompDate(txtComp.Text) & strDTEnc & ")" & _
    " or ( cat=" & "'" & CatCode(cboCat.Text) & "'" & " and compensdt=" & _
    strDTEnc & DateCompDate(txtComp.Text) & strDTEnc & ")" _
    , ConMain, adOpenKeyset, adLockOptimistic
    If Not (adrsTemp.BOF And adrsTemp.EOF) Then bytDH = bytDH + 1
End If
CheckDuplicates = bytDH
Exit Function
ERR_P:
    ShowError ("CheckDuplicates :: " & Me.Caption)
    CheckDuplicates = 1
End Function

Private Function ValidateModMaster() As Boolean
On Error GoTo ERR_P
ValidateModMaster = True
If Trim(txtDesc.Text) = "" Then
    MsgBox NewCaptionTxt("19016", adrsC), vbExclamation
    txtDesc.SetFocus
    ValidateModMaster = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateModMaster :: " & Me.Caption)
    ValidateModMaster = False
End Function

Private Sub txtComp_GotFocus()
    Call GF(txtComp)
End Sub

Private Sub txtComp_Validate(Cancel As Boolean)
If Not ValidDate(txtComp) Then
    txtComp.SetFocus
    Cancel = True
End If
End Sub

Private Sub txtDate_Click()
varCalDt = ""
varCalDt = Trim(txtDate.Text)
txtDate.Text = ""
Call ShowCalendar
End Sub

Private Sub txtDate_GotFocus()
    Call GF(txtDate)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    Call CDK(txtDate, KeyAscii)
End Sub

Private Sub txtComp_Click()
varCalDt = ""
varCalDt = Trim(txtComp.Text)
txtComp.Text = ""
Call ShowCalendar
End Sub

Private Sub txtComp_KeyPress(KeyAscii As Integer)
    Call CDK(txtComp, KeyAscii)
End Sub

Private Function SaveAddMaster() As Boolean
On Error GoTo ERR_P
Dim bytCat As Byte
Dim strAs As String
If optWO.Value = True Then
    strAs = pVStar.WosCode
Else
    strAs = pVStar.HlsCode
End If
SaveAddMaster = True        '' Insert
If chkCat.Value = 0 Then
    ConMain.Execute "insert into DeclWohl(Cat," & strKDate & " ," & strKDesc & " ,HCode," & _
    "Compensdt,Declas) Values('" & _
    CatCode(cboCat.Text) & "'," & strDTEnc & DateSaveIns(txtDate.Text) & strDTEnc & ",'" & _
    txtDesc.Text & "','" & pVStar.HlsCode & "'," & strDTEnc & DateSaveIns(txtComp.Text) & _
    strDTEnc & ",'" & strAs & "')"
    Call SetShift(txtDate.Text, CatCode(cboCat.Text), txtComp.Text)
Else
    For bytCat = 0 To cboCat.ListCount - 1
        ConMain.Execute "insert into DeclWohl(Cat," & strKDate & " ," & strKDesc & " ,HCode," & _
        "Compensdt,Declas) Values('" & _
        CatCode(cboCat.List(bytCat)) & "'," & strDTEnc & DateSaveIns(txtDate.Text) & strDTEnc & ",'" & _
        txtDesc.Text & "','" & pVStar.HlsCode & "'," & strDTEnc & DateSaveIns(txtComp.Text) & _
        strDTEnc & ",'" & strAs & "')"
        Call SetShift(txtDate.Text, CatCode(cboCat.List(bytCat)), txtComp.Text)
    Next
End If
Exit Function
ERR_P:
    SaveAddMaster = False
    ShowError ("SaveAddMaster :: " & Me.Caption)
End Function

Private Function SaveModMaster() As Boolean
On Error GoTo ERR_P
SaveModMaster = True        '' Update
ConMain.Execute "Update DeclWohl set " & strKDesc & " =" & "'" & Trim(txtDesc.Text) & _
"'" & " where Cat=" & "'" & CatCode(cboCat.Text) & "'" & " and " & strKDate & " =" & _
strDTEnc & DateSaveIns(txtDate.Text) & strDTEnc & " and Compensdt=" & strDTEnc & _
DateCompStr(txtComp.Text) & strDTEnc
Exit Function
ERR_P:
    SaveModMaster = False
    ShowError ("SaveModMaster :: " & Me.Caption)
End Function

Private Sub cmdDel_Click()
On Error GoTo ERR_P
'' Check for Rights
If Not DeleteRights Then
    MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
    Exit Sub
Else
    If TB1.TabEnabled(1) = False Then Exit Sub
    If TB1.Tab = 0 Then                         '' Do not Display Record if
        If TB1.TabEnabled(1) Then TB1.Tab = 1   '' Already Displayed
    End If
    If MsgBox(NewCaptionTxt("00015", adrsMod), vbYesNo + vbQuestion) _
    = vbYes Then        '' Delete the Record
        ConMain.Execute _
        "delete from Declwohl where cat='" & CatCode(cboCat.Text) & _
         "' and " & strKDesc & " ='" & txtDesc.Text & "'"
         Call AddActivityLog(lgDelete_Action, 1, 13)    '' Delete Log
         Call AuditInfo("DELETE", Me.Caption, "Deleted Declared Holiday For Category " & cboCat.Text & "as on " & txtDate.Text)
    End If
    Call FillGrid       '' Reflect the Grid
    bytMode = 1
    Call ChangeMode
End If
Exit Sub
ERR_P:
    ShowError ("Delete Record :: " & Me.Caption)
End Sub

Private Sub cmdAddSave_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        '' Check for Rights
        If Not AddRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        bytMode = 2
        Call ChangeMode
    Case 2          '' Add Mode
        If Not ValidateAddmaster Then Exit Sub  '' Validate For Add
        If Not SaveAddMaster Then Exit Sub      '' Save for Add
        Call SaveAddLog                         '' Save the Add Log
        Call FillGrid                           '' Reflect the Grid
        bytMode = 1
        Call ChangeMode
    Case 3          '' Edit Mode
        If Not ValidateModMaster Then Exit Sub  '' Validate for Edit
        If Not SaveModMaster Then Exit Sub      '' Save for Edit
        Call SaveModLog                         '' Save the Edit Log
        Call FillGrid                           '' Reflect the Grid
        bytMode = 1
        Call ChangeMode
End Select
Exit Sub
ERR_P:
    ShowError ("AddSave :: " & Me.Caption)
End Sub

Private Sub cmdEditCan_Click()
On Error GoTo ERR_P
Select Case bytMode
    Case 1          '' View Mode
        If TB1.TabEnabled(1) = False Then Exit Sub
        '' Check for Rights
        If Not EditRights Then
            MsgBox NewCaptionTxt("00001", adrsMod), vbInformation
            Exit Sub
        End If
        bytMode = 3
        Call ChangeMode
    Case 2       '' Add Mode
        If MSF1.Rows = 1 Then
            TB1.TabEnabled(1) = False
            TB1.Tab = 0
        End If
        bytMode = 1
        Call ChangeMode
    Case 3      '' Edit Mode
        bytMode = 1
        Call ChangeMode
End Select
Exit Sub
ERR_P:
    ShowError ("EditCance :: " & Me.Caption)
End Sub

Private Sub TB1_Click(PreviousTab As Integer)
If bytMode = 2 Then Exit Sub
If PreviousTab = 1 Then Exit Sub
MSF1.Col = 0

    If MSF1.Text = NewCaptionTxt("00051", adrsMod) Then Exit Sub

Call Display
End Sub

Private Sub MSF1_DblClick()
If TB1.TabEnabled(1) = False Then Exit Sub
TB1.Tab = 1
End Sub

Private Sub SetShift(ByVal strDateText As String, ByVal strCatText As String, _
            ByVal strCompDate As String)
On Error GoTo ERR_P
Dim adrsEmpCnt As New ADODB.Recordset
Dim strMonShf As String, strTempShfFile As String, strTempShift As String, Ns As String
''
strMonShf = Left(MonthName(Month(DateCompDate(strDateText))), 3)
strMonShf = strMonShf & Right(Year(DateCompDate(strDateText)), 2) & "shf"
''
strTempShfFile = Left(MonthName(Month(DateCompDate(strCompDate))), 3)
strTempShfFile = strTempShfFile & Right(Year(DateCompDate(strCompDate)), 2) & "shf"
''
If FindTable(strMonShf) Then
    If adrsEmpCnt.State = 1 Then adrsEmpCnt.Close

        adrsEmpCnt.Open "Select Empcode from empmst where cat='" & strCatText & "'" _
        , ConMain
  
    Do While Not adrsEmpCnt.EOF
        Ns = Day(DateCompDate(strDateText))
        If adrsRits.State = 1 Then adrsRits.Close   ' 12-08 can't compensate if both date having shift
        adrsRits.Open "Select d" & Trim(Day(DateCompDate(strDateText))) & ",d" & Trim(Day(DateCompDate(strCompDate))) & " from " & strMonShf & _
        " where Empcode=" & "'" & adrsEmpCnt(0) & "'", ConMain
        If Not (adrsRits.EOF And adrsRits.BOF) Then
            If ((adrsRits(0) <> pVStar.HlsCode And adrsRits(0) <> pVStar.WosCode) And (adrsRits(1) <> pVStar.HlsCode And adrsRits(1) <> pVStar.WosCode)) Or (adrsRits(0) = pVStar.WosCode Or adrsRits(0) = pVStar.HlsCode) Then
                adrsEmpCnt.MoveNext
            Else
                If adrsRits.State = 1 Then adrsRits.Close
                adrsRits.Open "Select d" & Ns & " from " & strMonShf & _
                " where Empcode=" & "'" & adrsEmpCnt(0) & "'", ConMain
                If Not (adrsRits.EOF And adrsRits.BOF) Then
                    strTempShift = adrsRits(0)
                    If optWO.Value = True Then
                        ConMain.Execute "Update " & strMonShf & " set d" & Ns & _
                        "='" & pVStar.WosCode & "' where Empcode='" & adrsEmpCnt(0) & "'"
                    End If
                    If optHL.Value = True Then
                        ConMain.Execute "Update " & strMonShf & " set d" & Ns & _
                        "='" & pVStar.HlsCode & "' where Empcode='" & adrsEmpCnt(0) & "'"
                    End If
                    If FindTable(strTempShfFile) Then
                        Ns = "D" & Trim(Day(DateCompDate(strCompDate)))
                        ConMain.Execute "update " & strTempShfFile & " set " & _
                        Ns & "=" & "'" & strTempShift & "'" & _
                        " where Empcode=" & "'" & adrsEmpCnt(0) & "'"
                    End If
                End If
                adrsEmpCnt.MoveNext
            End If
        Else
            adrsEmpCnt.MoveNext
        End If
    Loop
    adrsEmpCnt.Close
End If
Exit Sub
ERR_P:
    ShowError ("RetShift :: " & Me.Caption)
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
    If Not ValidDate(txtDate) Then txtDate.SetFocus: Cancel = True
End Sub

Private Sub txtDesc_GotFocus()
    Call GF(txtDesc)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = Asc(UCase(Chr(KeyPressCheck(KeyAscii, 3))))
End If
End Sub

Private Sub SaveAddLog()            '' Procedure to Save the Add Log
On Error GoTo ERR_P
Call AddActivityLog(lgADD_MODE, 1, 13)     '' Add Activity
Call AuditInfo("ADD", Me.Caption, "Add Declared Holiday For Category " & cboCat.Text & "as on " & txtDate.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub

Private Sub SaveModLog()            '' Procedure to Save the Edit Log
On Error GoTo ERR_P
Call AddActivityLog(lgEdit_Mode, 1, 13)     '' Edit Activity
Call AuditInfo("UPDATE", Me.Caption, "Edit Declared Holiday For Category " & cboCat.Text & "as on " & txtDate.Text)
Exit Sub
ERR_P:
    ShowError ("Log Error :: " & Me.Caption)
End Sub
