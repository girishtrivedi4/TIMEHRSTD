VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmODAvail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Post OD"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
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
   ScaleHeight     =   8265
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Updation"
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   6360
      Width           =   9615
      Begin VB.ListBox lstError 
         Height          =   1500
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   8055
      End
      Begin LVbuttons.LaVolpeButton cmdUpdate 
         Height          =   450
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   794
         BTYPE           =   3
         TX              =   "&Update"
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
         MICON           =   "frmODAviail.frx":0000
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
      Begin LVbuttons.LaVolpeButton cmdExit 
         Cancel          =   -1  'True
         Height          =   450
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   794
         BTYPE           =   3
         TX              =   "&Exit"
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
         MICON           =   "frmODAviail.frx":001C
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
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.Frame Frame2 
         Caption         =   "&Selection Criteria"
         Height          =   6015
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9375
         Begin VB.Frame frmVisible 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   465
            Left            =   3000
            TabIndex        =   11
            Top             =   2160
            Visible         =   0   'False
            Width           =   2760
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Caption         =   "Please Wait...."
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   0
               TabIndex        =   12
               Top             =   120
               Width           =   2745
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Time Setting"
            Height          =   1575
            Left            =   120
            TabIndex        =   7
            Top             =   4320
            Width           =   9135
            Begin VB.Frame Frame6 
               Height          =   855
               Left            =   5400
               TabIndex        =   18
               Top             =   240
               Width           =   3615
               Begin VB.ComboBox cmbStatus 
                  Height          =   360
                  Left            =   2040
                  TabIndex        =   21
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.Label Label5 
                  Caption         =   "For Non Selected Replace OD With"
                  Height          =   495
                  Left            =   120
                  TabIndex        =   20
                  Top             =   240
                  Width           =   1815
               End
            End
            Begin VB.Frame Frame5 
               Height          =   1215
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   5175
               Begin VB.CheckBox chkStatusUpdate 
                  Caption         =   "S&tatus Update"
                  Height          =   375
                  Left            =   3000
                  TabIndex        =   22
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.ComboBox cmbODTo 
                  Height          =   360
                  Left            =   1200
                  TabIndex        =   16
                  Top             =   600
                  Width           =   1695
               End
               Begin VB.ComboBox cmbODFrom 
                  Height          =   360
                  Left            =   1200
                  TabIndex        =   14
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.Label Label4 
                  Caption         =   "OD To"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   17
                  Top             =   600
                  Width           =   855
               End
               Begin VB.Label Label3 
                  Caption         =   "OD From"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   15
                  Top             =   240
                  Width           =   975
               End
            End
            Begin LVbuttons.LaVolpeButton cmdSubmit 
               Height          =   330
               Left            =   8040
               TabIndex        =   19
               Top             =   1200
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   582
               BTYPE           =   3
               TX              =   "&Submit"
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
               MICON           =   "frmODAviail.frx":0038
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
         End
         Begin VB.TextBox txtDate 
            Height          =   375
            Left            =   4680
            TabIndex        =   4
            Top             =   360
            Width           =   2535
         End
         Begin MSFlexGridLib.MSFlexGrid Flex 
            Height          =   3495
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   6165
            _Version        =   393216
            AllowUserResizing=   1
         End
         Begin MSForms.ComboBox cmbDept 
            Height          =   345
            Left            =   1320
            TabIndex        =   6
            Top             =   360
            Width           =   2595
            VariousPropertyBits=   612390939
            DisplayStyle    =   3
            Size            =   "4577;609"
            TextColumn      =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            Caption         =   "D&ate"
            Height          =   255
            Left            =   4080
            TabIndex        =   3
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "&Department"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   1455
         End
      End
   End
   Begin VB.Menu mnuSelection 
      Caption         =   "selection"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectAll 
         Caption         =   "SelectAll"
      End
      Begin VB.Menu mnuDeSelectAll 
         Caption         =   "DeselectAll"
      End
      Begin VB.Menu mnuSelectLate 
         Caption         =   "SelectLate"
      End
      Begin VB.Menu mnuSelectEarly 
         Caption         =   "SelectEarly"
      End
   End
End
Attribute VB_Name = "frmODAvail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmODAvail
' DateTime  : 11/06/2008 10:33
' Author    : jagdish
' Purpose   : For Client RoboSoft
'---------------------------------------------------------------------------------------


Option Explicit
Const strShiftStart = "SHIFT START"
Const strDept = "DEPARTUTE TIME"
Const strShitEnd = "SHIFT END"
Const strArrival = "ARRIVAL TIME"
Const indexOfArrTime As Integer = 3
Const indexOfDepTime As Integer = 6
Const indexOfOdFrom As Integer = 9
Const indexOfOfTo As Integer = 10
Const indexOfLate As Integer = 4
Const indexOfErly As Integer = 7
Const indexOfPresabs As Integer = 11
Const indexOfEmpcode As Integer = 1
Const strODCode As String = "OD"

Private Enum Selection
    SelectAll = 1
    DeselectAll = 2
    SelectLate = 3
    SelectEarly = 4
End Enum
Private Enum Side
    Left_Side = 1
    Right_Side = 2
    Full = 3
End Enum
Private Type ODInOut
    Od_From As Single
    Od_To As Single
End Type
Private Type ArriDeptTime
    Arr_Time As Single
    Dept_Time As Single
End Type
Dim mSelection As Selection
Dim adrsC As Recordset
Dim strShift() As String

Private Sub cmbDept_Change()
    cmbDept_Click
End Sub


Private Sub cmbDept_Click()
   On Error GoTo cmbDept_Click_Error

    frmVisible.Visible = True
    frmVisible.Refresh
    Call FillGridWithArray
    frmVisible.Visible = False

   On Error GoTo 0
   Exit Sub

cmbDept_Click_Error:

    ShowError "Error in procedure cmbDept_Click of Form frmODAvail"
End Sub

Private Function GetShift() As String()
    Dim adrsShift As Recordset
    Dim strShift() As String
    Dim intRecord As Integer
    Dim walker As Integer
   On Error GoTo GetShift_Error

    Set adrsShift = OpenRecordSet("SELECT shift,shf_in,shf_out FROM instshft WHERE shift<>'100'")
    If Not (adrsShift.EOF And adrsShift.BOF) Then
        intRecord = adrsShift.RecordCount - 1
        ReDim strShift(intRecord, 2)
        For walker = 0 To intRecord
            strShift(walker, 0) = FilterNull(adrsShift.Fields("shift"))
            strShift(walker, 1) = FilterNull(adrsShift.Fields("shf_in"))
            strShift(walker, 2) = FilterNull(adrsShift.Fields("shf_out"))
            adrsShift.MoveNext
        Next
    End If
    GetShift = strShift
    Erase strShift

   On Error GoTo 0
   Exit Function

GetShift_Error:

    ShowError "Error in procedure GetShift of Form frmODAvail"
End Function

Private Function GetShiftCode(strEmpCode As String, strDate As String) As String
    Dim strTableName As String
    Dim adrsSh As Recordset
   On Error GoTo GetShiftCode_Error

    strTableName = MakeName(MonthName(Month(DateCompDate(strDate))), _
        Year(DateCompDate(strDate)), "Shf")
    Set adrsSh = OpenRecordSet("SELECT D" & Day(DateCompDate(strDate)) & _
        " FROM " & strTableName & " WHERE empcode='" & strEmpCode & "'")
    If Not (adrsSh.EOF And adrsSh.BOF) Then
        If FilterNull(adrsSh.Fields(0)) = EmptyString Then
            GetShiftCode = "ShiftNotFound"
        Else
            GetShiftCode = adrsSh.Fields(0)
        End If
    End If

   On Error GoTo 0
   Exit Function

GetShiftCode_Error:

    ShowError "Error in procedure GetShiftCode of Form frmODAvail"
End Function

Private Function GetShiftTime(strShiftCode As String) As ODInOut
    Dim row As Integer
    Dim column As Integer
   On Error GoTo GetShiftTime_Error
    If strShiftCode = "ShiftNotFound" Then
        'do your logic
        Exit Function
    End If
    For row = 0 To UBound(strShift) - 1
        If strShift(row, 0) = strShiftCode Then
            GetShiftTime.Od_From = strShift(row, 1)
            GetShiftTime.Od_To = strShift(row, 2)
            Exit Function
        End If
    Next

   On Error GoTo 0
   Exit Function

GetShiftTime_Error:

    ShowError "Error in procedure GetShiftTime of Form frmODAvail"
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSubmit_Click()
    
   On Error GoTo cmdSubmit_Click_Error

    lstError.Clear
    frmVisible.Visible = True
    frmVisible.Refresh
    Dim row As Integer
    Dim column As Integer
    
    Dim ODInOut As ODInOut
    Dim ArrTDeptT As ArriDeptTime
    
    If SubmitValid Then
        Call AddToList(lstError, "Validation Completed")
        strShift = GetShift
        Call AddToList(lstError, "Data Structure Created For Shift")
        For row = 1 To Flex.Rows - 1
            If Flex.TextMatrix(row, 0) = strChecked Then
                Me.Refresh
                Call AddToList(lstError, Flex.TextMatrix(row, indexOfEmpcode) & "|Work With Checked employee")
                If cmbODFrom.Text = strShiftStart Or cmbODTo.Text = strShitEnd Then
                    ODInOut = GetShiftTime(GetShiftCode(Flex.TextMatrix(row, indexOfEmpcode), Format(txtDate.Text, "DD/MMM/YYYY")))
                End If
                If cmbODFrom.Text = strDept Then
                    ODInOut.Od_From = Flex.TextMatrix(row, indexOfDepTime)
                ElseIf cmbODFrom.Text = strShiftStart Then
                    'do nothing
                Else
                    ODInOut.Od_From = cmbODFrom.Text
                End If
                Call AddToList(lstError, "  Od From Time Changed")
                If cmbODTo.Text = strArrival Then
                    ODInOut.Od_To = Flex.TextMatrix(row, indexOfArrTime)
                ElseIf cmbODTo.Text = strShitEnd Then
                    'do nothing
                Else
                    ODInOut.Od_To = cmbODTo.Text
                End If
                Call AddToList(lstError, "  Od To Time Changed")
                ArrTDeptT.Arr_Time = Flex.TextMatrix(row, indexOfArrTime)
                ArrTDeptT.Dept_Time = Flex.TextMatrix(row, indexOfDepTime)
                'If ValidUpdate(Flex.TextMatrix(row, indexOfEmpcode), ODInOut, ArrTDeptT) Then
                    Flex.TextMatrix(row, indexOfOdFrom) = ODInOut.Od_From
                    Flex.TextMatrix(row, indexOfOfTo) = ODInOut.Od_To
                    
                    If chkStatusUpdate.Value = vbChecked Then
                        If cmbODFrom.Text = strDept And _
                            cmbODTo.Text = strShitEnd Then
                            Flex.TextMatrix(row, indexOfPresabs) = _
                                ChangeStatus(Flex.TextMatrix(row, indexOfPresabs), Right_Side)
                        ElseIf cmbODFrom.Text = strShiftStart And _
                            cmbODTo.Text = strArrival Then
                            Flex.TextMatrix(row, indexOfPresabs) = _
                                ChangeStatus(Flex.TextMatrix(row, indexOfPresabs), Left_Side)
                        Else
                            Flex.TextMatrix(row, indexOfPresabs) = ChangeStatus(Flex.TextMatrix(row, indexOfPresabs), Full)
                        End If
                        Call AddToList(lstError, "  Status Changed")
                    End If
                'End If
                Call AddToList(lstError, Flex.TextMatrix(row, indexOfEmpcode) & "|End Work With Checked employee")
                Call AddToList(lstError, "-----------------------------------------------------")
            Else
                'If chkStatusUpdate.Value = vbChecked Then
                    Call AddToList(lstError, Flex.TextMatrix(row, indexOfEmpcode) & "|Work With Unchecked employee")
                    If cmbStatus.Text <> "As It Is" And InStr(1, _
                        Flex.TextMatrix(row, _
                        indexOfPresabs), strODCode) > 0 Then
                        Flex.TextMatrix(row, indexOfPresabs) = Replace(Flex.TextMatrix(row, _
                        indexOfPresabs), strODCode, _
                        cmbStatus.Text, InStr(1, Flex.TextMatrix(row, _
                        indexOfPresabs), strODCode), 2)
                        Call AddToList(lstError, "  Status Changed")
                    End If
                    Call AddToList(lstError, Flex.TextMatrix(row, indexOfEmpcode) & "|End Work With Unchecked employee")
                    Call AddToList(lstError, "-----------------------------------------------------")
                'End If
            End If
        Next
        cmdUpdate.Enabled = True
        'EnableDisable True
    End If
    frmVisible.Visible = False
    MsgBox "Local Updation Completed", vbInformation
   On Error GoTo 0
   Exit Sub

cmdSubmit_Click_Error:

    ShowError "Error in procedure cmdSubmit_Click of Form frmODAvail"
End Sub

Private Function ChangeStatus(strStatus As String, MakeSide As Side) As String
   On Error GoTo ChangeStatus_Error

    Select Case MakeSide
        Case Left_Side
            ChangeStatus = strODCode & Right(strStatus, 2)
        Case Right_Side
            ChangeStatus = Left(strStatus, 2) & strODCode
        Case Full
            ChangeStatus = strODCode & strODCode
    End Select

   On Error GoTo 0
   Exit Function

ChangeStatus_Error:

    ShowError "Error in procedure ChangeStatus of Form frmODAvail"
End Function
Private Function SubmitValid() As Boolean
   On Error GoTo SubmitValid_Error

    SubmitValid = True
    
    Set adrsC = OpenRecordSet("Select * From NewCaptions Where ID Like '13%'")
    If cmbODFrom.Text <> strShiftStart And cmbODFrom.Text <> strDept Then
        If Val(Right(cmbODFrom.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            cmbODFrom.SetFocus
            SubmitValid = False
            Exit Function
        End If
        If Val(cmbODFrom.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            cmbODFrom.SetFocus
            SubmitValid = False
            Exit Function
        End If
    End If
    If cmbODTo.Text <> strShitEnd And cmbODTo.Text <> strArrival Then
        If Val(Right(cmbODTo.Text, 2)) > 59 Then
            MsgBox NewCaptionTxt("13042", adrsC), vbExclamation
            cmbODTo.SetFocus
            SubmitValid = False
            Exit Function
        End If
        If Val(cmbODTo.Text) > 48 Then
            MsgBox NewCaptionTxt("13043", adrsC), vbExclamation
            cmbODTo.SetFocus
            SubmitValid = False
            Exit Function
        End If
    End If

   On Error GoTo 0
   Exit Function

SubmitValid_Error:

    ShowError "Error in procedure SubmitValid of Form frmODAvail"
End Function

Private Function ValidUpdate(strEmpCode As String, ODFromTo As ODInOut, ArrTDeptT As ArriDeptTime) As Boolean
    
   On Error GoTo ValidUpdate_Error

    If Val(ODFromTo.Od_To) > 0 And Val(ODFromTo.Od_To) < Val(ODFromTo.Od_From) Then
        Call AddToList(lstError, strEmpCode & "|" & NewCaptionTxt("13048", adrsC))
        ValidUpdate = False
        Exit Function
    End If
    
    If Val(ODFromTo.Od_To) > 0 And Val(ODFromTo.Od_From) > Val(ODFromTo.Od_To) Then
        Call AddToList(lstError, strEmpCode & "|" & NewCaptionTxt("13049", adrsC))
        ValidUpdate = False
        Exit Function
    End If
    
    If Val(ODFromTo.Od_To) > 0 And Val(ODFromTo.Od_From) = 0 Then
        Call AddToList(lstError, strEmpCode & "|" & NewCaptionTxt("13050", adrsC))
        ValidUpdate = False
        Exit Function
    End If
    
    If Val(ODFromTo.Od_From) > 0 Then
        If Val(ODFromTo.Od_From) < Val(ArrTDeptT.Arr_Time) Then
            Call AddToList(lstError, strEmpCode & "|" & NewCaptionTxt("13051", adrsC))
            ValidUpdate = False
            Exit Function
        End If
        If Val(ODFromTo.Od_From) > Val(ArrTDeptT.Dept_Time) Then
            Call AddToList(lstError, strEmpCode & "|" & NewCaptionTxt("13051", adrsC))
            ValidUpdate = False
            Exit Function
        End If
    End If
    
    If Val(ODFromTo.Od_To) > 0 Then
        If Val(ODFromTo.Od_To) < Val(ArrTDeptT.Arr_Time) Then
            Call AddToList(lstError, strEmpCode & "|" & NewCaptionTxt("13052", adrsC))
            ValidUpdate = False
            Exit Function
        End If
        If Val(ODFromTo.Od_To) > Val(ArrTDeptT.Dept_Time) Then
            Call AddToList(lstError, strEmpCode & "|" & NewCaptionTxt("13052", adrsC))
            ValidUpdate = False
            Exit Function
        End If
    End If

   On Error GoTo 0
   Exit Function

ValidUpdate_Error:

    ShowError "Error in procedure ValidUpdate of Form frmODAvail"
End Function

Private Sub AddToList(lstError As ListBox, strMessage As String)
   On Error GoTo AddToList_Error

    lstError.AddItem strMessage
    lstError.ListIndex = lstError.ListCount - 1

   On Error GoTo 0
   Exit Sub

AddToList_Error:

    ShowError "Error in procedure AddToList of Form frmODAvail"
End Sub

Private Sub cmdUpdate_Click()
    Dim row As Integer
   On Error GoTo cmdUpdate_Click_Error

    strMon_Trn = MakeName(MonthName(Month(DateCompDate(txtDate.Text))), _
        Year(DateCompDate(txtDate.Text)), "Trn")
    If Not FindTable(strMon_Trn) Then
        MsgBox strMon_Trn & "  This file not found", vbInformation
        Exit Sub
    End If
    With Flex
        For row = 1 To Flex.Rows - 1
            If .TextMatrix(row, 0) = strChecked Then
                VstarDataEnv.cnDJConn.Execute "UPDATE " & strMon_Trn & _
                " SET od_from=" & .TextMatrix(row, indexOfOdFrom) & _
                ",od_to=" & .TextMatrix(row, indexOfOfTo) & _
                ",presabs='" & .TextMatrix(row, indexOfPresabs) & _
                "' WHERE empcode='" & .TextMatrix(row, indexOfEmpcode) & _
                "' AND " & strKDate & "=" & strDTEnc & "" & Format(CDate(txtDate.Text), "DD/MMM/YYYY") & _
                "" & strDTEnc & ""
                Call AddToList(lstError, .TextMatrix(row, indexOfEmpcode) & "|Updated")
            End If
        Next
    End With
    MsgBox "Database Server Update Completed", vbInformation
    'EnableDisable False

   On Error GoTo 0
   Exit Sub

cmdUpdate_Click_Error:

    ShowError "Error in procedure cmdUpdate_Click of Form frmODAvail"
End Sub


Private Sub Flex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo Flex_MouseDown_Error

    If Button = vbRightButton Then
        With Flex
            If .MouseCol = 5 Or .MouseCol = 8 Or .MouseCol = 0 Then
                Call PopupMenu(mnuSelection)
                If .MouseCol = 5 Or .MouseCol = 8 Or .MouseCol = 0 Then
                    Selection (.MouseCol)
                End If
            End If
        End With
    End If

   On Error GoTo 0
   Exit Sub

Flex_MouseDown_Error:

    ShowError "Error in procedure Flex_MouseDown of Form frmODAvail"
End Sub

Private Sub Selection(mCol As Integer)
    Dim iRow As Integer
    Dim iCol As Integer
    'MsgBox mCol
   On Error GoTo Selection_Error

    With Flex
        For iRow = 1 To .Rows - 1
            Select Case mSelection
                Case SelectAll
                    .TextMatrix(iRow, mCol) = strChecked
                Case DeselectAll
                    .TextMatrix(iRow, mCol) = strUnChecked
                Case SelectLate
                    If Val(.TextMatrix(iRow, indexOfLate)) > 0 And Val(.TextMatrix(iRow, indexOfErly)) = 0 Then
                        .TextMatrix(iRow, mCol) = strChecked
                    End If
                Case SelectEarly
                    If Val(.TextMatrix(iRow, indexOfErly)) < 0 And Val(.TextMatrix(iRow, indexOfLate)) = 0 Then
                        .TextMatrix(iRow, mCol) = strChecked
                    End If
            End Select
        Next
    End With

   On Error GoTo 0
   Exit Sub

Selection_Error:

    ShowError "Error in procedure Selection of Form frmODAvail"
End Sub
Private Sub Flex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo Flex_MouseUp_Error

    If Button = 1 Then
        With Flex
            If Not .MouseRow = 0 Then
                If .MouseCol = 5 Or .MouseCol = 8 Or .MouseCol = 0 Then
                    Call TriggerCheckbox(.MouseRow, .MouseCol)
                End If
            End If
        End With
    End If

   On Error GoTo 0
   Exit Sub

Flex_MouseUp_Error:

    ShowError "Error in procedure Flex_MouseUp of Form frmODAvail"
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    Call ComboFill(cmbDept, 2, 2)
    cmbDept.AddItem "ALL"
    cmbDept.Text = cmbDept.List(cmbDept.ListCount - 1)
    txtDate.Text = DateDisp(Date)
    Call SetFormIcon(Me)
    Call FillGridWithArray
    Call SetGridProperty
    Call ODComboFill
    Call StatusComboFill
    EnableDisable False

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    ShowError "Error in procedure Form_Load of Form frmODAvail"
End Sub

Private Sub StatusComboFill()
    Dim adrsTemp As Recordset
   On Error GoTo StatusComboFill_Error

    cmbStatus.AddItem "As It Is"
    Set adrsTemp = OpenRecordSet("SELECT DISTINCT(lvcode) FROM leavdesc WHERE lvcode<>'" & _
        strODCode & "' ORDER BY lvcode")
    Do While Not adrsTemp.EOF
        cmbStatus.AddItem adrsTemp.Fields("lvcode")
        adrsTemp.MoveNext
    Loop
    cmbStatus.Text = cmbStatus.List(0)

   On Error GoTo 0
   Exit Sub

StatusComboFill_Error:

    ShowError "Error in procedure StatusComboFill of Form frmODAvail"
End Sub


Private Sub EnableDisable(MI As Boolean)
   On Error GoTo EnableDisable_Error

    cmdUpdate.Enabled = MI
    cmdSubmit.Enabled = Not MI

   On Error GoTo 0
   Exit Sub

EnableDisable_Error:

    ShowError "Error in procedure EnableDisable of Form frmODAvail"
End Sub

Private Sub ODComboFill()
   On Error GoTo ODComboFill_Error

    cmbODFrom.AddItem strShiftStart ' "SHIFT START"
    cmbODFrom.AddItem strDept ' "DEPARTUTE TIME"
    cmbODFrom.Text = cmbODFrom.List(0)
    cmbODTo.AddItem strShitEnd ' "SHIFT END"
    cmbODTo.AddItem strArrival ' "ARRIVAL TIME"
    cmbODTo.Text = cmbODTo.List(0)

   On Error GoTo 0
   Exit Sub

ODComboFill_Error:

    ShowError "Error in procedure ODComboFill of Form frmODAvail"
End Sub
Private Sub SetGridProperty()
   On Error GoTo SetGridProperty_Error

    With Flex
        .ColWidth(0) = .ColWidth(0)
        .ColWidth(1) = .ColWidth(1)
        .ColWidth(2) = .ColWidth(2) * 2
        .ColWidth(3) = .ColWidth(3) / 2 + 50
        .ColWidth(4) = .ColWidth(4) / 2 + 50
'        .ColWidth(5) = .ColWidth(5) / 2 + 50
        .ColWidth(5) = 0
        .ColWidth(6) = .ColWidth(6) / 2 + 50
        .ColWidth(7) = .ColWidth(7) / 2 + 50
'        .ColWidth(8) = .ColWidth(8) / 2 + 50
        .ColWidth(8) = 0
        .ColWidth(9) = .ColWidth(9) / 2 + 50
        .ColWidth(10) = .ColWidth(10) / 2 + 50
    End With

   On Error GoTo 0
   Exit Sub

SetGridProperty_Error:

    ShowError "Error in procedure SetGridProperty of Form frmODAvail"
End Sub

Private Sub TriggerCheckbox(iRow As Integer, iCol As Integer)
   On Error GoTo TriggerCheckbox_Error

    With Flex
        If .TextMatrix(iRow, iCol) = strUnChecked Then
            .TextMatrix(iRow, iCol) = strChecked
        Else
            .TextMatrix(iRow, iCol) = strUnChecked
        End If
    End With

   On Error GoTo 0
   Exit Sub

TriggerCheckbox_Error:

    ShowError "Error in procedure TriggerCheckbox of Form frmODAvail"
End Sub

Private Sub Flex_KeyPress(KeyAscii As Integer)
   On Error GoTo Flex_KeyPress_Error

If KeyAscii = 13 Or KeyAscii = 32 Then
    With Flex
        If Not .row = 0 Then
            If .Col = 5 Or .Col = 8 Or .Col = 0 Then
                Call TriggerCheckbox(.row, .Col)
            End If
        End If
    End With
End If

   On Error GoTo 0
   Exit Sub

Flex_KeyPress_Error:

    ShowError "Error in procedure Flex_KeyPress of Form frmODAvail"
End Sub

Private Sub FillGridWithArray()
    Dim strTableName As String
    Dim strSql As String
    Dim intCounter As Integer
    Dim rs As Recordset
   On Error GoTo FillGridWithArray_Error

    strTableName = MakeName(MonthName(Month(DateCompDate(txtDate.Text))), _
        Year(DateCompDate(txtDate.Text)), "Trn")
    Flex.Clear
    Flex.Cols = 12
    With Flex
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = "Select"
        .TextMatrix(0, 1) = "Employee No."
        .TextMatrix(0, 2) = "Name"
        .TextMatrix(0, 3) = "Arrival Time"
        .TextMatrix(0, 4) = "Late Come"
        .TextMatrix(0, 5) = "Correction"
        .TextMatrix(0, 6) = "Dept Time"
        .TextMatrix(0, 7) = "Early Time"
        .TextMatrix(0, 8) = "Correction"
        .TextMatrix(0, 9) = "OD From"
        .TextMatrix(0, 10) = "OD To"
        .TextMatrix(0, 11) = "Presabs"
    End With
    If Not FindTable(strTableName) Then
        Exit Sub
    End If
    If cmbDept.Text = "ALL" Then
        strSql = "SELECT " & strTableName & ".Empcode," & strName & _
        ",arrtim,latehrs,deptim,earlhrs,od_to,od_from,presabs FROM " & strTableName & _
        ",Empmst WHERE Empmst.Empcode=" & strTableName & ".Empcode AND " & _
        strTableName & "." & strKDate & "=" & strDTEnc & "" & _
        Format(txtDate.Text, "DD/MMM/YYYY") & "" & strDTEnc & " AND (latehrs>0 OR earlhrs>0)"
    Else
        strSql = "SELECT " & strTableName & ".Empcode," & strName & _
        ",arrtim,latehrs,deptim,earlhrs,od_to,od_from,presabs FROM " & strTableName & _
        ",Empmst WHERE Empmst.Empcode=" & strTableName & ".Empcode AND " & _
        strTableName & "." & strKDate & "=" & strDTEnc & "" & _
        Format(txtDate.Text, "DD/MMM/YYYY") & "" & strDTEnc & " AND Empmst.dept=" & _
        cmbDept.Text & " AND (latehrs>0 OR earlhrs>0) "
    End If
    Set rs = OpenRecordSet(strSql)
    If Not (rs.EOF And rs.BOF) Then
        Flex.Rows = rs.RecordCount + 1
        rs.MoveFirst
        For intCounter = 1 To rs.RecordCount
            With Flex
                Call FlexChecBox(intCounter, 0)
                .TextMatrix(intCounter, 1) = rs.Fields("Empcode")
                .TextMatrix(intCounter, 2) = rs.Fields("name")
                .TextMatrix(intCounter, 3) = FilterNull(rs.Fields("arrtim"))
                .TextMatrix(intCounter, 4) = IIf(FilterNull( _
                    rs.Fields("latehrs")) < 0, 0, FilterNull(rs.Fields("latehrs")))
                Call FlexChecBox(intCounter, 5)
                .TextMatrix(intCounter, 6) = FilterNull(rs.Fields("deptim"))
                .TextMatrix(intCounter, 7) = IIf(FilterNull( _
                    rs.Fields("earlhrs")) >= 0, 0, FilterNull(rs.Fields("earlhrs")))
                '.TextMatrix(intCounter, 7) = rs.Fields("earlhrs")
                Call FlexChecBox(intCounter, 8)
                .TextMatrix(intCounter, 9) = FilterNull(rs.Fields("od_from"))
                .TextMatrix(intCounter, 10) = FilterNull(rs.Fields("od_to"))
                .TextMatrix(intCounter, 11) = rs.Fields("presabs")
            End With
            rs.MoveNext
        Next
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

FillGridWithArray_Error:

    ShowError "Error in procedure FillGridWithArray of Form frmODAvail"
End Sub

Private Sub FlexChecBox(mRow As Integer, mCol As Integer)
   On Error GoTo FlexChecBox_Error

    With Flex
        .row = mRow
        .Col = mCol
        .CellFontName = "Wingdings"
        .CellFontSize = 14
        .CellAlignment = flexAlignCenterCenter
        .TextMatrix(mRow, mCol) = strUnChecked
    End With

   On Error GoTo 0
   Exit Sub

FlexChecBox_Error:

    ShowError "Error in procedure FlexChecBox of Form frmODAvail"
End Sub

Private Sub mnuDeSelectAll_Click()
    mSelection = DeselectAll
End Sub

Private Sub mnuSelectAll_Click()
    mSelection = SelectAll
End Sub

Private Sub mnuSelectEarly_Click()
    mSelection = SelectEarly
End Sub

Private Sub mnuSelectLate_Click()
    mSelection = SelectLate
End Sub

Private Sub txtDate_Click()
   On Error GoTo txtDate_Click_Error

    varCalDt = ""
    varCalDt = Trim(txtDate.Text)
    txtDate.Text = ""
    Call ShowCalendar
    cmbDept_Click
   On Error GoTo 0
   Exit Sub

txtDate_Click_Error:

    ShowError "Error in procedure txtDate_Click of Form frmODAvail"
End Sub

