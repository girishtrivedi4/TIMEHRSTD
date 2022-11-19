Attribute VB_Name = "mdl2007"
'---------------------------------------------------------------------------------------
' Module    : mdl2007
' DateTime  : 13/07/07 10:11
' Author    :
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : GetVersion
' DateTime  : 29/07/2008 14:08
' Author    :
' Purpose   : To Get Current Version Of MIS
' Pre       :
' Post      :
' Return    : String
'---------------------------------------------------------------------------------------
'
Public Function GetVersion(Optional blnWithTital As Boolean = False) As String
On Error GoTo Err
    GetVersion = "Version-" & App.Major & "." & App.Minor & "." & App.Revision
    If blnWithTital = True Then
        GetVersion = App.Title & ": " & GetVersion
    End If
Exit Function
Err:
    Call ShowError("GetVersion")
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetValueForTag
' DateTime  : 29/07/2008 14:08
' Author    :
' Purpose   : To Get All Tag Value in one string.
' Pre       :
' Post      :
' Return    : Variant
'---------------------------------------------------------------------------------------
'
Public Function GetValueForTag()
On Error GoTo Err
    Dim FileObj As New FileSystemObject
    If Not FileObj.FileExists(App.path & "\Data\Tag.kab") Then Exit Function
    Open App.path & "\Data\Tag.kab" For Input As #1
    'strTags = Input(LOF(1), #1)
    strTags = Split(Input(LOF(1), #1), vbCrLf)
    Close #1
    If ArrayDimensions(strTags) = 1 Then
        blnTagArray = True
    End If
Exit Function
Err:
    Call ShowError("GetValueForTag")
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFlagStatus
' DateTime  : 29/07/2008 14:08
' Author    :
' Purpose   : Use To check whether your tag present in Tag file or not.
' Pre       :
' Post      :
' Return    : Boolean
'---------------------------------------------------------------------------------------
'
Public Function GetFlagStatus(strFlag) As Boolean
On Error GoTo Err
    'GetFlagStatus = InStr(1, strTags, strFlag)
    Dim strTemp
    If blnTagArray Then
        For Each strTemp In strTags
            If strTemp = strFlag Then
                GetFlagStatus = True
                Exit Function
            End If
        Next
    End If
 Exit Function
Err:
    Call ShowError("GetFlagStatus")
End Function

' Returns 0 for unintialized array, -1 for non-array,1 for intialized array
'---------------------------------------------------------------------------------------
' Procedure : ArrayDimensions
' DateTime  : 23/Sep/2008 10:20
' Author    :
' Purpose   : To check whether Array Initialize
' Return    : Long
'---------------------------------------------------------------------------------------
'
Public Function ArrayDimensions(pvarArray As Variant) As Long
    Dim lngTemp As Long
    Dim i As Long
    
    On Error Resume Next
    Do
        i = i + 1
        lngTemp = UBound(pvarArray, i)
        Select Case Err.Number
            Case 13: ArrayDimensions = -1
            Case 9: ArrayDimensions = i - 1
        End Select
    Loop Until Err.Number
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetLeaveCode
' DateTime  : 29/07/2008 14:09
' Author    :
' Purpose   : To Get All Leave Code present in Leave master.
' Pre       :
' Post      :
' Return    : Variant
'---------------------------------------------------------------------------------------
'
Public Function GetLeaveCode(strSeperator As String, _
    Optional strPaidOrUn As String = "", Optional strNotIn As String)
    
On Error GoTo Err
Dim strType As String
Dim adrsLeave As New ADODB.Recordset
If adrsLeave.State = 1 Then adrsLeave.Close
If strPaidOrUn <> "" Then
    Set adrsLeave = OpenRecordSet("SELECT DISTINCT lvcode FROM Leavdesc " & _
        " Where paid = " & strPaidOrUn & " AND " & strNotIn & "")
Else
    Set adrsLeave = OpenRecordSet("SELECT DISTINCT lvcode FROM Leavdesc " & _
    "where" & strNotIn & " ")
End If
Do While Not adrsLeave.EOF
    GetLeaveCode = GetLeaveCode & adrsLeave.Fields("lvcode") & strSeperator
    adrsLeave.MoveNext
Loop
Exit Function
Err:
    Call ShowError("GetLeaveCode")
End Function

Public Function GetLEaveCodeFromBal(strY As String, strD As String) As String
On Error GoTo Err
Dim adrT As New ADODB.Recordset
Set adrT = OpenRecordSet("SELECT * FROM " & strY & "")
For i = 0 To adrT.Fields.Count - 1
    GetLEaveCodeFromBal = GetLEaveCodeFromBal & strY & "." & adrT.Fields(i).name & strD
Next
Exit Function
Err:
    Call ShowError("Error in GetLEaveCodeFromBal")
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetMonthEnd
' DateTime  : 29/07/2008 14:10
' Author    :
' Purpose   : To Get Month End Date.
' Pre       :
' Post      :
' Return    : Variant
'---------------------------------------------------------------------------------------
'
Public Function GetMonthEnd(strMonth As String, strYear As String)
On Error GoTo Err
    GetMonthEnd = Format(DateAdd("d", -1, Format(DateAdd("m", 1, _
        Format(CDate("" & strMonth & " " & strYear & ""), "dd/mmm/yyyy")), "dd/mmm/yyyy")), "dd/MMM/yyyy")
    Exit Function
Err:
    Call ShowError("GetMonthEnd " & vbCrLf & Err.Description)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetMonthStart
' DateTime  : 29/07/2008 14:10
' Author    :
' Purpose   : To Get month start Date.
' Pre       :
' Post      :
' Return    : Variant
'---------------------------------------------------------------------------------------
'
Public Function GetMonthStart(strMonth As String, strYear As String)
On Error GoTo Err
    GetMonthStart = Format(CDate("" & strMonth & " " & strYear & ""), "DD/MMM/YYYY")
Exit Function
Err:
    Call ShowError("GetMonthStart" & vbCrLf & Err.Description)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetLeaveCodeFromList
' DateTime  : 11/06/2008 12:32
' Author    :
' Purpose   : To Get
' Pre       : Get List Box as a parameter
' Post      : And stored all selected item in array of string
' Return    : Return that Array
'---------------------------------------------------------------------------------------
'
Public Function GetSelectedItemsFromList(mlst As ListBox) _
    As String()
    Dim ListItemWalker As Integer
    Dim ArrayWalker As Integer
    Dim strCode() As String
   On Error GoTo GetLeaveCodeFromList_Error
    ReDim strCode(mlst.ListCount - 1)
    For ListItemWalker = 0 To mlst.ListCount - 1
        If mlst.Selected(ListItemWalker) = True Then
            strCode(ArrayWalker) = mlst.List(ListItemWalker)
            ArrayWalker = ArrayWalker + 1
        End If
    Next
    ReDim Preserve strCode(ArrayWalker - 1)
    GetSelectedItemsFromList = strCode
   On Error GoTo 0
   Exit Function

GetLeaveCodeFromList_Error:

    ShowError "Error in procedure GetLeaveCodeFromList of Module mdl2007"
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetTrnYear
' DateTime  : 11/06/2008 13:58
' Author    :
' Purpose   : To Get current Year
' Pre       : Input(Date)
' Post      : If month of input date is greater than Year Start then
'               return +1 current year or else return same year
' Return    :
'---------------------------------------------------------------------------------------
'
Public Function GetTrnYear(mDate As Date) As Integer
   On Error GoTo GetTrnYear_Error

    GetTrnYear = IIf(Month(mDate) < Val(pVStar.Yearstart), _
                pVStar.YearSel + 1, pVStar.YearSel)

   On Error GoTo 0
   Exit Function

GetTrnYear_Error:

    ShowError "Error in procedure GetTrnYear of Module mdl2007"
End Function


'---------------------------------------------------------------------------------------
' Procedure : PachClient
' DateTime  : 02/07/2008 14:58
' Author    :
' Purpose   : Add Extra Filed,Add Extra Table,Or Any Database change Throght query Please Check this function
' Pre       :
' Post      :
' Return    :
'---------------------------------------------------------------------------------------
'
Public Sub PatchClient()
On Error GoTo PatchClient_Error
    'to create database INI File
    Call CreateDatabaseINI
    'To Change Standard Database use this function
    Call StandardChange
    'To Change Client Specific database use this function
    Call ClientChange
On Error GoTo 0
Exit Sub
PatchClient_Error:
   If Erl = 0 Then
      ShowError "Error in procedure PatchClient of Module mdl2007"
   Else
      ShowError "Error in procedure PatchClient of Module mdl2007 And Line:" & Erl
   End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : StandardChange
' DateTime  : 26/07/2008 11:27
' Author    :
' Purpose   : All Standard related table or column please modified here.
' Pre       :
' Post      :
' Return    :
'---------------------------------------------------------------------------------------
'
Public Sub StandardChange()
'for standard
On Error GoTo ReportRelated_Error
    If bytBackEnd = 3 Then                  '' if Conditions Add By  20-12 From Error Rectification In Oracle
        If FieldExists("male", "fromdate") Then
            ConMain.Execute ParseQuery("ALTER TABLE male Modify fromdate ~TEXT~(50)")
        End If
        If FieldExists("male", "todate") Then
            ConMain.Execute ParseQuery("ALTER TABLE male Modify todate ~TEXT~(50)")
        End If
    Else
'    If FieldExists("male", "fromdate") Then
'        ConMain.Execute ParseQuery("ALTER TABLE male ALTER " & _
'            " COLUMN fromdate ~TEXT~(50)")
'    End If
'    If FieldExists("male", "todate") Then
'        ConMain.Execute ParseQuery("ALTER TABLE male ALTER " & _
'            " COLUMN todate ~TEXT~(50)")
'    End If
    End If
    If Not FieldExists("DPrAb", "offsT") Then
        ConMain.Execute ParseQuery("Alter table DPrAb " & _
            "Add offsT ~TEXT~(10)")
    End If
    If FieldExists("UserAccs", "UserModdUser") Then
        ConMain.Execute "ALTER TABLE UserAccs " & _
            " DROP COLUMN UserModdUser"
        ConMain.Execute ParseQuery("ALTER TABLE UserAccs " & _
            " ADD UserModUser ~TEXT~(20)")
    End If

    If Not FieldExists("CatDesc", "WeekOffPaid") Then               '' Add By  27-12
        ConMain.Execute ParseQuery("ALTER TABLE CatDesc " & _
            " ADD WeekOffPaid ~TEXT~(1)")
        ConMain.Execute "Update CatDesc Set WeekOffPaid = 'Y'"
    End If


    If Not FindTable("ECode") Then  ' 13-03
        Select Case bytBackEnd
            Case 1 ''SQL Server
                ConMain.Execute "Create table ECode(empcode nvarchar (8))"
            Case 2 ''MS-ACCESS
                ConMain.Execute "Create table ECode(empcode text (8))"
            Case 3 ''ORACLE
                ConMain.Execute "Create table ECode(empcode varchar2 (8))"
        End Select
    End If
   

        If Not FieldExists("Leavdesc", "CrMonthly") Then
            ConMain.Execute ParseQuery("ALTER TABLE Leavdesc " & _
                " ADD CrMonthly ~TEXT~(1)")
            ConMain.Execute "Update Leavdesc Set CrMonthly = 'N'"
        End If

 
    If blnIPAddress And Not FindTable("LocationIP") Then        ' 31-08
        Select Case bytBackEnd
            Case 1 ''SQL Server
                ConMain.Execute "Create table LocationIP(Location smallint,IP nvarchar (15))"
                ConMain.Execute "ALTER TABLE DailyPro ADD IP nvarchar (15)"
                ConMain.Execute "ALTER TABLE TblData ALTER COLUMN strF1 nvarchar(35)"
        End Select
    End If

       
On Error GoTo 0
Exit Sub
ReportRelated_Error:
   If Erl = 0 Then
      ShowError "Error in procedure ReportRelated of Module mdl2007"
   Else
      ShowError "Error in procedure ReportRelated of Module mdl2007 And Line:" & Erl
      Resume Next
   End If
End Sub

Public Sub ClientChange()
On Error GoTo ClientChange_Error


On Error GoTo 0
Exit Sub
ClientChange_Error:
   If Erl = 0 Then
      ShowError "Error in procedure ClientChange of Module mdl2007"
   Else
      ShowError "Error in procedure ClientChange of Module mdl2007 And Line:" & Erl
   End If
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : GetShift
' DateTime  : 21/06/2008 11:02
' Author    :
' Purpose   : To Get Shift Hours
' Pre       :
' Post      :
' Return    : Single
'---------------------------------------------------------------------------------------
'
Public Function GetShift(strSCode As String) As clsShift
    Dim adrsShift As ADODB.Recordset
On Error GoTo GetShift_Error
    Set adrsShift = OpenRecordSet("SELECT * FROM instshft " & _
    " WHERE Shift='" & strSCode & "'")
    If adrsShift.EOF Then Exit Function
    Dim mGetShift As New clsShift
    mGetShift.blnNight = IIf(adrsShift("Night") = 0, False, True)
    mGetShift.sngB1I = IIf(IsNull(adrsShift("Rst_In")), 0, adrsShift("Rst_In"))
    mGetShift.sngB1O = IIf(IsNull(adrsShift("Rst_Out")), 0, adrsShift("Rst_Out"))
    mGetShift.sngB2I = IIf(IsNull(adrsShift("Rst_In_2")), 0, adrsShift("Rst_In_2"))
    mGetShift.sngB2O = IIf(IsNull(adrsShift("Rst_Out_2")), 0, adrsShift("Rst_Out_2"))
    mGetShift.sngB3I = IIf(IsNull(adrsShift("Rst_In_3")), 0, adrsShift("Rst_In_3"))
    mGetShift.sngB3O = IIf(IsNull(adrsShift("Rst_Out_3")), 0, adrsShift("Rst_Out_3"))
    mGetShift.sngBH1 = IIf(IsNull(adrsShift("Rst_Brk")), 0, adrsShift("Rst_Brk"))
    mGetShift.sngBH2 = IIf(IsNull(adrsShift("Rst_Brk_2")), 0, adrsShift("Rst_Brk_2"))
    mGetShift.sngBH3 = IIf(IsNull(adrsShift("Rst_Brk_3")), 0, adrsShift("Rst_Brk_3"))
    mGetShift.sngHalfE = IIf(IsNull(adrsShift("HDEnd")), 0, adrsShift("HDEnd"))
    mGetShift.sngHalfS = IIf(IsNull(adrsShift("HDStart")), 0, adrsShift("HDStart"))
    mGetShift.sngHRS = IIf(IsNull(adrsShift("Shf_Hrs")), 0, adrsShift("Shf_Hrs"))
    mGetShift.sngIN = IIf(IsNull(adrsShift("Shf_In")), 0, adrsShift("Shf_In"))
    mGetShift.sngOut = IIf(IsNull(adrsShift("Shf_Out")), 0, adrsShift("Shf_Out"))
    mGetShift.strShift = IIf(IsNull(adrsShift("Shift")), "", adrsShift("Shift"))
    mGetShift.sngUPTO = IIf(IsNull(adrsShift("UPTO")), 0, adrsShift("UPTO"))
    Set GetShift = mGetShift
On Error GoTo 0
Exit Function
GetShift_Error:
   If Erl = 0 Then
      ShowError "Error in procedure GetShift of Module mdl2007"
   Else
      ShowError "Error in procedure GetShift of Module mdl2007 And Line:" & Erl
   End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateDatabaseINI
' DateTime  : 29/07/2008 10:23
' Author    :
' Purpose   : Create Database Keywold INI File
' Pre       :
' Post      :
' Return    :
'---------------------------------------------------------------------------------------
'
Public Sub CreateDatabaseINI()
On Error GoTo CreateDatabaseINI_Error
    If GetINIString("FirstTime", "Flag", App.path & _
                "\DBKeyWord.ini") = "TRUE" Then Exit Sub
    Call WriteINIString("1", "~TEXT~", "NVARCHAR", App.path & _
                "\DBKeyWord.ini")
    Call WriteINIString("1", "~SMALLINT~", "SMALLINT", App.path & _
                "\DBKeyWord.ini")
    Call WriteINIString("2", "~TEXT~", "TEXT", App.path & _
                "\DBKeyWord.ini")
    Call WriteINIString("2", "~SMALLINT~", "SMALLINT", App.path & _
                "\DBKeyWord.ini")
    Call WriteINIString("3", "~TEXT~", "VARCHAR2", App.path & _
                "\DBKeyWord.ini")
    Call WriteINIString("3", "~SMALLINT~", "SMALLINT", App.path & _
                "\DBKeyWord.ini")
    Call WriteINIString("FirstTime", "Flag", "TRUE", App.path & _
                "\DBKeyWord.ini")
On Error GoTo 0
Exit Sub
CreateDatabaseINI_Error:
   If Erl = 0 Then
      ShowError "Error in procedure CreateDatabaseINI of Module mdl2007"
   Else
      ShowError "Error in procedure CreateDatabaseINI of Module mdl2007 And Line:" & Erl
   End If
End Sub
