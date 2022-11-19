Attribute VB_Name = "mdlDate"
Option Explicit

Public Function CDK(Objin As Object, KeyAscii As Integer)
On Error GoTo ERR_P
Select Case KeyAscii            '' Function to Check Date TextBox Keypress
    Case 8:
    Case 13:
        If ValidDate(Objin) Then
            varCalDt = ""
            varCalDt = Trim(Objin.Text)
            Call ShowCalendar
        End If
    Case Else
        If Len(Objin.Text) = 2 Or Len(Objin.Text) = 5 Then
            If InStr("1234567890", Chr(KeyAscii)) <= 0 Then
                KeyAscii = 0
            Else
                Objin.Text = Objin.Text & "/"
                Objin.SelStart = Len(Objin.Text)
                Objin.SelLength = 0
                'SendKeys "{End}", True
            End If
        ElseIf Len(Objin.Text) = 7 Then
            If InStr("1234567890", Chr(KeyAscii)) <= 0 Then
                KeyAscii = 0
            Else
                SendKeys Chr(9)
            End If
        Else
            If InStr("1234567890", Chr(KeyAscii)) <= 0 Then KeyAscii = 0
        End If
End Select
Exit Function
ERR_P:
    ShowError ("CDK :: Date Module")
End Function

Private Function GetYear(ByVal BytY As Byte)
On Error GoTo ERR_P
Select Case BytY            '' Converts the Year into YYYY Format
    Case 0 To 9
        GetYear = CInt("200" & BytY)
    Case 10 To 29
        GetYear = CInt("20" & BytY)
    Case Else
        GetYear = CInt("19" & BytY)
End Select
Exit Function
ERR_P:
    ShowError ("Get Year :: Date Module")
End Function

Private Sub GetDMY(ByRef MEB As Object)
On Error GoTo ERR_P
If bytDateF = 1 Then                        '' Puts Values in the typDMY Variable
    typDMY.bytM = CByte(Left(MEB.Text, 2))
    typDMY.bytD = CByte(Mid(MEB.Text, 4, 2))
ElseIf bytDateF = 2 Then
    typDMY.bytD = CByte(Left(MEB.Text, 2))
    typDMY.bytM = CInt(Mid(MEB.Text, 4, 2))
End If
typDMY.BytY = CByte(Right(MEB.Text, 2))
typDMY.BytY = GetYear(typDMY.BytY)
Exit Sub
ERR_P:
    ShowError ("Get DMY :: Date Module")
End Sub

Private Function ValidYear(ByRef MEB As Object) As Boolean
    ValidYear = True            '' Checks for the Valid Year
End Function

Private Function ValidMonth(ByRef MEB As Object) As Boolean
ValidMonth = True                   '' Checks for the Valid Month
If typDMY.bytM > 12 Or typDMY.bytM = 0 Then
    ValidMonth = False
    MsgBox NewCaptionTxt("M4001", adrsMod), vbExclamation
End If
End Function

Private Function ValidDay(ByRef MEB As Object) As Boolean
ValidDay = True                     '' Checks for the Valid Day
Select Case typDMY.bytM
    Case 2                          '' If Month is 2 i.e. FEBRUARY
        If typDMY.bytD = 29 Then    '' If Day is 29
            If typDMY.BytY Mod 4 = 0 Then   '' If Divisible by 4
                ' Is it a century?
                If typDMY.BytY Mod 100 = 0 Then     '' if Divisible by 100
                    '' If a century, must be evenly divisible by 400.
                    If typDMY.BytY Mod 400 = 0 Then     '' If Divisible by 400
                        ValidDay = True                 '' Leap Year
                    Else
                        ValidDay = False                '' Non-Leap Year
                        MsgBox NewCaptionTxt("00073", adrsMod), vbExclamation
                    End If
                Else
                    ValidDay = True                     '' Leap Year
                End If
            Else
                ValidDay = False                        '' Non-Leap Year
                MsgBox NewCaptionTxt("00073", adrsMod), vbExclamation
            End If
        ElseIf typDMY.bytD > 29 Or typDMY.bytD = 0 Then
            ValidDay = False
            MsgBox NewCaptionTxt("M4002", adrsMod), vbExclamation
        End If
    Case 1, 3, 5, 7, 8, 10, 12          '' 31 Days Month
        If typDMY.bytD > 31 Or typDMY.bytD = 0 Then
            ValidDay = False
            MsgBox NewCaptionTxt("M4002", adrsMod), vbExclamation
        End If
    Case 4, 6, 9, 11                    '' 30 Days Month
        If typDMY.bytD > 30 Or typDMY.bytD = 0 Then
            ValidDay = False
            MsgBox NewCaptionTxt("M4002", adrsMod), vbExclamation
        End If
End Select
End Function

Public Function ValidDate(ByRef MEB As Object) As Boolean
On Error GoTo ERR_P
ValidDate = False                                   '' Checks for the Valid Date format
If Trim(MEB.Text) = "" Then                         '' If Nothing is Entered
    ValidDate = True
    Exit Function
Else
    If Not RetValidLength(MEB) Then Exit Function   '' If Invalid Length
    Call GetDMY(MEB)                                '' Get Day,Month and Year
    If Not ValidDay(MEB) Then Exit Function         '' If Invalid Day
    If Not ValidMonth(MEB) Then Exit Function       '' If Invalid Month
    If Not ValidYear(MEB) Then Exit Function        '' If Invalid Year
    If bytDateF = 1 Then    '' American
        MEB.Text = DateDisp(Format(MEB.Text, "MM/DD/YYYY"))
    Else                    '' British
        MEB.Text = DateDisp(Format(MEB.Text, "DD/MM/YYYY"))
    End If
    ValidDate = True
End If
Exit Function
ERR_P:
    ShowError ("Valid Date :: Date Module")
    ValidDate = False
End Function

Public Function RetValidLength(ByRef MEB As Object, Optional bytFlag As Byte = 0) As Boolean
Dim bytPos As Byte, bytCntTmp As Byte   '' Checks for the Valid Length and Structure
RetValidLength = True
Select Case Len(MEB.Text)
    Case 1 To 7
        '' Checks for Valid Length
        MsgBox NewCaptionTxt("M4003", adrsMod), vbExclamation
        RetValidLength = False
        Exit Function
    Case 8, 10
        '' Checks for Valid Number Of '/' Encounters
        bytCntTmp = 0
        For bytPos = 1 To Len(MEB.Text)
            If Mid(MEB.Text, bytPos, 1) = "/" Then bytCntTmp = bytCntTmp + 1
        Next
        If bytCntTmp <> 2 Then
            MsgBox NewCaptionTxt("M4004", adrsMod), vbExclamation
            RetValidLength = False
            Exit Function
        End If
        '' Checks for Valid '/' Position
        If InStr(MEB.Text, "/") <> 3 And InStr(MEB.Text, "/") <> 6 Then
            MsgBox NewCaptionTxt("M4004", adrsMod), vbExclamation
            RetValidLength = False
            Exit Function
        End If
    Case 0
        If bytFlag = 1 Then
            '' Checks for Valid Length
            MsgBox NewCaptionTxt("M4003", adrsMod), vbExclamation
            RetValidLength = False
            Exit Function
        End If
    Case Else
        '' Checks for Valid Length
        MsgBox NewCaptionTxt("M4003", adrsMod), vbExclamation
        RetValidLength = False
        Exit Function
End Select
End Function

Public Function DateSave(ByVal strDtSource As String) As String
On Error GoTo ERR_P
If Trim(strDtSource) = "" Then  '' Returns the String in the Format of Regional Settings
    DateSave = ""               '' I.e By default M/D/YY
    Exit Function
End If
If bytBackEnd = 2 And bytDateF = 2 Then
    If Day(strDtSource) > Month(strDtSource) Then
        DateSave = Format(strDtSource, "M/D/YYYY")
    Else
        DateSave = strDtSource
    End If
Else
    DateSave = strDtSource
End If
Exit Function
ERR_P:
    ShowError ("Date Save :: Date Module")
End Function

Public Function DateSaveIns(ByVal strDtSource As String) As String
On Error GoTo ERR_P
If Trim(strDtSource) = "" Then  '' Returns the String in the Format of Regional Settings
    DateSaveIns = ""               '' I.e By default M/D/YY
    Exit Function
End If
If bytBackEnd = 2 And bytDateF = 2 Then
    DateSaveIns = Format(strDtSource, "M/D/YYYY")
Else
    DateSaveIns = Format(strDtSource, "dd/mmm/yyyy")
End If
''DateSave = strDtSource
Exit Function
ERR_P:
    ShowError ("Date SaveIns :: Date Module")
End Function

Public Function DateCompStr(ByVal strDtSource As String) As String
On Error GoTo ERR_P
If Trim(strDtSource) = "" Then  '' Returns the String in the Format of Regional Settings
    DateCompStr = ""            '' I.e By default M/D/YY
    Exit Function
End If
If bytBackEnd = 2 And bytDateF = 2 Then
    DateCompStr = CStr(Format(strDtSource, "M/D/YYYY"))
Else
    DateCompStr = CStr(Format(strDtSource, "dd/mmm/yyyy"))
End If
Exit Function
ERR_P:
    ShowError ("Date Comp Str :: Date Module")
End Function

Public Function DateCompDate(ByVal strDtSource As String) As Date
On Error GoTo ERR_P
If Trim(strDtSource) = "" Then  '' Returns the Date in the Format of Regional Settings
    DateCompDate = Empty        '' I.e By default M/D/YY
    Exit Function
End If
DateCompDate = CDate(strDtSource)
Exit Function
ERR_P:
    ShowError ("Date Comp Date :: Date Module")
End Function

Public Function DateDisp(ByVal strDtSource As String) As String
On Error GoTo ERR_P
If Trim(strDtSource) = "" Then  '' Returns the String Formatted in Viewing Format
    DateDisp = ""
    Exit Function
End If
If bytDateF = 2 Then
    DateDisp = Format(strDtSource, "DD/MM/YYYY")
Else
    DateDisp = Format(strDtSource, "MM/DD/YYYY")
End If
Exit Function
ERR_P:
    ShowError ("Date Disp :: Date Module")
End Function

Public Function fncDateFormat() As Boolean      '' Function to Get the Regional date Format
On Error GoTo ERR_P
Dim lBuffSize As String
Dim sBuffer As String
Dim lRet As Long
lBuffSize = 256
sBuffer = String$(lBuffSize, vbNullChar)
lRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_SSHORTDATE, sBuffer, lBuffSize)
If lRet > 0 Then
    sBuffer = Left$(sBuffer, lRet - 1)
End If
fncDateFormat = True
If bytDateF = 1 And UCase(sBuffer) <> "M/D/YY" Then fncDateFormat = False
If bytDateF = 2 And UCase(sBuffer) <> "DD/MM/YY" Then fncDateFormat = False
Exit Function
ERR_P:
    ShowError ("fnc Date Format :: Date Module")
    fncDateFormat = False
End Function


Public Function MakeName(ByVal Month$, ByVal Year$, ByVal FileName$) As String
        MakeName = Left(Month$, 3) & Right(Year, 2) & strConv(FileName$, vbProperCase)
End Function

Public Function MonthNumber(ByVal strMonthName As String) As Byte
Select Case UCase(Left(strMonthName, 3))
    Case "JAN": MonthNumber = 1
    Case "FEB": MonthNumber = 2
    Case "MAR": MonthNumber = 3
    Case "APR": MonthNumber = 4
    Case "MAY": MonthNumber = 5
    Case "JUN": MonthNumber = 6
    Case "JUL": MonthNumber = 7
    Case "AUG": MonthNumber = 8
    Case "SEP": MonthNumber = 9
    Case "OCT": MonthNumber = 10
    Case "NOV": MonthNumber = 11
    Case "DEC": MonthNumber = 12
    Case Else: MonthNumber = 0
End Select
End Function
''For Mauritius 05-08-03
Public Function GetTO_Char(ByVal strDate As String) As String
Select Case bytBackEnd
    Case 1, 2 ''SQL-Server,MS-Access
        GetTO_Char = strDTEnc & DateCompStr(strDate) & strDTEnc
    Case 3  ''Oracle
         GetTO_Char = "'" & Format(strDate, "dd/mm/yyyy") & "'"
End Select
End Function

Public Function GetshiftFile(dttmp) As String
GetshiftFile = MonthName(Month(dttmp), True) & Right(Year(dttmp), 2) & "shf"
End Function
