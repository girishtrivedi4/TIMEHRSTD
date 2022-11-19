Attribute VB_Name = "ModFunctions"

' *********************************************************
Public Function AlignValue(ByVal pstrtext As String, _
                           ByVal Pstralign As String, _
                           ByVal Pintsize As Integer, _
                           Optional CutYN) As String

    Dim strrretval As String
    Dim i As Integer
    Dim j As Integer

    ' // ALIGN VALUES
    pstrtext = Trim(pstrtext)
    If Len(pstrtext) >= Pintsize Then
        If IsMissing(CutYN) Then
            AlignValue = pstrtext
        Else
            If Trim(UCase(CutYN)) = "Y" Then
                AlignValue = Left(pstrtext, Pintsize)
            Else
                AlignValue = pstrtext
            End If
        End If
        Exit Function
    End If
    
    i = Len(pstrtext)
    Select Case UCase(Pstralign)
        Case "C"
            j = Int((Pintsize - i) / 2)
            strrretval = Padl(pstrtext, j + Len(pstrtext))
            strrretval = "a" & strrretval
            strrretval = Padr(strrretval, Pintsize + 1)
            strrretval = Right(strrretval, Len(strrretval) - 1)
        Case "R"
            'j = PIntSize - i
            strrretval = Padl(pstrtext, Pintsize)
        Case "L"
            j = Pintsize - i
            strrretval = Padr(pstrtext, Pintsize)
    End Select
    AlignValue = strrretval
    
End Function

Private Function Padl(ByVal Psstr As String, _
                      ByVal Pintlen As Integer) As String

    Psstr = Trim(Psstr)
    If Len(Psstr) > Pintlen Then
        Padl = Left$(Psstr, Pintlen)
    Else
        Padl = Space(Pintlen - Len(Psstr)) & Psstr
    End If
    
End Function
Private Function Padr(ByVal sstr As String, _
                      ByVal intlen As Integer) As String

    Dim i As Integer
    sstr = Trim(sstr)
    If Len(sstr) > intlen Then
        Padr = Left$(sstr, intlen)
    Else
        Padr = sstr & Space(intlen - Len(sstr))
    End If
    
End Function




