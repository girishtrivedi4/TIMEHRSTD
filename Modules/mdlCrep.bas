Attribute VB_Name = "mdlCrep"

':
'Objects to be initalised for Crystal Reports
Public strOrderBy As String
Public RptChk As Integer
Public Report As CRAXDRT.Report
Public crxApp As New CRAXDRT.Application
Public CRPrinterSettings As CRAXDRT.Report
Public crGrp(1 To 7) As CRAXDRT.DatabaseFieldDefinition
Public crLbl(1 To 7) As CRAXDRT.DatabaseFieldDefinition
Public crHeader As CRAXDRT.DatabaseFieldDefinition
Public CRV As CrystalActiveXReportViewer
Public adrsCrep As New ADODB.Recordset
'Public rptCon As New ADODB.Connection
Public empstr3 As String   '' Used for storing commandtext temporary
Public blnIntz As Boolean
Public strAGrp(1 To 7) As String, strAlbl(1 To 7) As String, strAhead(1 To 7) As String
Public crFieldObject(7), crxGroupHeader(7), crlblObject(7), crxSetInReport(7)
Public StrCrUser As String, StrcrPwd As String, StrcrSvr As String
'Public StrLvCD(1 To 10) As String ''Previous
'Public strAlv(1 To 10) As String  ''Previous
Public StrLvCD(1 To 20) As String  '' 20-11
Public strAlv(1 To 20) As String   '' 20
Public bytPrint As Byte
Public strlstdt As String
Public StrTitle As String
Public RptDel As Byte
Public RptExp As Byte
Public ExpBl As Boolean
Public mlVGAWidth          As Long
    Public mlVGAHeight         As Long
    Public glVGAWidth       As Long '2000/09/22 Added
    Public glVGAHeight      As Long '2000/09/22 Added
    Public ml800x600Width   As Long '2000/12/30 Made public
    Public ml800x600Height  As Long '2001/01/01 Made public
    Public ml1024x768Width  As Long '2000/12/20 Made public
    Public ml1024x768Height As Long '2000/12/20 Made public
    Public ml1152x864Width      As Long     '2000/12/30 Made public
    Public ml1152x864Height     As Long     '2001/05/09 Made Public
    Public gl1280x768Width      As Long     '2003/01/17 Added to support Sony SDM-V72M
    Public gl1280x768Height     As Long     '2003/01/17 Added to support Sony SDM-V72M
    Public ml1280x1024Width     As Long     '2001/05/09 Made Public
    Public ml1280x1024Height    As Long     '2001/05/09 Made Public
    Public ml1400x1050Width     As Long     '2001/01/28 Added
    Public ml1400x1050Height    As Long     '2001/01/28 Added
    Public ml1600x1200Width     As Long     '2001/01/28 Added
    Public ml1600x1200Height    As Long     '2001/05/09 Made Public
    
    Public ml1680x1050Width As Long ' 24/Dec/2009
    Public ml1680x1050Height As Long ' 24/Dec/2009
    Public ml1792x1344Width As Long ' 24/Dec/2009
    Public ml1792x1344Height As Long ' 24/Dec/2009
    Public ml1800x1440Width As Long ' 24/Dec/2009
    Public ml1800x1440Height As Long ' 24/Dec/2009
    Public ml1856x1392Width As Long ' 24/Dec/2009
    Public ml1856x1392Height As Long ' 24/Dec/2009
    Public ml1920x1080Width As Long ' 24/Dec/2009
    Public ml1920x1080Height As Long ' 24/Dec/2009
    Public ml1920x1200Width As Long ' 24/Dec/2009
    Public ml1920x1200Height As Long ' 24/Dec/2009
    Public ml1920x1400Width As Long ' 24/Dec/2009
    Public ml1920x1400Height As Long ' 24/Dec/2009
    Public ml1920x1440Width As Long ' 24/Dec/2009
    Public ml1920x1440Height As Long ' 24/Dec/2009
    Public ml2048x1152Width As Long ' 24/Dec/2009
    Public ml2048x1152Height As Long ' 24/Dec/2009
    Public ml2048x1536Width As Long ' 24/Dec/2009
    Public ml2048x1536Height As Long ' 24/Dec/2009
      
        
    
    
    Public mntresREs As String


Public Const MONITORINFOF_PRIMARY = &H1
Public Const MONITOR_DEFAULTTONEAREST = &H2
Public Const MONITOR_DEFAULTTONULL = &H0
Public Const MONITOR_DEFAULTTOPRIMARY = &H1
' :

' API functions and constants used in Printer setting
Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
   "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
   ByVal pDefault As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" ( _
   ByVal hPrinter As Long) As Long
Private Declare Function DeviceCapabilities Lib "winspool.drv" _
   Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, _
   ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, _
   ByVal dev As Long) As Long
Public Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Public Declare Function EnumDisplayMonitors Lib "user32.dll" (ByVal hdc As Long, ByRef lprcClip As Any, ByVal lpfnEnum As Long, ByVal dwData As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long


Private Const DC_BINS = 6
Private Const DC_BINNAMES = 12

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type
Public Type POINT
    x As Long
    y As Long
End Type



':
Public Sub EnumPrinterBins(PrinterName As String, cbo As ComboBox)
On Error GoTo ERR_P
    Dim prn As Printer
    Dim hPrinter As Long                ' Handle of the selected printer
    Dim dwbins As Long                  ' The number of paperbins in the printer
    Dim i As Long                       ' counter
    Dim nameslist As String             ' The string returned with all the bin names
    Dim NameBin As String               ' The parsed bin name
    Dim numBin() As Integer             ' Used as part of the DeviceCapabilities API call
     
    cbo.clear
    For Each prn In Printers
        ' Look through all the currently installed printers
        If prn.DeviceName = PrinterName Then
            ' We've found our printer -- open a handle to it
            If OpenPrinter(prn.DeviceName, hPrinter, 0) <> 0 Then
                ' Get the available bin numbers
                dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, _
                                        DC_BINS, ByVal vbNullString, 0)
                ReDim numBin(1 To dwbins)
                nameslist = String(24 * dwbins, 0)
                dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, _
                                        DC_BINS, numBin(1), 0)
                dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, _
                                        DC_BINNAMES, ByVal nameslist, 0)
                For i = 1 To dwbins
                    ' For each bin number, add its corresponding name
                    ' to our combo box
                    NameBin = Mid(nameslist, 24 * (i - 1) + 1, 24)
                    NameBin = Left(NameBin, InStr(1, NameBin, Chr(0)) - 1)
                    cbo.AddItem NameBin
                    cbo.ItemData(cbo.NewIndex) = numBin(i)
                Next i
                ' Close the printer
                Call ClosePrinter(hPrinter)
            Else
                ' OpenPrinter failed, so we can't retrieve information about it
                cbo.AddItem prn.DeviceName & "  <Unavailable>"
            End If
        End If
    Next prn
    Exit Sub
ERR_P:
     MsgBox Err.Description, vbInformation
End Sub


Public Sub PrnAtul()
On Error GoTo ERR_P
': Printing function
Select Case frmReports.cboPaperOrientation.Text
  Case "Landscape"
        Report.PaperOrientation = crLandscape
  Case "Portrait"
        Report.PaperOrientation = crPortrait
 End Select

Select Case frmReports.cboPrinterDuplex.Text
  Case "Simplex"
  Report.PrinterDuplex = crPRDPSimplex
  Case "Horizontal"
  Report.PrinterDuplex = crPRDPHorizontal
  Case "Vertical"
  Report.PrinterDuplex = crPRDPVertical
End Select
  
 Select Case frmReports.cboPaperSize.Text
 Case "Default"
    Report.PaperSize = crDefaultPaperSize
  Case "Letter"
    Report.PaperSize = crPaperLetter
   Case "Small Letter"
    Report.PaperSize = crPaperLetterSmall
   Case "Legal"
    Report.PaperSize = crPaperLegal
   Case "10x14"
    Report.PaperSize = crPaper10x14
    Case "11x17"
    Report.PaperSize = crPaper11x17
    Case "A3"
    Report.PaperSize = crPaperA3
    Case "A4"
    Report.PaperSize = crPaperA4
    Case "A4 Small"
    Report.PaperSize = crPaperA4Small
    Case "A5"
    Report.PaperSize = crPaperA5
    Case "B4"
    Report.PaperSize = crPaperB4
    Case "B5"
    Report.PaperSize = crPaperB5
    Case "C Sheet"
    Report.PaperSize = crPaperCsheet
    Case "D Sheet"
    Report.PaperSize = crPaperDsheet
    Case "Envelope 9"
    Report.PaperSize = crPaperEnvelope9
    Case "Envelope 10"
    Report.PaperSize = crPaperEnvelope10
    Case "Envelope 11"
    Report.PaperSize = crPaperEnvelope11
    Case "Envelope 12"
    Report.PaperSize = crPaperEnvelope12
    Case "Envelope 14"
    Report.PaperSize = crPaperEnvelope14
    Case "Envelope B4"
    Report.PaperSize = crPaperEnvelopeB4
    Case "Envelope B5"
    Report.PaperSize = crPaperEnvelopeB5
    Case "Envelope B6"
    Report.PaperSize = crPaperEnvelopeB6
    Case "Envelope C3"
    Report.PaperSize = crPaperEnvelopeC3
    Case "Envelope C4"
    Report.PaperSize = crPaperEnvelopeC4
    Case "Envelope C5"
    Report.PaperSize = crPaperEnvelopeC5
    Case "Envelope C6"
    Report.PaperSize = crPaperEnvelopeC6
    Case "Envelope C65"
    Report.PaperSize = crPaperEnvelopeC65
    Case "Envelope DL"
    Report.PaperSize = crPaperEnvelopeDL
    Case "Envelope Italy"
    Report.PaperSize = crPaperEnvelopeItaly
    Case "Envelope Monarch"
    Report.PaperSize = crPaperEnvelopeMonarch
    Case "Envelope Personal"
    Report.PaperSize = crPaperEnvelopePersonal
    Case "E Sheet"
    Report.PaperSize = crPaperEsheet
    Case "Executive"
    Report.PaperSize = crPaperExecutive
    Case "Fanfold Legal German"
    Report.PaperSize = crPaperFanfoldLegalGerman
    Case "Fanfold Standard German"
    Report.PaperSize = crPaperFanfoldStdGerman
    Case "Fanfold US"
    Report.PaperSize = crPaperFanfoldUS
    Case "FanFold 8.5 * 12"
    Report.PaperSize = 119
    Case "Folio"
    Report.PaperSize = crPaperFolio
    Case "Ledger"
    Report.PaperSize = crPaperLedger
    Case "Note"
    Report.PaperSize = crPaperNote
    Case "Quarto"
    Report.PaperSize = crPaperQuarto
    Case "Statement"
    Report.PaperSize = crPaperStatement
    Case "Tabloid"
    Report.PaperSize = crPaperTabloid
End Select
Exit Sub
ERR_P:
    MsgBox Err.Description, vbOKOnly, "INFORMATION"
End Sub

Public Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, lprcMonitor As RECT, ByVal dwData As Long) As Long
': for mseting screen resolution wise
    Dim MI As MONITORINFO, R As RECT
   ' Debug.Print "Moitor handle: " + CStr(hMonitor)
      MI.cbSize = Len(MI)
    GetMonitorInfo hMonitor, MI
    mntresREs = Trim(CStr(MI.rcMonitor.Right - MI.rcMonitor.Left) + "x" + CStr(MI.rcMonitor.Bottom - MI.rcMonitor.Top))
    'MsgBox mntresREs
    MonitorEnumProc = 1
End Function



Public Function TB_FillDesktop32(frm As Form, Optional vntAlreadyFull As Variant) As Boolean
': for mseting screen resolution wise
    Dim tDesktopArea    As RECT
    Dim lScreenLeft     As Long
    Dim lScreenTop      As Long
    Dim lScreenWidth    As Long
    Dim lScreenHeight   As Long
    
    TB_GetDesktopWorkArea lScreenLeft, lScreenTop, lScreenWidth, lScreenHeight
    With frm
        If lScreenLeft <> .Left Or lScreenTop <> .Top Or lScreenWidth <> .Width Or lScreenHeight <> .Height Then
            If IsMissing(vntAlreadyFull) Then
                If .WindowState <> vbNormal Then
                    .WindowState = vbNormal
                End If
                .Move lScreenLeft, lScreenTop, lScreenWidth, lScreenHeight
                TB_FillDesktop32 = True
            End If
        Else
            If Not IsMissing(vntAlreadyFull) Then
                vntAlreadyFull = True
            End If
        End If
    End With
End Function

Public Sub TB_WindowItemSize(frm As Form, mntres As String, Optional vntSkipCenter As Variant)
': for setting screen resolution wise
    Dim L, t, w, h
    Dim bSkipCenter As Boolean
    
    If Not IsMissing(vntSkipCenter) Then
        bSkipCenter = vntSkipCenter
    End If
    
    TB_CalcWindowSizes
    Select Case mntres
        Case "640x480"
            w = mlVGAWidth
            h = mlVGAHeight
        Case "800x600"
            w = ml800x600Width
            h = ml800x600Height
        Case "1024x768"
            w = ml1024x768Width
            h = ml1024x768Height
        Case "1152x864"
            w = ml1152x864Width
            h = ml1152x864Height
        Case "1280x768"
            w = gl1280x768Width
            h = gl1280x768Height
        Case "1280x1024"
            w = ml1280x1024Width
            h = ml1280x1024Height
        Case "1400x1050"
            w = ml1400x1050Width
            h = ml1400x1050Height
        Case "1600x1200"
            w = ml1600x1200Width
            h = ml1600x1200Height
            
        Case "1680x1050"        '         '24/12/2009
            w = ml1680x1050Width
            h = ml1680x1050Height
        Case "1792x1344"
            w = ml1792x1344Width
            h = ml1792x1344Height
        Case "1800x1440"
            w = ml1800x1440Width
            h = ml1800x1440Height
        Case "1856x1392"
            w = ml1856x1392Width
            h = ml1856x1392Height
        Case "1920x1080"
            w = ml1920x1080Width
            h = ml1920x1080Height
        Case "1920x1200"
            w = ml1920x1200Width
            h = ml1920x1200Height
        Case "1920x1400"
            w = ml1920x1400Width
            h = ml1920x1400Height
        Case "1920x1440"
            w = ml1920x1440Width
            h = ml1920x1440Height
        Case "2048x1152"
            w = ml2048x1152Width
            h = ml2048x1152Height
        Case "2048x1536"
            w = ml2048x1536Width
            h = ml2048x1536Height
            
    End Select
    With frm
        If .WindowState = vbMaximized Then
            .WindowState = vbNormal
        End If
        .Move .Left, .Top, w, h
    End With
    If Not bSkipCenter Then
        TB_CenterForm32 frm
    Else
        With frm
            L = .Left
            t = .Top
            w = .Width
            h = .Height
        End With
    End If
    On Error Resume Next
    frm.ActiveForm.FormResize
End Sub


Public Sub TB_CalcWindowSizes()
': for setting screen resolution wise
    Static bHerebefore As Boolean
    Dim lTWX As Long
    Dim lTWY As Long
    
    If bHerebefore Then Exit Sub        'only need to do once
    bHerebefore = True                  'only once

    lTWX = Screen.TwipsPerPixelX        'local variable, faster
    lTWY = Screen.TwipsPerPixelY

    ' Convert to Twips

    mlVGAWidth = 640 * lTWX
    mlVGAHeight = 480 * lTWY
    glVGAWidth = mlVGAWidth
    glVGAHeight = mlVGAHeight

    ml800x600Width = 800 * lTWX
    ml800x600Height = 600 * lTWY

    ml1024x768Width = 1024 * lTWX
    ml1024x768Height = 768 * lTWY

    ml1152x864Width = 1152 * lTWX
    ml1152x864Height = 864 * lTWY
    
    gl1280x768Width = 1280 * lTWX
    gl1280x768Height = 768 * lTWY

    ml1280x1024Width = 1280 * lTWX
    ml1280x1024Height = 1024 * lTWY

    ml1400x1050Width = 1400 * lTWX
    ml1400x1050Height = 1050 * lTWY

    ml1600x1200Width = 1600 * lTWX
    ml1600x1200Height = 1200 * lTWY
    
    ml1680x1050Width = 1680 * lTWX ' 24/Dec/2009
    ml1680x1050Height = 1050 * lTWX
    ml1792x1344Width = 1792 * lTWX
    ml1792x1344Height = 1344 * lTWX
    ml1800x1440Width = 1800 * lTWX
    ml1800x1440Height = 1440 * lTWX
    ml1856x1392Width = 1856 * lTWX
    ml1856x1392Height = 1392 * lTWX
    ml1920x1080Width = 1920 * lTWX
    ml1920x1080Height = 1080 * lTWX
    ml1920x1200Width = 1920 * lTWX
    ml1920x1200Height = 1200 * lTWX
    ml1920x1400Width = 1920 * lTWX
    ml1920x1400Height = 1400 * lTWX
    ml1920x1440Width = 1920 * lTWX
    ml1920x1440Height = 1440 * lTWX
    ml2048x1152Width = 2048 * lTWX
    ml2048x1152Height = 1152 * lTWX
    ml2048x1536Width = 2048 * lTWX
    ml2048x1536Height = 1536 * lTWX

End Sub


Public Function TB_GetDesktopWorkArea(lScreenLeft As Long, lScreenTop As Long, lScreenWidth As Long, lScreenHeight As Long) As Boolean
'
' Get the desktop work area using API SystemParametersInfo
    Const SPI_GETWORKAREA = 48
    Dim tDesktopArea As RECT
    
    SystemParametersInfo SPI_GETWORKAREA, 0, tDesktopArea, 0    'issue the API
    lScreenLeft = tDesktopArea.Left * Screen.TwipsPerPixelX
    lScreenTop = tDesktopArea.Top * Screen.TwipsPerPixelY
    lScreenWidth = (tDesktopArea.Right - tDesktopArea.Left) * Screen.TwipsPerPixelX
    lScreenHeight = (tDesktopArea.Bottom - tDesktopArea.Top) * Screen.TwipsPerPixelY
End Function


Public Function TB_CenterForm32(frm As Form, Optional vntOffsetLeft As Variant, Optional vntOffsetTop As Variant, Optional vntAlreadyCentered As Variant) As Boolean
': Center within the desktop work area
    Dim tDesktopArea    As RECT
    Dim lOffsetLeft     As Long
    Dim lOffsetTop      As Long
    Dim lScreenLeft     As Long
    Dim lScreenTop      As Long
    Dim lScreenWidth    As Long
    Dim lScreenHeight   As Long
    Dim lLeft   As Long
    Dim lTop    As Long
    Dim lWidth  As Long
    Dim lHeight As Long
    
    If Not IsMissing(vntOffsetLeft) Then
        lOffsetLeft = vntOffsetLeft
    End If
    If Not IsMissing(vntOffsetTop) Then
        lOffsetTop = vntOffsetTop
    End If
    
    TB_GetDesktopWorkArea lScreenLeft, lScreenTop, lScreenWidth, lScreenHeight
    
    lWidth = lScreenWidth - 100
    lHeight = lScreenHeight - 100
    lLeft = (lScreenWidth - frm.Width) \ 2 + lScreenLeft + lOffsetLeft
    lTop = (lScreenHeight - frm.Height) \ 2 + lScreenTop + lOffsetTop
   
    
    With frm
        If .WindowState = vbNormal Then
            If .Left <> lLeft Or .Top <> lTop Then
                If IsMissing(vntAlreadyCentered) Then
                    .Move lLeft, lTop
                    TB_CenterForm32 = True
                End If
            Else
                If Not IsMissing(vntAlreadyCentered) Then
                    vntAlreadyCentered = True
                End If
            End If
        End If
    End With
End Function

Public Sub TB_CenterFormInMDI(mdi As MDIForm, frm As Form)
': for setting reolution screen
    Dim L As Long, t As Long
    With frm
        If .WindowState = vbNormal Then
            L = (mdi.ScaleWidth - .Width) \ 2
            t = (mdi.ScaleHeight - .Height) \ 2
            If L < 0 Then
                L = 0
            End If
            If t < 0 Then
                t = 0
            End If
            .Move L, t
        End If
    End With
End Sub


Public Sub TB_FillFormInMDI(mdi As MDIForm, frm As Form)
': Fill the mdi work area with the form
    If frm.WindowState = vbNormal Then
        
        frm.Move 0, 0, mdi.ScaleWidth, mdi.ScaleHeight
    End If
End Sub

Public Sub Rpt_Intialization()
      ': Procedure for intializing the report
10    On Error GoTo ERR_P
             
20      Report.DiscardSavedData
'   If rptCon.State = 1 Then rptCon.Close
'    rptCon = conmain
'    rptCon.Open
      '*********************************************************************************************
30        If adrsCrep.State = 1 Then adrsCrep.Close
40        adrsCrep.Open empstr3, ConMain, adOpenStatic, adLockOptimistic


          If bytRepMode = 3 And (typOptIdx.bytMon = 1 Or typOptIdx.bytMon = 23) Then
            If SubLeaveFlag = 1 And typOptIdx.bytMon = 1 Then   ' 07-11
                Call UpdateLvCode
            Else
                Call UpdateLeaveCode
            End If
           End If
          
50        If blnGoWithExportInExcell = True Then
              Dim strFile As String
60            Select Case bytRepMode
                  Case 1      'daily
70                    strFile = bytRepMode & typOptIdx.bytDly
80                Case 2      'weekly
90                    strFile = bytRepMode & typOptIdx.bytWek
100               Case 3      'monthly
110                   strFile = bytRepMode & typOptIdx.bytMon
120               Case 5      'master
130                   strFile = bytRepMode & typOptIdx.bytMst
140               Case 4      'yearly
150                   strFile = bytRepMode & typOptIdx.bytYer
160               Case 6      'preriodic
170                   strFile = bytRepMode & typOptIdx.bytPer
180           End Select
190           Call ExportIntoFile(strFile, empstr3)
200       End If
210      If (typOptIdx.bytMon <> 1 And bytRepMode <> 5 And typOptIdx.bytMon <> 28 And typOptIdx.bytDly <> 16 And typOptIdx.bytMon <> 39 And typOptIdx.bytYer <> 11 And typOptIdx.bytMon <> 34 And typOptIdx.bytDly <> 30 And typOptIdx.bytDly <> 32 And typOptIdx.bytDly <> 33) Then  'Changes done by  for Absenteeism   28-03 '16  28-04
220       grpno = 3
230       For i = 1 To 7
240           If strAGrp(i) <> "" Or strAGrp(i) <> Null Then
250               Set crGrp(i) = Report.Database.Tables.Item(1).Fields.GetItemByName(strAGrp(i))
260               Report.Areas.Item(grpno).GroupConditionField = crGrp(i)
270               Report.Areas.Parent.FormulaFields(i).Text = strAlbl(i)
280           Else
                  
290               If typOptIdx.bytMon = 1 Or typOptIdx.bytYer = 1 Or typOptIdx.bytMon = 32 Or typOptIdx.bytMon = 37 Or typOptIdx.bytMon = 39 Or typOptIdx.bytPer = 21 Or typOptIdx.bytPer = 7 Or typOptIdx.bytPer = 30 Or typOptIdx.bytMon = 40 Or typOptIdx.bytPer = 40 Or typOptIdx.bytMon = 42 Or typOptIdx.bytDly = 27 Or typOptIdx.bytMon = 43 Or typOptIdx.bytMon = 49 Or typOptIdx.bytDly = 31 Or typOptIdx.bytDly = 32 Or typOptIdx.bytMon = 50 Then 'index 21 added by  25-06
300               Else
310                   Set crGrp(i) = Report.Database.Tables.Item(1).Fields.GetItemByName("Empcode")
                      Report.Areas.Item(grpno).GroupConditionField = crGrp(i)
330               End If
                  ''
340               Report.Areas.Item(grpno).Suppress = True
350           End If
360          grpno = grpno + 1
370       Next
380     End If
      '  Dim ii As Integer
      '  For ii = 1 To Report.Areas.Count - 1
      '    MsgBox Report.Areas(ii).Name
      '  Next
        
        ' For Master Reports Which are
390     If bytRepMode = 5 Then
400       If typOptIdx.bytMst = 0 Or typOptIdx.bytMst = 1 Or typOptIdx.bytMst = 2 Or typOptIdx.bytMst = 13 Or typOptIdx.bytMst = 14 Then    'Changes by
410         grpno = 3
420       For i = 1 To 7
430           If strAGrp(i) <> "" Or strAGrp(i) <> Null Then
440               Set crGrp(i) = Report.Database.Tables.Item(1).Fields.GetItemByName(strAGrp(i))
450               Report.Areas.Item(grpno).GroupConditionField = crGrp(i)
460               Report.Areas.Parent.FormulaFields(i).Text = strAlbl(i)
470           Else
480               Report.Areas.Item(grpno).Suppress = True
490           End If
500          grpno = grpno + 1
510       Next
520     End If
530   End If
      '*************************************************
            'use this part
540       If Not (adrsCrep.BOF And adrsCrep.EOF) = False Then
550           If adrsCrep.State = 1 Then adrsCrep.Close
560           Call SetMSF1Cap(10)
570           MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
580           Err.clear
590           blnIntz = False
600       Else

610           Report.Database.SetDataSource adrsCrep
620           blnIntz = True
630       End If
          
640       Exit Sub

ERR_P:
650         ShowError ("Report_Initialize : :  MDlCrep And Line:" & Erl)
660         blnIntz = False
670         Resume Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : UpdateLeaveCode
' DateTime  : 26/07/2008 10:44
' Author    :
' Purpose   : To Display Total of leave in monthly attendance report
    'from source code.
    'For that please refer modofied report also.
    'other wise u get error.
' Pre       :
' Post      :
' Return    : Variant
'---------------------------------------------------------------------------------------
'
Private Function UpdateLeaveCode()
          Dim walkerField As Integer
          Dim adrsTemp As ADODB.Recordset
             
10    On Error GoTo UpdateLeaveCode_Error
20        Set adrsTemp = OpenRecordSet("SELECT DISTINCT lvcode FROM Leavdesc " & _
          " WHERE lvcode NOT IN ('" & _
          pVStar.AbsCode & "','" & pVStar.PrsCode & "','" & _
          pVStar.WosCode & "','" & pVStar.HlsCode & "')")
30        Do While Not adrsTemp.EOF
40            walkerField = walkerField + 1
50            If walkerField > 10 Then Exit Function
60            Report.FormulaFields.GetItemByName("lv" & walkerField).Text = _
                  "'" & FilterNull(adrsTemp.Fields(0)) & "'"
              'Command name always refes cmd please do not change in report.
70            Report.FormulaFields.GetItemByName("Tlv" & walkerField).Text = _
                 "{cmd." & FilterNull(adrsTemp.Fields(0)) & "}"
80            adrsTemp.MoveNext
90        Loop
100   On Error GoTo 0
110   Exit Function
UpdateLeaveCode_Error:
120      If Erl = 0 Then
130         ShowError "Error in procedure UpdateLeaveCode of Module mdlCrep"
140      Else
150         ShowError "Error in procedure UpdateLeaveCode of Module mdlCrep And Line:" & Erl
160      End If
End Function
Private Function UpdateLvCode() 'Added by
    Dim i As Integer, Flag As Integer, cnt As Integer, f As Integer, j As Integer
    Dim adrsTemp As Recordset
    Dim LvArr As Variant
    Dim FixLvArr() As String, TmpLvArr() As String, LvCode() As String, Temp As String
    LvArr = Array("", "LW", "OD", "CL", "CO", "CM", "HP", "EN", "NE")
    f = 1
    
    On Error GoTo UpdateLvCode_Error
    Set adrsTemp = OpenRecordSet("SELECT DISTINCT lvcode FROM Leavdesc " & " WHERE lvcode NOT IN ('" & _
    pVStar.AbsCode & "','" & pVStar.PrsCode & "','" & pVStar.WosCode & "','" & pVStar.HlsCode & "')")
    
    If adrsTemp.RecordCount > 10 Then Exit Function
    cnt = adrsTemp.RecordCount
    ReDim FixLvArr(cnt)
    ReDim LvCode(cnt)
    i = 1
    Do While Not adrsTemp.EOF
        LvCode(i) = adrsTemp.Fields(0)
        i = i + 1
        adrsTemp.MoveNext
    Loop
    adrsTemp.MoveFirst
    For i = 1 To UBound(LvArr)
        For j = 1 To cnt
            If LvArr(i) = LvCode(j) Then
                Temp = LvCode(j)
                FixLvArr(f) = LvCode(j)
                f = f + 1
                Exit For
            End If
        Next
    Next
    Set adrsTemp = OpenRecordSet("SELECT DISTINCT lvcode FROM Leavdesc " & " WHERE lvcode NOT IN ('" & _
    pVStar.AbsCode & "','" & pVStar.PrsCode & "','" & pVStar.WosCode & "','" & pVStar.HlsCode & "','OD','CL','CO','CM','HP','SL','EN','NE','EL','LW')")
    Do While Not adrsTemp.EOF
       strLvCode = strLvCode & "," & adrsTemp.Fields(0)
       adrsTemp.MoveNext
    Loop
    TmpLvArr = Split(strLvCode, ",")
    For i = 1 To UBound(TmpLvArr)
        FixLvArr(f) = TmpLvArr(i)
        f = f + 1
    Next
For i = 1 To UBound(FixLvArr)
    If UBound(FixLvArr) > 10 Then Exit Function
        If FixLvArr(i) <> "" Then
            Report.FormulaFields.GetItemByName("lv" & i).Text = _
                  "'" & FixLvArr(i) & "'"
              'Command name always refes cmd please do not change in report.
            Report.FormulaFields.GetItemByName("Tlv" & i).Text = _
                 "{cmd." & FixLvArr(i) & "}"
        End If
Next
100   On Error GoTo 0
110   Exit Function
UpdateLvCode_Error:
120      If Erl = 0 Then
130         ShowError "Error in procedure UpdateLvCode of Module mdlCrep"
140      Else
150         ShowError "Error in procedure UpdateLvCode of Module mdlCrep And Line:" & Erl
160      End If
End Function


