VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Begin VB.Form frmCRV 
   ClientHeight    =   5460
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   7365
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CRV 
      CausesValidation=   0   'False
      Height          =   5280
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7200
      _cx             =   12700
      _cy             =   9313
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1033
   End
End
Attribute VB_Name = "frmCRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
      'Added by  to change the screen resolutuion according to the monitor
      '*************************************************************
10    On Error GoTo Form_Load_Error
          'This Code Add by  MIS2007DF024
40        Me.WindowState = vbMaximized
50        CRV.Height = Screen.Height - 1000
60        CRV.Width = Screen.Width
           
      '*****************************************************
70    On Error GoTo 0
80    Exit Sub
Form_Load_Error:
90       If Erl = 0 Then
100         ShowError "Error in procedure Form_Load of Form frmCRV"
110      Else
120         ShowError "Error in procedure Form_Load of Form frmCRV And Line:" & Erl
130      End If
End Sub


Private Sub Form_Resize()
    CRV.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
