VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public sngTimeIn As Single      '' Arrival Time
Public sngTimeOut As Single     '' Departure Time
Public sngBreakS As Single      '' Break Start
Public sngBreakE As Single      '' Break End
Public sngBreakHrs As Single    '' Break Hours
Public sngTime5 As Single       '' Time 5
Public sngTime6 As Single       ''
Public sngTime7 As Single       ''
Public sngTime8 As Single       '' Till Time 8
Public sngODFrom As Single      '' OT From
Public sngODTo As Single        '' OT From To
Public sngOFDFrom As Single     '' Official Duty From
Public sngOFDTo As Single       '' Official Duty To
Public sngOFDHrs As Single      '' Official Duty Hours
