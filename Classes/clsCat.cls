VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public strCat As String     '' Category Code
Public strDesc As String    '' Category Description
Public sngLate As Single    '' if Late More than
Public sngEarl As Single    '' if Early More than
Public sngLateI As Single   '' Ignore Late
Public sngEarlI As Single   '' Ignore Early
Public sngCutL As Single    '' Cut If Late than
Public sngCutE As Single    '' Cut If Early than


Public sngLunchLtIgnore As Single   '' Ignore Lunch Late
