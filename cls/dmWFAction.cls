VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dmWFAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'工作流中动作的实体类对象
'Private A_Saved2Audit As String
'Private A_Audited2Print As String
'Private A_Printed As String

Public Property Get Saved2Audit() As String
    Saved2Audit = "Saved2Audit"
End Property

Public Property Get Audited2Print() As String
    Audited2Print = "Audited2Print"
End Property

Public Property Get Printed() As String
    Printed = "Printed"
End Property

