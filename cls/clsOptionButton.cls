VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'清除所有OptionButton的Value
Public Sub ClareOptionValue(ByRef frm1 As Object)
    Dim o As Control
    
    For Each o In frm1.Controls
        If TypeName(o) = "OptionButton" Then
            o.Value = False
        End If
    Next
    
End Sub
