VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCJComboLinker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private rs As RecordSet
Private szDisplayField As String
Private szValueField As String
Private oCombo As XtremeSuiteControls.ComboBox

Public Sub InitCls(ByVal vCombo As XtremeSuiteControls.ComboBox, ByVal vDisplayField As String, _
    ByVal vValueField As String, ByRef vRs As RecordSet)
    
    Set oCombo = vCombo
    Set rs = vRs.Clone
    szValueField = vValueField
    szDisplayField = vDisplayField
    
    '获取初始化的变量后，立即绑定到UI控件上
    Bind
End Sub

Private Sub Bind()
    If rs.RecordCount <= 0 Then
        Exit Sub
    End If
    
    oCombo.Clear
    'oCombo.AddItem ""
    Do While Not rs.EOF
        oCombo.AddItem rs(Trim$(szDisplayField))
        rs.MoveNext
    Loop

    If rs.EOF = True Then
        rs.MoveLast
    End If
End Sub

Public Function GetData(ByVal vDisplayValue As String) As String
    If Len(vDisplayValue) <= 0 Then
        GetData = ""
        Exit Function
    End If
    

    Dim szFilterString As String
    szFilterString = " [" & szDisplayField & "]='" & vDisplayValue & "'"
 
    rs.Filter = szFilterString
    
    GetData = rs(Trim$(szValueField))
End Function

  
Public Function SetComboText(ByRef MyCombo As ComboBox, ByVal oriStr As String) As Long
    Dim i As Long
       
    For i = 0 To MyCombo.ListCount - 1
        If MyCombo.List(i) = oriStr Then
            SetComboText = i
            Exit Function
        End If
    Next
       
    SetComboText = -1
End Function

