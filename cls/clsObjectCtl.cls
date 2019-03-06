VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjectCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'设置控件是否可用：
'vCtlName：控件名称
'vBool：TRUE为可用，FALSE为不可用
'注意：只可用于有Enabled属性的控件
Public Sub setCtlEnabled(ByRef vFrm As Form, ByVal vCtlName As String, ByVal vBool As Boolean)
    Dim oCtl As Object
    vCtlName = Trim$(vCtlName)
    
    
    Set oCtl = getCtl(vFrm, vCtlName)
    If Len(oCtl.name) <= 0 Then
        Exit Sub
    End If
    
    oCtl.Enabled = vBool
End Sub

'根据控件名称找到FORM中的控件对象并且返回
'支持控件数组
Public Function getCtl(ByRef vFrm As Form, ByVal vCtlName As String) As Object
    Dim oCtl As Object
    Dim lLeft As Long
    Dim lRight As Long
    Dim lIndex As Long
    Dim szCtlName As String
    
    lLeft = InStr(1, vCtlName, "(")
    If lLeft > 0 Then
        lRight = InStr(1, vCtlName, ")")
        lIndex = Mid$(vCtlName, lLeft + 1, lRight - lLeft - 1)
        
        szCtlName = Left$(vCtlName, lLeft - 1)
    Else
        szCtlName = vCtlName
    End If
    For Each oCtl In vFrm.Controls
        If oCtl.name = szCtlName Then
            If lLeft > 0 Then  '是控件数组
                If oCtl.Index = lIndex Then
                    Set getCtl = oCtl
                    Exit Function
                End If
            Else
                Set getCtl = oCtl
                Exit Function
            End If
        End If
    Next
End Function

'根据控件的Tag属性找到FORM中的控件对象并且返回
'支持控件数组
'返回值在使用的时候必须判断nothing之后再使用
Public Function getCtl_ByTag(ByRef vFrm As Form, ByVal vTagStr As String) As Object
    On Error Resume Next
    Dim oCtl As Object
    
    Set getCtl_ByTag = Nothing
    For Each oCtl In vFrm.Controls
        If Trim(oCtl.Tag) = vTagStr Then
            Set getCtl_ByTag = oCtl
            Exit Function
        End If
    Next
End Function

Public Function GetValue(ByRef vCtl) As String
    Select Case TypeName(vCtl)
        Case "CheckBox"
            GetValue = Abs(vCtl.Value)
        Case "TextBox"
            GetValue = vCtl.Text
        Case Else
            GetValue = vCtl.Text
    End Select
End Function


