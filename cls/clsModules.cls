VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'控制模块和TABS的显示

Private strSQL As String


'根据模块ID显示对应的模块窗体
Public Sub ShowModule(ByVal vModuleID As String)
    Dim rs As Recordset
    strSQL = "exec dbo.usp_GetSpeModule '" & vModuleID & "'"
    Debug.Print strSQL
    Set rs = New Recordset
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        MsgBox "在配置文件中未能找到模块[" & vModuleID & "]对应的页面！", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    
    Dim szFrmName As String
    szFrmName = IIf(IsNull(rs!B_ObjectID), "", rs!B_ObjectID)
    rs.Close
    Set rs = Nothing
    
    If Len(szFrmName) <= 0 Then
        MsgBox "在配置文件中未能正确配置[" & vModuleID & "]页面名称！", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    
    Dim frm1 As Object
    Set frm1 = Gm.CacheFrms.GetFrm(szFrmName)

    frm1.Show

    BringWindow2Top frm1.hwnd
End Sub


Public Sub ShowTabs(ByVal vModuleID As String, ByRef vC1Tab As C1Tab)
    Dim rs As New Recordset
    strSQL = "exec dbo.usp_GetModuleTabs '" & vModuleID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    If vC1Tab Is Nothing Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    Dim lTabIndex As Long
    Dim lVisible As Long
    Dim lSum As Long
    
    Do While Not rs.EOF
        lTabIndex = IIf(IsNull(rs!B_TabIndex), 0, rs!B_TabIndex)
        lVisible = IIf(IsNull(rs!B_Visible), 0, rs!B_Visible)
        vC1Tab.TabVisible(lTabIndex) = lVisible
        
        lSum = lSum + lVisible
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    
    If lSum <= 0 Then
        vC1Tab.TabHeight = 0
    End If
End Sub


Public Sub ShowButton(ByVal vModID As String, ByRef vFrm As Form, ByVal vCtlName As String)
    Dim rs As New Recordset
    strSQL = "SELECT * FROM G_BusinessModulesBTV WHERE B_ModID='" & vModID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    
    Dim o As Object
    For Each o In vFrm.Controls
        If o.name = vCtlName Then
        rs.Filter = " B_ButtonCaption='" & o.Caption & "'"
        If rs.RecordCount > 0 Then
            If rs!B_UnVisible = 1 Then
                o.Visible = False
            Else
                o.Visible = True
            End If
        End If
        End If
    Next
    
    rs.Close
    Set rs = Nothing

End Sub

'隐藏左侧模块按钮
Public Sub HideLeftMod(ByRef vFrm As Form)
    On Error Resume Next
    Dim rs As New Recordset
    strSQL = "SELECT * FROM G_BusinessModules"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    Dim lUnVisible As Long
    Dim lVisible As Long
    Dim o As Object
    Dim szTag As String
    
    
    For Each o In vFrm.Controls
        szTag = o.Tag
        rs.Filter = " B_ModID='" & szTag & "'"
        If rs.RecordCount > 0 Then
            lUnVisible = IIf(IsNull(rs!B_UnVisible), 0, rs!B_UnVisible)
            If lUnVisible = 0 Then
                lVisible = 1
            End If
            If lUnVisible = 1 Then
                lVisible = 0
            End If
            o.Visible = lVisible
        End If
    Next
End Sub
