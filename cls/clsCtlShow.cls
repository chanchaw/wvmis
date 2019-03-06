VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtlShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'保持属性值的局部变量
Private mvarObjectID As String '局部复制
'保持属性值的局部变量
Private mvarFieldType As Integer '局部复制

Private rsField As New Recordset
Private rsCtl As New Recordset
Private rsBand As New Recordset

Public WithEvents AB1 As ActiveBar2
Attribute AB1.VB_VarHelpID = -1
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


'Public Property Set rsCtlPara(ByVal vData As Recordset)
'    Set rsCtl = vData.Clone
'End Property
'
'Public Property Get rsCtlPara() As Recordset
'    Set rsCtlPara = rsCtl.Clone
'End Property



Public Property Let ObjectID(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ObjectID = 5
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ObjectID
    ObjectID = mvarObjectID
End Property

Public Sub InitClass(ByRef o As Object, ByVal m_FieldType As Integer)
    mvarFieldType = m_FieldType
    InitCtlField
    Set AB1 = o
End Sub

'-----以下为网格显示部件
Private Sub InitCtlField()
    Dim strSQL As String
    Dim strSQL1 As String
    Dim strSQL2 As String
    Dim i As Long
    Dim rs As New Recordset
    
    Set rs = New Recordset
    
    
    Gm.cnnTool.IniConnection8DM Gm.SysID.DBInfo
    
    Select Case mvarFieldType
        Case 1
            strSQL = "Select * From G_BLSFormTools Where B_ObjectID='" & Trim(mvarObjectID) & "' Order By B_BandIndex,B_ToolIndex"
            
            strSQL1 = "Select * From G_ItemObject Where B_ObjectID In ("
            strSQL1 = strSQL1 & "Select B_SubObject From G_BLSFormTools Where B_ObjectID='" & Trim(mvarObjectID) & "'"
            strSQL1 = strSQL1 & ")"
            
            strSQL2 = "Select  B_ItemID,B_BandIndex,B_DockingArea From G_BLSFormTools Where B_ObjectID='" & Trim(mvarObjectID) & "' Order By B_BandIndex,B_ToolIndex"
        Case 2
            strSQL = "Select * From G_BLFormTools Where B_ObjectID='" & Trim(mvarObjectID) & "' Order By B_BandIndex,B_ToolIndex"
            
            strSQL1 = "Select * From G_ItemObject Where B_ObjectID In ("
            strSQL1 = strSQL1 & "Select B_SubObject From G_BLFormTools Where B_ObjectID='" & Trim(mvarObjectID) & "'"
            strSQL1 = strSQL1 & ")"
            
            strSQL2 = "Select  B_ItemID,B_BandIndex,B_DockingArea From G_BLFormTools Where B_ObjectID='" & Trim(mvarObjectID) & "' Order By B_BandIndex,B_ToolIndex"
        Case 3
            strSQL = "Select * From G_BLRFormTools Where B_ObjectID='" & Trim(mvarObjectID) & "' Order By B_BandIndex,B_ToolIndex"
            strSQL1 = "Select * From G_ItemObject Where B_ObjectID In ("
            strSQL1 = strSQL1 & "Select B_SubObject From G_BLRFormTools Where B_ObjectID='" & Trim(mvarObjectID) & "'"
            strSQL1 = strSQL1 & ")"
            
            strSQL2 = "Select  B_ItemID,B_BandIndex,B_DockingArea From G_BLRFormTools Where B_ObjectID='" & Trim(mvarObjectID) & "' Order By B_BandIndex,B_ToolIndex"
    End Select
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    
    Set rsField = New Recordset
    '建立结构
    FillUnConnectRecordSet rs, rsField
    
    rs.Close
    Set rs = Nothing
    rs.Open strSQL1, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    Set rsCtl = New Recordset
    '建立结构
    FillUnConnectRecordSet rs, rsCtl
    
    rs.Close
    Set rs = Nothing
    rs.Open strSQL2, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    Set rsBand = New Recordset
    '建立结构
    FillUnConnectRecordSet rs, rsBand

    rs.Close
    Set rs = Nothing
End Sub

Public Sub Refresh()
    FillCtl
End Sub

'清除建立的控件
Private Sub ClearCtl(ByRef o As Object)
    On Error GoTo IFERR
    Dim ctl As Object
    For Each ctl In o
        If Mid(ctl.name, 1, 2) = "B_" Then
            o.Controls.Remove (ctl.name)
        End If
    Next
    
    Exit Sub
IFERR:
    MsgBox "在清空页面中的控件时出错" & vbNewLine & Err.Description, vbOKOnly + vbInformation, "提示"
End Sub

'主表上添加控件
Private Sub FillCtl()
    On Error Resume Next
    Dim frm1 As Object
    Dim oControl As Object

    Set frm1 = AB1.Parent
    
    ClearCtl frm1
    
    With AB1
        .Bands.RemoveAll
        .Tools.RemoveAll
        .RecalcLayout
    
        Dim oBand As band
        Dim OldBandName As String
        Dim BandName As String
    
        OldBandName = ""
        '建立 Band
        rsBand.MoveFirst
        Do While Not rsBand.EOF
            '增加Bands
            BandName = "栏目" & Trim(rsBand("B_BandIndex"))
            If OldBandName <> BandName Then
                Set oBand = Nothing
                Set oBand = .Bands.Add(BandName)
                With oBand
                    .Type = 0
                    .DisplayMoreToolsButton = False
                    .flags = 895
                    .GrabHandleStyle = ddGSNone
                    .AutoSizeForms = True
                    .DockLine = IIf(IsNull(rsBand("B_BandIndex")), 0, rsBand("B_BandIndex"))
                    .DockingArea = rsBand("B_DockingArea")
                End With
                
                OldBandName = BandName
                rsBand.MoveNext
            Else
                rsBand.MoveNext
            End If
        Loop
        
        Dim oTool As Tool
        rsField.MoveFirst
        '建立 Tools
        Do While Not rsField.EOF
            Set oTool = .Bands("栏目" & Trim(rsField("B_BandIndex"))).Tools.Add(rsField("B_ItemID"), rsField("B_CtlName"))
            
            oTool.Caption = rsField("B_FieldName")
            oTool.Style = 0
            oTool.ControlType = 6
            oTool.Width = rsField("B_Width")
            oTool.Height = rsField("B_Height")
            
            rsCtl.Filter = "B_ObjectID='" & rsField("B_SubObject") & "'"
            If Not rsCtl.EOF Then
                Set oControl = frm1.Controls.Add(rsCtl("B_CtlType"), Trim(rsField("B_CtlName")))
                
                ControlParaSet oControl, rsCtl("B_ObjectID")
                oTool.hwnd = oControl.hwnd
            End If
            rsField.MoveNext
        Loop
        .RecalcLayout
    End With
End Sub

Private Sub FillUnConnectRecordSet(ByRef sRs As Recordset, ByRef tRs As Recordset)
    On Error Resume Next
    Dim i As Long
    
    Set tRs = New Recordset
    For i = 0 To sRs.Fields.Count - 1
        tRs.Fields.Append sRs.Fields(i).name, sRs.Fields(i).Type, sRs.Fields(i).DefinedSize, sRs.Fields(i).Attributes
    Next
    
    tRs.Open
    Do While Not sRs.EOF
        tRs.AddNew
        For i = 0 To sRs.Fields.Count - 1
            tRs.Fields(i).Value = IIf(IsNull(sRs.Fields(i).Value), "", sRs.Fields(i).Value)
        Next
        tRs.Update
        sRs.MoveNext
    Loop
    
End Sub

Public Sub LoadObject(ByRef rs As Recordset)
    On Error Resume Next
    If rs Is Nothing Then
        Exit Sub
    End If
    
    Dim i As Integer
    Dim sFieldName As String
    Dim o As Object
    
    
    For Each o In AB1.Parent.Controls
        If Mid(o.name, 1, 2) = "B_" Then
            sFieldName = o.name
            '如果找到数据
            If rs.Fields(sFieldName).name = o.name Then
                Debug.Print o.name & "=" & o.Text
                Select Case rs.Fields(sFieldName).Type
                    Case adVarChar, adChar
                        o.Text = ""
                        o.Text = IIf(IsNull(rs.Fields(sFieldName).Value), "", rs.Fields(sFieldName).Value)       ' 1--String型
                    Case adInteger, adSingle, adSmallInt, adDouble, adNumeric
                        o.Text = 0
                        o.Text = IIf(IsNull(rs.Fields(sFieldName).Value), 0, rs.Fields(sFieldName).Value)   ' 2--数值型
                    Case adDate, adDBDate, adDBTime, adDBTimeStamp
                        o.Text = IIf(IsNull(rs.Fields(sFieldName).Value), Format(Now, "YYYY-MM-DD"), Format(rs.Fields(sFieldName).Value, "YYYY-MM-DD"))   ' 3--日期型
                    Case Else
                        o.Text = ""
                        o.Text = IIf(IsNull(rs.Fields(sFieldName).Value), "", rs.Fields(sFieldName).Value)       ' 1--String型
                End Select
                Debug.Print o.Text
            End If
        End If




        '单选控件的数据赋值
        '=================================
        Select Case TypeName(o)
        
            Case "OptionButton"
                If Len(Trim(o.Tag)) > 0 Then
                    If o.Caption = rs(o.Tag).Value Then
                        o.Value = True
                    End If
                End If
            Case "CheckBox"
                o.Value = Abs(IIf(IsNull(rs.Fields(sFieldName).Value), "", rs.Fields(sFieldName).Value))
        End Select
        '=================================
    Next

End Sub

Public Sub GetObject(ByVal sCtlName As String, ByRef o As Object)
    On Error Resume Next
    Dim oCtl As Object
    For Each oCtl In AB1.Parent.Controls
        If oCtl.name = sCtlName Then
            Set o = oCtl
            Exit Sub
        End If
    Next
End Sub

Public Function SaveObject(ByRef rs As ADODB.Recordset) As Boolean


'Public Function SaveObject(ByRef rs As RecordSet) As Boolean

    On Error Resume Next
    'On Error GoTo IFERR
    If rs Is Nothing Then
        SaveObject = False
        Exit Function
    End If
    

    Dim i As Integer
    Dim o As Object
    Dim sFieldName As String
    
    Dim cls1 As New clsObjectCtl
    
    
    
    '为单据日期添加时间
    '======================
    Dim m_Time As String
    m_Time = Format(Now, "HH:MM:SS")
    '======================
    
    
    'rs.MoveFirst
    For Each o In AB1.Parent.Controls
        If Mid(o.name, 1, 2) = "B_" Then
            '如果找到数据
            'Debug.Print o.name
            sFieldName = Trim(o.name)
            If IsInField(sFieldName, rs) = True Then
                
                If Trim(rs.Fields(sFieldName).name) = Trim(o.name) Then
                    Debug.Print o.name
                    Select Case rs.Fields(sFieldName).Type
                        Case adVarChar, adChar
                            rs.Fields(sFieldName).Value = IIf(IsNull(o.Text), "", o.Text)      ' 1--String型
                        Case adInteger, adSingle, adSmallInt, adDouble, adNumeric
                            rs.Fields(sFieldName).Value = cls1.GetValue(o)
'                            If Len(o.Text) <= 0 Then
'                                rs.Fields(sFieldName).Value = 0
'                            Else
'                                rs.Fields(sFieldName).Value = Val(o.Text)
'                            End If
                            'rs.Fields(sFieldName).Value = IIf(IsNull(o.Text), 0, o.Text)   ' 2--数值型
                        Case adDate, adDBDate, adDBTime, adDBTimeStamp
                            rs.Fields(sFieldName).Value = IIf(IsNull(o.Text), Format(Now, "YYYY-MM-DD HH:MM:SS"), Format(o.Text & m_Time, "YYYY-MM-DD HH:MM:SS"))   ' 3--日期型
                        Case Else
                            rs.Fields(sFieldName).Value = IIf(IsNull(o.Text), "", o.Text)
                    End Select
                    'Debug.Print rs.Fields(sFieldName).Value
                    'rs.Update
                End If
            End If
        End If
        
        
        
        '以下为单选控件数据的保存
        '=================================
'        If TypeName(o) = "OptionButton" Then
'            If Len(Trim(o.Tag)) > 0 Then
'                If o.Value = True Then
'                    rs.Fields(Trim(o.Tag)).Value = o.Caption
'                End If
'            End If
'        End If
        
        Select Case TypeName(o)
            Case "OptionButton"
                If Len(Trim(o.Tag)) > 0 Then
                    If o.Value = True Then
                        rs.Fields(Trim(o.Tag)).Value = o.Caption
                    End If
                End If
            Case "CheckBox"
                rs.Fields(Trim(o.name)).Value = Abs(o.Value)
        End Select
        '=================================
        
    Next

    rs.Update
'
'    Exit Function
'IFERR:
'    Dim szErr As String
'    szErr = "在保存主表字段时出错！" & vbNewLine & Err.Description
'    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Function


Private Function IsInField(ByVal ObjectName As String, ByRef rs As Recordset) As Boolean
    Dim i As Long
    
    IsInField = False
    For i = 0 To rs.Fields.Count - 1
        If ObjectName = rs.Fields(i).name Then
            IsInField = True
        End If
    Next
End Function


Public Sub RemoveAll()
    On Error Resume Next
    Dim o As Object
    For Each o In AB1.Parent.Controls
        If Mid(o.name, 1, 2) = "B_" Then
            'SetParent o.hWnd, AB1.Parent.hWnd
            AB1.Parent.Controls.Remove (o.name)
        End If
    Next
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set AB1 = Nothing
End Sub



'2012-2-7添加
Private Function GetFrmName(ByVal m_ObjectID) As String
    Dim strSQL As String
    Dim rs As Recordset
    
    GetFrmName = ""
    Select Case Mid(m_ObjectID, 1, 2)
        Case "11" '单表
            strSQL = "Select * From G_BLS Where B_ObjectID='" & m_ObjectID & "'"
        Case "12" '单据
            strSQL = "Select * From G_BL Where B_ObjectID='" & ObjectID & "'"
        Case "19" '自定义对象
            strSQL = "Select * From G_CustomObject Where B_ObjectID='" & ObjectID & "'"
    End Select
    
    
    Set rs = New Recordset
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        GetFrmName = rs("B_FormName")
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function


'返回TRUE表示可以保存，否则为不可保存
Public Function JudgeBeforeSave() As Boolean
    'On Error GoTo IFERR
    On Error Resume Next
    JudgeBeforeSave = True
    
    Dim rsCtlset As New Recordset
    Dim lNotNull As Long  '不为空
    Dim lGreaterZ As Long  '大于0
    Dim szValue As String
    Dim szTip As String
    Dim sFieldName As String
    Dim o As Object
    
    Set rsCtlset = rsField.Clone
    
    For Each o In AB1.Parent.Controls
        If Mid(o.name, 1, 2) = "B_" Then
            '如果找到数据
            sFieldName = o.name
            Debug.Print sFieldName
            
            rsCtlset.Filter = " B_FieldName='" & sFieldName & "'"
            lNotNull = Abs(IIf(IsNull(rsCtlset!B_NotNull), 0, rsCtlset!B_NotNull))
            lGreaterZ = Abs(IIf(IsNull(rsCtlset!B_GreaterZ), 0, rsCtlset!B_GreaterZ))
            
            '不为空的判断
            szValue = o.Text
            If lNotNull = 1 Then
                If Len(Trim$(szValue)) <= 0 Then
                    JudgeBeforeSave = False
                    szTip = rsCtlset!B_Caption & "不可为空！"
                    rsCtlset.Close
                    Set rsCtlset = Nothing
                    MsgBox szTip, vbOKOnly + vbInformation, "提示"
                    Exit Function
                End If
            End If
            
            '大于0的判断
            If lGreaterZ = 1 Then
                If szValue <= 0 Then
                    JudgeBeforeSave = False
                    szTip = rsCtlset!B_Caption & "不可为空！"
                    rsCtlset.Close
                    Set rsCtlset = Nothing
                    MsgBox szTip, vbOKOnly + vbInformation, "提示"
                    Exit Function
                End If
            End If
        End If
    Next

'    Exit Function
'IFERR:
'    Dim szErr As String
'    szErr = "保存时验证主表字段有未知的非法" & vbNewLine & Err.Description
'    MsgBox szErr, vbOKOnly + vbInformation, "提示"

End Function


Private Sub SetListBoxDefault(ByRef oCtl As Object)
    Dim strSQL As String
    Dim rs As Recordset
    
    Select Case TypeName(oCtl)
        Case "TA_UCCommonCtls.UCListBox"
            strSQL = oCtl.SQL
            Set rs = New Recordset
            rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
            oCtl.Text = rs(0)
            rs.Close
            Set rs = Nothing
        Case "UCListBox"
            strSQL = oCtl.SQL
            Set rs = New Recordset
            Debug.Print strSQL
            rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
            oCtl.Text = rs(0)
            rs.Close
            Set rs = Nothing
    End Select
End Sub

'设置控件的表现样式
Private Sub SetCtlStyle(ByRef vCtl As Object, ByVal vCtlType As String)
    On Error Resume Next
    'rsField   单据对象的主表中字段控件的设置
    Select Case vCtlType
        Case "VB.TextBox"
            vCtl.BackColor = &HCEDFDE
            
            '先设置默认，再根据配置设置对应
            vCtl.Locked = 0
            vCtl.BorderStyle = 0
            vCtl.Font.Size = 9
            
            
            vCtl.Locked = Abs(IIf(IsNull(rsField!B_Locked), 0, rsField!B_Locked))
            vCtl.BorderStyle = Abs(IIf(IsNull(rsField!B_BorderStyle), 0, rsField!B_BorderStyle))
            vCtl.Font.Size = IIf(IsNull(rsField!B_FontSize), 9, rsField!B_FontSize) '默认字体大小为9
            vCtl.ForeColor = Val(rsField!B_ForeColor)
        Case "VB.CheckBox"
            vCtl.Alignment = Val(IIf(IsNull(rsField!B_Alignment), 0, rsField!B_Alignment))
            vCtl.BackColor = &HCEDFDE
            vCtl.ForeColor = Val(rsField!B_ForeColor)
            vCtl.Font.Size = IIf(IsNull(rsField!B_FontSize), 9, rsField!B_FontSize) '默认字体大小为9
    End Select
End Sub


'对自动填充的控件参数进行设置
'目的为适应非TA类型的控件
Private Sub ControlParaSet(ByRef oControl As Object, ByVal m_ObjectID As String)
    On Error Resume Next
    Dim strSQL As String
    Dim rs As Recordset
    
    Set rs = New Recordset
    strSQL = "Select * From G_ItemObject where B_ObjectID='" & m_ObjectID & "'"
    'Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    
    If IIf(IsNull(rs("B_TANot")), 0, rs("B_TANot")) = 1 Then
        With oControl
            .Width = rsField("B_Width")
            .Height = rsField("B_Height")
            .Visible = True
        End With
        
        '将TextBox设置为LABEL的显示样式
        Select Case rs!B_CtlType
            Case "VB.TextBox"
                oControl.Text = rs!B_Caption
            Case "VB.CheckBox"
                oControl.Caption = rs!B_Caption
        End Select
        
        '设置表现样式
        '不同的控件需要设置的表现样式的属性不同
        SetCtlStyle oControl, rs!B_CtlType
    Else
        With oControl
            .Width = rsField("B_Width")
            .Height = rsField("B_Height")
            .EdgeHeight = IIf(IsNull(rsField("B_EdgeHeight")), 0, rsField("B_EdgeHeight"))
            
            
            .ConnectionString = Gm.cnnTool.cnnStr
            Debug.Print Gm.cnnTool.cnnStr
            .SQL = rsCtl("B_SQL")
            .Caption = rsField("B_Caption")
            .FieldName = rsField("B_FieldName")
            .SubMenuObject = rsCtl("B_SubMenuObject")
            .DefaultValue = rsCtl("B_DefaultValue")
            
            .Visible = True
            
            If IIf(IsNull(rsField("B_Enabled")), 0, rsField("B_Enabled")) <> 1 Then
                .Enabled = True
            Else
                .Enabled = False
            End If
            '.Text = ""
            .Refresh
        End With
        
        '判断如果当前创建的控件的类型是TA_UCCommonCtls.UCListBox
        '那么自动默认下拉数据的第一个项目
        'SetListBoxDefault oControl
    End If
    
    rs.Close
    Set rs = Nothing
End Sub


Public Function GetCtlPara_BLS(ByVal vObjectID As String) As Recordset
    Dim rs As New Recordset
    strSQL = "SELECT * FROM G_BLSFormTools WHERE B_ObjectID='" & vObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Set GetCtlPara_BLS = rs.Clone
    rs.Close
    Set rs = Nothing
End Function
