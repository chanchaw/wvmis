VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGridShow"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public WithEvents TDBGrid1 As TrueOleDBGrid80.TDBGrid
Attribute TDBGrid1.VB_VarHelpID = -1

'0:普通
'1:Radio Button
'2:Combo Box
'3:Sorted Combo Box
'4:Check Box
'当前显示
'11:颜色                 2
'12:打开文件(文件全名)   2
'13:打开目录             2
'14:打开文件(只有文件名) 2
'15:编辑 TXT 字段        2
'16:编辑 RFT 字段        2
'17:编辑 SQL 字段        2
'18:单选字段 (不分值与显示)  2
'19:单选字段 (分值与显示)  2
'20:表选字段 (字段在表中)2
'21:弹出网格             2

'99 自定义控件

Private rsField As New RecordSet
'保持属性值的局部变量
Private mvarObjectID As String '局部复制
Private mvarFieldType As Integer '局部复制

Private o(0 To 10) As Object
Private aGrid(2, 10)
Private iNowCtl As Integer
Private iMax As Integer


'Public TDBDropDown1 As TrueOleDBGrid80.TDBDropDown
Public WithEvents TDBDropDown1 As TrueOleDBGrid80.TDBDropDown
Attribute TDBDropDown1.VB_VarHelpID = -1

Dim rsdetail As New RecordSet
Dim m_DropDownGridSQL As String
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Type POINTAPI
   X As Long
   Y As Long
End Type

Dim iLeft As Long
Dim iTop As Long

'vFieldName：当前下拉控件所在的列的字段
'sWhereStr：传址，将事件下的and子句带入下拉的数据源SQL中
'Public Event OnSetDropDownGrid(ByVal vFieldName As String, ByRef sWhereStr As String)
Public Event OnSetDropDownGrid(ByRef sWhereStr As String)

Public Event OnTDBDropDownClose() '在2017年6月11日添加

Public m_PopUp As Boolean
Public fatherFrm As Object

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

'--显示颜色
Private Function SelectColor() As Long
    On Error Resume Next
    Dim CommonDialog1

    Set CommonDialog1 = CreateObject("MSComDlg.CommonDialog")
    With CommonDialog1
        .ShowColor
        SelectColor = .color
    End With
    Set CommonDialog1 = Nothing
End Function

'---打开目录
Private Function SelectDir() As String
    On Error Resume Next
    Dim o, o1, o2
    Set o = CreateObject("Shell.Application")
    Set o2 = o.BrowseForFolder(0, "选择文件夹", &H400, 0)
    Set o1 = o2.Items().item()
    SelectDir = o1.Path
    
    Set o = Nothing
    Set o2 = Nothing
    Set o1 = Nothing
End Function

'---打开文件
Private Function SelectFullFile() As String
    Dim CommonDialog1

    Set CommonDialog1 = CreateObject("MSComDlg.CommonDialog")
    With CommonDialog1
        .ShowOpen
        SelectFullFile = .FileName
    End With
    Set CommonDialog1 = Nothing
End Function

'--打开文件(只有文件名)
Private Function SelectFile() As String
    SelectFile = GetFileName(SelectFullFile)
End Function

'----取得文件名
Private Function GetFileName(ByVal sFilename As String) As String
    '取得目录
    If Len(sFilename) < 1 Then
        GetFileName = ""
        Exit Function
    End If
    Dim mIntStrLen As Integer
    Dim mStrTmp As String
    sFilename = Trim(sFilename)

    mIntStrLen = InStrRev(sFilename, "\", Len(Trim(sFilename)), vbTextCompare)
    mStrTmp = Mid(sFilename, mIntStrLen + 1, Len(Trim(sFilename)) - mIntStrLen)
    GetFileName = mStrTmp
End Function

'--编辑 TXT 字段        2
Private Function SelectTXT(ByVal mTxt As String) As String
    Dim frm1 As New frmSetTXT
    With frm1
        .mTxt = mTxt
        .Show vbModal

        If .OK = True Then
            SelectTXT = .mTxt
        Else
            SelectTXT = ""
        End If
    End With

    Unload frm1
    Set frm1 = Nothing
End Function

'--编辑 RTF 字段
Private Function SelectRTxt(ByVal mTxt As String) As String
    Dim frm1 As New frmSetRTxt
    With frm1
        .mRTxt = mTxt
        .Show vbModal

        If .OK = True Then
            SelectRTxt = .mRTxt
        Else
            SelectRTxt = ""
        End If
    End With

    Unload frm1
    Set frm1 = Nothing
End Function

'--编辑 SQL 字段        2
Private Function SelectSQL(ByVal mTxt As String) As String
    Dim frm1 As New frmSetSQL
    With frm1
        .mTxt = mTxt
        .Show vbModal

        If .OK = True Then
            SelectSQL = .mTxt
        Else
            SelectSQL = ""
        End If
    End With

    Unload frm1
    Set frm1 = Nothing
End Function

'--单选字段
'--单元分格符为 |  子分格符为 ,
Private Sub SetTComboItem(ByRef o As Object, ByVal sValueItem As String, ByVal iMaxComboItems As Integer)
    Dim aStr
    Dim aSub
    Dim i As Long
    Dim oValueItem As New TrueOleDBGrid80.ValueItem
    aStr = Split(sValueItem, "|")
    o.ValueItems.Clear
    For i = 0 To UBound(aStr)
        aSub = Split(aStr(i), ",")
        With oValueItem
            .Value = Val(aSub(0))
            .DisplayValue = Trim(aSub(1))
        End With
        Dim a As Column
        o.ValueItems.add oValueItem
    Next
    o.ValueItems.MaxComboItems = iMaxComboItems
    Set oValueItem = Nothing
End Sub

'--单选字段
'--单元分格符为 |
Private Sub SetComboItem(ByRef o As Object, ByVal sValueItem As String, ByVal iMaxComboItems As Integer)
    Dim aStr
    Dim i As Long
    Dim oValueItem As New TrueOleDBGrid80.ValueItem

    o.ValueItems.Clear
    aStr = Split(sValueItem, "|")

    For i = 0 To UBound(aStr)
        With oValueItem
            .Value = Trim(aStr(i))
            .DisplayValue = Trim(aStr(i))
        End With
        o.ValueItems.add oValueItem
    Next
    o.ValueItems.MaxComboItems = iMaxComboItems
    Set oValueItem = Nothing
End Sub

'--表选字段
Private Sub SetTableSelectItem(ByRef o As Object, ByVal sSQL As String, ByVal sValueField As String, ByVal sDisplayValue As String, ByVal iMaxComboItems As Integer)
    On Error Resume Next
    Dim strSQL As String
    Dim rs As New RecordSet
    Dim oValueItem As ValueItem
    Dim sWhereStr As String

    o.ValueItems.Clear
    Gm.cnnTool.IniConnection8DM Gm.SysID.DBInfo
    Set rs = New RecordSet

    strSQL = sSQL & sWhereStr
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly

    Do While Not rs.EOF
        Set oValueItem = New ValueItem
        With oValueItem
            .Value = rs(sValueField)
            .DisplayValue = rs(sDisplayValue)
        End With
        o.ValueItems.add oValueItem

        rs.movenext
    Loop

    If iMaxComboItems = 0 Then
        o.ValueItems.MaxComboItems = 2
    Else
        o.ValueItems.MaxComboItems = iMaxComboItems
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    rsdetail.Close
    Set rsdetail = Nothing
    rsField.Close
    Set rsField = Nothing
End Sub


Private Sub TDBDropDown1_Click()
    'MsgBox "点击了"
End Sub

Private Sub TDBDropDown1_DropDownClose()
    RaiseEvent OnTDBDropDownClose
'    Dim szTip As String
'    Dim i As Long
'
'    szTip = ""
'    For i = 0 To TDBDropDown1.Columns.Count - 1
'        szTip = szTip & TDBDropDown1.Columns(i).DataField & "=" & TDBDropDown1.Columns(i).Value
'    Next
'    MsgBox szTip
End Sub



Private Sub TDBGrid1_ButtonClick(ByVal colIndex As Integer)
    OnContrlButtonClick colIndex
End Sub

Private Sub OnContrlButtonClick(ByVal colIndex As Integer)
    On Error Resume Next
    
    If TDBGrid1.AllowUpdate = False Then
        Exit Sub
    End If
    'rsField
    Dim vData As Variant
    Dim iOperation As Integer

    rsField.Filter = "B_FieldName='" & TDBGrid1.Columns(colIndex).DataField & "'"
    
    If rsField.RecordCount < 1 Then
        Exit Sub
    Else
        iOperation = IIf(IsNull(rsField("B_ControlValue")), 0, rsField("B_ControlValue"))
    End If
    
    
    '通过网格弹出窗体
    '=======================
    PopUpFrm fatherFrm, TDBGrid1.Columns(colIndex).name
    '=======================
    
    
    Select Case iOperation
        Case 11
            '11:颜色                 2
            vData = SelectColor
        Case 12
            '12:打开文件(文件全名)   2
            vData = SelectFullFile
        Case 13
            '13:打开目录             2
            vData = SelectDir
        Case 14
            '14:打开文件(只有文件名) 2
            vData = SelectFile
        Case 15
            '15:编辑 TXT 字段        2
            vData = SelectTXT(TDBGrid1.Columns(colIndex).Value)
        Case 16
            '16:编辑 RFT 字段        2
            vData = SelectRTxt(TDBGrid1.Columns(colIndex).Value)
        Case 17
            '17:编辑 SQL 字段        2
            vData = SelectSQL(TDBGrid1.Columns(colIndex).Value)
                    
    End Select
    If Len(vData) > 0 Then
        TDBGrid1.Columns(colIndex).Value = vData
        TDBGrid1.Update
    End If
End Sub

Private Sub TDBGrid1_DblClick()
    If m_PopUp = True Then
        Exit Sub
    End If
    
    If TDBGrid1.ApproxCount > 0 Then
        OnContrlButtonClick TDBGrid1.Col
    End If
End Sub

'-----以下为网格显示部件
Private Sub InitGridField()
    On Error Resume Next
    Dim strSQL As String
    Dim i As Long
    Dim rs As New RecordSet
    
    Set rs = New RecordSet
    
    
    Gm.cnnTool.IniConnection8DM Gm.SysID.DBInfo
    
    Select Case mvarFieldType
        Case 1
            strSQL = "Select * From G_FieldSystem"
        Case 2
            strSQL = "Select * From G_FieldUser"
        Case 3
            strSQL = "Select * From G_BLSField Where B_ObjectID='" & Trim(mvarObjectID) & "'"
        Case 4
            strSQL = "Select * From G_BLField  Where B_ObjectID='" & Trim(mvarObjectID) & "'"
        Case 5
            strSQL = "Select * From G_BLRField  Where B_ObjectID='" & Trim(mvarObjectID) & "'"
    End Select
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    Set rsField = New RecordSet
    
    '建立结构
    FillUnConnectRecordSet rs, rsField
    rs.Close
    Set rs = Nothing
End Sub

Private Sub FillUnConnectRecordSet(ByRef sRs As RecordSet, ByRef tRs As RecordSet)
    On Error Resume Next
    Dim i As Long
    
    Set tRs = New RecordSet
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
        sRs.movenext
    Loop
    
End Sub



Public Sub InitClass(ByRef TDBGrid As TrueOleDBGrid80.TDBGrid, ByVal m_FieldType As Integer)
    mvarFieldType = m_FieldType
    Set TDBGrid1 = TDBGrid
    InitGridField
    iMax = 0
    Dim a As Long
    Dim iP As POINTAPI
    Dim iC As POINTAPI
    
    a = ClientToScreen(TDBGrid1.hwnd, iP)
    a = ClientToScreen(TDBGrid1.Parent.hwnd, iC)
    
    iLeft = iP.X - iC.X
    iTop = iP.Y - iC.Y
    
    iLeft = TDBGrid1.Parent.ScaleY(iLeft, vbPixels, vbTwips)
    iTop = TDBGrid1.Parent.ScaleX(iTop, vbPixels, vbTwips)


    '设置网格显示样式
    SetGridDisplay
End Sub

'显示分组
'-----B_Split
Public Sub ShowGridSplit()
    
End Sub

'-----B_FieldName,B_CnName,B_GridWidth,B_Alignment,B_FieldFormat,B_BackColor,B_ForeColor
Public Sub ShowGridFormat()
    On Error Resume Next
    Dim i As Integer

    For i = 0 To TDBGrid1.Columns.Count - 1
        rsField.Filter = "B_FieldName='" & Trim(TDBGrid1.Columns(i).DataField) & "'"
        If rsField.RecordCount > 0 Then
            
            TDBGrid1.Columns(i).Caption = IIf(IsNull(rsField("B_CnName")) = True, "", rsField("B_CnName"))
            TDBGrid1.Columns(i).Alignment = IIf(IsNull(rsField("B_Alignment")) = True, 0, rsField("B_Alignment"))
            TDBGrid1.Columns(i).NumberFormat = IIf(IsNull(rsField("B_FieldFormat")) = True, "", rsField("B_FieldFormat"))
            
            If rsField("B_Locked") <> 0 Then
                TDBGrid1.Columns(i).Locked = True
            End If
            'TDBGrid1.Columns(i).BackColor = IIf(IsNull(rsField("B_BackColor")) = True, &H80000005, rsField("B_GridWidth"))
            'TDBGrid1.Columns(i).ForeColor = IIf(IsNull(rsField("B_ForeColor")) = True, &H80000008, rsField("B_ForeColor"))
            
            If rsField("B_GridWidth") = 0 Or IsNull(rsField("B_GridWidth")) Then
                TDBGrid1.Columns(i).Visible = False
            Else
                TDBGrid1.Columns(i).width = rsField("B_GridWidth")
            End If
        
        Else
            TDBGrid1.Columns(i).Visible = False
            TDBGrid1.Columns(i).width = 0
            TDBGrid1.Columns(i).Locked = True
        End If
    Next
    
    SetGridCategory
    
    
    TDBGrid1.HoldFields
End Sub

'建立客户控件
Private Sub CreateObjectA(ByVal m_SubObjectID As String, ByVal m_FieldName As String)
    Dim rs As New RecordSet
    Dim strSQL As String
    Dim oCtl As Object
    
    Set rs = New RecordSet
    strSQL = "Select * From G_ItemObject Where B_ObjectID='" & Trim(m_SubObjectID) & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        '删除未用的控件
        For Each oCtl In TDBGrid1.Parent.Controls
            If oCtl.name = rs("B_CtlName") Then
                TDBGrid1.Parent.Controls.Remove (rs("B_CtlName"))
            End If
        Next
        
        '建立用户控件
        
        Set o(iMax) = TDBGrid1.Parent.Controls.add(rs("B_CtlType"), rs("B_CtlName"))
        With o(iMax)
            
            .ConnectionString = Gm.cnnTool.cnnStr
            .sql = rs("B_SQL")
            .FieldName = rs("B_FieldName")
            .SubMenuObject = IIf(IsNull(rs("B_SubMenuObject")), "", rs("B_SubMenuObject"))
            .DefaultValue = IIf(IsNull(rs("B_DefaultValue")), "", rs("B_DefaultValue"))
            Set .TDBGrid1 = TDBGrid1
            .Visible = False
            .ZOrder 0
            .Refresh
        End With
        aGrid(0, iMax) = m_FieldName
        aGrid(1, iMax) = iMax
        aGrid(2, iMax) = GetDataFieldIndex(m_FieldName)
        
        iMax = iMax + 1
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Function GetDataFieldIndex(ByVal m_FieldName As String) As Integer
    Dim i As Integer
    For i = 0 To TDBGrid1.Columns.Count - 1
        If TDBGrid1.Columns(i).DataField = m_FieldName Then
            GetDataFieldIndex = i
            Exit Function
        End If
    Next
End Function

'-----B_ControlValue,B_DisplayField,B_ItemValues,B_SQL,B_DropWidth,B_MaxComboItems
Public Sub ShowGridCtl()
    
    On Error Resume Next
    Dim i As Integer

    'rsFieldField
    For i = 0 To TDBGrid1.Columns.Count - 1
        rsField.Filter = "B_FieldName='" & Trim(TDBGrid1.Columns(i).DataField) & "'"
        If rsField.RecordCount > 0 Then
            If rsField("B_ControlValue") >= 10 Then
                If rsField("B_ControlValue") < 90 Then
                    TDBGrid1.Columns(i).ValueItems.Presentation = dbgComboBox
                    TDBGrid1.Columns(i).ValueItems.CycleOnClick = True
                    TDBGrid1.Columns(i).ValueItems.Translate = True
                End If
                Select Case rsField("B_ControlValue")
                    Case 18
                        '18:单选字段 (不分值与显示)  2

                        SetComboItem TDBGrid1.Columns(i), rsField("B_ItemValues"), IIf(IsNull(rsField("B_MaxComboItems")), 0, rsField("B_MaxComboItems"))
                    Case 19
                        '19:单选字段 (分值与显示)  2
                    
                        SetTComboItem TDBGrid1.Columns(i), rsField("B_ItemValues"), IIf(IsNull(rsField("B_MaxComboItems")), 0, rsField("B_MaxComboItems"))
                    Case 20
                        '20:表选字段 (字段在表中)2
                        SetTableSelectItem TDBGrid1.Columns(i), rsField("B_SQL"), rsField("B_FieldName"), rsField("B_DisplayField"), IIf(IsNull(rsField("B_MaxComboItems")), 0, rsField("B_MaxComboItems"))
                    Case 21
                        '21:弹出网格             2
                        Dim oCtl As Object
                        For Each oCtl In TDBGrid1.Parent.Controls
                            If oCtl.name = "TDBDropDownPop" & rsField("B_FieldName") Then
                                TDBGrid1.Parent.Controls.Remove ("TDBDropDownPop" & rsField("B_FieldName"))
                            End If
                        Next
                        
                        Set TDBDropDown1 = TDBGrid1.Parent.Controls.add("TrueOleDBGrid80.TDBDropDown", "TDBDropDownPop" & rsField("B_FieldName"))
                        'Set TDBDropDown1 = TDBGrid1.Parent.Controls.Add("TrueOleDBGrid80.TDBDropDown", "TDBDropDownPop")
                        
                        '显示的项目数量大于0了才显示下拉网格
                        If rsField!B_MaxComboItems > 0 Then
                            TDBGrid1.Columns(i).DropDown = TDBDropDown1
                        End If
                        TDBGrid1.Columns(i).AutoDropDown = True
                        TDBGrid1.Columns(i).AutoCompletion = True

                        SetDropDownGrid rsField("B_SQL"), rsField("B_FieldName"), rsField("B_DisplayField")
                        m_DropDownGridSQL = rsField("B_SQL")
                        Debug.Print m_DropDownGridSQL
                    Case 99
                        '9: 自定义控件
                        CreateObjectA rsField("B_SubObjectID"), rsField("B_FieldName")
                        
                End Select

            Else
                TDBGrid1.Columns(i).ValueItems.Presentation = IIf(IsNull(rsField("B_ControlValue")), 0, rsField("B_ControlValue"))
            End If
        End If
    Next

    TDBGrid1.HoldFields
End Sub


Public Sub ShowGridCtl_ZDGird(ByRef vTDBGrid As TDBGrid)
    
    On Error Resume Next
    Dim i As Integer

    'rsFieldField
    For i = 0 To TDBGrid1.Columns.Count - 1
        rsField.Filter = "B_FieldName='" & Trim(TDBGrid1.Columns(i).DataField) & "'"
        If rsField.RecordCount > 0 Then
            If rsField("B_ControlValue") >= 10 Then
                If rsField("B_ControlValue") < 90 Then
                    TDBGrid1.Columns(i).ValueItems.Presentation = dbgComboBox
                    TDBGrid1.Columns(i).ValueItems.CycleOnClick = True
                    TDBGrid1.Columns(i).ValueItems.Translate = True
                End If
                Select Case rsField("B_ControlValue")
                    Case 18
                        '18:单选字段 (不分值与显示)  2

                        SetComboItem TDBGrid1.Columns(i), rsField("B_ItemValues"), IIf(IsNull(rsField("B_MaxComboItems")), 0, rsField("B_MaxComboItems"))
                    Case 19
                        '19:单选字段 (分值与显示)  2
                    
                        SetTComboItem TDBGrid1.Columns(i), rsField("B_ItemValues"), IIf(IsNull(rsField("B_MaxComboItems")), 0, rsField("B_MaxComboItems"))
                    Case 20
                        '20:表选字段 (字段在表中)2
                        SetTableSelectItem TDBGrid1.Columns(i), rsField("B_SQL"), rsField("B_FieldName"), rsField("B_DisplayField"), IIf(IsNull(rsField("B_MaxComboItems")), 0, rsField("B_MaxComboItems"))
                    Case 21
                        '21:弹出网格             2
                        Dim oCtl As Object
                        For Each oCtl In vTDBGrid.Parent.Controls
                            If oCtl.name = "TDBDropDownPop" & rsField("B_FieldName") Then
                                vTDBGrid.Parent.Controls.Remove ("TDBDropDownPop" & rsField("B_FieldName"))
                            End If
                        Next
                        
                        Set TDBDropDown1 = vTDBGrid.Parent.Controls.add("TrueOleDBGrid80.TDBDropDown", "TDBDropDownPop" & rsField("B_FieldName"))
                        
                        
                        vTDBGrid.Columns(i).DropDown = TDBDropDown1
                        vTDBGrid.Columns(i).AutoDropDown = True
                        vTDBGrid.Columns(i).AutoCompletion = True

                        SetDropDownGrid_ZDTDBGrid vTDBGrid, rsField("B_SQL"), rsField("B_FieldName"), rsField("B_DisplayField")
                        m_DropDownGridSQL = rsField("B_SQL")
                    Case 99
                        '9: 自定义控件
                        CreateObjectA rsField("B_SubObjectID"), rsField("B_FieldName")
                        
                End Select

            Else
                TDBGrid1.Columns(i).ValueItems.Presentation = IIf(IsNull(rsField("B_ControlValue")), 0, rsField("B_ControlValue"))
            End If
        End If
    Next

    TDBGrid1.HoldFields
End Sub

'Public Sub RefreshDropDown8Para(ByVal vWhere As String)
'    On Error Resume Next
'    Dim strSQL As String
'    Dim sWhereStr As String
'    Dim i As Integer
'    Dim sSQL As String
'    sSQL = m_DropDownGridSQL
'
'    'RaiseEvent OnSetDropDownGrid(sWhereStr)
'
'    rsDetail.Close
'    Set rsDetail = New RecordSet
'    rsDetail.Open sSQL & vWhere, Gm.cnnTool.cnn, adOpenKeyset, adLockReadOnly
'    Set TDBDropDown1.DataSource = rsDetail
'End Sub


Public Sub RefreshDropDown()
    On Error Resume Next
    Dim strSQL As String
    Dim sWhereStr As String
    Dim i As Integer
    Dim sSQL As String
    sSQL = m_DropDownGridSQL
    
    RaiseEvent OnSetDropDownGrid(sWhereStr)

    rsdetail.Close
    Set rsdetail = New RecordSet
    rsdetail.Open sSQL & sWhereStr, Gm.cnnTool.cnn, adOpenKeyset, adLockReadOnly
    Set TDBDropDown1.DataSource = rsdetail
End Sub

Public Sub SetDropDownGrid(ByRef sSQL As String, ByVal sValueField As String, _
    ByVal sDisplayField As String)
    
    On Error GoTo IFERR
    Dim strSQL  As String
    Dim sWhereStr As String
    Dim i As Integer
    
    '在创建某列的下拉框时触发使用端的事件
    '向其传输当前列的字段名sValueField
    '传址的方式sWhereStr从使用端获取下拉数据源的and子句
    'RaiseEvent OnSetDropDownGrid(sValueField, sWhereStr)
    RaiseEvent OnSetDropDownGrid(sWhereStr)
    
    
    Debug.Print sSQL & sWhereStr
    Set rsdetail = New RecordSet
    rsdetail.Open sSQL & sWhereStr, Gm.cnnTool.cnn, adOpenKeyset, adLockReadOnly
    Debug.Print rsdetail.RecordCount
    Set TDBDropDown1.DataSource = rsdetail
    
    With TDBDropDown1
        .DataField = sValueField
        .ListField = sDisplayField
        .ColumnHeaders = False
        .ScrollBars = dbgVertical
        .Appearance = dbgFlat
        .BackColor = &H80000005
        .DeadAreaBackColor = &H80000005
        .EmptyRows = True
        .Styles(5).BackColor = &H31CFFF
        .Styles(5).ForeColor = &H80000008
        .RowDividerStyle = dbgNoDividers
        
        .RowHeight = 380
        .Style.Alignment = dbgCenter
        .Style.VerticalAlignment = dbgVertCenter
        
        .Font.Size = TDBGrid1.Font.Size
    End With

    Dim rs As New RecordSet
    Dim maxWidth As Long
    '从系统显示列表中获取列宽
    'strSQL = "Select * From G_FieldSystem Where 1=1"
    
    '在2017年1月2日 15:37:38修改为从当前单据获取
    strSQL = "SELECT * FROM G_BLField WHERE B_ObjectID='" & mvarObjectID & "'"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockReadOnly
    
    For i = 0 To TDBDropDown1.Columns.Count - 1
        rs.Filter = "B_FieldName='" & Trim(TDBDropDown1.Columns(i).DataField) & "'"
        If rs.RecordCount > 0 Then
            Debug.Print TDBDropDown1.Columns(i).DataField
            TDBDropDown1.Columns(i).width = IIf(IsNull(rs("B_GridWidth")) = True, 0, rs("B_GridWidth"))
            maxWidth = maxWidth + TDBDropDown1.Columns(i).width
            TDBDropDown1.Columns(i).Caption = IIf(IsNull(rs("B_CnName")) = True, "", rs("B_cnName"))
            TDBDropDown1.Columns(i).Alignment = IIf(IsNull(rs("B_Alignment")) = True, 0, rs("B_Alignment"))
            TDBDropDown1.Columns(i).NumberFormat = IIf(IsNull(rs("B_FieldFormat")) = True, "", rs("B_FieldFormat"))
        Else
            TDBDropDown1.Columns(i).Visible = False
        End If
    Next

    rs.Close
    Set rs = Nothing
    TDBDropDown1.height = TDBDropDown1.RowHeight * 12
    TDBDropDown1.width = maxWidth + 250
    TDBDropDown1.HoldFields
    
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "在设置网格中的下拉数据时出错！" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub


Public Sub SetDropDownGrid_ZDTDBGrid(ByRef vTDBGrid As TDBGrid, _
    ByRef sSQL As String, ByVal sValueField As String, _
    ByVal sDisplayField As String)
    
    Dim strSQL  As String
    Dim sWhereStr As String
    Dim i As Integer
    Dim dWidthTDBGridCol As Double  '下拉网格所属的TDBGrid网格列的宽度

    RaiseEvent OnSetDropDownGrid(sWhereStr)
    
    dWidthTDBGridCol = vTDBGrid.Columns(Trim$(sValueField)).width
    
    Set rsdetail = New RecordSet
    rsdetail.Open sSQL & sWhereStr, Gm.cnnTool.cnn, adOpenKeyset, adLockReadOnly
    Set TDBDropDown1.DataSource = rsdetail
    
    With TDBDropDown1
        .DataField = sValueField
        .ListField = sDisplayField
        .ColumnHeaders = False
        .ScrollBars = dbgVertical
        .Appearance = dbgFlat
        .BackColor = &H80000005
        .DeadAreaBackColor = &H80000005
        .EmptyRows = True
        .Styles(5).BackColor = &H31CFFF
        .Styles(5).ForeColor = &H80000008
        .RowDividerStyle = dbgNoDividers
    End With

    Dim rs As New RecordSet
    Dim maxWidth As Long
    strSQL = "SELECT * FROM G_BLField WHERE B_ObjectID='" & mvarObjectID & "'"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockReadOnly
    
    For i = 0 To TDBDropDown1.Columns.Count - 1
        rs.Filter = "B_FieldName='" & Trim(TDBDropDown1.Columns(i).DataField) & "'"
        If rs.RecordCount > 0 Then
            Debug.Print TDBDropDown1.Columns(i).DataField
            TDBDropDown1.Columns(i).width = IIf(IsNull(rs("B_GridWidth")) = True, 0, rs("B_GridWidth"))
            maxWidth = maxWidth + TDBDropDown1.Columns(i).width
            TDBDropDown1.Columns(i).Caption = IIf(IsNull(rs("B_CnName")) = True, "", rs("B_cnName"))
            TDBDropDown1.Columns(i).Alignment = IIf(IsNull(rs("B_Alignment")) = True, 0, rs("B_Alignment"))
            TDBDropDown1.Columns(i).NumberFormat = IIf(IsNull(rs("B_FieldFormat")) = True, "", rs("B_FieldFormat"))
        Else
            TDBDropDown1.Columns(i).Visible = False
        End If
    Next

    rs.Close
    Set rs = Nothing
    TDBDropDown1.height = TDBDropDown1.RowHeight * 12
    'TDBDropDown1.Width = maxWidth + 250
    TDBDropDown1.width = dWidthTDBGridCol + 250
    
    TDBDropDown1.HoldFields
End Sub


Private Sub TDBGrid1_BeforeRowColChange(Cancel As Integer)
    On Error Resume Next
    If TDBGrid1.AllowUpdate = True And iMax > 0 Then
        Dim i As Integer
        Dim iOldShow As Integer
        For i = 0 To iMax - 1
            If o(i).Visible = True Then
                iOldShow = i
                Exit For
            Else
                iOldShow = -1
            End If
        Next
        
        If iOldShow >= 0 Then
            o(iOldShow).Visible = False
    
            TDBGrid1.Columns(TDBGrid1.Col).Value = o(iOldShow).Text
            TDBGrid1.Update
            TDBGrid1.PostMsg 81
        End If
    End If
End Sub

Private Sub TDBGrid1_PostEvent(ByVal MsgId As Integer)

    Select Case MsgId
        Case 97
            If TDBGrid1.bookmark > 1 Then
                o(iNowCtl).Visible = False
                TDBGrid1.SetFocus
                TDBGrid1.Columns(TDBGrid1.Col).Value = o(iNowCtl).Text
                SendKeys "{UP}"
            End If
        Case 98
            If Not TDBGrid1.EOF Then
                o(iNowCtl).Visible = False
                TDBGrid1.SetFocus
                TDBGrid1.Columns(TDBGrid1.Col).Value = o(iNowCtl).Text
                SendKeys "{DOWN}"
            End If
        Case 99 '回车
            o(iNowCtl).Visible = False
            TDBGrid1.SetFocus
            TDBGrid1.Columns(TDBGrid1.Col).Value = o(iNowCtl).Text
            TDBGrid1.PostMsg 81
            SendKeys "{Enter}"
    End Select
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    WriteLog "进入了TDBGrid1_RowColChange"
    On Error Resume Next
    If TDBGrid1.AllowUpdate = True And iMax > 0 Then
        WriteLog "即将开始ShowCustomCtl"
        ShowCustomCtl
        WriteLog "ShowCustomCtl结束"
    End If
    
    WriteLog "即将退出TDBGrid1_RowColChange"
End Sub

Private Sub TDBGrid1_DataSourceChanged()
    WriteLog "进入TDBGrid1_DataSourceChanged"
    On Error Resume Next
    If TDBGrid1.AllowUpdate = False And iMax > 0 Then
        WriteLog "即将开始设置控件的可见性"
        o(iNowCtl).Visible = False
    End If
    
    WriteLog "即将退出TDBGrid1_DataSourceChanged"
End Sub

Private Sub ShowCustomCtl()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    Dim Icol As Integer
    Dim iCtl As Integer
    Dim iCtlIndex As Integer
    '当前网格是否需要显示,iCol 是网格,iCtl 是控件
    o(iNowCtl).Visible = False
    If IsNull(TDBGrid1.bookmark) Then
        o(iNowCtl).Visible = False
    Else
        For i = 0 To iMax - 1
            For j = TDBGrid1.Col To TDBGrid1.Columns.Count - 1
                If TDBGrid1.Columns(j).Visible = True Then
                    If TDBGrid1.Columns(j).DataField = aGrid(0, i) Then
                        Icol = j
                        iCtl = i
                        '当前有控件
                        Exit For
                    Else
                        Icol = -1
                    End If
                End If
            Next
            If Icol >= 0 Then
                Exit For
            End If
        Next
        
        If Icol >= 0 And TDBGrid1.Col = Icol Then
            iCtlIndex = iCtl
            'iCtlIndex = aGrid(1, iCol)
            iNowCtl = iCtlIndex
            With o(iCtlIndex)
                .width = TDBGrid1.Columns(TDBGrid1.Col).width + 10
                .height = TDBGrid1.RowHeight - 15
                .Text = TDBGrid1.Columns(Icol).Value
                .Move TDBGrid1.Columns(TDBGrid1.Col).Left + iLeft, TDBGrid1.RowTop(TDBGrid1.Row) + iTop
                .Visible = True
                .ZOrder 0
                .SetFocus
            End With
        End If
    End If
End Sub



'根据数据表G_PopUpDataSendBLDetail的设置弹出窗体
'参数解释:
'frm    单据所在的窗体(传址)
'm_FieldName   当前按钮所在列的字段名称(传值)
Private Function PopUpFrm(ByRef frm As Object, ByVal m_FieldName As String) As Boolean
    Dim strSQL As String
    Dim rs As RecordSet
    Dim rs1 As RecordSet
    Dim m_ParaNumber As Long  '参数的个数
    Dim aValue
    Dim i As Long
    
    PopUpFrm = True
    
    
    Set rs = New RecordSet
    strSQL = "Select * From G_PopUpDataSendBLDetail Where B_ObjectID='" & mvarObjectID & "'"
    strSQL = strSQL & " And B_FieldName='" & m_FieldName & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        PopUpFrm = False
        Exit Function
    End If
    
    
    
    
    '获取参数
    '=====================
    Set rs1 = New RecordSet
    strSQL = "Select * From G_PopUpDataSendBLDetailPara where B_ObjectID='" & mvarObjectID & "' And B_FieldName='" & m_FieldName & "'"
    rs1.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    If rs1.RecordCount > 0 Then
        aValue = ""
        m_ParaNumber = rs1("B_ParaNumber")

        For i = 1 To m_ParaNumber
            rs1.Filter = " B_ParaIndex=" & i
            If rs1.RecordCount > 0 Then
                If rs1("B_BillOrDetail") = 1 Then
                    aValue = aValue & frm.Controls(Trim(rs1("B_ParaFieldName"))).Text & ","
                Else
                    aValue = aValue & frm.TDBGrid1.Columns(rs1("B_ParaFieldName")).Value & ","
                End If
            Else
                aValue = aValue & ","
            End If
            
            rs1.Filter = ""
        Next
        
        
        aValue = Left(aValue, Len(aValue) - 1)
    End If
    rs1.Close
    Set rs1 = Nothing
    '=====================
    
    
    
    
    Dim clsCommand1 As New clsCommand
    clsCommand1.InitClass
    clsCommand1.ExecutePopUp 1, frm, rs("B_ObjectID"), m_FieldName, rs("B_fFieldName"), rs("B_fObjectID"), " ", "LoadObject", Nothing, aValue
    
End Function



'设置网格控件列的分类显示111
Public Sub SetGridCategory()
    On Error Resume Next
    Dim strSQL As String
    Dim rs As RecordSet
    Dim rs1 As RecordSet
    Dim i As Long
    Dim j As Long
    Dim s() As Object
    
    
    Set rs = New RecordSet
    strSQL = "Select * From G_TDBGColCategoryBill Where B_ObjectID='" & mvarObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount <= 1 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    Set rs1 = New RecordSet
    strSQL = "Select * From G_TDBGColCategoryDetail Where B_ObjectID='" & mvarObjectID & "'"
    Debug.Print strSQL
    rs1.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    

    
    
    
    '如果已经存在Split，那么不再设置网格显示格式
    If TDBGrid1.Splits.Count > 1 Then
    
        '当网格中已经设置了Splits
        '这里仅仅设置列的显示样式
        '=========================
        ReDim s(rs.RecordCount) As Object
        Do While Not rs.EOF
            Set s(i) = TDBGrid1.Splits(i)
            i = i + 1
            rs.movenext
        Loop
    
        
    
        TDBGrid1.Splits(0).ScrollBars = dbgHorizontal
        '逐个Split设置Caption,Size以及隐藏不该显示的字段
        For i = 0 To UBound(s) - 1
            '除了第一个split需要recordselector,其他的都不需要
            'If i > 0 Then
                TDBGrid1.Splits(i).RecordSelectors = False
            'End If
        
            rs.Filter = " B_SIndex=" & i
            If Len(IIf(IsNull(rs!B_Caption), "", rs!B_Caption)) > 0 Then
                TDBGrid1.Splits(i).Caption = rs!B_Caption
            End If
            TDBGrid1.Splits(i).SizeMode = dbgNumberOfColumns
            TDBGrid1.Splits(i).Size = rs!B_size
            
            rs1.Filter = " B_SIndex=" & i
            For j = 0 To TDBGrid1.Splits(i).Columns.Count - 1
                rs1.Filter = " B_SIndex=" & i & " And B_FieldName='" & TDBGrid1.Splits(i).Columns(j).DataField & "'"
                Debug.Print " B_SIndex=" & i & " And B_FieldName='" & TDBGrid1.Splits(i).Columns(j).DataField & "'"
                If rs1.RecordCount <= 0 Then
                    TDBGrid1.Splits(i).Columns(j).width = 0
                    TDBGrid1.Splits(i).Columns(j).Visible = False
                    TDBGrid1.Splits(i).Columns(j).AllowSizing = False
                End If
                rs1.Filter = ""
            Next
            
            rs.Filter = ""
            rs1.Filter = ""
        Next
        '=========================
    
    
    
        rs.Close
        Set rs = Nothing
        
        rs1.Close
        Set rs1 = Nothing
        Exit Sub
    End If
    
    
    ReDim s(rs.RecordCount) As Object
    For i = 0 To rs.RecordCount - 1
        Set s(i) = TDBGrid1.Splits.add(i)
        'TDBGrid1.Splits(i).ScrollGroup = 2
    Next i
    
    
    '逐个Split设置Caption,Size以及隐藏不该显示的字段
    For i = 0 To UBound(s) - 1
        rs.Filter = " B_SIndex=" & i
        If Len(IIf(IsNull(rs!B_Caption), "", rs!B_Caption)) > 0 Then
            TDBGrid1.Splits(i).Caption = rs!B_Caption
        End If
        
        

        If i = 0 Then
            TDBGrid1.Splits(i).SizeMode = dbgNumberOfColumns
            TDBGrid1.Splits(i).Size = rs!B_size
            TDBGrid1.Splits(i).ExtendRightColumn = False
        Else
            'TDBGrid1.Splits(i).RecordSelectors = False
            TDBGrid1.Splits(i).ScrollBars = dbgBoth
            TDBGrid1.Splits(i).ExtendRightColumn = True
        End If
        
        
        rs1.Filter = " B_SIndex=" & i
        For j = 0 To TDBGrid1.Splits(i).Columns.Count - 1
            rs1.Filter = " B_SIndex=" & i & " And B_FieldName='" & TDBGrid1.Splits(i).Columns(j).DataField & "'"
            
            If rs1.RecordCount <= 0 Then
                TDBGrid1.Splits(i).Columns(j).width = 0
                TDBGrid1.Splits(i).Columns(j).Visible = False
                TDBGrid1.Splits(i).Columns(j).AllowSizing = False
                
                Debug.Print " 隐藏：B_SIndex=" & i & " And B_FieldName='" & TDBGrid1.Splits(i).Columns(j).DataField & "'"
            Else
                Debug.Print " 显示：B_SIndex=" & i & " And B_FieldName='" & TDBGrid1.Splits(i).Columns(j).DataField & "'"
            End If
            rs1.Filter = ""
        Next
        
        rs.Filter = ""
        rs1.Filter = ""
    Next
    
    
    
    '删除掉最后一个无用的分类
    TDBGrid1.Splits.Remove rs.RecordCount
    
    

    
    rs.Close
    Set rs = Nothing
    
    rs1.Close
    Set rs1 = Nothing
End Sub

Public Sub SetColLockedExcept(ByRef vTDBGrid As TDBGrid, ByVal vFields As String, ByVal vSymbSplit As String)
    Dim arr
    Dim i As Long, j As Long
    
    arr = Split(vFields, vSymbSplit)
    
    For i = 0 To vTDBGrid.Columns.Count - 1
        vTDBGrid.Columns(i).Locked = True
    Next
    
    
    For i = 0 To vTDBGrid.Columns.Count - 1
        For j = 0 To UBound(arr)
            If vTDBGrid.Columns(i).DataField = arr(j) Then
                vTDBGrid.Columns(i).Locked = False
                Debug.Print "设置字段" & vTDBGrid.Columns(i).DataField & "为可编辑"
            End If
        Next
    Next
End Sub

'为TDBGrid网格控件的某列创建下拉列表框
'可以根据当前页面的具体数据再次刷新下拉控件中的数据
Public Sub CreateDropDownGrid(ByRef vFrm As Object, ByRef vTDBGrid As TDBGrid, _
    ByVal vSQL As String, ByVal vDisplayField As String, ByVal vFieldName As String)
    
    Dim oCtl As Object
    Dim i As Long
    
    vFieldName = Trim$(vFieldName)
    
    For Each oCtl In vTDBGrid.Parent.Controls
        If oCtl.name = "TDBDropDownPop" & vFieldName Then
            vFrm.Controls.Remove ("TDBDropDownPop" & vFieldName)
        End If
    Next
    
    Set TDBDropDown1 = vTDBGrid.Parent.Controls.add("TrueOleDBGrid80.TDBDropDown", "TDBDropDownPop" & vFieldName)
    
    '显示的项目数量大于0了才显示下拉网格
    vTDBGrid.Columns(vDisplayField).DropDown = TDBDropDown1
    vTDBGrid.Columns(vDisplayField).AutoDropDown = True
    vTDBGrid.Columns(vDisplayField).AutoCompletion = True
    
    SetDropDownGrid vSQL, vFieldName, vDisplayField
End Sub


Public Sub WriteLog(ByVal vContent As String)
    Dim cls1 As New clsFile
    Dim szLogFile As String

    szLogFile = App.Path & "\BLSaveErr.txt"
    cls1.WriteFileContent szLogFile, vContent
    
End Sub

'设置网格控件的显示样式
Private Sub SetGridDisplay()
    TDBGrid1.SelectedBackColor = RGB(0, 128, 255)
    TDBGrid1.SelectedForeColor = vbWhite
    'TDBGrid1.EditBackColor = vbRed
    TDBGrid1.RowHeight = 500
End Sub






Public Sub SetDropDownGrid4TDBGrid(ByRef vFrm As Form, ByRef vTDBGrid As TDBGrid, _
    ByRef sSQL As String, ByVal sValueField As String, _
    ByVal sDisplayField As String)
    
    Dim strSQL  As String
    Dim sWhereStr As String
    Dim i As Integer
    Dim dWidthTDBGridCol As Double  '下拉网格所属的TDBGrid网格列的宽度
    Dim szDropDownCtlName As String

    RaiseEvent OnSetDropDownGrid(sWhereStr)
    
    dWidthTDBGridCol = vTDBGrid.Columns(Trim$(sValueField)).width
    
    
    szDropDownCtlName = "TDBDropDownPop" & sValueField
    ClearHisDropDownCtl szDropDownCtlName
    
    '创建下拉网格控件
    Set TDBDropDown1 = vFrm.Controls.add("TrueOleDBGrid80.TDBDropDown", szDropDownCtlName)

    '生成数据源记录集并且绑定到下拉控件作为数据源
    Set rsdetail = New RecordSet
    rsdetail.Open sSQL & sWhereStr, Gm.cnnTool.cnn, adOpenKeyset, adLockReadOnly
    Set TDBDropDown1.DataSource = rsdetail
    
    '将下拉网格控件绑定到显示数据用的网格控件中
    vTDBGrid.Columns(Trim$(sValueField)).DropDown = TDBDropDown1
    vTDBGrid.Columns(Trim$(sValueField)).AutoDropDown = True
    vTDBGrid.Columns(Trim$(sValueField)).AutoCompletion = True
                        
                        
    With TDBDropDown1
        .DataField = sValueField
        .ListField = sDisplayField
        .ColumnHeaders = False
        .ScrollBars = dbgVertical
        .Appearance = dbgFlat
        .BackColor = &H80000005
        .DeadAreaBackColor = &H80000005
        .EmptyRows = True
        .Styles(5).BackColor = &H31CFFF
        .Styles(5).ForeColor = &H80000008
        .RowDividerStyle = dbgNoDividers
    End With

    Dim rs As New RecordSet
    Dim maxWidth As Long
    strSQL = "SELECT * FROM G_BLField WHERE B_ObjectID='" & mvarObjectID & "'"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockReadOnly
    
    For i = 0 To TDBDropDown1.Columns.Count - 1
        rs.Filter = "B_FieldName='" & Trim(TDBDropDown1.Columns(i).DataField) & "'"
        If rs.RecordCount > 0 Then
            Debug.Print TDBDropDown1.Columns(i).DataField
            TDBDropDown1.Columns(i).width = IIf(IsNull(rs("B_GridWidth")) = True, 0, rs("B_GridWidth"))
            maxWidth = maxWidth + TDBDropDown1.Columns(i).width
            TDBDropDown1.Columns(i).Caption = IIf(IsNull(rs("B_CnName")) = True, "", rs("B_cnName"))
            TDBDropDown1.Columns(i).Alignment = IIf(IsNull(rs("B_Alignment")) = True, 0, rs("B_Alignment"))
            TDBDropDown1.Columns(i).NumberFormat = IIf(IsNull(rs("B_FieldFormat")) = True, "", rs("B_FieldFormat"))
        Else
            TDBDropDown1.Columns(i).Visible = False
        End If
    Next

    rs.Close
    Set rs = Nothing
    TDBDropDown1.height = TDBDropDown1.RowHeight * 12
    TDBDropDown1.width = maxWidth + 250
    'TDBDropDown1.Width = dWidthTDBGridCol + 250
    
    TDBDropDown1.HoldFields
End Sub


'清除当前已经存在的同名下拉网格控件
Private Sub ClearHisDropDownCtl(ByVal vCtlName As String)
    Dim oCtl As Object
    For Each oCtl In TDBGrid1.Parent.Controls
       If oCtl.name = vCtlName Then
           TDBGrid1.Parent.Controls.Remove (vCtlName)
       End If
    Next
End Sub


Public Sub SaveColWidth()
    Dim rs As New RecordSet
    Dim szTableName As String
    Dim cls1 As New clsIDETool
    Dim i As Long
    Dim szFieldName As String
    Dim dWidth As Double
    
    szTableName = cls1.GetFieldTableName(mvarObjectID)
    strSQL = "Select * From " & szTableName & " where B_ObjectID='" & mvarObjectID & "'"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    For i = 0 To TDBGrid1.Columns.Count - 1
        szFieldName = TDBGrid1.Columns(i).DataField
        dWidth = TDBGrid1.Columns(i).width
        
        If dWidth > 0 Then
            rs.Filter = " B_FieldName='" & szFieldName & "'"
            If rs.RecordCount <= 0 Then
                rs.AddNew
            End If
            
            rs!B_GridWidth = dWidth
            rs.Update
        End If
    Next
    
    rs.Close
    Set rs = Nothing
End Sub
