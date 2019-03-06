VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSGridShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'显示 SharpGrid
Public WithEvents SGrid1 As SGGrid
Attribute SGrid1.VB_VarHelpID = -1
'保持属性值的局部变量
Private mvarObjectID As String '局部复制
Private mvarFieldType As Integer '局部复制
Private rsField As New RecordSet

Public A_rsCellColor As New RecordSet
Private rsCellColor As New RecordSet
Public A_FieldsCellColor As String


Private Const A_Symbol01 As String = ","   '分组设置中的第一层间隔符
Private Const A_Symbol02 As String = "+"   '分组设置中的第二层间隔符
Private A_arrGroup() As String   '分组字段形成的数组
Private A_arrHeadFields() As String  '每个分组中要显示在GroupHeader上的字段
Private A_arrGroupFields() As String  '每个分组中要显示在GroupHeader上的字段



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
        tRs.Fields.Append sRs.Fields(i).name, sRs.Fields(i).Type, sRs.Fields(i).DefinedSize, adFldIsNullable
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

Public Sub InitClass(ByRef SGrid As Object, ByVal m_FieldType As Integer)
    mvarFieldType = m_FieldType
    Set SGrid1 = SGrid
    InitGrid
    InitGridField
    
    '获取对于单元格颜色的配置文件
    GetConfig_CellColor
End Sub

Private Sub InitGrid()
    With SGrid1
        .RowNumbering = True
        .Styles("Heading").ForeColor = RGB(0, 0, 0)
        .Styles("Heading").Padding = 50
        Set .Styles("Heading").BkgPicture = LoadPicture(App.Path & "\image\GridHead.bmp")
        
        
        .Rows.At(0).height = 500
        .Styles("Heading").TextAlignment = sgAlignCenterCenter
        .Styles("Heading").WordWrap = True
        .Styles("Selection").BackColor = RGB(255, 207, 49)
        .Styles("Selection").ForeColor = RGB(0, 0, 0)
        .Styles("InactiveSelection").BackColor = RGB(255, 207, 49)
        
        .Styles("GroupHeader").Font.Size = 9
        .Styles("GroupHeader").Font.Bold = True
        
        .Styles("GroupHeader").BkgStyle = sgCellBkgSolid
        .Styles("GroupHeader").Borders = sgCellBorderBottom
        
        .Styles("GroupHeader").BorderSize = 1
        
        .Styles("GroupHeader").BackColor = RGB(222, 223, 206)
        .Styles("GroupHeader").ForeColor = RGB(0, 0, 0)
        .Styles("GroupHeader").Padding = 0
        
        
        .Styles("Normal").WordWrap = True
        .Styles("Normal").TextAlignment = sgAlignCenterCenter

        .DefaultRowHeight = 380

        .BackColor = RGB(255, 255, 247)
        .SpecialMode = sgModeListBox
        .AutoResize = sgNoAutoResize
        .RedrawEnabled = True

        .Redraw
    End With
End Sub

'-----B_FieldName,B_CnName,B_GridWidth,B_Alignment,B_FieldFormat,B_BackColor,B_ForeColor
Public Sub ShowGridFormat()
    On Error GoTo IFERR
    Dim i As Integer

    For i = 1 To SGrid1.Columns.Count - 1

        rsField.Filter = "B_FieldName='" & Trim(SGrid1.Columns(i).Key) & "'"
        If rsField.RecordCount > 0 Then
            
            SGrid1.Columns(i).Caption = IIf(IsNull(rsField("B_CnName")) = True, "", rsField("B_CnName"))
                       
            If rsField("B_GridWidth") = 0 Or IsNull(rsField("B_GridWidth")) Then
                SGrid1.Columns(i).Hidden = True
                SGrid1.Columns(i).width = 0
            Else
                SGrid1.Columns(i).width = rsField("B_GridWidth")
                SGrid1.Columns(i).Style.Format = rsField("B_FieldFormat")
            End If
        
        Else
            SGrid1.Columns(i).AllowSizing = False
            SGrid1.Columns(i).width = 0
            SGrid1.Columns(i).Tag = "Hide"
            SGrid1.Columns(i).Hidden = True
        End If

    Next
    SGrid1.Redraw
    
    Exit Sub
IFERR:
    MsgBox Err.Description, vbOKOnly + vbInformation, "提示"
End Sub

Public Sub FillGrid(ByRef vRs As RecordSet) 'ByRef vRs As RecordSet
    On Error Resume Next
    Dim iMaxRow As Long
    Dim iMaxCol As Long
    Dim i As Long
    Dim j As Long
    
    Dim rs As New RecordSet
    Set rs = vRs.Clone
    rs.Filter = vRs.Filter
    
    iMaxRow = rs.RecordCount
    iMaxCol = rs.Fields.Count
    
    SGrid1.DataColCount = iMaxCol
    SGrid1.DataRowCount = iMaxRow
    
    '填充表头
    For i = 1 To rs.Fields.Count
        If SGrid1.Columns(i).Key <> rs.Fields(i - 1).name Then
            SGrid1.Columns(i).Key = rs.Fields(i - 1).name
            Select Case rs.Fields(i - 1).Type
                Case adBigInt, adDouble, adInteger, adSmallInt, adNumeric
                    SGrid1.Columns(i).DataType = sgtDouble
                Case Else
                    SGrid1.Columns(i).DataType = sgtString
            End Select
            
        Else
            Exit For
        End If
    Next
    
    i = 0
    '填充表数据
    Do While Not rs.EOF
        For j = 0 To rs.Fields.Count - 1
            SGrid1.Array(i, j) = rs(j).Value
        Next
        i = i + 1
        rs.movenext
    Loop
    
    
    
    '填充控件
    SetControls
    
    
    '设置数据行颜色
    SetCellColor

End Sub

'分组
Public Sub GroupGrid(ByVal m_GroupFields As String, ByVal m_SumFields As String)
    On Error Resume Next
    If Len(m_GroupFields) < 1 Then
        Exit Sub
    End If
    SGrid1.Groups.RemoveAll
    
    '在2017年2月8日添加的新的分组样式
    GetGroupArr m_GroupFields
    
    
    Dim aGroup
    Dim aSum
    Dim iMax As Integer
    'aGroup = Split(m_GroupFields, ",")
    aGroup = A_arrGroupFields
    aSum = Split(m_SumFields, ",")
    
    iMax = UBound(aGroup)
    
    ReDim Grp(0 To iMax) As SGGroup
    Dim i As Integer
    Dim j As Integer
    Dim sFormula As String
    For i = 0 To iMax
        Set Grp(i) = SGrid1.Groups.add(aGroup(i), sgNoSorting, , True, False)
        
        Grp(i).FetchHeaderStyle = True
        Grp(i).HeaderTextSource = sgGrpHdrColCaptionAndValue
        'Grp(i).HeaderTextSource = sgGrpHdrFireFetchText
        Grp(i).FooterTextSource = sgGrpFooterFormula
        
        
        sFormula = ""
        For j = 0 To UBound(aSum)
            Grp(i).Calculations.add sgCalcSum, aSum(j)
            sFormula = sFormula & " '  [" & GetGridCaption(SGrid1, aSum(j)) & "小计]: ' & Format((GroupCalc(" & j + 1 & ")),'0.00')  &"

        Next
        sFormula = Mid(sFormula, 1, Len(sFormula) - 1)
        Grp(i).FooterFormula = sFormula
        
    Next
    SGrid1.RefreshGroups sgCollapseGroups
    SGrid1.ExpandAll
    
    For i = 0 To SGrid1.Rows.Count - 1
        If SGrid1.Rows.At(i).Type <> sgSimpleRow Then
            SGrid1.Rows.At(i).height = 300
        End If
    Next


    'FillRowKey
    
    
    'Debug.Print SGrid1.GroupHeadings.Count
End Sub

Private Function GetGridCaption(ByRef SGrid1 As SGGrid, ByVal m_FieldName As String) As String
    Dim i As Integer
    For i = 0 To SGrid1.Columns.Count - 1
        If SGrid1.Columns(i).Key = m_FieldName Then
            GetGridCaption = SGrid1.Columns(i).Caption
            Exit Function
        End If
    Next
End Function

'求和
Public Sub SumGrid(ByRef ActiveBar21 As ActiveBar2, ByRef rs As RecordSet, ByVal m_SumFields As String)
    On Error Resume Next
    If Len(m_SumFields) < 1 Then
        Exit Sub
    End If
    If rs.RecordCount < 1 Then
        Exit Sub
    End If
    
    Dim aSum, iSum
    Dim iMax As Integer
    Dim i As Integer
    
    aSum = Split(m_SumFields, ",")
    iMax = UBound(aSum)
    ReDim iSum(0 To iMax)
    
    For i = 0 To iMax
        iSum(i) = 0
    Next
    '生成总计数
    rs.MoveFirst
    Do While Not rs.EOF
        For i = 0 To iMax
            'iSum(i) = 0
            iSum(i) = iSum(i) + IIf(IsNull(rs(aSum(i))), 0, rs(aSum(i)))
        Next
        rs.movenext
    Loop
    
    '填充总计
    Dim strCount As String
    strCount = "   报表总计: "
    For i = 0 To iMax
        strCount = strCount & GetGridCaption(SGrid1, aSum(i)) & ":" & Trim(Format(iSum(i), "#,###.00")) & "  "
    Next
    ActiveBar21.Parent.状态.Caption = strCount
End Sub


'根据字段名称(存储过程中的字段名)获取当前列的列数
Public Function GetColumnIndexByFieldName(ByRef oSGrid As Object, ByVal m_FieldName As String) As Long
    Dim SGCol As SGColumn
    Dim i As Long
    
    
    For Each SGCol In oSGrid.Columns
        If SGCol.Key = m_FieldName Then
            i = SGCol.ColIndex
            Exit For
        End If
    Next
    GetColumnIndexByFieldName = i
End Function



'设置网格内部的控件
Private Sub SetControls()
    Dim strSQL As String
    Dim rs As RecordSet
    Dim ctl
    
    
    Set rs = New RecordSet
    strSQL = "Select * From G_SGGridControls Where B_ObjectID='" & mvarObjectID & "' Order by B_ID"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        
        Exit Sub
    End If
    
    
    
    Do While Not rs.EOF
        Select Case rs("B_ControlName")
            Case "CheckBox"
                SGrid1.Columns(Trim(rs("B_FieldName"))).Control.Type = sgCellCheckBox
            Case "DropDown"
                SGrid1.Columns(Trim(rs("B_FieldName"))).Control.Type = sgCellDropDown
            Case "DropList"
                SGrid1.Columns(Trim(rs("B_FieldName"))).Control.Type = sgCellDropList
            Case "General"
                SGrid1.Columns(Trim(rs("B_FieldName"))).Control.Type = sgCellGeneral
            Case "TextBox"
                SGrid1.Columns(Trim(rs("B_FieldName"))).Control.Type = sgCellTextBox
                
            Case Else
            
        End Select
        
    
        '设置控件的对齐方式
        '遵循字段的对齐方式
        rsField.Filter = " B_FieldName='" & SGrid1.Columns(Trim(rs("B_FieldName"))).Key & "'"
        If rsField.RecordCount > 0 Then
            Select Case rsField!B_Alignment
                Case 0
                    SGrid1.Columns(Trim(rs("B_FieldName"))).Style.PictureAlignment = sgPicAlignLeftCenter
                Case 1
                    SGrid1.Columns(Trim(rs("B_FieldName"))).Style.PictureAlignment = sgPicAlignRightCenter
                Case 2
                    SGrid1.Columns(Trim(rs("B_FieldName"))).Style.PictureAlignment = sgPicAlignCenterCenter
                Case Else
                    SGrid1.Columns(Trim(rs("B_FieldName"))).Style.PictureAlignment = sgPicAlignCenterCenter
            End Select
        End If
        
        rs.movenext
    Loop
    
    
    rs.Close
    Set rs = Nothing

    SGrid1.RedrawEnabled = True
End Sub


Private Sub GetConfig_CellColor()
    Dim strSQL As String
    Dim rs As RecordSet
    
    
    A_FieldsCellColor = ""
    Set rs = New RecordSet
    strSQL = "Select * From G_SGGridCellColor Where B_ObjectID='" & mvarObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    
    A_FieldsCellColor = rs("B_EffectFields")
        
    rs.Close
    Set rs = Nothing
    
    
    Set A_rsCellColor = New RecordSet
    strSQL = "Select * From G_SGGridCellColor Where B_ObjectID='" & mvarObjectID & "'"
    A_rsCellColor.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    StringToRecordset A_FieldsCellColor, ",", rsCellColor
End Sub

'将带有分割符号的字符串转移到记录集中
Private Sub StringToRecordset(ByVal m_str As String, ByVal m_Filter As String, ByRef rs As RecordSet)
    Dim i As Long
    Dim m_Array
    
    
    m_Array = Split(m_str, m_Filter)
    
    Set rs = New RecordSet
    rs.Fields.Append "B_Field1", adVarChar, 100, adFldIsNullable
    rs.Open
    
    
    For i = 0 To UBound(m_Array)
        If Len(Trim(m_Array(i))) > 0 Then
        
            rs.AddNew
            rs("B_Field1") = m_Array(i)
            rs.Update
        End If
    Next
    
End Sub


Private Sub SetCellColor()
    If Len(Trim(A_FieldsCellColor)) <= 0 Then
        Exit Sub
    End If
    
    
    'rsCellColor
    
    Dim i As Long
    '设置可编辑样式
    
    If A_FieldsCellColor = "*" Then
        For i = 0 To SGrid1.Columns.Count - 1
            SGrid1.Columns(i).FetchCellStyle = True
        Next
    Else
        For i = 0 To SGrid1.Columns.Count - 1
            rsCellColor.Filter = " B_Field1='" & SGrid1.Columns(i).Key & "'"
            
            If rsCellColor.RecordCount > 0 Then
                SGrid1.Columns(i).FetchCellStyle = True
            End If
        Next
    End If
    
    
    'ChangeBackColor
End Sub

Private Sub ChangeBackColor()
    On Error GoTo IFERR
    Dim oRow As SGRow
    For Each oRow In SGrid1.Rows
        If oRow.Type = sgSimpleRow And Not oRow.Heading Then
            A_rsCellColor.MoveFirst
            Do While Not A_rsCellColor.EOF
                If Trim(str(oRow.Cells(SGrid1.Columns(Trim$(A_rsCellColor!B_KeyFieldName)).Position).Value)) = Trim(str(A_rsCellColor!B_KeyValue)) Then
                    oRow.Style.BackColor = A_rsCellColor!B_Color
                End If
                A_rsCellColor.movenext
            Loop
        
        End If
    Next
    SGrid1.Redraw
    Exit Sub
IFERR:
    Dim strERR As String
    strERR = Err.Description
    MsgBox strERR, vbOKOnly + vbInformation, "提示"
End Sub



'导出作为列CAPTION的行
'遇到分组头和分组脚不会产生空行
'根据数据类型，数值型的字符前面添加@符号转换为字符。
'              数值型的数据保持数值，可以合计
'参数介绍：
'vFields2Show：需要导出到EXCEL的列的字段值
'vColIndex：SGGrid控件上ColIndex=0表示最左侧的带有序号的列，
'           真正的数据从ColIndex=1开始。默认1，从第一列数据开始
'vExcelStartColIndex：导出的数据到EXCEL中从第几列开始，默认为1
'vCheckBox2Show：在控件SGGRID中显示为CheckBox的列是否导出到EXCEL中。0表示不导出，1导出。
Public Sub ExportToExcel(ByVal vSGGrid As SGGrid, _
    ByVal vFields2Show As String, Optional ByVal vColIndex As Long = 2, _
    Optional ByVal vExcelStartColIndex As Long = 0, _
    Optional ByVal vCheckBox2Show As Long = 0)
       
       
    On Error GoTo IFERR
    Dim Irow, Icol As Integer
    Dim Irowcount, Icolcount As Integer
    Dim Fieldlen() '存字段长度值
'
'    Dim xlApp As Object
'    Dim xlBook As Object
'    Dim xlSheet As Object
    
    Dim xlApp As excel.Application
    Dim xlBook As excel.Workbook
    Dim xlSheet As excel.Worksheet
       
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.add
    Set xlSheet = xlBook.Worksheets(1)
    Dim oSGRow As SGRow
       
    Dim szTip As String
    Dim lExcelStartColIndex As Long
    Dim szTemp As String
       
    Dim l2Show_CheckBox As Long '1表示要导出，0为不要导出
       
       
    '===================================================
    Dim szFields2Show As String
    Dim rsFields2Show As RecordSet
    Dim clsString1 As New clsString
       
    '获取需要显示的列的字段形成的字符串
    szFields2Show = vFields2Show
    Set rsFields2Show = clsString1.String2Recordset(szFields2Show, ",").Clone
    '===================================================
       
    Irowcount = vSGGrid.Rows.Count
       
    Icolcount = 0
    For Icolcount = 0 To vSGGrid.Columns.Count - 1
        If vSGGrid.Columns(Icolcount).Visible = True And _
            vSGGrid.Columns(Icolcount).width > 0 Then
               
            Icolcount = Icolcount + 1
        End If
    Next
       
    '如果指定开始的列号大于了列总数
    '则为非法，那么退出
    If vColIndex > Icolcount Then
        MsgBox "指定的列序号超出了边界！", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    ReDim Fieldlen(Icolcount)
    xlApp.Visible = False '显示表格
       
    '逐行扫描
    Irow = 0
    Icol = 0
    For Each oSGRow In vSGGrid.Rows
        If oSGRow.Type = sgSimpleRow Then
        If oSGRow.Type <> sgGroupFooter And oSGRow.Type <> sgGroupHeader Then
           
           
        lExcelStartColIndex = vExcelStartColIndex
        For Icol = vColIndex To vSGGrid.Columns.Count - 1
'            If vSGGrid.Columns(Icol).Visible = True And _
'                vSGGrid.Columns(Icol).Width > 0 Then
                   
            szTemp = vSGGrid.Columns(Icol).Key
            rsFields2Show.Filter = " B_Field0='" & szTemp & "'"
               
            '显示为CheckBox的列是否导出
            If vSGGrid.Columns(Icol).Control.Type = sgCellCheckBox Then
                l2Show_CheckBox = vCheckBox2Show
            Else
                l2Show_CheckBox = 1
            End If
               
            If rsFields2Show.RecordCount > 0 And l2Show_CheckBox = 1 Then
                Select Case vSGGrid.Columns(Icol).DataType
                    Case sgtDecimal, sgtDouble, sgtFloat, sgtLong, sgtShort
                    Case sgtString
                        xlSheet.Cells(Irow + 1, lExcelStartColIndex + 1).NumberFormatLocal = "@"
                    Case Else
                        xlSheet.Cells(Irow + 1, lExcelStartColIndex + 1).NumberFormatLocal = "@"
                End Select
                xlSheet.Cells(Irow + 1, lExcelStartColIndex + 1).Value = oSGRow.Cells(vSGGrid.Columns(Icol).Position).Value
                   
                lExcelStartColIndex = lExcelStartColIndex + 1
            End If
               
        Next
        Irow = Irow + 1
        End If
        End If
    Next
       
       
    xlApp.Visible = True '显示表格
    'xlBook.Save '保存"
    Set xlApp = Nothing '交还控制给Excel
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = Err.Description & "Excel导出时不正确!"
    MsgBox szErr, vbExclamation, "Excel"
    Exit Sub
End Sub

'判断当前行是否是数据行
Public Function IsDataRow(ByRef vRow As SGRow) As Boolean
    If vRow.Type = sgSimpleRow And vRow.Heading = False Then
        IsDataRow = True
    Else
        IsDataRow = False
    End If
End Function

'获取当前行的某个字段上的值，返回字符串
Public Function GetCurRowFieldValue(ByRef vSGGrid As SGGrid, ByVal vField As String) As String
    vField = Trim$(vField)
    GetCurRowFieldValue = IIf(IsNull(vSGGrid.Rows.Current.Cells(vSGGrid.Columns(vField).Position).Value), "", vSGGrid.Rows.Current.Cells(vSGGrid.Columns(vField).Position).Value)
End Function

Private Sub SGrid1_FetchGroupHeaderData(ByVal GroupIndex As Integer, _
    ByVal RowKey As Long, Text As String, PictureExpanded As Variant, _
    PictureCollapsed As Variant)
    
    On Error Resume Next
    Dim lCount As Long, i As Long, gh As SGGroupHeading
    Dim szTip As String
    Dim DataRowKey As Long
    Dim Index As Long
    
    
    Index = SGrid1.Columns("B_CodeID").Position

    For i = 0 To SGrid1.GroupHeadings.Count
        If RowKey = SGrid1.GroupHeadings(i).Row.Key Then
            'Text = SGrid1.GroupHeadings(i).ChildRows(3).Cells(Index).Text
            Debug.Print SGrid1.GroupHeadings(i).ChildRows.First.Cells(Index).Text
            Text = SGrid1.GroupHeadings(i).ChildRows.First.Cells(Index).Text
            
        End If
    Next
End Sub

Private Sub SGrid1_FetchGroupHeaderStyle(ByVal GroupIndex As Integer, ByVal RowKey As Long, ByVal HeaderStyle As DDSharpGridOLEDB2U.IsgStyle)
'     With HeaderStyle
'      '.BackColor = QBColor(SGrid1.Rows(RowKey).GroupHeading.GroupingValue)
'      .ForeColor = vbWhite
'      .TextAlignment = sgAlignCenterCenter
'      .Font.Bold = True
'   End With

End Sub

Private Sub FillRowKey()
    On Error GoTo IFERR
    Dim oRow As SGRow
    Dim IndexCol As Long
    IndexCol = SGrid1.Columns("B_MemoBill").Position
    
'    For Each oRow In SGrid1.Rows
'        If oRow.Heading = False Then
'            Debug.Print oRow.Cells(IndexCol).Value
'            oRow.Cells(IndexCol).Value = oRow.Cells(IndexCol).Value & ",RowKey=" & oRow.Key
'        End If
'    Next

    Dim szTip As String
    For Each oRow In SGrid1.Rows
        szTip = ""
        If oRow.Heading = True Then
            szTip = "是Heading"
        End If
        
        If oRow.Type = sgGroupHeader Then
            szTip = szTip & ",是sgGroupHeader"
        End If
        
        If oRow.Type = sgGroupFooter Then
            szTip = szTip & ",是sgGroupFooter"
        End If
        
        If oRow.Type = sgSimpleRow Then
            szTip = szTip & ",是sgSimpleRow"
        End If
        
        szTip = szTip & ",RowKey=" & oRow.Key
        
        Debug.Print szTip
    Next
    
    Exit Sub
IFERR:
    MsgBox Err.Description, vbOKOnly + vbInformation, "提示"
End Sub

'传递进来的分组字段字符串遵循的格式：
'B_CodeID+B_ClientName+B_Date,B_GoodsName+B_Specification
'英文的逗号,作为一级分隔符。+作为二级分隔符
'通过一级分隔符获取的数组，再通过二级分隔符获取数组的第一个元素为SGGRID分组的字段
'如上面的例子则有两层分组，依次是：B_CodeID，B_GoodsName
'其中第一层分组B_CodeID的GroupHeader显示的文字中追加B_ClientName和B_Date
'第二层分组B_GoodsName，出了B_GoodsName还要追加字段B_Specification作为GroupHeader的Text
Private Function GetGroupArr(ByVal vGroupFields As String)
    Dim i As Long, j As Long
    Dim arr2() As String
    
    A_arrGroup = Split(vGroupFields, A_Symbol01)
    ReDim A_arrGroupFields(0 To UBound(A_arrGroup)) As String
    
    For i = 0 To UBound(A_arrGroup)
        arr2 = Split(A_arrGroup(i), A_Symbol02)
        A_arrGroupFields(i) = arr2(0)
    Next
    
'    For i = 0 To UBound(A_arrGroupFields)
'        Debug.Print A_arrGroupFields(i)
'    Next
End Function
