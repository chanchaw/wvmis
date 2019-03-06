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
'��ʾ SharpGrid
Public WithEvents SGrid1 As SGGrid
Attribute SGrid1.VB_VarHelpID = -1
'��������ֵ�ľֲ�����
Private mvarObjectID As String '�ֲ�����
Private mvarFieldType As Integer '�ֲ�����
Private rsField As New RecordSet

Public A_rsCellColor As New RecordSet
Private rsCellColor As New RecordSet
Public A_FieldsCellColor As String


Private Const A_Symbol01 As String = ","   '���������еĵ�һ������
Private Const A_Symbol02 As String = "+"   '���������еĵڶ�������
Private A_arrGroup() As String   '�����ֶ��γɵ�����
Private A_arrHeadFields() As String  'ÿ��������Ҫ��ʾ��GroupHeader�ϵ��ֶ�
Private A_arrGroupFields() As String  'ÿ��������Ҫ��ʾ��GroupHeader�ϵ��ֶ�



Public Property Let ObjectID(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ObjectID = 5
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ObjectID
    ObjectID = mvarObjectID
End Property

'-----����Ϊ������ʾ����
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
    
    '�����ṹ
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
    
    '��ȡ���ڵ�Ԫ����ɫ�������ļ�
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
    MsgBox Err.Description, vbOKOnly + vbInformation, "��ʾ"
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
    
    '����ͷ
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
    '��������
    Do While Not rs.EOF
        For j = 0 To rs.Fields.Count - 1
            SGrid1.Array(i, j) = rs(j).Value
        Next
        i = i + 1
        rs.movenext
    Loop
    
    
    
    '���ؼ�
    SetControls
    
    
    '������������ɫ
    SetCellColor

End Sub

'����
Public Sub GroupGrid(ByVal m_GroupFields As String, ByVal m_SumFields As String)
    On Error Resume Next
    If Len(m_GroupFields) < 1 Then
        Exit Sub
    End If
    SGrid1.Groups.RemoveAll
    
    '��2017��2��8�����ӵ��µķ�����ʽ
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
            sFormula = sFormula & " '  [" & GetGridCaption(SGrid1, aSum(j)) & "С��]: ' & Format((GroupCalc(" & j + 1 & ")),'0.00')  &"

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

'���
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
    '�����ܼ���
    rs.MoveFirst
    Do While Not rs.EOF
        For i = 0 To iMax
            'iSum(i) = 0
            iSum(i) = iSum(i) + IIf(IsNull(rs(aSum(i))), 0, rs(aSum(i)))
        Next
        rs.movenext
    Loop
    
    '����ܼ�
    Dim strCount As String
    strCount = "   �����ܼ�: "
    For i = 0 To iMax
        strCount = strCount & GetGridCaption(SGrid1, aSum(i)) & ":" & Trim(Format(iSum(i), "#,###.00")) & "  "
    Next
    ActiveBar21.Parent.״̬.Caption = strCount
End Sub


'�����ֶ�����(�洢�����е��ֶ���)��ȡ��ǰ�е�����
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



'���������ڲ��Ŀؼ�
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
        
    
        '���ÿؼ��Ķ��뷽ʽ
        '��ѭ�ֶεĶ��뷽ʽ
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

'�����зָ���ŵ��ַ���ת�Ƶ���¼����
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
    '���ÿɱ༭��ʽ
    
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
    MsgBox strERR, vbOKOnly + vbInformation, "��ʾ"
End Sub



'������Ϊ��CAPTION����
'��������ͷ�ͷ���Ų����������
'�����������ͣ���ֵ�͵��ַ�ǰ������@����ת��Ϊ�ַ���
'              ��ֵ�͵����ݱ�����ֵ�����Ժϼ�
'�������ܣ�
'vFields2Show����Ҫ������EXCEL���е��ֶ�ֵ
'vColIndex��SGGrid�ؼ���ColIndex=0��ʾ�����Ĵ�����ŵ��У�
'           ���������ݴ�ColIndex=1��ʼ��Ĭ��1���ӵ�һ�����ݿ�ʼ
'vExcelStartColIndex�����������ݵ�EXCEL�дӵڼ��п�ʼ��Ĭ��Ϊ1
'vCheckBox2Show���ڿؼ�SGGRID����ʾΪCheckBox�����Ƿ񵼳���EXCEL�С�0��ʾ��������1������
Public Sub ExportToExcel(ByVal vSGGrid As SGGrid, _
    ByVal vFields2Show As String, Optional ByVal vColIndex As Long = 2, _
    Optional ByVal vExcelStartColIndex As Long = 0, _
    Optional ByVal vCheckBox2Show As Long = 0)
       
       
    On Error GoTo IFERR
    Dim Irow, Icol As Integer
    Dim Irowcount, Icolcount As Integer
    Dim Fieldlen() '���ֶγ���ֵ
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
       
    Dim l2Show_CheckBox As Long '1��ʾҪ������0Ϊ��Ҫ����
       
       
    '===================================================
    Dim szFields2Show As String
    Dim rsFields2Show As RecordSet
    Dim clsString1 As New clsString
       
    '��ȡ��Ҫ��ʾ���е��ֶ��γɵ��ַ���
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
       
    '���ָ����ʼ���кŴ�����������
    '��Ϊ�Ƿ�����ô�˳�
    If vColIndex > Icolcount Then
        MsgBox "ָ��������ų����˱߽磡", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    ReDim Fieldlen(Icolcount)
    xlApp.Visible = False '��ʾ����
       
    '����ɨ��
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
               
            '��ʾΪCheckBox�����Ƿ񵼳�
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
       
       
    xlApp.Visible = True '��ʾ����
    'xlBook.Save '����"
    Set xlApp = Nothing '�������Ƹ�Excel
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = Err.Description & "Excel����ʱ����ȷ!"
    MsgBox szErr, vbExclamation, "Excel"
    Exit Sub
End Sub

'�жϵ�ǰ���Ƿ���������
Public Function IsDataRow(ByRef vRow As SGRow) As Boolean
    If vRow.Type = sgSimpleRow And vRow.Heading = False Then
        IsDataRow = True
    Else
        IsDataRow = False
    End If
End Function

'��ȡ��ǰ�е�ĳ���ֶ��ϵ�ֵ�������ַ���
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
            szTip = "��Heading"
        End If
        
        If oRow.Type = sgGroupHeader Then
            szTip = szTip & ",��sgGroupHeader"
        End If
        
        If oRow.Type = sgGroupFooter Then
            szTip = szTip & ",��sgGroupFooter"
        End If
        
        If oRow.Type = sgSimpleRow Then
            szTip = szTip & ",��sgSimpleRow"
        End If
        
        szTip = szTip & ",RowKey=" & oRow.Key
        
        Debug.Print szTip
    Next
    
    Exit Sub
IFERR:
    MsgBox Err.Description, vbOKOnly + vbInformation, "��ʾ"
End Sub

'���ݽ����ķ����ֶ��ַ�����ѭ�ĸ�ʽ��
'B_CodeID+B_ClientName+B_Date,B_GoodsName+B_Specification
'Ӣ�ĵĶ���,��Ϊһ���ָ�����+��Ϊ�����ָ���
'ͨ��һ���ָ�����ȡ�����飬��ͨ�������ָ�����ȡ����ĵ�һ��Ԫ��ΪSGGRID������ֶ�
'���������������������飬�����ǣ�B_CodeID��B_GoodsName
'���е�һ�����B_CodeID��GroupHeader��ʾ��������׷��B_ClientName��B_Date
'�ڶ������B_GoodsName������B_GoodsName��Ҫ׷���ֶ�B_Specification��ΪGroupHeader��Text
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