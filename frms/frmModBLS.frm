VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{B45CF12E-4E4F-487D-8096-DB3BFE63F435}#1.0#0"; "sg20ou.ocx"
Begin VB.Form frmModBLS 
   BackColor       =   &H00CEDFDE&
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   Icon            =   "frmModBLS.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   8145
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   5790
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8145
      _LayoutVersion  =   1
      _ExtentX        =   14367
      _ExtentY        =   10213
      _DataPath       =   ""
      Bands           =   "frmModBLS.frx":038A
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2280
         Top             =   4800
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin DDSharpGridOLEDB2U.SGGrid SGGrid1 
         Height          =   4635
         Left            =   180
         TabIndex        =   1
         Top             =   900
         Width           =   7755
         _cx             =   13679
         _cy             =   8176
         DataMember      =   ""
         DataMode        =   1
         AutoFields      =   -1  'True
         Enabled         =   -1  'True
         GridBorderStyle =   0
         ScrollBars      =   3
         FlatScrollBars  =   1
         ScrollBarTrack  =   0   'False
         DataRowCount    =   0
         BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataColCount    =   2
         HeadingRowCount =   1
         HeadingColCount =   1
         TextAlignment   =   5
         WordWrap        =   -1  'True
         Ellipsis        =   1
         HeadingBackColor=   -2147483633
         HeadingForeColor=   -2147483630
         HeadingTextAlignment=   0
         HeadingWordWrap =   0   'False
         HeadingEllipsis =   1
         GridLines       =   1
         HeadingGridLines=   2
         GridLinesColor  =   -2147483633
         HeadingGridLinesColor=   -2147483632
         EvenOddStyle    =   1
         ColorEven       =   -2147483624
         ColorOdd        =   16777215
         UserResizeAnimate=   1
         UserResizing    =   3
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         UserDragging    =   2
         UserHiding      =   2
         CellPadding     =   15
         CellBkgStyle    =   1
         CellBackColor   =   16252927
         CellForeColor   =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   1
         FocusRectColor  =   0
         FocusRectLineWidth=   1
         TabKeyBehavior  =   2
         EnterKeyBehavior=   2
         NavigationWrapMode=   1
         SkipReadOnly    =   -1  'True
         DefaultColWidth =   1200
         DefaultRowHeight=   380
         CellsBorderColor=   0
         CellsBorderVisible=   -1  'True
         RowNumbering    =   0   'False
         EqualRowHeight  =   0   'False
         EqualColWidth   =   0   'False
         HScrollHeight   =   0
         VScrollWidth    =   0
         Appearance      =   2
         FitLastColumn   =   0   'False
         SelectionMode   =   2
         MultiSelect     =   2
         AllowAddNew     =   0   'False
         AllowDelete     =   0   'False
         AllowEdit       =   0   'False
         ScrollBarTips   =   0
         CellTips        =   0
         CellTipsDelay   =   1000
         SpecialMode     =   1
         OutlineLines    =   1
         CacheAllRecords =   -1  'True
         ColumnClickSort =   0   'False
         PreviewPaneType =   0
         PreviewPanePosition=   2
         PreviewPaneSize =   2000
         GroupIndentation=   225
         InactiveSelection=   1
         AutoScroll      =   -1  'True
         AutoResize      =   1
         AutoResizeHeadings=   -1  'True
         OLEDragMode     =   0
         OLEDropMode     =   0
         MaxRows         =   4194304
         MaxColumns      =   8192
         NewRowPos       =   1
         CustomBkgDraw   =   0
         AutoGroup       =   -1  'True
         GroupByBoxVisible=   0   'False
         AlphaBlendEnabled=   0   'False
         DragAlphaLevel  =   206
         AutoSearch      =   0
         AutoSearchDelay =   2000
         Format          =   "frmModBLS.frx":6DAE
         Caption         =   "frmModBLS.frx":6DE0
         ScrollTipColumn =   "frmModBLS.frx":6E04
         GroupByBoxText  =   "frmModBLS.frx":6E28
         StylesCollection=   "frmModBLS.frx":6EAE
         ColumnsCollection=   "frmModBLS.frx":AA50
         ValueItems      =   "frmModBLS.frx":C73A
      End
   End
End
Attribute VB_Name = "frmModBLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit
Private m_SQL As String
Private m_OrderBySQL As String
Private m_EditFormName As String
Private clsSGridShow1 As clsSGridShow
Private mvarObjectID As String '局部复制
Private sFilter As String
Private iKeyIndex As Integer
Public m_KeyID As Variant

Private m_FieldID As String
Private m_TableName As String


'单据窗体获取客户资料中用的参数
'==============================
Public frmName As String
Public frm1 As Object
'==============================


'==============================
Public fatherFrm As Object   '单据窗体(传址)
Private mvarfObjectID As String
Private mvarfFieldName As String  '单据窗体上主表的字段名
Private mvarSendIndex As Integer  '被弹出窗体记录集中的即将被传输出去的数据的Index
Private mvarBillOrDetail As Integer '0 为主表  1为明细表



Private A_Order As String
Private A_Expand As Long   '1为展开，0为折叠
Private A_IsGroup As Long  '0为非分组，1为分组


Public Property Let BillOrDetail(ByVal vData As Integer)
    mvarBillOrDetail = vData
End Property

Public Property Get BillOrDetail() As Integer
    BillOrDetail = mvarBillOrDetail
End Property


Public Property Let fObjectID(ByVal vData As String)
    mvarfObjectID = vData
End Property

Public Property Get fObjectID() As String
    fObjectID = mvarfObjectID
End Property

Public Property Let fFieldName(ByVal vData As String)
    mvarfFieldName = vData
End Property

Public Property Get fFieldName() As String
    fFieldName = mvarfFieldName
End Property


Public Property Let SendIndex(ByVal vData As String)
    mvarSendIndex = vData
End Property

Public Property Get SendIndex() As String
    SendIndex = mvarSendIndex
End Property

'==============================



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


'获取Sharp Grid中被选中的多行中某列的VALUE，中间以逗号做间隔
Private Function GetSGGridMulRowsSingleColValue(ByRef vGrid As SGGrid, ByVal vFieldName As String) As String
    Dim Row As SGRow
    Dim szReturn As String
    
    szReturn = ""
    
    For Each Row In vGrid.Rows
    If Row.Type = sgSimpleRow And Row.Heading = False Then
    If vGrid.Selection.IsRowSelected(Row.Position) = True Then '当前行被选中
        '对于当前行的操作使用Row.XXX
          '在循环结束之后再进行网格数据的刷新
        szReturn = szReturn & Trim(Row.Cells(vGrid.Columns(vFieldName).Position).Text) & ","
    End If
    End If
    Next

    If Len(szReturn) > 0 Then
        szReturn = Left$(szReturn, Len(szReturn) - 1)
    End If
    
    GetSGGridMulRowsSingleColValue = szReturn
End Function


'以下为2012-2-12之后修改
'选择数据到父窗体
'========================
Private Sub SelectTo()
    Dim rs As RecordSet
    Dim strSQL As String
    Dim m_ToString As String
    Dim oSGRow  As SGRow
    
    Dim m_IndexObject As Long
    Dim i As Long
    Dim j As Long
    
    Dim szSelected As String  '传输到主表的时候可以多选，多个元素间使用英文逗号做间隔
    
    
    m_ToString = ""
    If Len(Trim(mvarfFieldName)) <= 0 Then
        Exit Sub
    End If
    
    
    '传输数据到主表
    If mvarBillOrDetail = 0 Then
        i = 0
        i = InStr(1, mvarfFieldName, "(")
        j = InStr(1, mvarfFieldName, ")")
        
        '获取多选的行的中的某列的VALUE
        szSelected = GetSGGridMulRowsSingleColValue(SGGrid1, Trim(SGGrid1.Columns(mvarSendIndex).Key))
        
        
        If i > 0 Then
            m_IndexObject = Val(Trim(Mid(mvarfFieldName, i + 1, j - i - 1)))
            'fatherFrm.Controls(left(mvarfFieldName, i - 1))(m_IndexObject).Text = Trim(SGGrid1.Rows.Current.Cells(mvarSendIndex).Text)
            fatherFrm.Controls(Left(mvarfFieldName, i - 1))(m_IndexObject).Text = szSelected
        Else
            'fatherFrm.Controls(mvarfFieldName).Text = Trim(SGGrid1.Rows.Current.Cells(mvarSendIndex).Text)
            fatherFrm.Controls(mvarfFieldName).Text = szSelected
        End If
    Else
    '传输数据到明细表
        Set rs = New RecordSet
        strSQL = "Select * From G_PopUpDataSendBLDetail Where B_ObjectID='" & mvarfObjectID & "'"
        strSQL = strSQL & " And B_FieldName='" & mvarfFieldName & "'"
        
        'Debug.Print strSQL
        rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        
        Do While Not rs.EOF
            m_ToString = ""
            For Each oSGRow In SGGrid1.Selection.Grid.Rows
                If SGGrid1.Selection.IsRowSelected(oSGRow.Position) = True Then
                    m_ToString = m_ToString & oSGRow.Cells(rs!B_fFieldName).Value & ","
                End If
            Next
            
            m_ToString = Left(m_ToString, Len(m_ToString) - 1)
            
            
            fatherFrm.TDBGrid1.Columns(rs("B_tFieldName")).Value = m_ToString
            
            rs.MoveNext
        Loop
        
        
        rs.Close
        Set rs = Nothing
        
    End If
    
    
    Unload Me
End Sub
'========================


Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "新增"
            AddNewObject
        Case "编辑"
            EditObject m_KeyID
        Case "删除"
            If MsgBox("是否要删除?", vbExclamation + vbOKCancel + vbDefaultButton2, "删除") = vbOK Then
                DeleteObject
            End If
        Case "导入"
            
        Case "导出"
            ExportToExcelA
        Case "刷新"
            RefreshGrid
            'LoadObject
            
        Case "过滤"
            FilterForm
        Case "关闭"
            Unload Me
            
        Case "选择"
            SelectTo
    End Select
    
End Sub

Private Sub ExportToExcelA()
    On Error Resume Next
    Dim Irow, Icol As Integer
    Dim Irowcount, Icolcount As Integer
    Dim Fieldlen() '存字段长度值
    Dim xlApp As excel.Application
    Dim xlBook As excel.Workbook
    Dim xlSheet As excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)

    Irowcount = SGGrid1.Rows.Count
    
    Icolcount = 0
    For Icolcount = 0 To SGGrid1.Columns.Count - 1
        If SGGrid1.Columns(Icol).Visible = True Then
            Icolcount = Icolcount + 1
        End If
    Next

    ReDim Fieldlen(Icolcount)
    xlApp.Visible = True '显示表格
    With SGGrid1
        For Irow = 0 To Irowcount - 1
            .Row = Irow
            
            For Icol = 0 To Icolcount - 1
                
                .Col = Icol
                If .Columns(Icol).Visible = True Then
                    xlSheet.Cells(Irow + 1, Icol + 1).NumberFormatLocal = "@"
                    xlSheet.Cells(Irow + 1, Icol + 1).Value = .CellAt(Irow, Icol).Text
                End If
                
            Next
    
        Next
    End With
    xlApp.Visible = True '显示表格
    'xlBook.Save '保存"
    Set xlApp = Nothing '交还控制给Excel
    Exit Sub
IFERR:
    MsgBox "Excel导出时不正确!", vbExclamation, "Excel"
    Exit Sub

End Sub


Private Sub Form_Load()

    ActiveBar21.ClientAreaControl = SGGrid1
    ActiveBar21.RecalcLayout
    GetObjectParameter
    
    Me.Left = 0
    Me.Top = 0
    
    AnimateForm Me
End Sub

Public Sub RefreshGrid()
    Adodc1.RecordSet.Filter = ""
    SGridShow
End Sub

'新增对象
Private Sub AddNewObject()
    On Error Resume Next
    
    
    '判断是否有新增的权限
    If Gm.PI.JudgeNew(Me.ObjectID) = False Then
        Exit Sub
    End If
    
    '刷新网格
    Dim o As Object
    
    Set o = GetFormNew(m_EditFormName)
    With o
        .AddNewObject ObjectID
        Set .HighterlevelFrm = Me
        .Show vbModal
    End With
    LoadObject
    Adodc1.RecordSet.MoveLast
End Sub

'编辑对象
Private Sub EditObject(ByVal m_KeyID As Variant)
    On Error Resume Next
    
    '判断是否有修改的权限
    If Gm.PI.JudgeUpdate(Me.ObjectID) = False Then
        Exit Sub
    End If
    
    
    Dim o As Object
    Dim Irow As Long
    Dim sKey As Variant
    Irow = SGGrid1.Row
    Set o = GetFormNew(m_EditFormName)
    With o
        sKey = SGGrid1.Rows.Current.Cells(iKeyIndex + 1).Text

        .m_KeyID = sKey
        .EditObject ObjectID
        .Show vbModal
    End With
    LoadObject
    SGGrid1.Row = Irow
End Sub

Private Sub GetKeyIndex()
    'iKeyIndex
    Dim i As Integer
    For i = 0 To Adodc1.RecordSet.Fields.Count - 1
        If Adodc1.RecordSet.Fields(i).Properties.Item(4).Value = True Then
            iKeyIndex = i
            Exit Sub
        End If
    Next
End Sub

Private Sub DeleteObject()
    On Error GoTo IFERR
    
    '判断是否有删除的权限
    If Gm.PI.JudgeDelete(Me.ObjectID) = False Then
        Exit Sub
    End If
    
    
    Dim sKey As String
    Dim strSQL As String
    sKey = SGGrid1.Rows.Current.Cells(iKeyIndex + 1).Text

    GetField
    
    strSQL = "Delete From " & m_TableName & " Where " & m_FieldID & "='" & sKey & "'"
    Gm.cnnTool.cnn.Execute strSQL
    SGGrid1.Delete
    SGGrid1.Update
    
    
    Exit Sub
IFERR:
    Dim szTip As String
    szTip = "存在对应业务数据，不可删除！"
    MsgBox szTip, vbOKOnly + vbInformation, "提示"
End Sub


'取得记录
Public Sub LoadObject()
    Dim m_SQLEx As String
    If Len(Trim(m_OrderBySQL)) > 0 Then
        m_SQLEx = m_SQL & " " & m_OrderBySQL
    Else
        m_SQLEx = m_SQL
    End If
    
    
    With Adodc1
        .ConnectionString = Gm.cnnTool.cnnStr
        .CommandType = adCmdText
        .RecordSource = m_SQLEx
        .Refresh
        GetKeyIndex
    End With
    SGridShow

End Sub

Private Sub SGridShow()
    Set clsSGridShow1 = New clsSGridShow
    'Adodc1.Recordset.Filter = sFilter
    With clsSGridShow1
        .ObjectID = ObjectID
        .InitClass SGGrid1, 3
        
        .FillGrid Adodc1.RecordSet
'        SGGrid1.DataMode = sgBound
'        Set SGGrid1.DataSource = Adodc1
'        SGGrid1.ReBind
        .ShowGridFormat
    End With
End Sub

Private Sub FilterForm()
    Dim frm1 As New frmUCTSearch
    With frm1
        Set .rs = Adodc1.RecordSet.Clone
        .ObjectID = ObjectID
        .FieldType = 3
        .Show vbModal
    End With
    If frm1.OK = True Then

        sFilter = frm1.strResult
        Adodc1.RecordSet.Filter = sFilter
        SGridShow
    End If
    Unload frm1
    Set frm1 = Nothing
End Sub


Private Sub GetField()
    Dim rs As New RecordSet
    Dim strSQL As String
    
    Set rs = New RecordSet
    Debug.Print m_SQL
    strSQL = m_SQL & " And 1=0"
    
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    
    rs.AddNew
    
    m_FieldID = rs.Fields(0).name
    m_TableName = rs.Fields(0).Properties(1).Value
    
    rs.CancelUpdate
    rs.Close
    Set rs = Nothing
End Sub


'取得参数
Private Sub GetObjectParameter()
    Dim rs As New RecordSet
    Dim strSQL As String
    
    Set rs = New RecordSet
    strSQL = "Select * From G_BLS Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    m_SQL = rs("B_SQL")
    m_OrderBySQL = IIf(IsNull(rs!B_OrderBySQL), "", rs!B_OrderBySQL)
    m_EditFormName = rs("B_EditFormName")
    
    Me.Width = rs("B_Width")
    Me.Height = rs("B_Height")
    Me.Caption = rs("B_BillName")
    
    rs.Close
    Set rs = Nothing
End Sub

'Private Sub SGGrid1_DblClick()
'    If Len(frmName) > 0 Then
'        Call SelectTo
'    Else
'        EditObject m_KeyID
'    End If
'End Sub


Private Sub SGGrid1_DblClick()
        
    '如果双击的不是正常的数据行，则直接跳出
    If SGGrid1.Rows.Current.Type <> 0 Then
        Exit Sub
    End If
    
    If Len(mvarfObjectID) > 0 Then
        SelectTo
    Else
        EditObject m_KeyID
    End If
    
End Sub


Private Sub SGGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ActiveBar21.Bands("Band3").PopupMenu
    End If
    Exit Sub
End Sub


Private Sub clsGridShow1_OnTDBDropDownClose()
    'SetTheOthersAfterDropDownClose clsGridShow1.TDBDropDown1
End Sub

'在下拉网格控件关闭的时候设置除第一个元素之外的需要设置的元素到父网格控件上
'下拉网格在close时默认填充了第一个元素，并且只填充第一个。所以之后的需要填充的要手动制作。
'切记要结合开发环境中针对明细表的SQL，从基础资料中获取的字段不可填充
'因为在使用adodc2.requery的时候会自动填充。否则在requery时候会出错，提示已经被更改。
Private Sub SetTheOthersAfterDropDownClose(ByRef vTDBDropDownCtl As TrueOleDBGrid80.TDBDropDown)
'    TDBGrid1.Columns("B_GoodsNameAlias").Value = vTDBDropDownCtl.Columns("B_GoodsNameAlias").Value
'    TDBGrid1.Columns("B_Specification").Value = vTDBDropDownCtl.Columns("B_Specification").Value
'    TDBGrid1.Columns("B_Producer").Value = vTDBDropDownCtl.Columns("B_Producer").Value
'
End Sub

