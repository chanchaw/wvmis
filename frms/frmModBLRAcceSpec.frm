VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmModBLRAcceSpec 
   Caption         =   "指定合同的所有辅料计划数据"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmModBLRAcceSpec.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   6195
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10200
      _cx             =   17992
      _cy             =   10927
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   800
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   3
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmModBLRAcceSpec.frx":058A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
         Height          =   5580
         Left            =   0
         TabIndex        =   1
         Top             =   420
         Width           =   10200
         _LayoutVersion  =   1
         _ExtentX        =   17992
         _ExtentY        =   9843
         _DataPath       =   ""
         Bands           =   "frmModBLRAcceSpec.frx":05E2
         Begin VB.PictureBox PctBack 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   480
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   180
            Visible         =   0   'False
            Width           =   255
            Begin VB.CommandButton btPopUpWindow 
               Caption         =   "7"
               BeginProperty Font 
                  Name            =   "Marlett"
                  Size            =   9
                  Charset         =   2
                  Weight          =   500
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   4
               Top             =   0
               Visible         =   0   'False
               Width           =   255
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   2955
            Left            =   1440
            TabIndex        =   2
            Top             =   1080
            Width           =   6435
            _cx             =   11351
            _cy             =   5212
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16744448
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar2 
         Height          =   420
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   10200
         _LayoutVersion  =   1
         _ExtentX        =   17992
         _ExtentY        =   741
         _DataPath       =   ""
         Bands           =   "frmModBLRAcceSpec.frx":07AA
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   3900
            Top             =   5640
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
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
      End
      Begin VB.Label 状态 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   0
         TabIndex        =   6
         Top             =   6000
         Width           =   5040
      End
   End
End
Attribute VB_Name = "frmModBLRAcceSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mvarObjectID As String
Public aValue
Private clsReport1 As New clsReport
Private clsCtlShow1 As New clsCtlShow
Private clsVsFlexGrid1 As New clsVsFlexGrid
Private m_ReportObjectID As String

Private m_GroupFields As String
Private m_SumFields As String
Private m_SearchString As String
Private m_NextObject As String

Private iKeyIndex As Integer
Private iCommandIndex As Integer


'==============================
Public fatherFrm As Object
Private mvarfObjectID As String
Private mvarfFieldName As String
Private mvarSendIndex As Integer

Private mvarBillOrDetail As Integer '0 为主表  1为明细表
'==============================


'2012-7-31报表打印权限设置的参数
'==================================
Private m_IsPrintObject As Boolean
Private m_FieldNamePermission As String '数据表中设置可打印权限的字段名称
Private m_RoleBill As String  '主表的角色名称
Private m_DetailName As String  '明细表名称
'==================================

Private A_rsCloneCellColor As RecordSet
Private A_rsObjectPara As New RecordSet

Private strSQL As String

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


Private Sub ActiveBar2_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "导出Excel"
            ExportExcel
        Case "刷新"
            PriviewGrid
        Case "打印"
            PrintPreview
        Case "收缩"
            clsVsFlexGrid1.SetCollapsed VSFlexGrid1
        Case "展开"
            clsVsFlexGrid1.SetExpanded VSFlexGrid1
        Case "筛选"
            FilterForm
        Case "关闭"
            Unload Me
            
        Case "保存样式"
            clsVsFlexGrid1.SaveColWidth
    End Select
End Sub


Private Sub FilterForm()
    Dim frm1 As New frmUCTSearch
    With frm1
        Set .rs = Adodc1.RecordSet
        .ObjectID = mvarObjectID
        .FieldType = 5
        .Show vbModal
    End With
    If frm1.OK = True Then
        Adodc1.RecordSet.Filter = ""
        Adodc1.RecordSet.Filter = frm1.strResult
        Debug.Print Adodc1.RecordSet.RecordCount
        ShowSGGrid
    End If
    Unload frm1
    Set frm1 = Nothing
End Sub

Private Sub PrintPreview()
    clsReport1.PrintPreview
End Sub



Private Sub Form_Load()
    ActiveBar21.ClientAreaControl = VSFlexGrid1
    ActiveBar21.RecalcLayout

    GetObjectParameter
    AnimateForm Me

    InitCtl
    
    
    CreatePopUpButton
    
    InitDateTime
End Sub

Public Sub LoadObject()
    LoadCtlParameter
    PriviewGrid
    
    
    '有最大化导航的时候才用到
    Me.WindowState = 2
End Sub

'取得Ctl参数
Private Sub LoadCtlParameter()
    'On Error GoTo IFERR
    Dim rs As New RecordSet
    Dim strSQL As String
    Dim i As Long
    Dim o As Object
    
    Set rs = New RecordSet
    strSQL = "Select * From G_BLRFormTools Where B_ObjectID='" & mvarObjectID & "'"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    i = 0
    Do While Not rs.EOF
        For Each o In Me.Controls
            If o.name = rs("B_CtlName") Then
                If i <= UBound(aValue) Then
                    o.Text = aValue(i)
                    i = i + 1
                Else
                    Exit Sub
                End If
            End If
        Next
        rs.movenext
    Loop
    
    Set o = Nothing
    rs.Close
    Set rs = Nothing
    Exit Sub
IFERR:
    Set rs = Nothing
End Sub

Private Sub PriviewGrid()
    On Error GoTo IFERR
    With clsReport1
        .InitClass Me, mvarObjectID, m_ReportObjectID
        .Refresh
    End With
    
    With Adodc1
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .CommandType = adCmdText
        
        
        .ConnectionString = Gm.cnnTool.cnnStr
        .CommandType = adCmdText
        .CommandTimeout = 0
        .RecordSource = clsReport1.sql
        Debug.Print .RecordSource
        .Refresh
    End With
    ShowSGGrid
    GetKeyIndex
    GetCommandIndex


    Exit Sub
IFERR:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub ShowSGGrid()
    Set clsVsFlexGrid1 = New clsVsFlexGrid
    With clsVsFlexGrid1
        .InitCls mvarObjectID, VSFlexGrid1
        .FillGrid Adodc1.RecordSet
    End With
End Sub

Private Sub InitCtl()
    With clsCtlShow1
        .ObjectID = mvarObjectID
        .InitClass ActiveBar21, 3
        .Refresh
    End With
End Sub

'取得参数
Private Sub GetObjectParameter()
    Dim rs As New RecordSet
    Dim strSQL As String
    
    Set rs = New RecordSet
    strSQL = "Select * From G_BLR Where B_ObjectID='" & mvarObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    m_ReportObjectID = rs("B_ReportObject")
    
    Me.width = rs("B_Width")
    Me.height = rs("B_Height")
    Me.Caption = rs("B_ReportName")
    m_NextObject = rs("B_NextObject")
    
    
    m_GroupFields = IIf(IsNull(rs("B_GroupFields")), 0, rs("B_GroupFields"))
    m_SumFields = IIf(IsNull(rs("B_SumFields")), 0, rs("B_SumFields"))
    
    
    Set A_rsObjectPara = rs.Clone
    
    
    rs.Close
    Set rs = Nothing
End Sub


Public Sub Change(ByVal sCtl As String, ByVal sCommand As String)
    'PriviewGrid
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dispose
End Sub


Private Sub GetKeyIndex()
    'iKeyIndex
    Dim i As Integer
    For i = 0 To Adodc1.RecordSet.Fields.Count - 1
        If Adodc1.RecordSet.Fields(i).Properties.item(4).Value = True Then
            iKeyIndex = i
            Exit Sub
        End If
    Next
End Sub

Private Sub GetCommandIndex()
    'iKeyIndex
    Dim i As Integer
    For i = 0 To Adodc1.RecordSet.Fields.Count - 1
        If Adodc1.RecordSet.Fields(i).name = "Command" Then
            iCommandIndex = i
            Exit Sub
        End If
    Next
End Sub


'初始化时间
Private Sub InitDateTime()
    On Error Resume Next
    Dim oControl As Object
    Dim rs As RecordSet
    Dim strSQL As String
    
    
    Set rs = New RecordSet
    strSQL = "Select * From G_CJDefaultTime Where B_GroupName='报表默认时间'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    '没有设置参数的话,直接退出本函数
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    
    
    For Each oControl In Me.Controls
        Select Case oControl.name
            Case "B_SDate"
                If Len(IIf(IsNull(rs("B_SDate")), "", rs("B_SDate"))) > 0 Then
                    oControl.Text = Format(rs("B_SDate"), "YYYY-MM-DD")
                Else
                    oControl.Text = Format(Now, "YYYY-MM-DD")
                End If
                
            Case "B_STime"
                If Len(IIf(IsNull(rs("B_STime")), "", rs("B_STime"))) > 0 Then
                    oControl.Value = Format(rs("B_STime"), "HH:MM:SS")
                Else
                    oControl.Value = Format(Now, "HH:MM:SS")
                End If
                
            Case "B_EDate"
                If Len(IIf(IsNull(rs("B_EDate")), "", rs("B_EDate"))) > 0 Then
                    oControl.Text = Format(rs("B_EDate"), "YYYY-MM-DD")
                Else
                    oControl.Text = Format(Now, "YYYY-MM-DD")
                End If
                
            Case "B_ETime"
                If Len(IIf(IsNull(rs("B_ETime")), "", rs("B_ETime"))) > 0 Then
                    oControl.Value = Format(rs("B_ETime"), "HH:MM:SS")
                Else
                    oControl.Value = Format(Now, "HH:MM:SS")
                End If
                
        End Select
    Next
    
    
    rs.Close
    Set rs = Nothing
    
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



'选择性弹出右键菜单
Private Sub PopUpRightMenu()
    On Error Resume Next
    Dim strSQL As String
    Dim rs As RecordSet
    Dim szPopUpMenuName As String
    
    '1. 获取设置的右键菜单
    Set rs = New RecordSet
    strSQL = "Select * From G_PopUpMenuOnBLR Where B_ObjectID='" & mvarObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    
    szPopUpMenuName = rs("B_BandName")
    
    rs.Close
    Set rs = Nothing
    
    
    
    '2. 弹出菜单
    If Len(Trim$(szPopUpMenuName)) <= 0 Then
        Exit Sub
    End If
    
    ActiveBar2.Bands(szPopUpMenuName).PopupMenu
End Sub


'以下为2012-2-12添加代码:
'自动填充弹出窗体的按钮
'======================================
'生成按钮
Private Sub CreatePopUpButton()
    Dim strSQL As String
    Dim rs As RecordSet
    Dim rs1 As RecordSet
    Dim i As Long
    Dim szTemp As String
    
    Set rs = New RecordSet
    strSQL = "Select * From G_PopUpWindowSet Where B_fObjectID='" & mvarObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    
    'GetLeft
    
    
    i = 0
    Do While Not rs.EOF
        i = i + 1
        Load btPopUpWindow(i)
        Load PctBack(i)
        
        PctBack(i).Visible = True
        btPopUpWindow(i).Visible = True
        
        
        Set rs1 = New RecordSet
        szTemp = IIf(IsNull(rs("B_PositionField")), "", rs("B_PositionField"))
        If Len(szTemp) <= 0 Then
            MsgBox "该对象被设置有弹出窗体的查询接口没有设置按钮所在位置！", vbOKOnly + vbInformation, "提示"
            Exit Sub
        End If
        
        'strSQL = "Select * From G_BLRFormTools Where B_FieldName='" & rs("B_fControlName") & "' And B_ObjectID='" & mvarObjectID & "'"
        strSQL = "Select * From G_BLRFormTools Where B_FieldName='" & szTemp & "' And B_ObjectID='" & mvarObjectID & "'"
        Debug.Print strSQL
        rs1.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        
        SetParent PctBack(i).hwnd, ActiveBar21.Bands("栏目" & Trim(rs1("B_BandIndex"))).Tools(Trim(rs1("B_FieldName"))).hwnd
        SetParent btPopUpWindow(i).hwnd, PctBack(i).hwnd

        btPopUpWindow(i).Top = 0
        btPopUpWindow(i).Left = 0
        
        PctBack(i).Top = 20
        'PctBack(i).left = Me.Controls(rs("B_fControlName")).left + Me.Controls(rs("B_fControlName")).Width - btPopUpWindow(0).Width
        PctBack(i).Left = Me.Controls(rs("B_PositionField")).Left + Me.Controls(rs("B_PositionField")).width - btPopUpWindow(0).width
    
    
        PctBack(i).width = PctBack(0).width
        PctBack(i).height = PctBack(0).height
        btPopUpWindow(i).height = btPopUpWindow(0).height
        btPopUpWindow(i).width = btPopUpWindow(0).width
        
        btPopUpWindow(i).Tag = rs("B_tObjectID") & "," & rs("B_fObjectID") & "," & rs("B_fControlName") & "," & rs("B_SendIndex")
        
        
        rs1.Close
        Set rs1 = Nothing
        
            
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub


Private Sub btPopUpWindow_Click(Index As Integer)
    Dim clsCommand1 As New clsCommand
    Dim m_Array
    
    m_Array = Split(btPopUpWindow(Index).Tag, ",")
    
    
    clsCommand1.InitClass
    '数据表G_PopUpWindowSet  B_fObjectID  B_fControlName B_SendIndex  B_tObjectID
    clsCommand1.ExecutePopUp 0, Me, m_Array(1), m_Array(2), m_Array(3), m_Array(0), " ", "LoadObject", Nothing, ""
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
   
   

'在2016年12月26日 13:01:39制作的替换掉原来的PrintPreview
'需要在打印预览的时候对特性控件进行修改
Private Sub PrintPreviewEx()
    Dim rpt1 As New ActiveReport1
    Dim cls1 As New clsPrint
    
    
    Dim szReportFile As String
    szReportFile = cls1.DownloadReport(m_ReportObjectID)
    rpt1.Caption = "打印预览"
    rpt1.WindowState = 2
    With rpt1
        .Refresh
        Set .DataControl1.RecordSet = Adodc1.RecordSet.Clone
        .LoadLayout szReportFile
        .Show
    End With
    
    Set rpt1 = Nothing
End Sub

Private Sub PrintPreviewEx01()
    Dim frm1 As New frmModBLRPreviewOri
    frm1.ObjectID = m_ReportObjectID
    Set frm1.RecordSet = Adodc1.RecordSet.Clone
    
    frm1.Show
End Sub

Private Sub ExportExcel()
    Dim cls1 As New clsFile
    Dim szFilePath As String
    szFilePath = cls1.ShowSaveFileDialog("Excel文件(*.xls)|*.xls")
    
    
    clsVsFlexGrid1.ExportExcel szFilePath
End Sub

Private Sub VSFlexGrid1_DblClick()
    OpenBL
End Sub

Private Sub OpenBL()
    On Error Resume Next
    
    Dim clsCommand1 As New clsCommand
    Dim aCommand
    Dim szCommand As String
    szCommand = VSFlexGrid1.cell(flexcpText, VSFlexGrid1.Row, clsVsFlexGrid1.GetColIndex("Command"))
    aCommand = Split(szCommand, ",")
    
    If UBound(aCommand) > 3 Then
        clsCommand1.InitClass
        clsCommand1.Execute aCommand(0), aCommand(1), aCommand(2), Nothing, aCommand(4)
    End If
    
End Sub

Private Sub VSFlexGrid1_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    If VSFlexGrid1.cell(flexcpText, Row, Col) = "0" Then
        VSFlexGrid1.cell(flexcpText, Row, Col) = ""
    End If
End Sub


Private Sub Dispose()
    clsVsFlexGrid1.Dispose
    Gm.CacheFrms.DelFrm mvarObjectID
End Sub

