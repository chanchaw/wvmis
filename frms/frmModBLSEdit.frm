VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{0E8071F7-A7DD-47AC-95A9-365FFFE096DF}#1.0#0"; "UCButton.ocx"
Begin VB.Form frmModBLSEdit 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   Icon            =   "frmModBLSEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   4860
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5760
      _cx             =   10160
      _cy             =   8573
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
      BackColor       =   13557726
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   12
      ChildSpacing    =   2
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
      GridRows        =   1
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmModBLSEdit.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
         Height          =   4500
         Left            =   180
         TabIndex        =   2
         Top             =   180
         Width           =   3750
         _LayoutVersion  =   1
         _ExtentX        =   6615
         _ExtentY        =   7938
         _DataPath       =   ""
         Bands           =   "frmModBLSEdit.frx":03CF
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00CEDFDE&
         BorderStyle     =   0  'None
         Height          =   4500
         Left            =   3960
         ScaleHeight     =   4500
         ScaleWidth      =   1620
         TabIndex        =   1
         Top             =   180
         Width           =   1620
         Begin TA_UCButton.UCButton UCButton3 
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   1005
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   661
            Caption         =   "取消退出  "
            Icon            =   "frmModBLSEdit.frx":0597
            IconMask        =   "frmModBLSEdit.frx":082D
            CaptionAlignment=   1
         End
         Begin TA_UCButton.UCButton UCButton2 
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   495
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   661
            Caption         =   "保存退出  "
            Icon            =   "frmModBLSEdit.frx":0AC3
            IconMask        =   "frmModBLSEdit.frx":0E5D
            CaptionAlignment=   1
         End
         Begin TA_UCButton.UCButton UCButton1 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   0
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   661
            Caption         =   "保存新增  "
            Icon            =   "frmModBLSEdit.frx":11F7
            IconMask        =   "frmModBLSEdit.frx":1591
            CaptionAlignment=   1
         End
      End
   End
End
Attribute VB_Name = "frmModBLSEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_KeyID As Variant

Dim m_BillName As String
Dim m_SQL As String
Dim m_FWidth As Long
Dim m_FHeight As Long
Dim m_PKey As String        '主键
Dim m_Table As String    '单表中基础数据表（如果有外链接的情况）的表名

Dim clsCtlShow1 As New clsCtlShow

Private IsChange As Boolean
Private mvarObjectID As String '局部复制
Private A_strUnique As String  '多字段验证唯一性
Private A_rsFieldsNotNull As RecordSet  '多字段非空验证
Private A_HighterlevelFrm As Object

Private A_AutoFillRs As New RecordSet
Private A_rsCtl As RecordSet

Public Property Set AutoFillRs(ByVal vData As RecordSet)
    Set A_AutoFillRs = vData.Clone
End Property

Public Property Get AutoFillRs() As RecordSet
    Set AutoFillRs = A_AutoFillRs.Clone
End Property


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

Public Property Set HighterlevelFrm(ByVal vData As Object)

'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ObjectID = 5
    Set A_HighterlevelFrm = vData
End Property

Public Property Get HighterlevelFrm() As Object
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ObjectID
    Set HighterlevelFrm = A_HighterlevelFrm
End Property

Private Sub FillCtl()
    With clsCtlShow1
        .ObjectID = ObjectID
        .InitClass ActiveBar21, 1
        .Refresh
        
        Set A_rsCtl = .GetCtlPara_BLS(mvarObjectID)
    End With
    
End Sub

Public Sub AddNewObject(ByVal ObjectID As String)
    mvarObjectID = ObjectID
    
    FillCtl
    
    GetPTableKey
    IsChange = False
    
    
    '在新增的时候自动填充树形结构控件中的内容
    AutoFillCtls
End Sub

Public Sub EditObject(ByVal ObjectID As String)
    mvarObjectID = ObjectID
    FillCtl
    LoadObject
End Sub

'多字段唯一性确认
Private Function UniqueJudge() As Boolean
    If Len(A_strUnique) <= 0 Then
        UniqueJudge = True
        Exit Function
    End If
    
    Dim strSQL As String
    Dim rs As RecordSet
    Dim o As Object
    Dim m_Array
    Dim i As Long
    Dim m_ProcessName As String
    Dim m_TableName As String
    
    m_TableName = GetTableName(m_SQL)
    
    strSQL = "Select * From " & m_TableName & " Where 1=1"
    m_Array = Split(A_strUnique, ",")
    
    For i = 0 To UBound(m_Array)
        m_ProcessName = Me.Controls(Trim(m_Array(i))).Text
        'Debug.Print m_ProcessName
        strSQL = strSQL & " And " & m_Array(i) & "='" & m_ProcessName & "'"
    Next
    'Debug.Print strSQL
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount > 0 Then
        UniqueJudge = False
    Else
        UniqueJudge = True
    End If
    
    rs.Close
    Set rs = Nothing
End Function

'字段非空判断的初始化
Private Sub NotNullInit()
    Set A_rsFieldsNotNull = New RecordSet
    A_rsFieldsNotNull.Fields.Append "B_FieldName", adVarChar, 100
    A_rsFieldsNotNull.Fields.Append "B_Caption", adVarChar, 100
    A_rsFieldsNotNull.Open
    
    
    Dim rs As New RecordSet
    strSQL = "Select * From G_BLSFormTools "
    strSQL = strSQL & " Where B_ObjectID='" & ObjectID & "'"
    strSQL = strSQL & " And abs(isnull(B_NotNull,0))=1"
    
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    Do While Not rs.EOF
        A_rsFieldsNotNull.AddNew
        A_rsFieldsNotNull!B_FieldName = rs!B_FieldName
        A_rsFieldsNotNull!B_Caption = rs!B_Caption
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

'初始化多字段的唯一性（除了主键外的）
'例如：110009工序编辑中
'在染色工序和印花工序中可同时存在“缝头”工序
'但是在染色工序中不可同时存在两个或者两个以上的“缝头”工序
Private Sub UniqueInit()
    Dim strSQL As String
    Dim rs As RecordSet
    Dim lTemp As Integer
    
    Set rs = New RecordSet
    strSQL = "Select * From G_BLSFormTools "
    strSQL = strSQL & " Where B_ObjectID='" & ObjectID & "'"
    'strSQL = strSQL & " And abs(isnull(B_Unique,0))=1"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        A_strUnique = ""
        Exit Sub
    End If
    
    A_strUnique = ""
    Do While Not rs.EOF
        lTemp = Abs(IIf(IsNull(rs!B_Unique), 0, rs!B_Unique))
        If lTemp = 1 Then
            A_strUnique = A_strUnique & rs!B_FieldName & ","
        End If
        
        rs.movenext
    Loop
    If Right(A_strUnique, 1) = "," Then
        A_strUnique = Left(A_strUnique, Len(A_strUnique) - 1)
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

'从查询SQL中获取数据表名
Private Function GetTableName(ByVal m_str As String) As String
    Dim i, j As Long
    i = InStr(1, m_str, "from", vbTextCompare) + 4
    j = InStr(1, m_str, "where", vbTextCompare) - 1
    GetTableName = Trim(Mid(m_str, i, j - i))
End Function

Private Sub GetObjectParameter()
    Dim rs As New RecordSet
    Dim strSQL As String
    
    
    Set rs = New RecordSet
    strSQL = "Select * From G_BLS Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    m_BillName = rs("B_BillName")
    m_SQL = rs("B_SQL")
    m_FWidth = rs("B_FWidth")
    m_FHeight = rs("B_FHeight")
    
    Me.Caption = m_BillName
    Me.width = m_FWidth
    Me.height = m_FHeight
    
    rs.Close
    Set rs = Nothing
    
    GetPKey
    
    
    
    '多字段唯一性确认
    UniqueInit
    NotNullInit
End Sub

'调用对象
Private Sub LoadObject()
    On Error Resume Next
    Dim o As Object
    Dim strSQL As String
    Dim rs As New RecordSet

    '取得主键
    GetPTableKey
        
    Set rs = New RecordSet
    '" & m_Table & "."
    strSQL = m_SQL & " And " & m_Table & "." & m_PKey & "='" & m_KeyID & "'"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    clsCtlShow1.LoadObject rs

    rs.Close
    Set rs = Nothing
    
    IsChange = False
    
    
    
End Sub

Private Sub GetPTableKey()
    Dim strSQL As String
    Dim rs As New RecordSet
    
    '取得主键
    strSQL = m_SQL & " And 1=0"
    Debug.Print strSQL
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    m_PKey = rs.Fields(0).name
    m_Table = rs.Fields(0).Properties.item(1).Value
    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetPKey()
    Dim strSQL As String
    Dim rs As New RecordSet
    '取得主键
    Debug.Print m_SQL
    strSQL = m_SQL & " And 1=0"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    m_PKey = rs.Fields(0).name
    rs.Close
    Set rs = Nothing
End Sub
'保存对象
Private Function SaveObject() As Boolean
    On Error GoTo IFERR
    Dim o As Object
    Dim strSQL As String
    Dim rs As New RecordSet
    Dim rsForm As New RecordSet
    Dim i As Long
    
    If CheckInputItem = False Then
        SaveObject = False
        Exit Function
    End If
    
    
    If CtlIsNull = True Then
        SaveObject = False
        Exit Function
    End If
    
    '验证多字段唯一性
'    If UniqueJudge = False Then
'        MsgBox "违背多字段唯一性原则！", vbOKOnly + vbInformation, "提示"
'        SaveObject = False
'        Exit Function
'    End If
    
    
    strSQL = m_SQL & " And " & m_Table & "." & m_PKey & "='" & Trim(m_KeyID) & "'"
    Debug.Print strSQL
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount < 1 Then
        rs.AddNew
    End If
    
    clsCtlShow1.SaveObject rs

    rs.Close
    SaveObject = True
    IsChange = False
    Exit Function
IFERR:
    MsgBox Err.Description
    Set rs = Nothing
    Exit Function
End Function


Private Sub Form_Load()
    IsChange = False
    GetObjectParameter
    AnimateForm Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If ExitObject = False Then
        Cancel = 1
    End If
    clsCtlShow1.RemoveAll
    Set clsCtlShow1 = Nothing
    
    
    A_rsFieldsNotNull.Close
    Set A_rsFieldsNotNull = Nothing
End Sub

'保存新增
Private Sub UCButton1_Click()
    If SaveObject = True Then
        '由于编辑窗体显示为模态，所以下面代码不起效果，干脆注释掉。
'        If Not IsNull(A_HighterlevelFrm) Then
'            A_HighterlevelFrm.RefreshGrid
'        End If
        AddNewObject ObjectID
    End If
End Sub

Private Sub UCButton2_Click()
    If SaveObject = True Then
        Unload Me
    End If
End Sub

Private Function ExitObject() As Boolean
    Dim iReturn As Integer
    If IsChange = True Then
        IsChange = False
        iReturn = MsgBox("数据已经改变是否要保存?", vbExclamation + vbYesNoCancel + vbDefaultButton2, "保存")
        Select Case iReturn
            Case vbYes
                UCButton2_Click
                ExitObject = True
            Case vbNo
                ExitObject = True
        End Select
    Else
        ExitObject = True
    End If
End Function

Private Sub UCButton3_Click()
    Unload Me
End Sub

Private Function CheckInputItem() As Boolean
    Dim i As Long
    Dim o As Object
    For Each o In Me.Controls
        'Debug.Print o.name
        If Mid(o.name, 1, 2) = "B_" Then
            i = i + 1
            If Len(o.Text) < 1 Then
                MsgBox "请输入 - " & o.Caption
                CheckInputItem = False
                Set o = Nothing
                Exit Function
            End If
            If i = 1 Then
                Exit For
            End If
        End If
    Next
    CheckInputItem = True
    Set o = Nothing
End Function


'根据数据表G_UserCTData获取当前用户不可见的SQL语句
Private Function GetPBString() As String
    Dim strSQL As String
    Dim rs As RecordSet
    
    Set rs = New RecordSet
End Function

'在弹出编辑页面的时候自动填充某些字段
'例如：在树形结构中，选中某节点则弹出时候自动填充该节点
Private Sub AutoFillCtls()
    If A_AutoFillRs.State <> adStateOpen Then
        Exit Sub
    End If
    
    If A_AutoFillRs.RecordCount <= 0 Then
        Exit Sub
    End If
    
    
    Dim strSQL As String
    Dim rs As New RecordSet
    Dim szFieldName As String
    
    strSQL = "SELECT * FROM G_BLSFormTools WHERE B_ObjectID='" & mvarObjectID & "' AND abs(isnull(B_AutoFillParent,0))=1"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    Do While Not rs.EOF
        szFieldName = rs!B_FieldName
        Me.Controls(szFieldName).Text = A_AutoFillRs(0)
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub


'用于子控件
Public Sub Change(ByVal ObjectName As String, ByVal Operation As String)
    'AutoFillPY ObjectName
    
    'IsChange = True
    
    
'    Dim sz As String
'    Debug.Print "Operation=" & Operation
'    sz = Me.Controls("B_ClientName").Text
'    Debug.Print sz
'    sz = GetPYFirst(sz)
'    Me.Controls("B_ClientID").Text = sz

    AutoFillPY ObjectName
End Sub

'自动填充拼音
Private Sub AutoFillPY(ByVal ObjectName As String)
    On Error Resume Next
    Dim rs As New RecordSet
    
    Set rs = A_rsCtl.Clone
    
    If rs.RecordCount <= 0 Then
        Exit Sub
    End If
    
    rs.Filter = " B_FieldName='" & ObjectName & "'"
    If rs.RecordCount <= 0 Then
        Exit Sub
    End If
    
    Dim szPYTo As String
    Dim szPY As String
    Dim szCtlString As String
    
    szCtlString = Me.Controls(ObjectName).Text
    szPYTo = IIf(IsNull(rs!B_PYTo), "", rs!B_PYTo)
    If Len(szPYTo) > 0 Then
        szPY = GetPYFirst(szCtlString)
        Me.Controls(szPYTo).Text = szPY
    End If
    
End Sub

'判断控件是否为空
'空则返回TRUE，否则返回FALSE
Private Function CtlIsNull() As Boolean
    CtlIsNull = False
    
    '没有需要判断空的字段
    If A_rsFieldsNotNull.State <> adStateOpen Then
        CtlIsNull = False
        Exit Function
    End If
    
    If A_rsFieldsNotNull.RecordCount <= 0 Then
        CtlIsNull = False
        Exit Function
    End If
    

    Dim szValue As String, oCtl As Object
    Dim szField As String, szTip As String
    
        
    For Each oCtl In Me.Controls
        szField = oCtl.name
        If Left$(szField, 2) = "B_" Then
            A_rsFieldsNotNull.Filter = " B_FieldName='" & szField & "'"
            If A_rsFieldsNotNull.RecordCount > 0 Then
                szValue = Me.Controls(Trim(szField)).Text
                
                If Len(szValue) <= 0 Then
                    szTip = "[" & A_rsFieldsNotNull!B_Caption & "]不可为空！"
                    CtlIsNull = True
                    
                    MsgBox szTip, vbOKOnly + vbInformation, "提示"
                    Exit Function
                End If
            End If
            
        End If
    Next
    
End Function
