VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'保持属性值的局部变量
Private mvarSQL As String '局部复制
Public frm1 As Object
'保持属性值的局部变量
Private mvarObjectID As String '局部复制 报表编号
'保持属性值的局部变量
Private mvarReportObjectID As String '局部复制

Private m_tmpFileName As String
Private m_rtpFileText() As Byte
Private rs As New RecordSet
'保持属性值的局部变量
Private mvarKeyID As Variant '局部复制
Private m_ReportFileName As String
Public Property Let KeyID(ByVal vData As Variant)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.KeyID = 5
    mvarKeyID = vData
End Property


Public Property Set KeyID(ByVal vData As Variant)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.KeyID = Form1
    Set mvarKeyID = vData
End Property


Public Property Get KeyID() As Variant
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.KeyID
    If IsObject(mvarKeyID) Then
        Set KeyID = mvarKeyID
    Else
        KeyID = mvarKeyID
    End If
End Property

Public Property Let ReportObjectID(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ReportObjectID = 5
    mvarReportObjectID = vData
End Property

Public Property Get ReportObjectID() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ReportObjectID
    ReportObjectID = mvarReportObjectID
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

'保持属性值的局部变量
Public Property Let sql(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SQL = 5
    mvarSQL = vData
End Property
Public Property Get sql() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SQL
    sql = mvarSQL
End Property

Public Sub InitClass(ByRef o As Object, ByVal m_ObjectID As String, ByVal m_ReportObjectID As String)
    ReportObjectID = m_ReportObjectID
    ObjectID = m_ObjectID
    Set frm1 = o
End Sub

Public Sub Refresh()
    BuildSQL
End Sub

'由窗体中得到控件数据,生成SQL语句
Private Sub BuildSQL()
    On Error GoTo IFERR
    
    Dim strSQL As String
    Dim rs As New RecordSet
    Dim o As Object
    Dim sSQL As String
    
    Set rs = New RecordSet
    strSQL = "Select B_ExeSQL,B_ReportType From G_Report Where B_ObjectID='" & Trim(mvarReportObjectID) & "'"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        'MsgBox "报表未设置!", vbExclamation, "报表"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    sSQL = rs("B_ExeSQL") & " "
    
    If rs("B_ReportType") = 1 Then
        '单据类报表
        mvarSQL = sSQL & KeyID
        Debug.Print mvarSQL
    Else
        '普通报表
        Set rs = New RecordSet
        strSQL = "Select * From G_BLRFormTools Where B_ObjectID='" & mvarObjectID & "'"
        strSQL = strSQL & " Order By B_BandIndex,B_ToolIndex"
        
        rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
        'Debug.Print strSQL
        Do While Not rs.EOF
            Debug.Print rs!B_FieldName & "=" & rs!B_SubObject
            Select Case IIf(IsNull(rs!B_USET), 0, rs!B_USET)
                Case 1 '链接：00:00:00
                    sSQL = sSQL & "'" & Trim(frm1.Controls(rs("B_CtlName")).Text) & " 00:00:00',"
                Case 2 '链接：23:59:59
                    sSQL = sSQL & "'" & Trim(frm1.Controls(rs("B_CtlName")).Text) & " 23:59:59',"
                Case Else
                    sSQL = sSQL & "'" & Trim(frm1.Controls(rs("B_CtlName")).Text) & "',"
            End Select
            Debug.Print sSQL
            rs.movenext
        Loop
        
        mvarSQL = Mid(sSQL, 1, Len(sSQL) - 1)
        Debug.Print mvarSQL
    
        rs.Close
        Set rs = Nothing
    
    End If
    Exit Sub
    
IFERR:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Public Function GetValueParameter() As String
    Dim sSQL As String
    Dim strSQL As String
    Dim o As Object
    
    Set rs = New RecordSet
    strSQL = "Select B_CtlName From G_BLRFormTools Where B_ObjectID='" & mvarObjectID & "'"
    strSQL = strSQL & " Order By B_BandIndex,B_ToolIndex"
    
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        For Each o In frm1
            If Trim(rs("B_CtlName")) = Trim(o.name) Then
                sSQL = sSQL & Trim(o.Text) & ","
            End If
        Next
        rs.movenext
    Loop
    
    GetValueParameter = Mid(sSQL, 1, Len(sSQL) - 1)
    
    rs.Close
    Set rs = Nothing
End Function

'打印预览
Public Sub PrintPreview()
'    Dim frm1 As New frmModBLRPreview
'    With frm1
'        BuildSQL
'        .m_ObjectID = ReportObjectID
'        .m_SQL = SQL
'        .Show vbModal
'    End With

    Dim frm1 As New frmModBLRPreviewOri
    Dim rs As New RecordSet
    rs.Open mvarSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    With frm1
        Set .RecordSet = rs.Clone
        .ObjectID = mvarReportObjectID
        '.Show vbModal
        .Show
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Public Sub PrintPreviewBrs(ByRef vRs As RecordSet)

    Dim frm1 As New frmModBLRPreviewOri
    With frm1
        Set .RecordSet = vRs.Clone
        .ObjectID = mvarReportObjectID
        .Show vbModal
    End With
End Sub

'直接打印
Public Sub PrintDirectMine(ByVal m_SQL As String, ByVal ReportObjectID As String)
    On Error GoTo IFERR
    Dim rs1 As New RecordSet
    Dim strSQL As String
    Dim m_SQLString As String
    Dim m_rptUType As Integer
    
    'BuildSQL
    
    Set rs1 = New RecordSet
    strSQL = "Select Top 1 * From G_Report Where B_ObjectID='" & Trim(ReportObjectID) & "'"
    Debug.Print strSQL
    rs1.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs1.EOF Then
        m_rptUType = rs1("B_ReportUType")
        If m_rptUType = True Then
            m_ReportFileName = Trim(App.Path & "\" & rs1("B_ReportFileName"))
        Else
            m_rtpFileText = rs1("B_ReportImage")
        End If
    Else
        MsgBox "报表未设置!", vbExclamation, "报表"
        rs1.Close
        Set rs = Nothing
        Exit Sub
    End If

    Set rs = New RecordSet
    Set rs = Gm.cnnTool.cnn.Execute(m_SQL)
    
    If m_rptUType = 1 Then
        ShowReport
    Else
        ShowReportAsText
    End If
    
    rs1.Close
    Set rs1 = Nothing
    Exit Sub
IFERR:
    MsgBox Err.Description
    Exit Sub
End Sub



'直接打印
Public Sub PrintDirect()
    On Error GoTo IFERR
    Dim rs1 As New RecordSet
    Dim strSQL As String
    Dim m_SQLString As String
    Dim m_rptUType As Integer
    
    BuildSQL
    
    Set rs1 = New RecordSet
    strSQL = "Select Top 1 * From G_Report Where B_ObjectID='" & Trim(ReportObjectID) & "'"
    rs1.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs1.EOF Then
        m_rptUType = rs1("B_ReportUType")
        If m_rptUType = True Then
            m_ReportFileName = Trim(App.Path & "\" & rs1("B_ReportFileName"))
        Else
            m_rtpFileText = rs1("B_ReportImage")
        End If
    Else
        MsgBox "报表未设置!", vbExclamation, "报表"
        rs1.Close
        Set rs = Nothing
        Exit Sub
    End If

    Set rs = New RecordSet
    Set rs = Gm.cnnTool.cnn.Execute(sql)
    
    If m_rptUType = 1 Then
        ShowReport
    Else
        ShowReportAsText
    End If
    
    rs1.Close
    Set rs1 = Nothing
    Exit Sub
IFERR:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub ShowReport()
    On Error GoTo IFERR
    Dim rpt1 As New ActiveReport1
    With rpt1
        .DataControl1.RecordSet = rs
        .LoadLayout m_ReportFileName
        .PrintReport False
    End With
    Exit Sub
IFERR:
    Exit Sub
End Sub

Private Sub ShowReportAsText()
    On Error GoTo IFERR
    Dim m_tmpFileName As String
    Dim rpt1 As New ActiveReport1
    
    m_tmpFileName = Trim(App.Path) & "\" & mvarObjectID & ".xml"

    Open m_tmpFileName For Binary As #1
    Put #1, , m_rtpFileText()
    Close #1

    With rpt1
        .Refresh
        .DataControl1.RecordSet = rs
        .LoadLayout m_tmpFileName
        .PrintReport False
    End With
    

    Exit Sub
IFERR:
    Exit Sub
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    fso.DeleteFile (m_tmpFileName)
End Sub




'直接打印    可自选报表窗体
'vID:计划单表G_CJBill中的B_ID
Public Sub PrintDirectMine_ARFrmName(ByVal m_SQL As String, _
    ByVal ReportObjectID As String, ByVal vARFrmName As String, _
    ByVal vID As Long, Optional ByVal vPrintPreview As Boolean = False)
    
    On Error GoTo IFERR
    Dim rs1 As New RecordSet
    Dim strSQL As String
    Dim m_SQLString As String
    Dim m_rptUType As Integer
    
    'BuildSQL
    
    Set rs1 = New RecordSet
    strSQL = "Select Top 1 * From G_Report Where B_ObjectID='" & Trim(ReportObjectID) & "'"
    Debug.Print strSQL
    rs1.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs1.EOF Then
        m_rptUType = rs1("B_ReportUType")
        If m_rptUType = True Then
            m_ReportFileName = Trim(App.Path & "\" & rs1("B_ReportFileName"))
        Else
            m_rtpFileText = rs1("B_ReportImage")
        End If
    Else
        MsgBox "报表未设置!", vbExclamation, "报表"
        rs1.Close
        Set rs = Nothing
        Exit Sub
    End If

    Set rs = New RecordSet
    Set rs = Gm.cnnTool.cnn.Execute(m_SQL)
    
    
    If m_rptUType = 1 Then
        ShowReport
    Else
        ShowReportAsText_ARFrmName vARFrmName, vID, vPrintPreview
    End If
    
    rs1.Close
    Set rs1 = Nothing
    Exit Sub
IFERR:
    MsgBox Err.Description
    Exit Sub
End Sub


Private Sub ShowReportAsText_ARFrmName(ByVal vARFrmName As String, ByVal vID As Long, _
    Optional ByVal vPrintPreview As Boolean = False)
    
    On Error GoTo IFERR
    Dim m_tmpFileName As String
    Dim rpt1 As Object
    
    'GetFormObject rpt1, vARFrmName
    Set rpt1 = GetFormNew(vARFrmName)
    
    m_tmpFileName = Trim(App.Path) & "\" & mvarObjectID & ".xml"

    Open m_tmpFileName For Binary As #1
    Put #1, , m_rtpFileText()
    Close #1

    With rpt1
        .W_ID = vID
        .Refresh
        .DataControl1.RecordSet = rs
        .LoadLayout m_tmpFileName
        
        '.Show vbModal
        
        If vPrintPreview = True Then
            .Show vbModal
        Else
            .PrintReport False
        End If
    End With
    

    Exit Sub
IFERR:
    Dim szErr As String
    MsgBox "错误发生于打印报表中。" & Err.Description
    Exit Sub
End Sub

'获取公用报表的查询SQL
Public Function GetReportSQL(ByVal vObjectID As String) As String
Attribute GetReportSQL.VB_Description = "获取开发环境中公用报表对象的查询SQL"
    Dim rs As New RecordSet
    Dim strSQL As String
    
    strSQL = "Select * From G_Report where B_ObjectID='" & vObjectID & "'"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        GetReportSQL = ""
    Else
        GetReportSQL = rs!B_ExeSQL
    End If
    
    rs.Close
    Set rs = Nothing
End Function


