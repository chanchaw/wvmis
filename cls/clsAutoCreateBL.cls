VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoCreateBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'本类用于自动生成一个单据对象(开发环境中开头为12的单据对象)
'本代码的初次使用见珍盈针织

Private A_ObjectID As String   '即将被生成的单据对象的对象编号
Private A_TNBill As String     '正式主表表名
Private A_TNDetail As String   '正式明细表表名
Private A_TNDraftBill As String '草稿主表表名
Private A_TNDraftDetail As String '草稿明细表表名
Private A_BID As String    '单据编号的英文前缀
Private A_ID As Long   '创建的主表的B_ID
Private A_strItemID As String   '通过CreateOneDetail和通过CreateDetailsByRs新建的明细数据的B_ItemID
                                '都会添加到该字符串内

Private strSQL As String
Private A_Inited As Boolean '初始化标记


'获取套表以及主键的实体类
Public Function GetSuitTable() As dmSuitTables
    Dim o As dmSuitTables
    With o
        .Bill = A_TNBill
        .DraftBill = A_TNDraftBill
        .Detail = A_TNDetail
        .DraftDetail = A_TNDraftDetail
    End With
    
    Set GetSuitTable = o
End Function


'初始化本类
Public Sub InitCls(ByVal vObjectID As String)
    A_strItemID = ""
    A_ObjectID = vObjectID
    GetObjectPara
    
    A_Inited = True
End Sub

'获取目标对象的一些参数
Private Sub GetObjectPara()
    Dim rs As RecordSet
    strSQL = "Select * From G_BL Where B_ObjectID='" & A_ObjectID & "'"
    Set rs = New RecordSet
    'rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        A_TNBill = ""
        A_TNDetail = ""
        A_TNDraftBill = ""
        A_TNDraftDetail = ""
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    
    A_TNBill = rs!B_MainTable
    A_TNDetail = rs!B_DetailTable
    A_TNDraftBill = rs!B_DraftMainTable
    A_TNDraftDetail = rs!B_DraftDetailTable
    A_BID = rs!B_BID
    
    rs.Close
    Set rs = Nothing
End Sub


'将草稿数据生成为正式数据
Public Sub Update()
    '1.先将草稿主表数据拷贝到正式主表
    strSQL = "Insert Into " & A_TNBill
    strSQL = strSQL & " Select * From " & A_TNDraftBill & " Where B_ID=" & A_ID
    Gm.cnnTool.cnn.Execute strSQL
    
    '创建草稿主表的时候B_Closed=0，那么登帐之后需要修改为1
    strSQL = "Update " & A_TNBill & " set B_Closed=1 where B_ID=" & A_ID
    Gm.cnnTool.cnn.Execute strSQL
    
    
    '2.再将草稿明细表的数据拷贝到正式明细表
    If Right$(A_strItemID, 1) = "," Then
        A_strItemID = Left$(A_strItemID, Len(A_strItemID) - 1)
    End If
    
    strSQL = "Insert Into " & A_TNDetail
    strSQL = strSQL & " Select * From " & A_TNDraftDetail & " Where B_ItemID In"
    strSQL = strSQL & " (" & A_strItemID & ")"
    Debug.Print strSQL
    Gm.cnnTool.cnn.Execute strSQL
    
    
    '3.删除草稿表的数据
    strSQL = "Delete From " & A_TNDraftDetail & " Where B_ItemID In (" & A_strItemID & ")"
    Gm.cnnTool.cnn.Execute strSQL
    
    strSQL = "Delete From " & A_TNDraftBill & " Where B_ID=" & A_ID
    Gm.cnnTool.cnn.Execute strSQL
    
End Sub

'在生成后,自动打开出库码单
Private Sub OpenBL(ByVal vObjectID As String, ByVal vBillName As String, ByVal vID As String)
    Dim clsCommand1 As New clsCommand
    clsCommand1.InitClass
    clsCommand1.Execute vObjectID, vBillName, "LoadObject", Nothing, vID
End Sub

Public Sub openbill()
    OpenBL A_ObjectID, "单据", A_ID
End Sub

'创建一个草稿主表
'vData：带有两种分隔符的数据。类似：B_ClientID="0001",B_Memo="暂时入库单"
'       第一个分隔符间隔每个字段，第二个分隔符间隔字段名称和VALUE，将这些数据写入新建的主表
'本类根据初始化时提供的A_ObjectID获取的数据有：B_ObjectID，B_CodeID，B_Date（当前时间），B_BID，B_Closed
Public Function CreateOneBill(ByVal vData As String, ByVal vSymFilter01 As String, ByVal vSymFilter02 As String) As Long
    Dim i As Long
    Dim rs As New RecordSet
    Set rs = String2Recordset(vData, vSymFilter01, vSymFilter02)
    
    Dim dtDate As String
    Dim rsBill As RecordSet
    Set rsBill = New RecordSet
    strSQL = "Select * From " & A_TNDraftBill & " Where 1=0"
    rsBill.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    dtDate = Format(Now, "YYYY-MM-DD HH:MM:SS")
    
    
    If rs.RecordCount > 0 Then
        '创建新的一个草稿主表之前清空下列2个变量
        A_strItemID = ""
        A_ID = -999999
        
        rsBill.AddNew
        rsBill!B_ObjectID = A_ObjectID
        rsBill!B_Codeid = GetFrameCodeDetail(A_ObjectID)
        rsBill!B_Date = dtDate
        rsBill!B_BID = A_BID
        rsBill!B_Closed = 1
            
        rs.MoveFirst
        Do While Not rs.EOF
            rsBill(Trim$(rs!B_Field01)) = rs!B_Field02
            rs.movenext
        Loop
        rsBill.Update
        A_ID = rsBill!B_id
    End If
    
    CreateOneBill = A_ID
    
    rsBill.Close
    Set rsBill = Nothing
End Function

'根据一个记录集创建一个主表
'在草稿主表A_TNDraftBill中创建一行数据，除了：B_ObjectID，B_CodeID，B_Date，B_BID，B_Closed
'系统级的字段，其他的字段按照参数vData来写入数据
'即遍历vData的字段在草稿报表中寻找对应的字段写入数据
Public Function CreateOneBillByRs(ByRef vData As RecordSet) As Long
    Dim i As Long
    Dim rs As New RecordSet
    Set rs = vData.Clone
    
    Dim dtDate As String
    Dim rsBill As RecordSet
    Set rsBill = New RecordSet
    strSQL = "Select * From " & A_TNDraftBill & " Where 1=0"
    rsBill.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    dtDate = Format(Now, "YYYY-MM-DD HH:MM:SS")
    
    
    If rs.RecordCount > 0 Then
        '创建新的一个草稿主表之前清空下列2个变量
        A_strItemID = ""
        A_ID = -999999
        
        rsBill.AddNew
        rsBill!B_ObjectID = A_ObjectID
        rsBill!B_Codeid = GetFrameCodeDetail(A_ObjectID)
        rsBill!B_Date = dtDate
        rsBill!B_BID = A_BID
        rsBill!B_Closed = 0
            
            
            
        '只创建一个主表
        rs.MoveFirst
        For i = 0 To rs.Fields.Count - 1
            If Not IsNull(rs(i)) Then
                rsBill(Trim$(rs(i).name)) = rs(i)
            End If
        Next
        rsBill.Update
        A_ID = rsBill!B_id
    End If
    
    CreateOneBillByRs = A_ID
    
    rsBill.Close
    Set rsBill = Nothing
End Function

'创建一条草稿明细,返回被创建的明细记录的B_ItemID
'一般本方法和CreateDetailsByRs二者选其一使用
Public Function CreateOneDetail(ByVal vData As String, ByVal vSymFilter01 As String, _
    ByVal vSymFilter02 As String) As Long
    
    If A_ID <= 0 Then
        Exit Function
    End If
    
    Dim rs As New RecordSet
    Set rs = String2Recordset(vData, vSymFilter01, vSymFilter02)
    If rs.RecordCount <= 0 Then
        Exit Function
    End If
    
    Dim rsdetail As New RecordSet
    strSQL = "Select * From " & A_TNDraftDetail & " Where 1=0"
    rsdetail.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rsdetail.AddNew
    rsdetail!B_id = A_ID
    
    rs.MoveFirst
    Do While Not rs.EOF
        rsdetail(Trim$(rs!B_Field01)) = rs!B_Field02
        rs.movenext
    Loop
    rsdetail.Update
    A_strItemID = A_strItemID & rsdetail!B_itemid & ","
    CreateOneDetail = rsdetail!B_itemid
    rsdetail.Close
    Set rsdetail = Nothing
End Function


'创建一条草稿明细数据。一般本方法和CreateOneDetail二者选其一使用
'vDataRs：需要写入的数据。字段对应DB中的表的字段
'返回值：返回创建的N条明细记录的B_ItemID字符串，中间以英文的逗号间隔。
Public Function CreateDetailsByRs(ByVal vDataRs As RecordSet) As String
    If A_ID <= 0 Then
        Exit Function
    End If
    
    If vDataRs.State <> adStateOpen Then
        Exit Function
    End If
    
    If vDataRs.RecordCount <= 0 Then
        Exit Function
    End If
    
    
    Dim i As Long
    Dim rsdetail As New RecordSet
    strSQL = "Select * From " & A_TNDraftDetail & " Where 1=0"
    rsdetail.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic

    
    vDataRs.MoveFirst
    Do While Not vDataRs.EOF
        rsdetail.AddNew
        rsdetail!B_id = A_ID
    
        For i = 0 To vDataRs.Fields.Count - 1
            rsdetail(Trim$(vDataRs(i).name)) = vDataRs(i)
            Debug.Print vDataRs(i).name & "=" & vDataRs(i)
        Next
        rsdetail.Update
        A_strItemID = A_strItemID & rsdetail!B_itemid & ","
        vDataRs.movenext
    Loop
    CreateDetailsByRs = A_strItemID
    rsdetail.Close
    Set rsdetail = Nothing
End Function

'将一个有2个分隔符的字符串转换为一个记录集
Private Function String2Recordset(ByVal vData As String, ByVal vSymFilter01 As String, ByVal vSymFilter02 As String) As RecordSet
    Dim rs As New RecordSet
    Dim arrArr01, arrArr02
    Dim i As Long, j As Long
    
    rs.Fields.Append "B_Field01", adVarChar, 100
    rs.Fields.Append "B_Field02", adVarChar, 100
    rs.Open
    
    
    arrArr01 = Split(vData, vSymFilter01)
    For i = 0 To UBound(arrArr01)
        arrArr02 = Split(arrArr01(i), vSymFilter02)
        
        rs.AddNew
        rs(0) = arrArr02(0)
        rs(1) = arrArr02(1)
        rs.Update
        
    Next
    
    Set String2Recordset = rs
End Function


'生成单据编号B_CodeID
Public Function GetFrameCodeDetail(ByVal m_ObjectID As String) As String
    On Error Resume Next
    Dim strTmpBH As String
    Dim strTmpBHLast As String
    Dim strTmpMonth As String
    Dim strTmpDay As String
    Dim mstrSQL As String
    Dim rs As New RecordSet
    Dim gdateSystemDat As Date
    Dim rstemp As RecordSet
    Dim mvarm_BID As String
    Dim strSQL As String
    Dim m_DraftMainTable As String
    Dim m_MainTable As String
    
    Set rstemp = New RecordSet
    strSQL = "Select * From G_BL Where B_ObjectID='" & m_ObjectID & "'"
    Debug.Print strSQL
    
    mvarm_BID = A_BID
    m_DraftMainTable = A_TNDraftBill
    m_MainTable = A_TNBill
    
    
    rstemp.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    mvarm_BID = rstemp("B_BID")
    m_DraftMainTable = rstemp("B_DraftMainTable")
    m_MainTable = rstemp("B_MainTable")
    rstemp.Close
    Set rstemp = Nothing


   

    gdateSystemDat = Now

    Set rs = New RecordSet
    strTmpMonth = Trim(Month(gdateSystemDat))
    If Len(Trim(strTmpMonth)) = 1 Then
        strTmpMonth = "0" & strTmpMonth
    End If
    strTmpDay = Trim(Day(gdateSystemDat))
    If Len(Trim(strTmpDay)) = 1 Then
        strTmpDay = "0" & strTmpDay
    End If

    strTmpBH = Trim(mvarm_BID) & Trim(Year(gdateSystemDat)) & strTmpMonth & strTmpDay
    Debug.Print strTmpBH

    mstrSQL = "Select CASE WHEN ISNULL(P1.B_CodeID,0)>ISNULL(P2.B_CodeID,0) THEN P1.B_CodeID"
    mstrSQL = mstrSQL & " Else P2.B_CodeID End as B_PCodeID"
    mstrSQL = mstrSQL & " From (Select Max(B_CodeID) as B_CodeID"
    mstrSQL = mstrSQL & " From " & m_DraftMainTable
    mstrSQL = mstrSQL & " Where B_CodeID Like '" & Trim(strTmpBH) & "%') as P1,"

    mstrSQL = mstrSQL & " (Select Max(B_CodeID) as B_CodeID"
    mstrSQL = mstrSQL & " From " & m_MainTable
    mstrSQL = mstrSQL & " Where B_CodeID Like '" & Trim(strTmpBH) & "%') as P2"
    Debug.Print mstrSQL
    rs.Open mstrSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    If IsNull(rs(0)) Then
        '如果没有以前的记录
        strTmpBH = strTmpBH & "0001"
        GetFrameCodeDetail = strTmpBH
    Else
        strTmpBHLast = Trim(str(Val(Mid(Trim(rs(0)), 12, 4)) + 1))
        If Len(Trim(strTmpBHLast)) < 4 Then
            strTmpBH = strTmpBH & String(4 - Len(Trim(strTmpBHLast)), "0") & strTmpBHLast
        Else
            strTmpBH = strTmpBH & strTmpBHLast
        End If
        GetFrameCodeDetail = strTmpBH
    End If
    rs.Close
    Set rs = Nothing
End Function

'已经登帐的单据通过本方法可以向正式表中追加一行明细记录
'Private Function AppendDraftDetailByRs(ByVal vID As Long, ByVal vRs As Recordset) As Recordset
'    Dim rtnRs As Recordset
'    Set rtnRs = New Recordset
'    rtnRs.Fields.Append "B_ItemID", adInteger
'    rtnRs.Open
'
'
'
'    Dim i As Long
'    Dim rsDetail As New Recordset
'    strSQL = "Select * From " & A_TNDraftDetail & " Where 1=0"
'    rsDetail.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
'
'
'    vRs.MoveFirst
'    Do While Not vRs.EOF
'        rsDetail.AddNew
'        rsDetail!B_ID = vID
'
'        For i = 0 To vRs.Fields.Count - 1
'            rsDetail(Trim$(vRs(i).name)) = vRs(i)
'            'Debug.Print vRs(i).name & "=" & vRs(i)
'        Next
'        rsDetail.Update
'
'
'        '保存新增的行的B_ItemID
'        rtnRs.AddNew
'        rtnRs!B_ItemID = rsDetail!B_ItemID
'
'        vRs.MoveNext
'    Loop
'    rsDetail.Close
'    Set rsDetail = Nothing
'
'
'    Set AppendDraftDetailByRs = rtnRs.Clone
'End Function
'
'
'Public Function AppendDetailByRs(ByVal vID As Long, ByVal vRs As Recordset) As Boolean
'
'    If A_Inited = False Then
'        AppendDetailByRs = False
'        MsgBox "初始化后才可以使用！", vbOKOnly + vbInformation, "提示"
'        Exit Function
'    End If
'
'    Dim rs As Recordset
'    Set rs = AppendDraftDetailByRs(vID, vRs).Clone
'
'
'End Function
'
