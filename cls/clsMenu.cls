VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ab2 As ActiveBar2
Attribute ab2.VB_VarHelpID = -1
'保持属性值的局部变量
Private mvarObjectID As String '局部复制
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

Public Sub InitClass(ByRef o As Object, ByVal m_ObjectID As String)
    ObjectID = m_ObjectID
    Set ab2 = o
End Sub

Public Sub LoadObject()
    On Error Resume Next
    Dim rs As New RecordSet
    Dim strSQL As String
    Dim mnuImage() As Byte
    
    
    Gm.cnnTool.IniConnection8DM Gm.SysID.DBInfo
    
    Set rs = New RecordSet
    
    strSQL = "Select * From G_Menu Where B_ObjectID='" & ObjectID & "'"
    
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then
        mnuImage() = rs("B_MenuImage")
        
        ab2.Load "", mnuImage, ddSOByteArray

        ab2.RecalcLayout
        FillCommand
    Else
        MsgBox "菜单未设置!", vbExclamation, "设置"
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

'填充命令
Private Sub FillCommand()
    Dim rs As New RecordSet
    Dim strSQL As String
    Dim i As Integer
    Dim j As Integer
    Dim isRoot As Boolean
    
    Set rs = New RecordSet
    
    If Gm.IsDeveloper = True Then
        strSQL = "Select * From G_MenuItems"
        isRoot = True
    Else
        strSQL = "Select * From G_MenuItems Where B_ObjectID In ("
        strSQL = strSQL & " Select B_ObjectID From G_UserPro Where B_UserName='" & Gm.SysID.SystemUser & "'"
        strSQL = strSQL & " )"
        Debug.Print strSQL
    End If
    Debug.Print strSQL


    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockReadOnly
    
    Dim oCtl As Object
    With ab2
        For i = 0 To .Bands.Count - 1
            For j = 0 To .Bands(i).Tools.Count - 1
            
                rs.Filter = "B_MenuItem='" & Trim(.Bands(i).Tools(j).name) & "'"
                Debug.Print .Bands(i).Tools(j).name
                If rs.RecordCount > 0 Then
                    Debug.Print "进入了设置" & Trim(.Bands(i).Tools(j).name)
                    Set oCtl = .Bands(i).Tools(j)
                    oCtl.Text = rs("B_ObjectID") & "," & rs("B_ObjectName") & "," & rs("B_Operation") & "," & rs("B_Parameter")
                    'Debug.Print oCtl.Text
                End If
            Next
        Next
    End With
End Sub

Private Sub ab2_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    On Error GoTo IFERR
    '当菜单点击时
    'ObjectID,
    Dim aList
    aList = Split(Tool.Text, ",")
    
    Gm.Authority.Execute aList(0), aList(1), aList(2), Tool, aList(3)
    Exit Sub
IFERR:
    'MsgBox "未能正确解析菜单命令！", vbOKOnly + vbInformation, "提示"
    Exit Sub
End Sub



'遍历获取所有子系统中菜单项的名称，对比DB端G_MenuItems中的数据
'将表中不存在于子系统中的项目删除
Public Sub ClearRepeatMenuItem(ByRef vAB As ActiveBar2)
    '0. 创建临时表
    Dim cls1 As New clsDataBase
    Dim szTempTableName As String
    Dim strSQL As String
    
    Set cls1 = New clsDataBase
    cls1.initCls Gm.SysID.DBInfo.DBName
    szTempTableName = "tempClearRepeatMenuItems"
    If cls1.JudgeTableExist(szTempTableName) = True Then
        strSQL = "Drop Table " & szTempTableName
        Gm.cnnTool.cnn.Execute strSQL
        
    End If
    
    '创建表
    strSQL = "Create Table " & szTempTableName & " (B_MenuItem VARCHAR(300))"
    Gm.cnnTool.cnn.Execute strSQL
    
    '1. 获取所有子系统的菜单项，整合为一个记录集
    Dim rsMenus As New RecordSet
    Dim mnuImage() As Byte
    strSQL = "SELECT * FROM G_Menu"
    rsMenus.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rsMenus.RecordCount <= 0 Then
        rsMenus.Close
        Set rsMenus = Nothing
        Exit Sub
    End If
    
    Dim rsAllMenuItems As New RecordSet
    strSQL = "Select B_MenuItem From " & szTempTableName
    rsAllMenuItems.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    Dim i As Long
    Dim j As Long
    Do While Not rsMenus.EOF
        mnuImage() = rsMenus("B_MenuImage")
        vAB.Load "", mnuImage, ddSOByteArray
        vAB.RecalcLayout
        
        '将菜单项写入临时表
        With vAB
            For i = 0 To .Bands.Count - 1
                For j = 0 To .Bands(i).Tools.Count - 1
                    rsAllMenuItems.Filter = " B_MenuItem='" & .Bands(i).Tools(j).name & "'"
                    Debug.Print .Bands(i).Tools(j).name
                    If rsAllMenuItems.RecordCount <= 0 Then
                        rsAllMenuItems.AddNew
                        rsAllMenuItems!B_MenuItem = .Bands(i).Tools(j).name
                        rsAllMenuItems.Update
                    End If
                Next
            Next
        End With
    
        rsMenus.MoveNext
    Loop
    rsMenus.Close
    Set rsMenus = Nothing
    

    
    '3. 删除2中的表中的对比于菜单项中多余的数据
    strSQL = "Delete"
    strSQL = strSQL & " From G_MenuItems"
    strSQL = strSQL & " WHERE B_MenuItem NOT IN"
    strSQL = strSQL & " ("
    strSQL = strSQL & "     SELECT  B_MenuItem FROM tempClearRepeatMenuItems"
    strSQL = strSQL & " )"
    Gm.cnnTool.cnn.Execute strSQL
    
    MsgBox "已清空所有不存在于子系统中的菜单项！", vbOKOnly + vbInformation, "提示"
End Sub


