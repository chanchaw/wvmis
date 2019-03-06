VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLanucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private udClient As New clsUpdate

'软件启动器


'启动软件
Public Function Lanuch() As Boolean
    '1. 判断不可多次打开本软件
    'Lanuch = JudgeLaunchValid
    Lanuch = True
    
'    WriteLog "加载本地配置文件"
    '2. 加载本地配置文件
    LoadIniFile
    
'    WriteLog "初始化硬件信息"
    '2.1. 初始化硬件信息
    InitGmHardWareID
    
'    WriteLog "创建数据库链接"
    '3. 创建数据库链接
    CreateDBConnection
    
'    WriteLog "检测客户端版本"
    '4. 检测客户端版本
    udClient.CheckNewVersion App.hInstance
    
'    WriteLog "登录"
    '5. 登录
    Login
    
    
    '5. 使用参数加载器加载系统参数
    ParaLoader.LoadPara
    LoadCompanyInfo
    '6. 同步服务器时间

    '7. 显示主控台
    ShowMain
    
End Function



Private Sub LoadIniFile()
    On Error Resume Next
    Dim clsIniFile1 As New clsIniFile
    With clsIniFile1
        .IniFileName = App.Path & "\" & CONSTINIFILENAME
        g_SwitchLog4Runtime = Val(.GetIniKey("日志", "运行时日志"))
    End With
    
    
    g_LogFile_Runtime = App.Path & "\runtimelog.text"
    
    Set Gm.IniFile = clsIniFile1

End Sub


Private Sub CreateDBConnection()
    Dim odmSysID As New dmSysID
    Dim odmDBInfo As New dmDBInfo
    Dim odmDBInfoSOB As New dmDBInfo
    Dim odmDBInfoImage As New dmDBInfo
    With odmDBInfo
        .Server = Gm.IniFile.GetIniKey(IF_DBSECTION, IF_DBSECTION_SERVERKEY)
        .DBName = Gm.IniFile.GetIniKey(IF_DBSECTION, IF_DBSECTION_DBKEY)
        .DBUser = Gm.IniFile.GetIniKey(IF_DBSECTION, IF_DBSECTION_USERKEY)
        .DBPassWord = Gm.IniFile.GetIniKey(IF_DBSECTION, IF_DBSECTION_PWKEY)
    End With
    
    '初始化账套集数据库链接参数
    With odmDBInfoSOB
        .Server = Gm.IniFile.GetIniKey(IF_DBSECTION_SOB, IF_DBSECTION_SERVERKEY_SOB)
        .DBName = Gm.IniFile.GetIniKey(IF_DBSECTION_SOB, IF_DBSECTION_DBKEY_SOB)
        .DBUser = Gm.IniFile.GetIniKey(IF_DBSECTION_SOB, IF_DBSECTION_USERKEY_SOB)
        .DBPassWord = Gm.IniFile.GetIniKey(IF_DBSECTION_SOB, IF_DBSECTION_PWKEY_SOB)
    End With
    
    With odmDBInfoImage
        .Server = Gm.IniFile.GetIniKey(IF_DBSECTION_Image, IF_DBSECTION_SERVERKEY_Image)
        .DBName = Gm.IniFile.GetIniKey(IF_DBSECTION_Image, IF_DBSECTION_DBKEY_Image)
        .DBUser = Gm.IniFile.GetIniKey(IF_DBSECTION_Image, IF_DBSECTION_USERKEY_Image)
        .DBPassWord = Gm.IniFile.GetIniKey(IF_DBSECTION_Image, IF_DBSECTION_PWKEY_Image)
    End With
    
    
    Set odmSysID.DBInfo = odmDBInfo
    Set Gm.SysID = odmSysID

    
    Dim cnnTool As New clsDBConnect
    With odmDBInfo
        cnnTool.InitializeConnection .Server, .DBName, .DBUser, .DBPassWord
    End With
    
    
    Dim cnnToolSOB As New clsDBConnect
    With odmDBInfoSOB
        cnnToolSOB.InitializeConnection .Server, .DBName, .DBUser, .DBPassWord
    End With
    
      Dim cnnToolImage As New clsDBConnect
    With odmDBInfoImage
        cnnToolImage.InitializeConnection .Server, .DBName, .DBUser, .DBPassWord
    End With
    
    Set Gm.cnnTool = cnnTool
    Set Gm.cnnToolSOB = cnnToolSOB
    Set Gm.cnnToolImage = cnnToolImage
End Sub


Private Sub Login()
    Dim frm1 As New frmLoginSOB
    Dim bLogin As Boolean
    frm1.Show vbModal
    bLogin = frm1.OK
    Unload frm1
    
    If bLogin = False Then
        End
    End If
End Sub

Private Sub Login2017年10月14日()
    Dim frm1 As New frmLogin
    Dim bLogin As Boolean
    frm1.Show vbModal
    bLogin = frm1.OK
    Unload frm1
    
    If bLogin = False Then
        End
    End If
End Sub

Private Sub ShowMain()
    Dim frm1 As New frmMain
    Set Gm.frmMain = frm1
    Gm.frmMain.Show
End Sub

Private Sub InitGmHardWareID()
    Dim dm As New dmHardWareID
    dm.CUN = GetComputerUserName
    dm.CN = GetComputerName
    Set Gm.HardWareID = dm
End Sub
Public Sub WriteLog(ByVal vContent As String)
    Dim cls1 As New clsFile
    Dim szLogFile As String

    szLogFile = App.Path & "\BLSaveErr.txt"
    cls1.WriteFileContent szLogFile, vContent
    
End Sub
