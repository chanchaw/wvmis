VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00CEDFDE&
   Caption         =   "织造企业MIS系统"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10980
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Align           =   1  'Align Top
      Height          =   7305
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10980
      _LayoutVersion  =   1
      _ExtentX        =   19368
      _ExtentY        =   12885
      _DataPath       =   ""
      Bands           =   "frmMain.frx":16AC2
      Begin TDBTime6Ctl.TDBTime TDBTime1 
         Height          =   315
         Left            =   2220
         TabIndex        =   1
         Top             =   4020
         Visible         =   0   'False
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "frmMain.frx":16C8A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmMain.frx":16CED
         Spin            =   "frmMain.frx":16D3D
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn:ss"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn:ss"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   .99999
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "11:16:59"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   .470127314814815
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   1200
         Top             =   2160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private clsMenu1 As New clsMenu


Private Sub MDIForm_Load()
    On Error Resume Next
    Dim m_MenuID As String
    Me.BackColor = RGB(156, 186, 206)
    
    
    '设置背景
'    If Len(clsSParameter1.GetParameterString("PictureName")) > 2 Then
'        Me.Picture = LoadPicture(clsSParameter1.GetParameterString("PictureName"))
'    End If
    
    
    m_MenuID = GetMenuID
    With clsMenu1
        .InitClass ActiveBar21, m_MenuID
        .LoadObject
    End With
    
    
    '设置用户
    ActiveBar21.Bands("P状态栏").Tools("P子系统").Caption = Gm.SysID.SubSystem
    ActiveBar21.Bands("P状态栏").Tools("P当前用户").Caption = Gm.SysID.SystemUserName
'    ActiveBar21.Bands("P状态栏").Tools("P当前用户").Caption = Gm.SysID.GetOperator
    ActiveBar21.Bands("P状态栏").Tools("P服务器").Caption = Gm.SysID.DBInfo.Server
    ActiveBar21.Bands("P状态栏").Tools("P帐套").Caption = Gm.SysID.DBInfo.DBName
    

    

    '加载左侧导航栏
    AddDockedForms

End Sub


Private Function GetMenuID() As String
    Dim m_MenuID As String
    Dim rs As New RecordSet
    Dim strSQL As String
    
    strSQL = "Select * From G_SubSystem Where B_SubSystem='" & Gm.SysID.SubSystem & "'"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        GetMenuID = rs("B_MenuObjectID")
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("是否要退出本系统?", vbExclamation + vbOKCancel + vbDefaultButton2, "退出") = vbCancel Then
        Cancel = 1
        Exit Sub
    End If
End Sub


Private Sub MDIForm_Terminate()
    Debug.Print "this is Terminate"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    DeleteTemp
    ClearDBLink
End Sub

Private Sub ClearDBLink()

End Sub

Private Sub DeleteTemp()
    On Error Resume Next

    'Shell "cmd.exe /c del " & App.Path & "\*.jpg"
End Sub



'===================================本代码块为制作导航
Private Sub AddDockedForms()
    Dim frm As IDock2AB
    Set frm = frmNavigatorLeft
    frm.DockYourselfTo ActiveBar21, True, ddDALeft, ddGSCaption
    
    ActiveBar21.RecalcLayout
End Sub
'===================================本代码块为制作导航END


