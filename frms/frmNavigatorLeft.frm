VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{22D1E30D-3561-406D-8495-1061AE20E101}#1.0#0"; "AlphaImageControl.ocx"
Begin VB.Form frmNavigatorLeft 
   BorderStyle     =   0  'None
   Caption         =   "导航栏"
   ClientHeight    =   8205
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   2925
   Icon            =   "frmNavigatorLeft.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8205
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2925
      _cx             =   5159
      _cy             =   14473
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
      BorderWidth     =   2
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmNavigatorLeft.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   15000
         Left            =   30
         Picture         =   "frmNavigatorLeft.frx":0409
         ScaleHeight     =   15000
         ScaleWidth      =   3000
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   3000
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "     财 务 系 统"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   8
            Left            =   240
            MouseIcon       =   "frmNavigatorLeft.frx":1681A
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Tag             =   "染化料仓库"
            Top             =   7800
            Width           =   2100
         End
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   480
            Index           =   8
            Left            =   240
            Tag             =   "染化料仓库"
            Top             =   7680
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Image           =   "frmNavigatorLeft.frx":16B24
            Props           =   5
         End
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   480
            Index           =   7
            Left            =   240
            Tag             =   "染化料仓库"
            Top             =   6960
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Image           =   "frmNavigatorLeft.frx":1BA3C
            Props           =   5
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "     五 金 仓 库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   7
            Left            =   240
            MouseIcon       =   "frmNavigatorLeft.frx":20954
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Tag             =   "染化料仓库"
            Top             =   6960
            Width           =   2100
         End
         Begin VB.Image Image1 
            Height          =   630
            Index           =   8
            Left            =   120
            Picture         =   "frmNavigatorLeft.frx":20C5E
            Tag             =   "五金仓库"
            Top             =   7560
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.Image Image1 
            Height          =   630
            Index           =   7
            Left            =   120
            Picture         =   "frmNavigatorLeft.frx":2629D
            Tag             =   "五金仓库"
            Top             =   6840
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "路腾纺织印染ERP"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   180
            TabIndex        =   9
            Top             =   600
            Width           =   2520
         End
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   480
            Index           =   6
            Left            =   270
            Tag             =   "染化料仓库"
            Top             =   6150
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Image           =   "frmNavigatorLeft.frx":2B8DC
            Props           =   5
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "     成 品 仓 库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   6
            Left            =   240
            MouseIcon       =   "frmNavigatorLeft.frx":307F4
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Tag             =   "染化料仓库"
            Top             =   6240
            Width           =   2100
         End
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   480
            Index           =   5
            Left            =   270
            Tag             =   "五金仓库"
            Top             =   5370
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Image           =   "frmNavigatorLeft.frx":30AFE
            Props           =   5
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "     色 布 仓 库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   5
            Left            =   240
            MouseIcon       =   "frmNavigatorLeft.frx":35A16
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Tag             =   "五金仓库"
            Top             =   5460
            Width           =   2100
         End
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   480
            Index           =   4
            Left            =   270
            Tag             =   "成品仓库"
            Top             =   4590
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Image           =   "frmNavigatorLeft.frx":35D20
            Props           =   5
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "     白 坯 仓 库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   4
            Left            =   240
            MouseIcon       =   "frmNavigatorLeft.frx":3AC38
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Tag             =   "成品仓库"
            Top             =   4680
            Width           =   2100
         End
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   480
            Index           =   3
            Left            =   270
            Tag             =   "白坯仓库"
            Top             =   3810
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Image           =   "frmNavigatorLeft.frx":3AF42
            Props           =   5
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "     原 料 仓 库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   3
            Left            =   240
            MouseIcon       =   "frmNavigatorLeft.frx":3FE5A
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Tag             =   "白坯仓库"
            Top             =   3900
            Width           =   2100
         End
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   480
            Index           =   2
            Left            =   270
            Tag             =   "原料仓库"
            Top             =   3030
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Image           =   "frmNavigatorLeft.frx":40164
            Props           =   5
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "     辅 料 仓 库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   2
            Left            =   240
            MouseIcon       =   "frmNavigatorLeft.frx":4507C
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Tag             =   "原料仓库"
            Top             =   3120
            Width           =   2100
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "     订 单 合 同"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   1
            Left            =   240
            MouseIcon       =   "frmNavigatorLeft.frx":45386
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Tag             =   "生产计划管理"
            Top             =   2340
            Width           =   2100
         End
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   480
            Index           =   1
            Left            =   300
            Tag             =   "生产计划管理"
            Top             =   2220
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Image           =   "frmNavigatorLeft.frx":45690
            Props           =   5
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "       基础资料"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   270
            MouseIcon       =   "frmNavigatorLeft.frx":484FD
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Tag             =   "基础资料"
            Top             =   1560
            Width           =   2355
         End
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   480
            Index           =   0
            Left            =   300
            Tag             =   "基础资料"
            Top             =   1470
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
            Image           =   "frmNavigatorLeft.frx":48807
            Props           =   5
         End
         Begin VB.Image Image1 
            Height          =   630
            Index           =   0
            Left            =   180
            Picture         =   "frmNavigatorLeft.frx":48D28
            Tag             =   "基础资料"
            Top             =   1380
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.Image Image1 
            Height          =   630
            Index           =   1
            Left            =   180
            Picture         =   "frmNavigatorLeft.frx":4E367
            Tag             =   "生产计划管理"
            Top             =   2160
            Width           =   2520
         End
         Begin VB.Image Image1 
            Height          =   630
            Index           =   2
            Left            =   180
            Picture         =   "frmNavigatorLeft.frx":539A6
            Tag             =   "原料仓库"
            Top             =   2940
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.Image Image1 
            Height          =   630
            Index           =   3
            Left            =   180
            Picture         =   "frmNavigatorLeft.frx":58FE5
            Tag             =   "白坯仓库"
            Top             =   3720
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.Image Image1 
            Height          =   630
            Index           =   4
            Left            =   180
            Picture         =   "frmNavigatorLeft.frx":5E624
            Tag             =   "成品仓库"
            Top             =   4500
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.Image Image1 
            Height          =   630
            Index           =   5
            Left            =   180
            Picture         =   "frmNavigatorLeft.frx":63C63
            Tag             =   "五金仓库"
            Top             =   5280
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.Image Image1 
            Height          =   630
            Index           =   6
            Left            =   180
            Picture         =   "frmNavigatorLeft.frx":692A2
            Tag             =   "染化料仓库"
            Top             =   6060
            Visible         =   0   'False
            Width           =   2520
         End
      End
   End
End
Attribute VB_Name = "frmNavigatorLeft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IDock2AB

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private strSQL As String

Private clsModules1 As New clsModules

Private Sub Form_Load()
    InitFrm
    
End Sub

Private Function IDock2AB_DockYourselfTo(ByVal ActiveBar As ActiveBar2LibraryCtl.IActiveBar2, Optional ByVal parmIsVisible As Boolean = True, Optional ByVal paramDockingarea As ActiveBar2LibraryCtl.DockingAreaTypes = 3&, Optional ByVal paramGrabHandleStyle As ActiveBar2LibraryCtl.GrabHandleStyles = 7&, Optional ByVal paramDockingOffset As Long = 0&) As ActiveBar2LibraryCtl.IBand
    Dim b As ActiveBar2LibraryCtl.band
    Dim T As ActiveBar2LibraryCtl.Tool
    Dim sBandName As String

    On Error GoTo eh_IWillDockToActiveBar_DockYourselfTo


    sBandName = DOCKABLEBANDPREFIXNAME + Me.name

    ' The Dockable Form band for this form does not exist, so create one.
    Set b = ActiveBar.Bands.add(sBandName)
        'b.Caption = Me.Caption
        b.Caption = ""

        b.DockingArea = paramDockingarea
        b.DockLine = 0
        b.DockingOffset = paramDockingOffset

        b.GrabHandleStyle = paramGrabHandleStyle

        b.AutoSizeForms = True
        b.Type = ddBTNormal
        b.DisplayMoreToolsButton = False

        ABAddFlag ddBFSizer, b

        b.Visible = parmIsVisible
        
        

    ' Add a DockableForm tool to dock this window to.
    Set T = b.Tools.add(Me.hwnd, DOCKABLETOOLPREFIXNAME + Me.name)
        T.ControlType = ddTTForm
        T.Caption = Me.Caption
        Set T.Custom = Me
        
        T.width = 2880
        


ex_IWillDockToActiveBar_DockYourselfTo:

    Exit Function
    

eh_IWillDockToActiveBar_DockYourselfTo:

    MsgBox "There was an error while docking form [" + Me.name + "]."
    Resume ex_IWillDockToActiveBar_DockYourselfTo
End Function

'=============上面代码为制作左侧导航栏


'打开基础资料的导航页面
Private Sub OpenDT()
    clsModules1.ShowModule MODULE_DATADICTIONARY
End Sub

'打开辅料导航页面
Private Sub OpenAccessory()
    clsModules1.ShowModule MODULE_ACCESSORY
End Sub

'原料导航
Private Sub OpenYarn()
    clsModules1.ShowModule MODULE_YARN
End Sub

Private Sub OpenWhite()
    clsModules1.ShowModule MODULE_WHITE
End Sub

Private Sub OpenColor()
    clsModules1.ShowModule MODULE_COLOR
End Sub
Private Sub OpenCP()
    clsModules1.ShowModule MODULE_CP
End Sub


Private Sub OpenOrder()
    clsModules1.ShowModule MODULE_ORDER
End Sub

Private Sub OpenGold()
    clsModules1.ShowModule MODULE_Gold
End Sub
Private Sub Openfinancial()
    clsModules1.ShowModule MODULE_financial
End Sub



'设置浮动效果
'当前vIndex的设置为浮动效果，其他的浮动效果隐藏
Private Sub SetFloat(ByVal vIndex As Long)
    Dim i As Long
    For i = 0 To 8
        If i = vIndex Then
            Image1(i).Visible = True
        Else
            Image1(i).Visible = False
        End If
    Next
End Sub

Private Sub Label1_Click(Index As Integer)
    SetFloat Index
    
    Select Case Index
        Case 0   '基础资料
            OpenDT
        Case 1 '生产计划
            OpenOrder
        Case 2  '辅料仓库
            OpenAccessory
        Case 3   '原料仓库
            OpenYarn
        Case 4
            OpenWhite
        Case 5
            OpenColor
        Case 6
            OpenCP
        Case 7
            OpenGold
        Case 8
            Openfinancial
    End Select
End Sub

Private Sub HideMod()
    Dim cls1 As New clsModules
    cls1.HideLeftMod Me
End Sub

Private Sub InitFrm()
    HideMod
    Label1_Click 1
End Sub
