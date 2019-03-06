VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmAdminTool01 
   Caption         =   "开发者工具"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10395
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdminTool01.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6870
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   6870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10395
      _LayoutVersion  =   1
      _ExtentX        =   18336
      _ExtentY        =   12118
      _DataPath       =   ""
      Bands           =   "frmAdminTool01.frx":038A
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   5835
         Left            =   420
         TabIndex        =   1
         Top             =   540
         Width           =   9375
         _cx             =   16536
         _cy             =   10292
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   2
         MousePointer    =   0
         Version         =   800
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "按钮尺寸与布局"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   0   'False
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   5460
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   9285
            _cx             =   16378
            _cy             =   9631
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   2
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
            GridRows        =   4
            GridCols        =   4
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmAdminTool01.frx":0552
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   5400
               Left            =   30
               ScaleHeight     =   5400
               ScaleWidth      =   9225
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   30
               Width           =   9225
               Begin XtremeSuiteControls.FlatEdit FlatEdit1 
                  Height          =   375
                  Index           =   0
                  Left            =   1680
                  TabIndex        =   21
                  Top             =   2640
                  Width           =   1155
                  _Version        =   1048578
                  _ExtentX        =   2037
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit1 
                  Height          =   375
                  Index           =   1
                  Left            =   1680
                  TabIndex        =   25
                  Top             =   3360
                  Width           =   1155
                  _Version        =   1048578
                  _ExtentX        =   2037
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "新按钮的坐标"
                  Height          =   195
                  Index           =   1
                  Left            =   4740
                  TabIndex        =   27
                  Top             =   3420
                  Width           =   1080
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "个按钮的坐标"
                  Height          =   195
                  Index           =   11
                  Left            =   2940
                  TabIndex        =   26
                  Top             =   3420
                  Width           =   1080
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "单据第："
                  Height          =   195
                  Index           =   10
                  Left            =   840
                  TabIndex        =   24
                  Top             =   3420
                  Width           =   720
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "新按钮的坐标"
                  Height          =   195
                  Index           =   0
                  Left            =   4740
                  TabIndex        =   23
                  Top             =   2700
                  Width           =   1080
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "个按钮的坐标"
                  Height          =   195
                  Index           =   9
                  Left            =   2940
                  TabIndex        =   22
                  Top             =   2700
                  Width           =   1080
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "单据第："
                  Height          =   195
                  Index           =   8
                  Left            =   840
                  TabIndex        =   20
                  Top             =   2700
                  Width           =   720
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "3360"
                  Height          =   195
                  Index           =   7
                  Left            =   7320
                  TabIndex        =   19
                  Top             =   1020
                  Width           =   360
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "第一报表TOP："
                  Height          =   195
                  Index           =   7
                  Left            =   6000
                  TabIndex        =   18
                  Top             =   1020
                  Width           =   1200
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "1200"
                  Height          =   195
                  Index           =   6
                  Left            =   7320
                  TabIndex        =   17
                  Top             =   600
                  Width           =   360
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "第一报表LEFT："
                  Height          =   195
                  Index           =   6
                  Left            =   6000
                  TabIndex        =   16
                  Top             =   600
                  Width           =   1245
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "960"
                  Height          =   195
                  Index           =   5
                  Left            =   4860
                  TabIndex        =   15
                  Top             =   960
                  Width           =   270
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "第一按钮TOP："
                  Height          =   195
                  Index           =   5
                  Left            =   3540
                  TabIndex        =   14
                  Top             =   960
                  Width           =   1200
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "1200"
                  Height          =   195
                  Index           =   4
                  Left            =   4860
                  TabIndex        =   13
                  Top             =   600
                  Width           =   360
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "第一按钮LEFT："
                  Height          =   195
                  Index           =   4
                  Left            =   3540
                  TabIndex        =   12
                  Top             =   600
                  Width           =   1245
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "645"
                  Height          =   195
                  Index           =   3
                  Left            =   1740
                  TabIndex        =   11
                  Top             =   2040
                  Width           =   270
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "纵向间隔："
                  Height          =   195
                  Index           =   3
                  Left            =   780
                  TabIndex        =   10
                  Top             =   2040
                  Width           =   900
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "945"
                  Height          =   195
                  Index           =   2
                  Left            =   1740
                  TabIndex        =   9
                  Top             =   1620
                  Width           =   270
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "横向间隔："
                  Height          =   195
                  Index           =   2
                  Left            =   780
                  TabIndex        =   8
                  Top             =   1620
                  Width           =   900
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "1335"
                  Height          =   195
                  Index           =   1
                  Left            =   1740
                  TabIndex        =   7
                  Top             =   960
                  Width           =   360
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "按钮高度："
                  Height          =   195
                  Index           =   1
                  Left            =   780
                  TabIndex        =   6
                  Top             =   960
                  Width           =   900
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "1395"
                  Height          =   195
                  Index           =   0
                  Left            =   1740
                  TabIndex        =   5
                  Top             =   600
                  Width           =   360
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "按钮宽度："
                  Height          =   195
                  Index           =   0
                  Left            =   780
                  TabIndex        =   4
                  Top             =   600
                  Width           =   900
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmAdminTool01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'下面是单据按钮的标准参数
Private Const theXO As Double = 1200  '第一个按钮的LEFT
Private Const theYO As Double = 960   '第一个按钮的TOP
Private Const theWidth As Double = 1395   '按钮宽度
Private Const theXG As Double = 945 '横向按钮间隔


'下面是报表按钮的标准参数：
Private Const rptXO As Double = 1200
Private Const rptYO As Double = 3360
Private Const rptWidth As Double = 1395
Private Const rptHeight As Double = 1335
Private Const rptXG As Double = 945  '横向按钮间隔
Private Const rptYG As Double = 645  '纵向按钮的间隔

Private Sub InitFrm()
    InitLayout
End Sub

Private Sub InitLayout()
    With ActiveBar21
        .ClientAreaControl = C1Tab1
        .RecalcLayout
    End With
End Sub

Private Sub Form_Load()
    InitFrm
End Sub

'获取第N个单据按钮的坐标
Private Function GetNewBLB(ByVal N As Long) As dmXY
    Dim dm As dmXY
    Set dm = New dmXY
    With dm
        .X = theXO + (N - 1) * theWidth + (N - 1) * theXG '左边距 + 第N个按钮的宽度 + 间隔
        .Y = theYO
    End With
End Function

'获取报表按钮的坐标
Private Function GetNewRptB(ByVal N As Long) As dmXY
    Dim dm As New dmXY
    With dm
'        .X =
        
    End With
End Function

