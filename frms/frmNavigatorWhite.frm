VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.2#0"; "Codejock.CommandBars.v16.2.4.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmNavigatorWhite 
   Caption         =   "白坯模块"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12990
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNavigatorWhite.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9555
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   9555
      Left            =   -240
      TabIndex        =   0
      Top             =   -360
      Width           =   12990
      _LayoutVersion  =   1
      _ExtentX        =   22913
      _ExtentY        =   16854
      _DataPath       =   ""
      Bands           =   "frmNavigatorWhite.frx":058A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   9255
         Left            =   540
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   12135
         _cx             =   21405
         _cy             =   16325
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
         _GridInfo       =   $"frmNavigatorWhite.frx":0752
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   9195
            Left            =   30
            TabIndex        =   2
            Top             =   30
            Width           =   12075
            _cx             =   21299
            _cy             =   16219
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
            BackColor       =   16777215
            ForeColor       =   -2147483630
            FrontTabColor   =   14270310
            BackTabColor    =   16777215
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "Tab&1|Tab&2|Tab&3"
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   6
            Position        =   0
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   0   'False
            TabsPerPage     =   0
            BorderWidth     =   0
            BoldCurrent     =   -1  'True
            DogEars         =   -1  'True
            MultiRow        =   0   'False
            MultiRowOffset  =   200
            CaptionStyle    =   0
            TabHeight       =   500
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Flags(0)        =   2
            Flags(1)        =   2
            Flags(2)        =   2
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   8610
               Left            =   45
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   540
               Width           =   11985
               _cx             =   21140
               _cy             =   15187
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
               BorderWidth     =   1
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
               _GridInfo       =   $"frmNavigatorWhite.frx":07D8
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.PictureBox Picture1 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   8580
                  Left            =   15
                  ScaleHeight     =   8580
                  ScaleWidth      =   11955
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   11955
                  Begin XtremeCommandBars.BackstageSeparator BackstageSeparator1 
                     Height          =   90
                     Left            =   960
                     TabIndex        =   20
                     Top             =   3498
                     Width           =   9435
                     _Version        =   1048578
                     _ExtentX        =   16642
                     _ExtentY        =   159
                     _StockProps     =   2
                     MarkupText      =   ""
                  End
                  Begin XtremeSuiteControls.GroupBox GroupBox1 
                     Height          =   5175
                     Left            =   -6780
                     TabIndex        =   10
                     Top             =   -2880
                     Visible         =   0   'False
                     Width           =   7815
                     _Version        =   1048578
                     _ExtentX        =   13785
                     _ExtentY        =   9128
                     _StockProps     =   79
                     Caption         =   "2018年1月26日弃用"
                     UseVisualStyle  =   -1  'True
                     Begin XtremeSuiteControls.PushButton btnObject 
                        Height          =   1335
                        Index           =   0
                        Left            =   600
                        TabIndex        =   11
                        Tag             =   "12B005"
                        Top             =   540
                        Width           =   1395
                        _Version        =   1048578
                        _ExtentX        =   2461
                        _ExtentY        =   2355
                        _StockProps     =   79
                        Caption         =   "白坯入库"
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        UseVisualStyle  =   -1  'True
                        TextImageRelation=   1
                        IconWidth       =   48
                        Icon            =   "frmNavigatorWhite.frx":085A
                     End
                     Begin XtremeSuiteControls.PushButton btnObject 
                        Height          =   1335
                        Index           =   4
                        Left            =   1500
                        TabIndex        =   12
                        Tag             =   "19B009"
                        Top             =   600
                        Width           =   1395
                        _Version        =   1048578
                        _ExtentX        =   2461
                        _ExtentY        =   2355
                        _StockProps     =   79
                        Caption         =   "白坯出库"
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        UseVisualStyle  =   -1  'True
                        TextImageRelation=   1
                        IconWidth       =   48
                        Icon            =   "frmNavigatorWhite.frx":2CC4
                     End
                     Begin XtremeSuiteControls.PushButton btnObject 
                        Height          =   1335
                        Index           =   7
                        Left            =   2520
                        TabIndex        =   13
                        Tag             =   "19B021"
                        Top             =   600
                        Width           =   1395
                        _Version        =   1048578
                        _ExtentX        =   2461
                        _ExtentY        =   2355
                        _StockProps     =   79
                        Caption         =   "白坯加工入库"
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        UseVisualStyle  =   -1  'True
                        TextImageRelation=   1
                        IconWidth       =   48
                        Icon            =   "frmNavigatorWhite.frx":512E
                     End
                     Begin XtremeSuiteControls.PushButton btnObject 
                        Height          =   1335
                        Index           =   3
                        Left            =   3480
                        TabIndex        =   14
                        Tag             =   "13B010"
                        Top             =   540
                        Width           =   1395
                        _Version        =   1048578
                        _ExtentX        =   2461
                        _ExtentY        =   2355
                        _StockProps     =   79
                        Caption         =   "收发存汇总表"
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        UseVisualStyle  =   -1  'True
                        TextImageRelation=   1
                        IconWidth       =   48
                        Icon            =   "frmNavigatorWhite.frx":7598
                     End
                     Begin XtremeSuiteControls.PushButton btnObject 
                        Height          =   1335
                        Index           =   6
                        Left            =   360
                        TabIndex        =   15
                        Tag             =   "19B017"
                        Top             =   1800
                        Width           =   1395
                        _Version        =   1048578
                        _ExtentX        =   2461
                        _ExtentY        =   2355
                        _StockProps     =   79
                        Caption         =   "白坯采购入库单"
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        UseVisualStyle  =   -1  'True
                        TextImageRelation=   1
                        IconWidth       =   48
                        Icon            =   "frmNavigatorWhite.frx":9A02
                     End
                     Begin XtremeSuiteControls.PushButton btnObject 
                        Height          =   1335
                        Index           =   9
                        Left            =   1500
                        TabIndex        =   16
                        Tag             =   "19B020"
                        Top             =   1980
                        Width           =   1395
                        _Version        =   1048578
                        _ExtentX        =   2461
                        _ExtentY        =   2355
                        _StockProps     =   79
                        Caption         =   "白坯无合同"
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        UseVisualStyle  =   -1  'True
                        TextImageRelation=   1
                        IconWidth       =   48
                        Icon            =   "frmNavigatorWhite.frx":BE6C
                     End
                     Begin XtremeSuiteControls.PushButton btnObject 
                        Height          =   1335
                        Index           =   8
                        Left            =   2580
                        TabIndex        =   17
                        Tag             =   "19B022"
                        Top             =   2160
                        Width           =   1395
                        _Version        =   1048578
                        _ExtentX        =   2461
                        _ExtentY        =   2355
                        _StockProps     =   79
                        Caption         =   "打印流传卡"
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        UseVisualStyle  =   -1  'True
                        TextImageRelation=   1
                        IconWidth       =   48
                        Icon            =   "frmNavigatorWhite.frx":E2D6
                     End
                     Begin XtremeSuiteControls.PushButton btnObject 
                        Height          =   1335
                        Index           =   1
                        Left            =   3180
                        TabIndex        =   18
                        Tag             =   "13B008"
                        Top             =   2580
                        Width           =   1395
                        _Version        =   1048578
                        _ExtentX        =   2461
                        _ExtentY        =   2355
                        _StockProps     =   79
                        Caption         =   "单据流水"
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        UseVisualStyle  =   -1  'True
                        TextImageRelation=   1
                        IconWidth       =   48
                        Icon            =   "frmNavigatorWhite.frx":10740
                     End
                     Begin XtremeSuiteControls.PushButton btnObject 
                        Height          =   1335
                        Index           =   2
                        Left            =   4440
                        TabIndex        =   19
                        Tag             =   "13B009"
                        Top             =   2400
                        Width           =   1395
                        _Version        =   1048578
                        _ExtentX        =   2461
                        _ExtentY        =   2355
                        _StockProps     =   79
                        Caption         =   "库存表"
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        UseVisualStyle  =   -1  'True
                        TextImageRelation=   1
                        IconWidth       =   48
                        Icon            =   "frmNavigatorWhite.frx":12BAA
                     End
                     Begin XtremeSuiteControls.PushButton btnObject 
                        Height          =   1335
                        Index           =   10
                        Left            =   0
                        TabIndex        =   31
                        Tag             =   "19B025"
                        Top             =   0
                        Width           =   1395
                        _Version        =   1048578
                        _ExtentX        =   2461
                        _ExtentY        =   2355
                        _StockProps     =   79
                        Caption         =   "综合查询"
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        UseVisualStyle  =   -1  'True
                        TextImageRelation=   1
                        IconWidth       =   48
                        Icon            =   "frmNavigatorWhite.frx":15014
                     End
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   5
                     Left            =   1200
                     TabIndex        =   5
                     Tag             =   "19B015"
                     Top             =   240
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯订单"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":1747E
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   12
                     Left            =   3540
                     TabIndex        =   6
                     Tag             =   "19B033"
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯加工出入库"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":198E8
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   14
                     Left            =   3540
                     TabIndex        =   7
                     Tag             =   "19B035"
                     Top             =   5511
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯入库汇总"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":1BD52
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   18
                     Left            =   5880
                     TabIndex        =   9
                     Tag             =   "19B040"
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯发货"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":1E1BC
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   19
                     Left            =   8280
                     TabIndex        =   21
                     Tag             =   "19B045"
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "染厂退白坯"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":20626
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   21
                     Left            =   5880
                     TabIndex        =   22
                     Tag             =   "19B055"
                     Top             =   5511
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯收发序时表"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":22A90
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   23
                     Left            =   3540
                     TabIndex        =   23
                     Tag             =   "19B058"
                     Top             =   7200
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "染厂收发存汇总"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":24EFA
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   24
                     Left            =   1200
                     TabIndex        =   24
                     Tag             =   "19B057"
                     Top             =   7200
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯库存表"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":27364
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   11
                     Left            =   5880
                     TabIndex        =   25
                     Tag             =   "19B027"
                     Top             =   3882
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯计划明细"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":297CE
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   13
                     Left            =   1200
                     TabIndex        =   26
                     Tag             =   "19B034"
                     Top             =   5511
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯入库明细"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":2BC38
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   15
                     Left            =   3540
                     TabIndex        =   27
                     Tag             =   "19B038"
                     Top             =   3882
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯车间次布报表"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":2E0A2
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   17
                     Left            =   1200
                     TabIndex        =   28
                     Tag             =   "19B039"
                     Top             =   3882
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯车间生产报表"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":3050C
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   20
                     Left            =   8280
                     TabIndex        =   29
                     Tag             =   "19B047"
                     Top             =   1920
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯调拨"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":32976
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   22
                     Left            =   8280
                     TabIndex        =   30
                     Tag             =   "19B056"
                     Top             =   5511
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯发货序时表"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":34DE0
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   25
                     Left            =   1200
                     TabIndex        =   32
                     Tag             =   "19B066"
                     Top             =   1869
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯上年度库存"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":3724A
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   26
                     Left            =   3540
                     TabIndex        =   33
                     Tag             =   "19B068"
                     Top             =   1869
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯采购入库"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":396B4
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   27
                     Left            =   5880
                     TabIndex        =   34
                     Tag             =   "19B069"
                     Top             =   1869
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯采购退货"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":3BB1E
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   28
                     Left            =   5880
                     TabIndex        =   35
                     Tag             =   "19B070"
                     Top             =   7200
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯采购序时表"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":3DF88
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   29
                     Left            =   8280
                     TabIndex        =   36
                     Tag             =   "19B073"
                     Top             =   7200
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯上年度报表"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":403F2
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   30
                     Left            =   8280
                     TabIndex        =   37
                     Tag             =   "19B078"
                     Top             =   3840
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯打卷查询"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":4285C
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   31
                     Left            =   10440
                     TabIndex        =   38
                     Tag             =   "19B108"
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯转库单"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":44CC6
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   32
                     Left            =   10440
                     TabIndex        =   39
                     Tag             =   "19B109"
                     Top             =   3840
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "转库单序时表"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":47130
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   33
                     Left            =   10440
                     TabIndex        =   40
                     Tag             =   "19B118"
                     Top             =   5511
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯收发存汇总"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":4959A
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   34
                     Left            =   10440
                     TabIndex        =   41
                     Tag             =   "19B119"
                     Top             =   7200
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "加工结算汇总表"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorWhite.frx":4BA04
                  End
               End
            End
         End
      End
   End
   Begin XtremeSuiteControls.PushButton btnObject 
      Height          =   1335
      Index           =   16
      Left            =   0
      TabIndex        =   8
      Tag             =   "19B038"
      Top             =   0
      Width           =   1395
      _Version        =   1048578
      _ExtentX        =   2461
      _ExtentY        =   2355
      _StockProps     =   79
      Caption         =   "白坯车间次布报表"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextImageRelation=   1
      IconWidth       =   48
      Icon            =   "frmNavigatorWhite.frx":4DE6E
   End
End
Attribute VB_Name = "frmNavigatorWhite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InitLayout()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
End Sub

Public Sub InitFrm()
    InitLayout
    
    ConfirmPermission
End Sub

Private Sub btnObject_Click(Index As Integer)
    OpenObject btnObject(Index).Tag, "通过导航中按钮打开"
    
End Sub

Private Sub Form_Load()
    InitFrm
    btnObject(8).Visible = False
End Sub


Private Sub OpenObject(ByVal m_ObjectID As String, ByVal m_BillName As String)
    Gm.Authority.Execute m_ObjectID, m_BillName, "LoadObject", Nothing
End Sub

'设置按钮的可用度
Public Sub ConfirmPermission()
    On Error Resume Next
    Dim i As Long
    Dim szObjectID As String
    
    
    For i = 0 To btnObject.Count - 1
        szObjectID = btnObject(i).Tag
        If Len(szObjectID) > 0 Then
            If Gm.PI.JudgeView(szObjectID) = True Then
                btnObject(i).Enabled = True
            Else
                btnObject(i).Enabled = False
            End If
        End If
    Next
End Sub




