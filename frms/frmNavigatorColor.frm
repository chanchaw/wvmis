VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.2#0"; "Codejock.CommandBars.v16.2.4.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmNavigatorColor 
   Caption         =   "色布模块"
   ClientHeight    =   9375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNavigatorColor.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   20160
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20160
      _LayoutVersion  =   1
      _ExtentX        =   35560
      _ExtentY        =   16536
      _DataPath       =   ""
      Bands           =   "frmNavigatorColor.frx":058A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   8955
         Left            =   540
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   19635
         _cx             =   34634
         _cy             =   15796
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
         _GridInfo       =   $"frmNavigatorColor.frx":0752
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   8895
            Left            =   30
            TabIndex        =   2
            Top             =   30
            Width           =   19575
            _cx             =   34528
            _cy             =   15690
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
               Height          =   8310
               Left            =   45
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   540
               Width           =   19485
               _cx             =   34369
               _cy             =   14658
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
               GridRows        =   5
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmNavigatorColor.frx":07D9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.Frame Frame8 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "转库"
                  Height          =   3885
                  Left            =   14625
                  TabIndex        =   42
                  Top             =   4410
                  Width           =   4845
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   5
                     Left            =   720
                     TabIndex        =   43
                     Tag             =   "19B110"
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布转库单"
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
                     Icon            =   "frmNavigatorColor.frx":0868
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   6
                     Left            =   720
                     TabIndex        =   44
                     Tag             =   "19B111"
                     Top             =   2160
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
                     Icon            =   "frmNavigatorColor.frx":2CD2
                  End
               End
               Begin VB.Frame Frame7 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "色布供应商往来"
                  Height          =   2070
                  Left            =   4875
                  TabIndex        =   35
                  Top             =   6225
                  Width           =   4845
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   28
                     Left            =   2640
                     TabIndex        =   36
                     Tag             =   "19B088"
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布采购发货"
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
                     Icon            =   "frmNavigatorColor.frx":513C
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   27
                     Left            =   480
                     TabIndex        =   37
                     Tag             =   "19B087"
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布采购入库"
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
                     Icon            =   "frmNavigatorColor.frx":75A6
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   1
                     Left            =   1680
                     TabIndex        =   39
                     Tag             =   "19B089"
                     Top             =   360
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布采购"
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
                     Icon            =   "frmNavigatorColor.frx":9A10
                  End
               End
               Begin VB.Frame Frame6 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "客户往来"
                  Height          =   2070
                  Left            =   15
                  TabIndex        =   31
                  Top             =   6225
                  Width           =   4830
                  Begin XtremeCommandBars.BackstageSeparator BackstageSeparator1 
                     Height          =   30
                     Left            =   480
                     TabIndex        =   38
                     Top             =   1920
                     Width           =   3735
                     _Version        =   1048578
                     _ExtentX        =   6588
                     _ExtentY        =   53
                     _StockProps     =   2
                     MarkupText      =   ""
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   9
                     Left            =   480
                     TabIndex        =   32
                     Tag             =   "19B042"
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布发货"
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
                     Icon            =   "frmNavigatorColor.frx":BE7A
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   11
                     Left            =   2640
                     TabIndex        =   33
                     Tag             =   "19B044"
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "客户退货"
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
                     Icon            =   "frmNavigatorColor.frx":E2E4
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   13
                     Left            =   1560
                     TabIndex        =   34
                     Tag             =   "19B060"
                     Top             =   360
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布发货"
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
                     Icon            =   "frmNavigatorColor.frx":1074E
                  End
               End
               Begin VB.Frame Frame5 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "综合报表"
                  Height          =   4365
                  Left            =   14625
                  TabIndex        =   28
                  Top             =   15
                  Width           =   4845
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   20
                     Left            =   600
                     TabIndex        =   29
                     Tag             =   "19B079"
                     Top             =   360
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布打卷查询"
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
                     Icon            =   "frmNavigatorColor.frx":12BB8
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   19
                     Left            =   2760
                     TabIndex        =   30
                     Tag             =   "19B077"
                     Top             =   360
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布库存表"
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
                     Icon            =   "frmNavigatorColor.frx":15022
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   29
                     Left            =   600
                     TabIndex        =   45
                     Tag             =   "19B117"
                     Top             =   2160
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布收发存汇总"
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
                     Icon            =   "frmNavigatorColor.frx":1748C
                  End
               End
               Begin VB.Frame Frame4 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "往期库存"
                  Height          =   6180
                  Left            =   9750
                  TabIndex        =   24
                  Top             =   15
                  Width           =   4845
                  Begin XtremeCommandBars.BackstageSeparator BackstageSeparator5 
                     Height          =   30
                     Left            =   600
                     TabIndex        =   27
                     Top             =   3600
                     Width           =   3855
                     _Version        =   1048578
                     _ExtentX        =   6800
                     _ExtentY        =   53
                     _StockProps     =   2
                     MarkupText      =   ""
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   18
                     Left            =   480
                     TabIndex        =   25
                     Tag             =   "19B074"
                     Top             =   3960
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布上年度报表"
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
                     Icon            =   "frmNavigatorColor.frx":198F6
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   17
                     Left            =   480
                     TabIndex        =   26
                     Tag             =   "19B067"
                     Top             =   360
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布上年度库存"
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
                     Icon            =   "frmNavigatorColor.frx":1BD60
                  End
               End
               Begin VB.Frame Frame3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "家纺厂往来"
                  Height          =   2070
                  Left            =   9750
                  TabIndex        =   21
                  Top             =   6225
                  Width           =   4845
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   26
                     Left            =   480
                     TabIndex        =   22
                     Tag             =   "19B085"
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "家纺厂成品入库"
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
                     Icon            =   "frmNavigatorColor.frx":1E1CA
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   25
                     Left            =   2520
                     TabIndex        =   23
                     Tag             =   "19B084"
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布发家纺厂"
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
                     Icon            =   "frmNavigatorColor.frx":20634
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   2
                     Left            =   1560
                     TabIndex        =   40
                     Tag             =   "19B090"
                     Top             =   360
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "家纺厂"
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
                     Icon            =   "frmNavigatorColor.frx":22A9E
                  End
               End
               Begin VB.Frame Frame2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "深加工入库"
                  Height          =   6180
                  Left            =   4875
                  TabIndex        =   14
                  Top             =   15
                  Width           =   4845
                  Begin XtremeCommandBars.BackstageSeparator BackstageSeparator4 
                     Height          =   30
                     Left            =   360
                     TabIndex        =   19
                     Top             =   3600
                     Width           =   3975
                     _Version        =   1048578
                     _ExtentX        =   7011
                     _ExtentY        =   53
                     _StockProps     =   2
                     MarkupText      =   ""
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   8
                     Left            =   2640
                     TabIndex        =   15
                     Tag             =   "19B041"
                     Top             =   1680
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "深加工入库"
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
                     Icon            =   "frmNavigatorColor.frx":24F08
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   21
                     Left            =   2640
                     TabIndex        =   16
                     Tag             =   "19B080"
                     Top             =   2280
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布外加工发货"
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
                     Icon            =   "frmNavigatorColor.frx":27372
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   24
                     Left            =   2640
                     TabIndex        =   17
                     Tag             =   "19B083"
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "深加工回修布入库"
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
                     Icon            =   "frmNavigatorColor.frx":297DC
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   23
                     Left            =   480
                     TabIndex        =   18
                     Tag             =   "19B082"
                     Top             =   2040
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布退深加工"
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
                     Icon            =   "frmNavigatorColor.frx":2BC46
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   22
                     Left            =   480
                     TabIndex        =   20
                     Tag             =   "19B081"
                     Top             =   360
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "深加工收发"
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
                     Icon            =   "frmNavigatorColor.frx":2E0B0
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   3
                     Left            =   2760
                     TabIndex        =   41
                     Tag             =   "19B099"
                     Top             =   3960
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布退和回修序时表"
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
                     Icon            =   "frmNavigatorColor.frx":3051A
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   30
                     Left            =   480
                     TabIndex        =   46
                     Tag             =   "19B122"
                     Top             =   3960
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布打卷手动入库"
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
                     Icon            =   "frmNavigatorColor.frx":32984
                  End
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "染厂往来"
                  Height          =   6180
                  Left            =   15
                  TabIndex        =   4
                  Top             =   15
                  Width           =   4830
                  Begin XtremeCommandBars.BackstageSeparator BackstageSeparator3 
                     Height          =   30
                     Left            =   360
                     TabIndex        =   9
                     Top             =   3600
                     Width           =   3375
                     _Version        =   1048578
                     _ExtentX        =   5953
                     _ExtentY        =   53
                     _StockProps     =   2
                     MarkupText      =   ""
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   0
                     Left            =   480
                     TabIndex        =   5
                     Tag             =   "19B029"
                     Top             =   1800
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "染厂入库"
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
                     Icon            =   "frmNavigatorColor.frx":34DEE
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   10
                     Left            =   2640
                     TabIndex        =   6
                     Tag             =   "19B037"
                     Top             =   1800
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布退货"
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
                     Icon            =   "frmNavigatorColor.frx":37258
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   14
                     Left            =   480
                     TabIndex        =   7
                     Tag             =   "19B061"
                     Top             =   1800
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "回修色布入库"
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
                     Icon            =   "frmNavigatorColor.frx":396C2
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   15
                     Left            =   2400
                     TabIndex        =   8
                     Tag             =   "19B062"
                     Top             =   1680
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布回修退染厂"
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
                     Icon            =   "frmNavigatorColor.frx":3BB2C
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   7
                     Left            =   480
                     TabIndex        =   10
                     Tag             =   "19B018"
                     Top             =   3960
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "入库细码单"
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
                     Icon            =   "frmNavigatorColor.frx":3DF96
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   4
                     Left            =   2400
                     TabIndex        =   11
                     Tag             =   "19B032"
                     Top             =   3960
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "染厂入库明细表"
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
                     Icon            =   "frmNavigatorColor.frx":40400
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   12
                     Left            =   480
                     TabIndex        =   12
                     Tag             =   "19B059"
                     Top             =   360
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布收发"
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
                     Icon            =   "frmNavigatorColor.frx":4286A
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   16
                     Left            =   2520
                     TabIndex        =   13
                     Tag             =   "19B063"
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布回修"
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
                     Icon            =   "frmNavigatorColor.frx":44CD4
                  End
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmNavigatorColor"
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





