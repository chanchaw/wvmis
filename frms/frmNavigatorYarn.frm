VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.2#0"; "Codejock.CommandBars.v16.2.4.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmNavigatorYarn 
   Caption         =   "原料仓库导航"
   ClientHeight    =   9180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNavigatorYarn.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   9180
      Left            =   -360
      TabIndex        =   0
      Top             =   -120
      Width           =   13320
      _LayoutVersion  =   1
      _ExtentX        =   23495
      _ExtentY        =   16193
      _DataPath       =   ""
      Bands           =   "frmNavigatorYarn.frx":038A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   8655
         Left            =   540
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   12795
         _cx             =   22569
         _cy             =   15266
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
         _GridInfo       =   $"frmNavigatorYarn.frx":0552
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   8595
            Left            =   30
            TabIndex        =   2
            Top             =   30
            Width           =   12735
            _cx             =   22463
            _cy             =   15161
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
               Height          =   8010
               Left            =   45
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   540
               Width           =   12645
               _cx             =   22304
               _cy             =   14129
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
               _GridInfo       =   $"frmNavigatorYarn.frx":05D8
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.PictureBox Picture1 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   7980
                  Left            =   15
                  ScaleHeight     =   7980
                  ScaleWidth      =   12615
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   12615
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   6
                     Left            =   1200
                     TabIndex        =   5
                     Tag             =   "19B014"
                     Top             =   840
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "原料订单"
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
                     Icon            =   "frmNavigatorYarn.frx":065A
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   8
                     Left            =   5340
                     TabIndex        =   6
                     Tag             =   "19B019"
                     Top             =   840
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "原料无合同"
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
                     Icon            =   "frmNavigatorYarn.frx":2AC4
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   10
                     Left            =   3270
                     TabIndex        =   7
                     Tag             =   "19B030"
                     Top             =   840
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "原料入库"
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
                     Icon            =   "frmNavigatorYarn.frx":4F2E
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   11
                     Left            =   3270
                     TabIndex        =   8
                     Tag             =   "19B031"
                     Top             =   4620
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "原料对比表"
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
                     Icon            =   "frmNavigatorYarn.frx":7398
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   12
                     Left            =   1200
                     TabIndex        =   9
                     Tag             =   "19B036"
                     Top             =   4620
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "原料明细表"
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
                     Icon            =   "frmNavigatorYarn.frx":9802
                  End
                  Begin XtremeCommandBars.BackstageSeparator BackstageSeparator1 
                     Height          =   90
                     Left            =   960
                     TabIndex        =   10
                     Top             =   4200
                     Width           =   9435
                     _Version        =   1048578
                     _ExtentX        =   16642
                     _ExtentY        =   159
                     _StockProps     =   2
                     MarkupText      =   ""
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   13
                     Left            =   9360
                     TabIndex        =   11
                     Tag             =   "19B043"
                     Top             =   2640
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "原料发货/领料"
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
                     Icon            =   "frmNavigatorYarn.frx":BC6C
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   14
                     Left            =   3270
                     TabIndex        =   12
                     Tag             =   "19B046"
                     Top             =   2640
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "采购退货"
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
                     Icon            =   "frmNavigatorYarn.frx":E0D6
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   15
                     Left            =   5340
                     TabIndex        =   13
                     Tag             =   "19B048"
                     Top             =   4620
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "原料入库序时表"
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
                     Icon            =   "frmNavigatorYarn.frx":10540
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   16
                     Left            =   1200
                     TabIndex        =   14
                     Tag             =   "19B049"
                     Top             =   6540
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "原料库存表"
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
                     Icon            =   "frmNavigatorYarn.frx":129AA
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   17
                     Left            =   3270
                     TabIndex        =   15
                     Tag             =   "19B050"
                     Top             =   6540
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "加工户收发存汇总"
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
                     Icon            =   "frmNavigatorYarn.frx":14E14
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   18
                     Left            =   5340
                     TabIndex        =   16
                     Tag             =   "19B051"
                     Top             =   6540
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "原料收发存汇总"
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
                     Icon            =   "frmNavigatorYarn.frx":1727E
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   19
                     Left            =   7410
                     TabIndex        =   17
                     Tag             =   "19B052"
                     Top             =   2640
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "加工户退原料"
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
                     Icon            =   "frmNavigatorYarn.frx":196E8
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   20
                     Left            =   5340
                     TabIndex        =   18
                     Tag             =   "19B053"
                     Top             =   2640
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "原料调拨"
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
                     Icon            =   "frmNavigatorYarn.frx":1BB52
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   22
                     Left            =   7440
                     TabIndex        =   19
                     Tag             =   "19B064"
                     Top             =   6540
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "加工户收发存明细"
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
                     Icon            =   "frmNavigatorYarn.frx":1DFBC
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   0
                     Left            =   7440
                     TabIndex        =   20
                     Tag             =   "19B065"
                     Top             =   840
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "上年度库存"
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
                     Icon            =   "frmNavigatorYarn.frx":20426
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   2
                     Left            =   7440
                     TabIndex        =   21
                     Tag             =   "19B054"
                     Top             =   4620
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "原料发货序时表"
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
                     Icon            =   "frmNavigatorYarn.frx":22890
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   3
                     Left            =   9360
                     TabIndex        =   22
                     Tag             =   "19B072"
                     Top             =   6540
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "上年度报表"
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
                     Icon            =   "frmNavigatorYarn.frx":24CFA
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   1
                     Left            =   9360
                     TabIndex        =   23
                     Tag             =   "19B106"
                     Top             =   4620
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "转库单报表"
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
                     Icon            =   "frmNavigatorYarn.frx":27164
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   4
                     Left            =   9360
                     TabIndex        =   24
                     Tag             =   "19B107"
                     Top             =   840
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "转库单"
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
                     Icon            =   "frmNavigatorYarn.frx":295CE
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   5
                     Left            =   11280
                     TabIndex        =   25
                     Tag             =   "19B120"
                     Top             =   4620
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "供应商对账单"
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
                     Icon            =   "frmNavigatorYarn.frx":2BA38
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   7
                     Left            =   1200
                     TabIndex        =   26
                     Tag             =   "19B123"
                     Top             =   2640
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "原料采购序时表"
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
                     Icon            =   "frmNavigatorYarn.frx":2DEA2
                  End
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmNavigatorYarn"
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
'    btnObject(0).Visible = False
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


