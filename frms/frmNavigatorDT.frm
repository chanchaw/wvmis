VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmNavigatorDT 
   Caption         =   "基础资料"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNavigatorDT.frx":0000
   LinkTopic       =   "基础资料"
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   15960
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   9660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15960
      _LayoutVersion  =   1
      _ExtentX        =   28152
      _ExtentY        =   17039
      _DataPath       =   ""
      Bands           =   "frmNavigatorDT.frx":038A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   9375
         Left            =   540
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   15435
         _cx             =   27226
         _cy             =   16536
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
         _GridInfo       =   $"frmNavigatorDT.frx":0552
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   9315
            Left            =   30
            TabIndex        =   2
            Top             =   30
            Width           =   15375
            _cx             =   27120
            _cy             =   16431
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
               Height          =   8730
               Left            =   45
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   540
               Width           =   15285
               _cx             =   26961
               _cy             =   15399
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
               _GridInfo       =   $"frmNavigatorDT.frx":05D9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.PictureBox Picture1 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   8700
                  Left            =   15
                  ScaleHeight     =   8700
                  ScaleWidth      =   15255
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   15255
                  Begin VB.Frame Frame1 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "工序工艺"
                     Height          =   5055
                     Left            =   12360
                     TabIndex        =   37
                     Top             =   3720
                     Width           =   2655
                     Begin XtremeSuiteControls.PushButton btnObject 
                        Height          =   1335
                        Index           =   32
                        Left            =   480
                        TabIndex        =   38
                        Tag             =   "11B035"
                        Top             =   240
                        Width           =   1395
                        _Version        =   1048578
                        _ExtentX        =   2461
                        _ExtentY        =   2355
                        _StockProps     =   79
                        Caption         =   "缝制工艺"
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
                        Icon            =   "frmNavigatorDT.frx":065C
                     End
                     Begin XtremeSuiteControls.PushButton btnObject 
                        Height          =   1335
                        Index           =   33
                        Left            =   480
                        TabIndex        =   39
                        Tag             =   "11B036"
                        Top             =   1920
                        Width           =   1395
                        _Version        =   1048578
                        _ExtentX        =   2461
                        _ExtentY        =   2355
                        _StockProps     =   79
                        Caption         =   "工序工艺"
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
                        Icon            =   "frmNavigatorDT.frx":2AC6
                     End
                     Begin XtremeSuiteControls.PushButton btnObject 
                        Height          =   1335
                        Index           =   34
                        Left            =   480
                        TabIndex        =   40
                        Tag             =   "11B037"
                        Top             =   3600
                        Width           =   1395
                        _Version        =   1048578
                        _ExtentX        =   2461
                        _ExtentY        =   2355
                        _StockProps     =   79
                        Caption         =   "缝制工序"
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
                        Icon            =   "frmNavigatorDT.frx":4F30
                     End
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   0
                     Left            =   720
                     TabIndex        =   5
                     Tag             =   "11B004"
                     Top             =   60
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "计量单位"
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
                     Icon            =   "frmNavigatorDT.frx":739A
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   2
                     Left            =   4800
                     TabIndex        =   6
                     Tag             =   "11B001"
                     Top             =   60
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "部门分类"
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
                     Icon            =   "frmNavigatorDT.frx":9804
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   3
                     Left            =   6840
                     TabIndex        =   7
                     Tag             =   "11B002"
                     Top             =   60
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "员工资料"
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
                     Icon            =   "frmNavigatorDT.frx":BC6E
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   20
                     Left            =   8880
                     TabIndex        =   8
                     Tag             =   "11B022"
                     Top             =   0
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "仓库设置"
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
                     Icon            =   "frmNavigatorDT.frx":E0D8
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   4
                     Left            =   720
                     TabIndex        =   9
                     Tag             =   "11B005"
                     Top             =   1980
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "辅料类别"
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
                     Icon            =   "frmNavigatorDT.frx":10542
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   5
                     Left            =   2820
                     TabIndex        =   10
                     Tag             =   "11B006"
                     Top             =   1980
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "辅料资料"
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
                     Icon            =   "frmNavigatorDT.frx":129AC
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   6
                     Left            =   4800
                     TabIndex        =   11
                     Tag             =   "11B007"
                     Top             =   1980
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "原料类别"
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
                     Icon            =   "frmNavigatorDT.frx":14E16
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   7
                     Left            =   6840
                     TabIndex        =   12
                     Tag             =   "11B008"
                     Top             =   1980
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "原料资料"
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
                     Icon            =   "frmNavigatorDT.frx":17280
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   21
                     Left            =   8880
                     TabIndex        =   13
                     Tag             =   "11B021"
                     Top             =   1980
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "单据类型"
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
                     Icon            =   "frmNavigatorDT.frx":196EA
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   8
                     Left            =   720
                     TabIndex        =   14
                     Tag             =   "11B009"
                     Top             =   3840
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯类别"
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
                     Icon            =   "frmNavigatorDT.frx":1BB54
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   9
                     Left            =   2820
                     TabIndex        =   15
                     Tag             =   "11B010"
                     Top             =   3840
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "白坯资料"
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
                     Icon            =   "frmNavigatorDT.frx":1DFBE
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   10
                     Left            =   4800
                     TabIndex        =   16
                     Tag             =   "11B011"
                     Top             =   3840
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布类别"
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
                     Icon            =   "frmNavigatorDT.frx":20428
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   11
                     Left            =   6840
                     TabIndex        =   17
                     Tag             =   "11B012"
                     Top             =   3840
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "色布资料"
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
                     Icon            =   "frmNavigatorDT.frx":22892
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   12
                     Left            =   720
                     TabIndex        =   18
                     Tag             =   "11B013"
                     Top             =   5640
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "成品类别"
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
                     Icon            =   "frmNavigatorDT.frx":24CFC
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   1
                     Left            =   2820
                     TabIndex        =   19
                     Tag             =   "11B003"
                     Top             =   60
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "班次设置"
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
                     Icon            =   "frmNavigatorDT.frx":27166
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   13
                     Left            =   2820
                     TabIndex        =   20
                     Tag             =   "11B014"
                     Top             =   5580
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "成品资料"
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
                     Icon            =   "frmNavigatorDT.frx":295D0
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   14
                     Left            =   4800
                     TabIndex        =   21
                     Tag             =   "11B015"
                     Top             =   5580
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "交货方式"
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
                     Icon            =   "frmNavigatorDT.frx":2BA3A
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   15
                     Left            =   6840
                     TabIndex        =   22
                     Tag             =   "11B016"
                     Top             =   5580
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "结算方式"
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
                     Icon            =   "frmNavigatorDT.frx":2DEA4
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   16
                     Left            =   720
                     TabIndex        =   23
                     Tag             =   "11B017"
                     Top             =   7440
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "成品做工"
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
                     Icon            =   "frmNavigatorDT.frx":3030E
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   24
                     Left            =   8820
                     TabIndex        =   24
                     Tag             =   "11B024"
                     Top             =   5580
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "颜色类别"
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
                     Icon            =   "frmNavigatorDT.frx":32778
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   25
                     Left            =   8820
                     TabIndex        =   25
                     Tag             =   "11B025"
                     Top             =   3840
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "颜色类别分类"
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
                     Icon            =   "frmNavigatorDT.frx":34BE2
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   26
                     Left            =   8820
                     TabIndex        =   26
                     Tag             =   "11B026"
                     Top             =   7440
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "包装方式"
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
                     Icon            =   "frmNavigatorDT.frx":3704C
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   27
                     Left            =   10800
                     TabIndex        =   27
                     Tag             =   "11B027"
                     Top             =   0
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "进度项目表"
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
                     Icon            =   "frmNavigatorDT.frx":394B6
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   28
                     Left            =   10800
                     TabIndex        =   28
                     Tag             =   "11B028"
                     Top             =   1980
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "进度工艺分类"
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
                     Icon            =   "frmNavigatorDT.frx":3B920
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   29
                     Left            =   10800
                     TabIndex        =   29
                     Tag             =   "11B029"
                     Top             =   3840
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "进度工艺"
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
                     Icon            =   "frmNavigatorDT.frx":3DD8A
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   17
                     Left            =   2820
                     TabIndex        =   30
                     Tag             =   "11B018"
                     Top             =   7440
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "成品加工方"
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
                     Icon            =   "frmNavigatorDT.frx":401F4
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   18
                     Left            =   4800
                     TabIndex        =   31
                     Tag             =   "11B019"
                     Top             =   7440
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "往来单位类型"
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
                     Icon            =   "frmNavigatorDT.frx":4265E
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   19
                     Left            =   6840
                     TabIndex        =   32
                     Tag             =   "11B020"
                     Top             =   7440
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "往来单位"
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
                     Icon            =   "frmNavigatorDT.frx":44AC8
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   22
                     Left            =   10800
                     TabIndex        =   33
                     Tag             =   "11B031"
                     Top             =   5520
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "合同结算方式"
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
                     Icon            =   "frmNavigatorDT.frx":46F32
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   23
                     Left            =   10800
                     TabIndex        =   34
                     Tag             =   "11B032"
                     Top             =   7440
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "应付方结算方式"
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
                     Icon            =   "frmNavigatorDT.frx":4939C
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   30
                     Left            =   12720
                     TabIndex        =   35
                     Tag             =   "11B033"
                     Top             =   0
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "五金类别"
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
                     Icon            =   "frmNavigatorDT.frx":4B806
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   31
                     Left            =   12720
                     TabIndex        =   36
                     Tag             =   "11B034"
                     Top             =   2040
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "五金情况"
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
                     Icon            =   "frmNavigatorDT.frx":4DC70
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   35
                     Left            =   14640
                     TabIndex        =   41
                     Tag             =   "11S085"
                     Top             =   0
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "加工类型"
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
                     Icon            =   "frmNavigatorDT.frx":500DA
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   36
                     Left            =   14640
                     TabIndex        =   42
                     Tag             =   "11S086"
                     Top             =   2040
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "办公项目类型"
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
                     Icon            =   "frmNavigatorDT.frx":52544
                  End
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmNavigatorDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RGB
        Red   As Byte
        Green   As Byte
        Blue   As Byte
End Type

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
'    Select Case Index
'
'        Case 23
'        frmOrder.Show
'     BringWindow2Top frmOrder.hwnd
'    End Select
    
End Sub





Private Sub Form_Load()
'    Dim lRed As Long
'    Dim lGreen As Long
'    Dim lBlue As Long
'    lRed = 255
'    lGreen = 182
'    lBlue = 193
'
'    'Picture2.BackColor = RGB(lRed, lGreen, lBlue)
'    'Picture2.BackColor = &HC1B6FF
'
'    Dim ys As RGB
'    Dim lys As Long
'    ys.Red = 255
'    ys.Green = 182
'    ys.Blue = 193
'    lys = RGBToLong(ys)
'    Picture2.BackColor = lys
    
    InitFrm
End Sub


'Private Function RGBToLong(ColorRGB As RGB) As Long
'        RGBToLong = RGB(ColorRGB.Red, ColorRGB.Green, ColorRGB.Blue)
'End Function


Private Sub OpenObject(ByVal m_ObjectID As String, ByVal m_BillName As String)
    Gm.log4Runtime "进入导航页面的OpenObject"
    Gm.Authority.Execute m_ObjectID, m_BillName, "LoadObject", Nothing
    Gm.log4Runtime "打开命令执行完毕"
End Sub


'设置按钮的可用度
Public Sub ConfirmPermission()
    On Error Resume Next
    Dim i As Long
    Dim szObjectID As String
    
    
    For i = 0 To btnObject.Count - 1
        szObjectID = btnObject(i).Tag
        Debug.Print szObjectID
        If Len(szObjectID) > 0 Then
            If Gm.PI.JudgeView(szObjectID) = True Then
                btnObject(i).Enabled = True
            Else
                btnObject(i).Enabled = False
            End If
        End If
    Next
End Sub

