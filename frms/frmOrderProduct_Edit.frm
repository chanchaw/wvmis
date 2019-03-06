VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmOrderProduct_Edit 
   Caption         =   "设置明细"
   ClientHeight    =   10125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOrderProduct_Edit.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10125
   ScaleWidth      =   18375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   10125
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   18375
      _LayoutVersion  =   1
      _ExtentX        =   32411
      _ExtentY        =   17859
      _DataPath       =   ""
      Bands           =   "frmOrderProduct_Edit.frx":058A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   10155
         Left            =   360
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   17835
         _cx             =   31459
         _cy             =   17912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
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
         BorderWidth     =   6
         ChildSpacing    =   4
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
         _GridInfo       =   $"frmOrderProduct_Edit.frx":1336
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   9975
            Left            =   90
            ScaleHeight     =   9975
            ScaleWidth      =   17655
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   90
            Width           =   17655
            Begin VB.Frame Frame4 
               Height          =   5055
               Left            =   7200
               TabIndex        =   100
               Top             =   960
               Width           =   3495
               Begin VB.PictureBox Picture6 
                  Height          =   375
                  Left            =   1320
                  ScaleHeight     =   315
                  ScaleWidth      =   1875
                  TabIndex        =   114
                  TabStop         =   0   'False
                  Top             =   2730
                  Width           =   1935
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit14 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   101
                  Top             =   240
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit19 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   102
                  Top             =   4575
                  Width           =   1575
                  _Version        =   1048578
                  _ExtentX        =   2778
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   14737632
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   0   'False
                  BackColor       =   14737632
               End
               Begin XtremeSuiteControls.PushButton PushButton5 
                  Height          =   375
                  Left            =   2880
                  TabIndex        =   103
                  Top             =   4575
                  Width           =   375
                  _Version        =   1048578
                  _ExtentX        =   661
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   ".."
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit23 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   104
                  Top             =   765
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit24 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   105
                  Top             =   1275
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit41 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   110
                  Top             =   1770
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit42 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   112
                  Top             =   2250
                  Width           =   1695
                  _Version        =   1048578
                  _ExtentX        =   2990
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   0   'False
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit43 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   116
                  Top             =   3210
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit44 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   117
                  Top             =   4050
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.ComboBox ComboBox4 
                  Height          =   345
                  Left            =   1320
                  TabIndex        =   118
                  Top             =   3660
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   609
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.PushButton PushButton12 
                  Height          =   375
                  Left            =   3000
                  TabIndex        =   122
                  Top             =   2250
                  Width           =   375
                  _Version        =   1048578
                  _ExtentX        =   661
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   ".."
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.Label Label51 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   121
                  Top             =   3270
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "数   量："
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label50 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   120
                  Top             =   3690
                  Width           =   1095
                  _Version        =   1048578
                  _ExtentX        =   1931
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "计算单位："
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label49 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   119
                  Top             =   4110
                  Width           =   1035
                  _Version        =   1048578
                  _ExtentX        =   1826
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "单   价："
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label48 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   115
                  Top             =   2790
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "色  块："
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label47 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   113
                  Top             =   2310
                  Width           =   855
                  _Version        =   1048578
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "颜   色："
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label46 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   111
                  Top             =   1830
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "色   号："
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label15 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   109
                  Top             =   300
                  Width           =   1035
                  _Version        =   1048578
                  _ExtentX        =   1826
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "背面面料:"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label21 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   108
                  Top             =   4650
                  Width           =   1035
                  _Version        =   1048578
                  _ExtentX        =   1826
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "背面工厂:"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label25 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   107
                  Top             =   825
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "背面门幅："
                  ForeColor       =   255
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label26 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   106
                  Top             =   1335
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "背面克重："
                  ForeColor       =   255
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
            Begin VB.Frame Frame3 
               Height          =   5055
               Left            =   3600
               TabIndex        =   77
               Top             =   960
               Width           =   3495
               Begin VB.PictureBox Picture4 
                  Height          =   375
                  Left            =   1320
                  ScaleHeight     =   315
                  ScaleWidth      =   1875
                  TabIndex        =   91
                  TabStop         =   0   'False
                  Top             =   2760
                  Width           =   1935
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit6 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   78
                  Top             =   240
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit18 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   79
                  Top             =   4605
                  Width           =   1575
                  _Version        =   1048578
                  _ExtentX        =   2778
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   14737632
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   0   'False
                  BackColor       =   14737632
               End
               Begin XtremeSuiteControls.PushButton PushButton4 
                  Height          =   375
                  Left            =   2880
                  TabIndex        =   80
                  Top             =   4605
                  Width           =   375
                  _Version        =   1048578
                  _ExtentX        =   661
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   ".."
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit21 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   81
                  Top             =   795
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit22 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   82
                  Top             =   1305
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit36 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   87
                  Top             =   1800
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit38 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   89
                  Top             =   2280
                  Width           =   1695
                  _Version        =   1048578
                  _ExtentX        =   2990
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   0   'False
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit39 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   93
                  Top             =   3240
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit40 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   94
                  Top             =   4080
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.ComboBox ComboBox3 
                  Height          =   345
                  Left            =   1320
                  TabIndex        =   95
                  Top             =   3720
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   609
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.PushButton PushButton11 
                  Height          =   375
                  Left            =   3000
                  TabIndex        =   99
                  Top             =   2280
                  Width           =   375
                  _Version        =   1048578
                  _ExtentX        =   661
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   ".."
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.Label Label45 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   98
                  Top             =   3300
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "数   量："
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label44 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   97
                  Top             =   3720
                  Width           =   1095
                  _Version        =   1048578
                  _ExtentX        =   1931
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "计算单位："
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label43 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   96
                  Top             =   4140
                  Width           =   1035
                  _Version        =   1048578
                  _ExtentX        =   1826
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "单   价："
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label42 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   92
                  Top             =   2820
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "色  块："
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label41 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   90
                  Top             =   2340
                  Width           =   855
                  _Version        =   1048578
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "颜   色："
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label39 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   88
                  Top             =   1860
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "色   号："
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label7 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   86
                  Top             =   240
                  Width           =   1215
                  _Version        =   1048578
                  _ExtentX        =   2143
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "中间面料："
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label20 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   85
                  Top             =   4605
                  Width           =   1215
                  _Version        =   1048578
                  _ExtentX        =   2143
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "中间工厂："
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label23 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   84
                  Top             =   855
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "中间门幅："
                  ForeColor       =   255
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label24 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   83
                  Top             =   1365
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "中间克重："
                  ForeColor       =   255
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
            Begin VB.Frame Frame2 
               Height          =   5055
               Left            =   0
               TabIndex        =   52
               Top             =   960
               Width           =   3495
               Begin VB.PictureBox Picture2 
                  Height          =   375
                  Left            =   1320
                  ScaleHeight     =   315
                  ScaleWidth      =   1875
                  TabIndex        =   53
                  TabStop         =   0   'False
                  Top             =   2730
                  Width           =   1935
               End
               Begin XtremeSuiteControls.PushButton PushButton1 
                  Height          =   375
                  Left            =   2940
                  TabIndex        =   54
                  Top             =   2250
                  Width           =   375
                  _Version        =   1048578
                  _ExtentX        =   661
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   ".."
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit3 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   55
                  Top             =   765
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit4 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   56
                  Top             =   1275
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit5 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   57
                  Top             =   240
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit11 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   58
                  Top             =   2250
                  Width           =   1695
                  _Version        =   1048578
                  _ExtentX        =   2990
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   0   'False
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit17 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   59
                  Top             =   4575
                  Width           =   1575
                  _Version        =   1048578
                  _ExtentX        =   2778
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   14737632
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   0   'False
                  BackColor       =   14737632
               End
               Begin XtremeSuiteControls.PushButton PushButton3 
                  Height          =   375
                  Left            =   2880
                  TabIndex        =   60
                  Top             =   4575
                  Width           =   375
                  _Version        =   1048578
                  _ExtentX        =   661
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   ".."
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit15 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   61
                  Top             =   1770
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit30 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   62
                  Top             =   3240
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.FlatEdit FlatEdit32 
                  DataField       =   "B_CodeID"
                  DataSource      =   "Adodc1"
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   63
                  Top             =   4080
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.ComboBox ComboBox1 
                  Height          =   345
                  Left            =   1320
                  TabIndex        =   64
                  Top             =   3690
                  Width           =   1935
                  _Version        =   1048578
                  _ExtentX        =   3413
                  _ExtentY        =   609
                  _StockProps     =   77
                  ForeColor       =   0
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label6 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   74
                  Top             =   300
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "正面面料："
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label5 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   73
                  Top             =   2310
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "颜   色："
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label4 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   72
                  Top             =   1335
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "正面克重："
                  ForeColor       =   255
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label3 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   71
                  Top             =   825
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "正面门幅："
                  ForeColor       =   255
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label12 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   70
                  Top             =   1830
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "色   号："
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label16 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   69
                  Top             =   2790
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "色  块："
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label19 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   68
                  Top             =   4635
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "正面工厂："
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label32 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   67
                  Top             =   3300
                  Width           =   975
                  _Version        =   1048578
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "数   量："
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label33 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   66
                  Top             =   3720
                  Width           =   1095
                  _Version        =   1048578
                  _ExtentX        =   1931
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "计算单位："
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin XtremeSuiteControls.Label Label34 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   65
                  Top             =   4140
                  Width           =   1035
                  _Version        =   1048578
                  _ExtentX        =   1826
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "单   价："
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   11.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   16320
               Top             =   240
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Frame Frame1 
               Height          =   855
               Left            =   10800
               TabIndex        =   38
               Top             =   0
               Width           =   6735
               Begin XtremeSuiteControls.PushButton PushButton7 
                  Height          =   495
                  Left            =   240
                  TabIndex        =   39
                  Top             =   240
                  Width           =   1215
                  _Version        =   1048578
                  _ExtentX        =   2143
                  _ExtentY        =   873
                  _StockProps     =   79
                  Caption         =   "预览图片"
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.PushButton PushButton8 
                  Height          =   495
                  Left            =   2160
                  TabIndex        =   40
                  Top             =   240
                  Width           =   1215
                  _Version        =   1048578
                  _ExtentX        =   2143
                  _ExtentY        =   873
                  _StockProps     =   79
                  Caption         =   "上传图片"
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.PushButton PushButton9 
                  Height          =   495
                  Left            =   3960
                  TabIndex        =   41
                  Top             =   240
                  Width           =   1215
                  _Version        =   1048578
                  _ExtentX        =   2143
                  _ExtentY        =   873
                  _StockProps     =   79
                  Caption         =   "删除图片"
                  UseVisualStyle  =   -1  'True
               End
            End
            Begin VB.PictureBox Picture3 
               Height          =   5655
               Left            =   10800
               ScaleHeight     =   5595
               ScaleWidth      =   6795
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   960
               Width           =   6855
               Begin VB.PictureBox Picture5 
                  Height          =   5595
                  Left            =   0
                  ScaleHeight     =   5535
                  ScaleWidth      =   6735
                  TabIndex        =   42
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   6795
               End
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit29 
               Height          =   375
               Left            =   8520
               TabIndex        =   36
               Top             =   60
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit12 
               Bindings        =   "frmOrderProduct_Edit.frx":13BD
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "0.0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
               Height          =   375
               Left            =   8280
               TabIndex        =   4
               Top             =   6180
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1560
               TabIndex        =   0
               Top             =   60
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit2 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   4920
               TabIndex        =   15
               Top             =   540
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BackColor       =   14737632
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit8 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   4920
               TabIndex        =   3
               Top             =   6180
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit10 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   4920
               TabIndex        =   16
               Top             =   6690
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit7 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1560
               TabIndex        =   2
               Top             =   6180
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit9 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1560
               TabIndex        =   5
               Top             =   6690
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.PushButton PushButton2 
               Height          =   375
               Left            =   6480
               TabIndex        =   25
               Top             =   540
               Width           =   375
               _Version        =   1048578
               _ExtentX        =   661
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.ComboBox ComboBox2 
               Height          =   345
               Left            =   1560
               TabIndex        =   11
               Top             =   555
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   609
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit20 
               Height          =   855
               Left            =   1320
               TabIndex        =   6
               Top             =   7680
               Width           =   4695
               _Version        =   1048578
               _ExtentX        =   8281
               _ExtentY        =   1508
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MultiLine       =   -1  'True
               ScrollBars      =   2
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit25 
               Height          =   855
               Left            =   1320
               TabIndex        =   7
               Top             =   8640
               Width           =   4695
               _Version        =   1048578
               _ExtentX        =   8281
               _ExtentY        =   1508
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MultiLine       =   -1  'True
               ScrollBars      =   2
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit16 
               Height          =   855
               Left            =   7320
               TabIndex        =   8
               Top             =   7680
               Width           =   4935
               _Version        =   1048578
               _ExtentX        =   8705
               _ExtentY        =   1508
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MultiLine       =   -1  'True
               ScrollBars      =   2
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit26 
               Height          =   855
               Left            =   7320
               TabIndex        =   9
               Top             =   8640
               Width           =   4935
               _Version        =   1048578
               _ExtentX        =   8705
               _ExtentY        =   1508
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MultiLine       =   -1  'True
               ScrollBars      =   2
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit27 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   8520
               TabIndex        =   1
               Top             =   540
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.PushButton PushButton6 
               Height          =   375
               Left            =   6420
               TabIndex        =   32
               Top             =   60
               Width           =   495
               _Version        =   1048578
               _ExtentX        =   873
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit28 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   4920
               TabIndex        =   33
               Top             =   60
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit13 
               Height          =   855
               Left            =   13440
               TabIndex        =   10
               Top             =   7680
               Width           =   4215
               _Version        =   1048578
               _ExtentX        =   7435
               _ExtentY        =   1508
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MultiLine       =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit33 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1560
               TabIndex        =   43
               Top             =   7200
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BackColor       =   14737632
            End
            Begin XtremeSuiteControls.PushButton PushButton10 
               Height          =   375
               Left            =   3120
               TabIndex        =   44
               Top             =   7200
               Width           =   375
               _Version        =   1048578
               _ExtentX        =   661
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit31 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   4920
               TabIndex        =   49
               Top             =   7200
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit34 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   8280
               TabIndex        =   50
               Top             =   7200
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit35 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   11520
               TabIndex        =   51
               Top             =   7200
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
            End
            Begin VB.Line Line1 
               X1              =   120
               X2              =   10320
               Y1              =   6120
               Y2              =   6120
            End
            Begin XtremeSuiteControls.Label Label38 
               Height          =   255
               Left            =   10440
               TabIndex        =   48
               Top             =   7260
               Width           =   1095
               _Version        =   1048578
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "加工金额："
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label37 
               Height          =   255
               Left            =   7080
               TabIndex        =   47
               Top             =   7260
               Width           =   1095
               _Version        =   1048578
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "加工单价："
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label36 
               Height          =   255
               Left            =   3720
               TabIndex        =   46
               Top             =   7260
               Width           =   1095
               _Version        =   1048578
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "加工数量："
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label35 
               Height          =   255
               Left            =   360
               TabIndex        =   45
               Top             =   7260
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "家纺厂："
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "款   号："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   7320
               TabIndex        =   35
               Top             =   120
               Width           =   1035
            End
            Begin XtremeSuiteControls.Label Label30 
               Height          =   255
               Left            =   3720
               TabIndex        =   34
               Top             =   120
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "源订单号："
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label29 
               Height          =   255
               Left            =   7320
               TabIndex        =   31
               Top             =   600
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "尺   寸："
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label28 
               Height          =   315
               Left            =   6240
               TabIndex        =   30
               Top             =   9000
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "打     箱:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label27 
               Height          =   315
               Left            =   240
               TabIndex        =   29
               Top             =   9000
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "缝      制:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label22 
               Height          =   315
               Left            =   240
               TabIndex        =   28
               Top             =   7920
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "裁     剪:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label18 
               Height          =   315
               Left            =   6240
               TabIndex        =   27
               Top             =   8040
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "包     装:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label17 
               Height          =   255
               Left            =   360
               TabIndex        =   26
               Top             =   600
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "手工品名："
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   255
               Left            =   360
               TabIndex        =   24
               Top             =   120
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "订单号："
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   255
               Left            =   3720
               TabIndex        =   23
               Top             =   600
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "品   名："
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label8 
               Height          =   255
               Left            =   360
               TabIndex        =   22
               Top             =   6240
               Width           =   1095
               _Version        =   1048578
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "数   量："
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label9 
               Height          =   255
               Left            =   7080
               TabIndex        =   21
               Top             =   6240
               Width           =   1095
               _Version        =   1048578
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "箱    数："
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label10 
               Height          =   255
               Left            =   360
               TabIndex        =   20
               Top             =   6720
               Width           =   1095
               _Version        =   1048578
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "单   价："
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label11 
               Height          =   255
               Left            =   3720
               TabIndex        =   19
               Top             =   6750
               Width           =   1095
               _Version        =   1048578
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "金   额："
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label13 
               Height          =   255
               Left            =   3720
               TabIndex        =   18
               Top             =   6240
               Width           =   1035
               _Version        =   1048578
               _ExtentX        =   1826
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "每箱数量:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label14 
               Height          =   315
               Left            =   12360
               TabIndex        =   17
               Top             =   8040
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "备      注:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
      End
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit37 
      DataField       =   "B_CodeID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6480
      TabIndex        =   75
      Top             =   -2520
      Width           =   1935
      _Version        =   1048578
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label40 
      Height          =   255
      Left            =   5760
      TabIndex        =   76
      Top             =   3900
      Width           =   975
      _Version        =   1048578
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "色   号："
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOrderProduct_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Valuation As String
Private num As Long
Public bool As Boolean
Public id As String
Public itemid As String
Public colorid As String
Public colorid2 As String
Public colorid3 As String
'验证身份和时间
Private clspI As New clspI

Public Productid As String
Public Positiveid As String
Public Middleid As String
Public backid As String
Public client As String

Private cls1 As New clsPicture
Private szFile As String

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    
        Case "保存"
            save
        Case "退出"
            Unload Me
        
    End Select
End Sub

Private Sub save()
    If clspI.authenticate(id) = False Then
            Exit Sub
      End If
    
    If Trim(FlatEdit1.Text) = "" Then
        MsgBox "订单号不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(FlatEdit2.Text) = "" And ComboBox2.Text = "" Then
        MsgBox "品名不能为全部为空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(FlatEdit3.Text) = "" Then
        MsgBox "门幅不能为空", vbInformation, "提示"
        Exit Sub
    End If
     If Trim(FlatEdit4.Text) = "" Then
        MsgBox "克重不能为空", vbInformation, "提示"
        Exit Sub
    End If
     If Trim(FlatEdit11.Text) = "" And Trim(FlatEdit15.Text) = "" Then
        MsgBox "颜色或色号不能都为空", vbInformation, "提示"
        Exit Sub
    End If
    If Len(FlatEdit5.Text) <= 0 And Len(FlatEdit6.Text) <= 0 And Len(FlatEdit14.Text) <= 0 Then
        MsgBox "3个位置面料不能为空", vbInformation, "提示"
        Exit Sub
    End If

 
     If Trim(FlatEdit7.Text) <= 0 Then
            MsgBox "数量不能为空", vbInformation, "提示"
            Exit Sub
    End If
 
    
     If Trim(FlatEdit27.Text) = "" Then
        MsgBox "尺寸不能为空", vbInformation, "提示"
        Exit Sub
    End If
    
    If Trim(FlatEdit11.Text) = "" Then
        colorid = ""
    End If
    
    savedetail
    
End Sub

Private Sub GotFocusBox1()
ComboBox1.Clear
ComboBox1.AddItem "公斤"
ComboBox1.AddItem "米数"

ComboBox3.AddItem "公斤"
ComboBox3.AddItem "米数"

ComboBox4.AddItem "公斤"
ComboBox4.AddItem "米数"

End Sub

Private Sub ComboBox2_GotFocus()
    ComboBox2.Clear
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select distinct B_GoodManual from G_BilldetailOrder WHERE B_GoodManual LIKE '%'+'" & ComboBox2.Text & "'+'%'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Do While Not rs.EOF
        ComboBox2.AddItem "" & rs!B_GoodManual & ""
        rs.movenext
    Loop
End Sub



Private Sub FlatEdit11_Change()
      On Error Resume Next
        Dim rs As New RecordSet
        Dim sql As String
        If colorid <> "" Then
            FlatEdit11.Enabled = True
        End If
        sql = "select * from G_Color where B_SID='" & colorid & "'"
        Debug.Print colorid
        rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        Picture2.BackColor = rs!B_hex
End Sub

Private Sub FlatEdit15_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 13
        If Len(FlatEdit11.Text) <= 0 Then
            PushButton1_Click
        Else
            FlatEdit5.SetFocus
        End If
    End Select
End Sub

Private Sub FlatEdit31_Change()
FlatEdit35.Text = Format(Val(FlatEdit31.Text) * Val(FlatEdit34), "0.00")
End Sub

Private Sub FlatEdit34_Change()
FlatEdit35.Text = Format(Val(FlatEdit31.Text) * Val(FlatEdit34), "0.00")
End Sub

Private Sub FlatEdit38_Change()
On Error Resume Next
        Dim rs As New RecordSet
        Dim sql As String
        If colorid2 <> "" Then
            FlatEdit38.Enabled = True
        End If
        sql = "select * from G_Color where B_SID='" & colorid2 & "'"
        Debug.Print colorid
        rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        Picture4.BackColor = rs!B_hex
End Sub

Private Sub FlatEdit42_Change()
On Error Resume Next
        Dim rs As New RecordSet
        Dim sql As String
        If colorid3 <> "" Then
            FlatEdit42.Enabled = True
        End If
        sql = "select * from G_Color where B_SID='" & colorid3 & "'"
        Debug.Print colorid
        rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        Picture6.BackColor = rs!B_hex
End Sub

Private Sub FlatEdit8_Change()
        Dim a As Long
        Dim b As Double
       If Val(FlatEdit8) <> 0 Then
            a = Val(FlatEdit7.Text) / Val(FlatEdit8)
            b = Val(FlatEdit7.Text) / Val(FlatEdit8)
            If b > a Then
                FlatEdit12.Text = a + 1
            Else
                FlatEdit12.Text = a
            End If
            
       End If
End Sub

Private Sub Form_Load()
    InitFrm
    num = 0
    bool = False
    HT_Name
   GotFocusBox1
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
End Sub
'将合同号名字替换（根据不同工厂不同命名）
Private Sub HT_Name()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "select * from G_Config_FormCtlShow where B_sid='订单号' "
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        Label1.Caption = rs!B_Caption
    End If
End Sub

'正面面料颜色
Private Sub PushButton1_Click()
    On Error Resume Next
    Dim sql As String
    Dim rs As New RecordSet
    Dim frm1 As New frmpopupColor
    frm1.Show vbModal
    FlatEdit11.Text = Trim(frm1.colorname)
    colorid = frm1.colorid
    If frm1.bsaved = True Then
        FlatEdit11.Enabled = True
    Else
        If colorid = "" Then
             FlatEdit11.Enabled = False
        End If
    End If
    
        sql = "select * from G_Color where B_SID='" & colorid & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

        If rs!B_hex <> "" Then
        
            Picture2.BackColor = rs!B_hex
        Else
            Picture2.BackColor = &H8000000F
        End If
      
    
    Unload frm1
End Sub
Private Sub FlatEdit7_Change()
        Dim a As Long
        Dim b As Double
        If Val(FlatEdit8) <> 0 Then
            a = Val(FlatEdit7.Text) / Val(FlatEdit8)
            b = Val(FlatEdit7.Text) / Val(FlatEdit8)
            If b > a Then
                FlatEdit12.Text = a + 1
            Else
                FlatEdit12.Text = a
            End If
       End If
End Sub

Private Sub FlatEdit9_Change()
        FlatEdit10.Text = Format(Val(FlatEdit7.Text) * Val(FlatEdit9), "0.00")
End Sub


Private Sub savedetail()

    If Len(itemid) > 0 Then
        savedetail_update
    Else
        Dim rs1 As New RecordSet
        Dim sql1 As String
        sql1 = "select *from G_BillOrder where B_id='" & id & "'"
        rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        
        If rs1.RecordCount > 0 Then
            Detail
        Else
            Dim rs3 As New RecordSet
            Dim sql3 As String
            sql3 = "exec usp_savedetailProduct '" & id & "','" & FlatEdit1.Text & "','" & Productid & "','" & FlatEdit3.Text & "','" & FlatEdit4.Text & "','" & FlatEdit15.Text & "','" & FlatEdit5.Text & "','" & FlatEdit6.Text & "','" & FlatEdit14.Text & "','" & FlatEdit7.Text & "','" & FlatEdit12.Text & "','" & FlatEdit8.Text & "','" & FlatEdit9.Text & "','" & FlatEdit10.Text & "','" & FlatEdit13.Text & "','" & FlatEdit11.Text & "','" & colorid & "'"
            sql3 = sql3 & ",'" & ComboBox2.Text & "','" & Positiveid & "','" & Middleid & "','" & backid & "','" & FlatEdit20.Text & "','" & FlatEdit16.Text & "'"
            sql3 = sql3 & ",'" & FlatEdit21.Text & "','" & FlatEdit22.Text & "','" & FlatEdit23.Text & "','" & FlatEdit24.Text & "','" & FlatEdit25.Text & "','" & FlatEdit26.Text & "','" & FlatEdit27.Text & "','" & Gm.SysID.SystemUser & "','1','" & FlatEdit28.Text & "','" & FlatEdit29.Text & "'"
            sql3 = sql3 & ",'" & FlatEdit30.Text & "','" & ComboBox1.Text & "','" & FlatEdit32.Text & "','" & FlatEdit33.Text & "','" & FlatEdit31.Text & "','" & FlatEdit34.Text & "','" & FlatEdit35.Text & "'"
            sql3 = sql3 & ",'" & FlatEdit36.Text & "','" & colorid2 & "','" & FlatEdit38.Text & "','" & FlatEdit39.Text & "','" & ComboBox3.Text & "','" & FlatEdit40.Text & "'"
            sql3 = sql3 & ",'" & FlatEdit41.Text & "','" & colorid3 & "','" & FlatEdit42.Text & "','" & FlatEdit43.Text & "','" & ComboBox4.Text & "','" & FlatEdit44.Text & "'"
   
            Debug.Print sql3
            rs3.Open sql3, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
         End If
        bool = True
        
    End If
 
    Me.Hide
End Sub
'成功保存后全部清空
Private Sub AddNew()
    On Error Resume Next
    Dim o As Object
    
    For Each o In Me.Controls
        Select Case TypeName(o)
        
            Case "FlatEdit"
                o.Text = ""
                
        End Select
    Next
    
End Sub

Private Sub savedetail_update()
    Dim rs As New RecordSet
    Dim sql1 As String
    sql1 = "select * from G_BillDetailOrder where B_itemid='" & itemid & "'"
    rs.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        Dim sql3 As String
        sql3 = "exec usp_updateAll '" & id & "','" & rs!B_ordercode & "','" & FlatEdit1.Text & "'"
        Debug.Print sql3
         Gm.cnnTool.cnn.Execute sql3
         
        Dim sql2 As String
         sql2 = "exec usp_saveProductdetail_update '" & itemid & "','" & FlatEdit1.Text & "','" & Productid & "','" & FlatEdit3.Text & "','" & FlatEdit4.Text & "','" & FlatEdit15.Text & "','" & FlatEdit5.Text & "','" & FlatEdit6.Text & "','" & FlatEdit14.Text & "','" & FlatEdit7.Text & "','" & FlatEdit12.Text & "','" & FlatEdit8.Text & "','" & FlatEdit9.Text & "','" & FlatEdit10.Text & "','" & FlatEdit13.Text & "','" & FlatEdit11.Text & "','" & colorid & "'"
         sql2 = sql2 & ",'" & ComboBox2.Text & "','" & Positiveid & "','" & Middleid & "','" & backid & "','" & FlatEdit20.Text & "','" & FlatEdit16.Text & "'"
         sql2 = sql2 & ",'" & FlatEdit21.Text & "','" & FlatEdit22.Text & "','" & FlatEdit23.Text & "','" & FlatEdit24.Text & "','" & FlatEdit25.Text & "','" & FlatEdit26.Text & "','" & FlatEdit27.Text & "','" & FlatEdit28.Text & "','" & FlatEdit29.Text & "'"
         sql2 = sql2 & ",'" & FlatEdit30.Text & "','" & ComboBox1.Text & "','" & FlatEdit32.Text & "','" & FlatEdit33.Text & "','" & FlatEdit31.Text & "','" & FlatEdit34.Text & "','" & FlatEdit35.Text & "'"
         sql2 = sql2 & ",'" & FlatEdit36.Text & "','" & colorid2 & "','" & FlatEdit38.Text & "','" & FlatEdit39.Text & "','" & ComboBox3.Text & "','" & FlatEdit40.Text & "'"
         sql2 = sql2 & ",'" & FlatEdit41.Text & "','" & colorid3 & "','" & FlatEdit42.Text & "','" & FlatEdit43.Text & "','" & ComboBox4.Text & "','" & FlatEdit44.Text & "'"
         
         Debug.Print sql2
        Gm.cnnTool.cnn.Execute sql2
        
    Else
        Dim sql As String
        sql = "exec usp_saveProductdetail_Draftupdate '" & itemid & "','" & FlatEdit1.Text & "','" & Productid & "','" & FlatEdit3.Text & "','" & FlatEdit4.Text & "','" & FlatEdit15.Text & "','" & FlatEdit5.Text & "','" & FlatEdit6.Text & "','" & FlatEdit14.Text & "','" & FlatEdit7.Text & "','" & FlatEdit12.Text & "','" & FlatEdit8.Text & "','" & FlatEdit9.Text & "','" & FlatEdit10.Text & "','" & FlatEdit13.Text & "','" & FlatEdit11.Text & "','" & colorid & "'"
        sql = sql & ",'" & ComboBox2.Text & "','" & Positiveid & "','" & Middleid & "','" & backid & "','" & FlatEdit20.Text & "','" & FlatEdit16.Text & "'"
        sql = sql & ",'" & FlatEdit21.Text & "','" & FlatEdit22.Text & "','" & FlatEdit23.Text & "','" & FlatEdit24.Text & "','" & FlatEdit25.Text & "','" & FlatEdit26.Text & "','" & FlatEdit27.Text & "','" & FlatEdit28.Text & "','" & FlatEdit29.Text & "'"
        sql = sql & ",'" & FlatEdit30.Text & "','" & ComboBox1.Text & "','" & FlatEdit32.Text & "','" & FlatEdit33.Text & "','" & FlatEdit31.Text & "','" & FlatEdit34.Text & "','" & FlatEdit35.Text & "'"
        sql = sql & ",'" & FlatEdit36.Text & "','" & colorid2 & "','" & FlatEdit38.Text & "','" & FlatEdit39.Text & "','" & ComboBox3.Text & "','" & FlatEdit40.Text & "'"
        sql = sql & ",'" & FlatEdit41.Text & "','" & colorid3 & "','" & FlatEdit42.Text & "','" & FlatEdit43.Text & "','" & ComboBox4.Text & "','" & FlatEdit44.Text & "'"
         
        Gm.cnnTool.cnn.Execute sql
    End If
End Sub

Private Sub FlatEdit7_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 And Not CBool(InStr(txbNumber, ".")) Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit8_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 And Not CBool(InStr(txbNumber, ".")) Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit9_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 And Not CBool(InStr(txbNumber, ".")) Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit12_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 And Not CBool(InStr(txbNumber, ".")) Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub

'明细表中有数据，进行删除
Private Sub Detail()
    Dim sql As String
    sql = "delete from G_DraftBillDetailOrder where B_ID='" & id & "'"
    Gm.cnnTool.cnn.Execute sql
    Dim rs1 As New RecordSet
    Dim sql1 As String
    sql1 = "exec usp_savedetailProduct '" & id & "','" & FlatEdit1.Text & "','" & Productid & "','" & FlatEdit3.Text & "','" & FlatEdit4.Text & "','" & FlatEdit15.Text & "','" & FlatEdit5.Text & "','" & FlatEdit6.Text & "','" & FlatEdit14.Text & "','" & FlatEdit7.Text & "','" & FlatEdit12.Text & "','" & FlatEdit8.Text & "','" & FlatEdit9.Text & "','" & FlatEdit10.Text & "','" & FlatEdit13.Text & "','" & FlatEdit11.Text & "','" & colorid & "'"
    sql1 = sql1 & ",'" & ComboBox2.Text & "','" & Positiveid & "','" & Middleid & "','" & backid & "','" & FlatEdit20.Text & "','" & FlatEdit16.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit21.Text & "','" & FlatEdit22.Text & "','" & FlatEdit23.Text & "','" & FlatEdit24.Text & "','" & FlatEdit25.Text & "','" & FlatEdit26.Text & "','" & FlatEdit27.Text & "','" & Gm.SysID.SystemUser & "','1','" & FlatEdit28.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit30.Text & "','" & ComboBox1.Text & "','" & FlatEdit32.Text & "','" & FlatEdit33.Text & "','" & FlatEdit31.Text & "','" & FlatEdit34.Text & "','" & FlatEdit35.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit36.Text & "','" & colorid2 & "','" & FlatEdit38.Text & "','" & FlatEdit39.Text & "','" & ComboBox3.Text & "','" & FlatEdit40.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit41.Text & "','" & colorid3 & "','" & FlatEdit42.Text & "','" & FlatEdit43.Text & "','" & ComboBox4.Text & "','" & FlatEdit44.Text & "'"
    Debug.Print sql1
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    Dim sql2 As String
    sql2 = "insert into G_BillDetailOrder (B_itemid,B_ID,B_OrderCode,B_GoodsID,B_Width,B_Weight,B_patterncode,B_Positivefabric,B_Middlefabric,B_Backfabric,B_Qty,B_BoxQty,B_QtyPerbox,B_Price,B_Sum,B_MemoDetail,B_color,B_colorid,B_GoodManual,B_PositiveFactory,B_MiddleFactory,B_BackFactory,B_Process,B_Packaging,B_Width2,B_Weight2,B_Width3,B_Weight3,B_seam,B_Beatbox,B_size,B_username,B_ContractLogo,B_SourceOrderCode)  select B_itemid,B_ID,B_OrderCode,B_GoodsID,B_Width,B_Weight,B_patterncode,B_Positivefabric,B_Middlefabric,B_Backfabric,B_Qty,B_BoxQty,B_QtyPerbox,B_Price,B_Sum,B_MemoDetail,B_color,B_colorid,B_GoodManual,B_PositiveFactory,B_MiddleFactory,B_BackFactory,B_Process,B_Packaging,B_Width2,B_Weight2,B_Width3,B_Weight3,B_seam,B_Beatbox,B_size,B_username,B_ContractLogo,B_SourceOrderCode from G_DraftBillDetailOrder  where B_ID='" & id & "'"
    Gm.cnnTool.cnn.Execute sql2
    Debug.Print sql2
    Dim sql3 As String
    sql3 = "delete from G_DraftBillDetailOrder where B_ID='" & id & "'"
    Gm.cnnTool.cnn.Execute sql3
End Sub
Private Sub FlatEdit1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 13
            Case 13
        If Len(FlatEdit2.Text) <= 0 Then
            PushButton2_Click
        Else
            FlatEdit3.SetFocus
        End If
    End Select
End Sub

Private Sub FlatEdit3_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            FlatEdit4.SetFocus
    End Select
End Sub
Private Sub FlatEdit4_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
         FlatEdit15.SetFocus
    End Select
End Sub
Private Sub FlatEdit5_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
             FlatEdit6.SetFocus
    End Select
End Sub
Private Sub FlatEdit6_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
             FlatEdit14.SetFocus
    End Select
End Sub
Private Sub FlatEdit14_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
             FlatEdit7.SetFocus
    End Select
End Sub
Private Sub FlatEdit7_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
             FlatEdit8.SetFocus
    End Select
End Sub
Private Sub FlatEdit8_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
             FlatEdit12.SetFocus
    End Select
End Sub

Private Sub FlatEdit13_KeyUp(KeyCode As Integer, Shift As Integer)
'     On Error Resume Next
'    Select Case KeyCode
'        Case 13
''        If MsgBox("是否保存数据", vbYesNoCancel + vbDefaultButton2 + vbInformation, "提示") = vbYes Then
''            save
''        if
''            Unload Me
''        End If
'        Dim szReturn As VbMsgBoxResult
'        szReturn = MsgBox("是否要保存？", vbYesNoCancel + vbDefaultButton1, "提示")
'        Select Case szReturn
'
'           Case vbYes
'               save
'           Case vbNo
'               Unload Me
'           Case vbCancel
'
'        End Select
'
'    End Select
End Sub

Private Sub PushButton10_Click()
    Dim frm1 As New frmPopupClient_Edit
     frm1.a = "家纺厂"
     frm1.Show vbModal
'    Positiveid = frm1.clientid
    FlatEdit33.Text = frm1.ClientName
    Unload frm1
End Sub
'中间面料颜色
Private Sub PushButton11_Click()
On Error Resume Next
    Dim sql As String
    Dim rs As New RecordSet
    Dim frm1 As New frmpopupColor
    frm1.Show vbModal
    FlatEdit38.Text = Trim(frm1.colorname)
    colorid2 = frm1.colorid
    If frm1.bsaved = True Then
        FlatEdit38.Enabled = True
    Else
        If colorid = "" Then
             FlatEdit38.Enabled = False
        End If
    End If
    
        sql = "select * from G_Color where B_SID='" & colorid2 & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

        If rs!B_hex <> "" Then
        
            Picture4.BackColor = rs!B_hex
        Else
            Picture4.BackColor = &H8000000F
        End If
      
    
    Unload frm1
End Sub
'背面面料颜色
Private Sub PushButton12_Click()
On Error Resume Next
    Dim sql As String
    Dim rs As New RecordSet
    Dim frm1 As New frmpopupColor
    frm1.Show vbModal
    FlatEdit42.Text = Trim(frm1.colorname)
    colorid3 = frm1.colorid
    If frm1.bsaved = True Then
        FlatEdit42.Enabled = True
    Else
        If colorid = "" Then
             FlatEdit42.Enabled = False
        End If
    End If
    
        sql = "select * from G_Color where B_SID='" & colorid3 & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

        If rs!B_hex <> "" Then
        
            Picture6.BackColor = rs!B_hex
        Else
            Picture6.BackColor = &H8000000F
        End If
      
    Unload frm1
End Sub

Private Sub PushButton2_Click()
    Dim frm1 As New frmpopupProduct
    frm1.Show vbModal
    FlatEdit2.Text = Trim(frm1.WhiteName)
    Productid = frm1.Whiteid
    Unload frm1
End Sub

Private Sub PushButton3_Click()
    Dim frm1 As New frmPopupClient_Edit
     frm1.a = "色布供应商"
     frm1.Show vbModal
    Positiveid = frm1.clientid
    FlatEdit17.Text = frm1.ClientName
    Unload frm1
End Sub

Private Sub PushButton4_Click()
    Dim frm1 As New frmPopupClient_Edit
     frm1.a = "色布供应商"
     frm1.Show vbModal
    Middleid = frm1.clientid
    FlatEdit18.Text = frm1.ClientName
    Unload frm1
End Sub

Private Sub PushButton5_Click()
    Dim frm1 As New frmPopupClient_Edit
     frm1.a = "色布供应商"
     frm1.Show vbModal
    backid = frm1.clientid
    FlatEdit19.Text = frm1.ClientName
    Unload frm1
End Sub

Private Sub PushButton6_Click()
    If client = "" Then
        MsgBox "先将主表选择一个客户", vbInformation, "提示"
        Exit Sub
        
    End If
    Dim frm1 As New frmPopupitemidb
    frm1.clientid = client
    frm1.Show vbModal
    FlatEdit28.Text = frm1.ordercode
    
    Unload frm1
End Sub
'预览
Private Sub PushButton7_Click()
On Error GoTo IFERR
    
    With CommonDialog1
        .ShowOpen
   
        szFile = .FileName
    End With
     
    If Len(szFile) <= 0 Then
        Exit Sub
    End If
    cls1.InitCls szFile, Picture5
    
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "图片格式不支持，请提供JPG或者BMP格式图片！" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub
'上传
Private Sub PushButton8_Click()
Dim sql As String
Dim rs As New RecordSet

    If szFile <> "" Then
'                 需要引用：Microsoft Scripting Runtime
        Dim fso As New FileSystemObject
    
        Dim lLength As Long
        Dim oFile As File
    
        Set oFile = fso.GetFile(szFile)
        
        '设置上传图片的大小
        sql = "select * from G_ImageSize"
        rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
        
        
        If oFile.Size / 1000000 > rs!B_size Then
            MsgBox "图片太大不能上传", vbInformation, "提示"
            Exit Sub
        End If
    
        '获取的长度的单位是：字节
        saveImage
       
            MsgBox "图片上传成功", vbInformation, "提示"
       
    End If
End Sub
'删除
Private Sub PushButton9_Click()
Dim sql As String
    If m_ID > 0 Then
        sql = "delete from WVAccountImage.dbo.G_image_NEW_FZ where B_OrderID='" & m_OrderID & "' and B_BDCItemID='" & m_ID & "'"
        Gm.cnnToolImage.cnn.Execute sql
    
    End If
    Picture5.Picture = Nothing
End Sub
Private Sub saveImage()

    Dim rs As New RecordSet
    Dim sql As String
    sql = "SELECT * FROM WVAccountImage.dbo.G_image_NEW_HT where 1=0"
    rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
    
     If szFile = "" Then
     
     Else
        Dim rs1 As New RecordSet
        Dim sql1 As String
        sql1 = "SELECT * FROM WVAccountImage.dbo.G_image_NEW_HT where B_OrderID='" & id & "' and B_BDOItemID='" & itemid & "'"
        rs1.Open sql1, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
        
        If rs1.RecordCount > 0 Then
            
            PicSaveToDB rs1!B_picture, szFile
            rs!B_KuanHao = FlatEdit29.Text
            rs1.Update
            rs1.Close
            Set rs1 = Nothing
        Else
            rs.AddNew
'            rs!B_id = theid
            PicSaveToDB rs!B_picture, szFile
            rs!B_BDOItemID = itemid  '缝制计划表 一个主键对应一张图片
            rs!B_OrderID = id '合同计划主键
            rs!B_KuanHao = FlatEdit29.Text
            rs.Update
            rs.Close
            Set rs = Nothing
        End If
        
    End If
End Sub
'上传图片到服务器
'fld：记录集中的字段
'vFilePath：图片文件的绝对路径，包含图片文件名和扩展名
Private Sub PicSaveToDB(ByRef fld As ADODB.Field, ByVal vFilePath As String)
    Const blocksize = 4096
    Dim bytedata() As Byte
    Dim numblocks As Long
    Dim filelength As Long
    Dim leftover As Long
    Dim sourcefile As Long
    Dim i As Long
    sourcefile = FreeFile
    
    Open Trim(vFilePath) For Binary Access Read As sourcefile
    filelength = LOF(sourcefile)
    
    If filelength = 0 Then
        Close sourcefile
        'MsgBox Trim(vFilePath) & "无内容或不存在！"
    Else
        numblocks = filelength \ blocksize
        leftover = filelength Mod blocksize
        fld.Value = Null
        
        ReDim bytedata(blocksize)
        
        For i = 1 To numblocks
            Get sourcefile, , bytedata()
            fld.AppendChunk bytedata()
        Next
        

        ReDim bytedata(leftover)
        Get sourcefile, , bytedata()
        fld.AppendChunk bytedata()
        Close sourcefile
    End If
End Sub
'从DB中下载图片并且显示到UI的图片控件上
'vRs：包含有图片资源的数据源
'vPicField：保存图片文件的字段名
'oCtl：用于显示的控件。PictureBox、Image
Private Sub PicShow2Ctl(ByRef vFld As ADODB.Field, ByRef oCtl As Object)
    'On Error GoTo IFERR
    Dim Stream As ADODB.Stream
    Set Stream = New ADODB.Stream
    

    oCtl.Picture = LoadPicture("")
    If Not IsNull(vFld) Then
        Stream.Type = adTypeBinary
        Stream.Open
        Stream.Write vFld
        Stream.SaveToFile "filename", adSaveCreateOverWrite
        'Stream.SaveToFile "c:\aaa.jpg", adSaveCreateOverWrite
        
        szFile = LoadPicture("filename")
'        Debug.Print FileName
        oCtl.Picture = LoadPicture("filename")
        Stream.Close
    End If
    
    Set Stream = Nothing
'    Exit Sub
'IFERR:
'    Dim szErr As String
'    szErr = "错误发生于下载图片中，" & Err.Description
'    MsgBox szErr
End Sub

Public Sub OpenImage()

                Dim rs1 As New RecordSet
                Dim sql As String
                Dim clsFile01 As New clsFile
                Dim szPic As String
                
                sql = "SELECT * FROM WVAccountImage.dbo.G_image_NEW_HT where B_OrderID='" & id & "' AND B_BDOItemID='" & itemid & "'"
                rs1.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
                
                If rs1.RecordCount > 0 Then
                    szPic = App.Path & "\temp\" & rs1!B_id & rs1!B_KuanHao & ".JPG"
                    Debug.Print szPic
                    
'                    clsFile01.DownloadPic rs1!B_picture, szPic
'                    cls1.InitCls szPic, frm1.Picture5
                    
                    PicShow2Ctl rs1!B_picture, Picture5
                    'PicShow2Ctl rs!B_picture, Picture3
                Else
                    Picture5.Picture = Nothing
                End If

End Sub
