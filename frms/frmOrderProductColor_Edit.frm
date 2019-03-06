VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmOrderProductColor_Edit 
   Caption         =   "色布计划"
   ClientHeight    =   9450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   9450
   ScaleWidth      =   19410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar22 
      Height          =   10950
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20250
      _LayoutVersion  =   1
      _ExtentX        =   35719
      _ExtentY        =   19315
      _DataPath       =   ""
      Bands           =   "frmOrderProductColor_Edit.frx":0000
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   10695
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   20175
         _cx             =   35586
         _cy             =   18865
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   800
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   3263743
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "色布采购|色布生产"
         Align           =   0
         CurrTab         =   1
         FirstTab        =   0
         Style           =   6
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   0   'False
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   3
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
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
         Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar23 
            Height          =   10695
            Left            =   -20760
            TabIndex        =   2
            Top             =   330
            Width           =   20175
            _LayoutVersion  =   1
            _ExtentX        =   35586
            _ExtentY        =   18865
            _DataPath       =   ""
            Bands           =   "frmOrderProductColor_Edit.frx":01C8
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   6435
               Left            =   480
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   960
               Width           =   10515
               _cx             =   18547
               _cy             =   11351
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
               _GridInfo       =   $"frmOrderProductColor_Edit.frx":1968
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.PictureBox Picture3 
                  BorderStyle     =   0  'None
                  Height          =   4665
                  Left            =   90
                  ScaleHeight     =   4665
                  ScaleWidth      =   10335
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   10335
                  Begin VB.PictureBox Picture4 
                     Height          =   375
                     Left            =   4680
                     ScaleHeight     =   315
                     ScaleWidth      =   1515
                     TabIndex        =   143
                     TabStop         =   0   'False
                     Top             =   2280
                     Width           =   1575
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit38 
                     Height          =   375
                     Left            =   8040
                     TabIndex        =   5
                     Top             =   2796
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
                  End
                  Begin XtremeSuiteControls.PushButton PushButton10 
                     Height          =   375
                     Left            =   6300
                     TabIndex        =   6
                     Top             =   180
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit39 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1440
                     TabIndex        =   7
                     Top             =   180
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
                  Begin XtremeSuiteControls.FlatEdit FlatEdit40 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1440
                     TabIndex        =   8
                     Top             =   1012
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
                     Left            =   8040
                     TabIndex        =   9
                     Top             =   1012
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
                  Begin XtremeSuiteControls.FlatEdit FlatEdit42 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   4740
                     TabIndex        =   10
                     Top             =   1012
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
                  Begin XtremeSuiteControls.FlatEdit FlatEdit43 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1440
                     TabIndex        =   11
                     Top             =   1904
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
                     Left            =   1440
                     TabIndex        =   12
                     Top             =   2796
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
                  Begin XtremeSuiteControls.FlatEdit FlatEdit45 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   4740
                     TabIndex        =   13
                     Top             =   180
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
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit46 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   4740
                     TabIndex        =   14
                     Top             =   2796
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
                  Begin XtremeSuiteControls.FlatEdit FlatEdit47 
                     Height          =   975
                     Left            =   4740
                     TabIndex        =   15
                     Top             =   3600
                     Width           =   5235
                     _Version        =   1048578
                     _ExtentX        =   9234
                     _ExtentY        =   1720
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     MultiLine       =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit48 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   4740
                     TabIndex        =   16
                     Top             =   1905
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
                  Begin XtremeSuiteControls.FlatEdit FlatEdit49 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   8040
                     TabIndex        =   17
                     Top             =   1904
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
                  Begin XtremeSuiteControls.DateTimePicker DateTimePicker1 
                     Height          =   375
                     Left            =   8040
                     TabIndex        =   18
                     Top             =   180
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   68
                     Format          =   1
                  End
                  Begin XtremeSuiteControls.PushButton PushButton11 
                     Height          =   375
                     Left            =   6300
                     TabIndex        =   19
                     Top             =   1904
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit50 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1440
                     TabIndex        =   144
                     Top             =   3720
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
                  Begin XtremeSuiteControls.Label Label63 
                     Height          =   255
                     Left            =   240
                     TabIndex        =   145
                     Top             =   3780
                     Width           =   615
                     _Version        =   1048578
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "单价："
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
                  Begin XtremeSuiteControls.Label Label50 
                     Height          =   255
                     Left            =   7080
                     TabIndex        =   32
                     Top             =   240
                     Width           =   855
                     _Version        =   1048578
                     _ExtentX        =   1508
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "交期："
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
                  Begin XtremeSuiteControls.Label Label51 
                     Height          =   255
                     Left            =   7080
                     TabIndex        =   31
                     Top             =   1964
                     Width           =   615
                     _Version        =   1048578
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "匹数："
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
                  Begin XtremeSuiteControls.Label Label52 
                     Height          =   255
                     Left            =   3480
                     TabIndex        =   30
                     Top             =   1965
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "颜     色："
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
                  Begin XtremeSuiteControls.Label Label53 
                     Height          =   255
                     Left            =   240
                     TabIndex        =   29
                     Top             =   240
                     Width           =   855
                     _Version        =   1048578
                     _ExtentX        =   1508
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
                  Begin XtremeSuiteControls.Label Label54 
                     Height          =   255
                     Left            =   240
                     TabIndex        =   28
                     Top             =   1072
                     Width           =   615
                     _Version        =   1048578
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "品名："
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
                  Begin XtremeSuiteControls.Label Label55 
                     Height          =   255
                     Left            =   7080
                     TabIndex        =   27
                     Top             =   1072
                     Width           =   615
                     _Version        =   1048578
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "门幅："
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
                  Begin XtremeSuiteControls.Label Label56 
                     Height          =   255
                     Left            =   3480
                     TabIndex        =   26
                     Top             =   1065
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "克     重："
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
                  Begin XtremeSuiteControls.Label Label57 
                     Height          =   255
                     Left            =   3480
                     TabIndex        =   25
                     Top             =   240
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "色布供应商："
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
                  Begin XtremeSuiteControls.Label Label58 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   24
                     Top             =   1904
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "色号："
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
                  Begin XtremeSuiteControls.Label Label59 
                     Height          =   255
                     Left            =   3480
                     TabIndex        =   23
                     Top             =   2850
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "米     数："
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
                  Begin XtremeSuiteControls.Label Label60 
                     Height          =   255
                     Left            =   240
                     TabIndex        =   22
                     Top             =   2856
                     Width           =   855
                     _Version        =   1048578
                     _ExtentX        =   1508
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "公斤数："
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
                  Begin XtremeSuiteControls.Label Label61 
                     Height          =   255
                     Left            =   7080
                     TabIndex        =   21
                     Top             =   2856
                     Width           =   555
                     _Version        =   1048578
                     _ExtentX        =   979
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "码数:"
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
                  Begin XtremeSuiteControls.Label Label62 
                     Height          =   315
                     Left            =   3480
                     TabIndex        =   20
                     Top             =   3690
                     Width           =   1335
                     _Version        =   1048578
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _StockProps     =   79
                     Caption         =   "备       注:"
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
               Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
                  Height          =   1530
                  Left            =   90
                  TabIndex        =   33
                  Top             =   4815
                  Width           =   10335
                  _ExtentX        =   18230
                  _ExtentY        =   2699
                  _LayoutType     =   0
                  _RowHeight      =   25
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).DataField=   ""
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).DataField=   ""
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   2
                  Splits(0)._UserFlags=   0
                  Splits(0).RecordSelectorWidth=   953
                  Splits(0)._SavedRecordSelectors=   0   'False
                  Splits(0).AllowColSelect=   0   'False
                  Splits(0).DividerColor=   15790320
                  Splits(0).SpringMode=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=2"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=4339"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
                  Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
                  Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(6)=   "Column(1).Width=4339"
                  Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
                  Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=17"
                  Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
                  Splits.Count    =   1
                  PrintInfos(0)._StateFlags=   3
                  PrintInfos(0).Name=   "piInternal 0"
                  PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
                  PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
                  PrintInfos(0).PageHeaderHeight=   0
                  PrintInfos(0).PageFooterHeight=   0
                  PrintInfos.Count=   1
                  AllowUpdate     =   0   'False
                  ColumnFooters   =   -1  'True
                  DefColWidth     =   0
                  HeadLines       =   1.5
                  FootLines       =   1.5
                  MultipleLines   =   0
                  CellTipsWidth   =   0
                  DeadAreaBackColor=   15790320
                  RowDividerColor =   15790320
                  RowSubDividerColor=   15790320
                  DirectionAfterEnter=   1
                  MaxRows         =   250000
                  ViewColumnCaptionWidth=   0
                  ViewColumnWidth =   0
                  _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
                  _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(8)   =   ":id=1,.fontname=宋体"
                  _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgpicMode=2,.bgbmp=1,.bold=0"
                  _StyleDefs(11)  =   ":id=2,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(12)  =   ":id=2,.fontname=宋体"
                  _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgpicMode=2,.bgbmp=2,.bold=0"
                  _StyleDefs(14)  =   ":id=3,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(15)  =   ":id=3,.fontname=宋体"
                  _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
                  _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
                  _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
                  _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
                  _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
                  _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
                  _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
                  _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.alignment=2,.valignment=2"
                  _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                  _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
                  _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                  _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                  _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
                  _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
                  _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
                  _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
                  _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
                  _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
                  _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
                  _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
                  _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
                  _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
                  _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
                  _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
                  _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
                  _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
                  _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
                  _StyleDefs(44)  =   "Named:id=33:Normal"
                  _StyleDefs(45)  =   ":id=33,.parent=0"
                  _StyleDefs(46)  =   "Named:id=34:Heading"
                  _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(48)  =   ":id=34,.wraptext=-1"
                  _StyleDefs(49)  =   "Named:id=35:Footing"
                  _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(51)  =   "Named:id=36:Selected"
                  _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(53)  =   "Named:id=37:Caption"
                  _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
                  _StyleDefs(55)  =   "Named:id=38:HighlightRow"
                  _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(57)  =   "Named:id=39:EvenRow"
                  _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                  _StyleDefs(59)  =   "Named:id=40:OddRow"
                  _StyleDefs(60)  =   ":id=40,.parent=33"
                  _StyleDefs(61)  =   "Named:id=41:RecordSelector"
                  _StyleDefs(62)  =   ":id=41,.parent=34"
                  _StyleDefs(63)  =   "Named:id=42:FilterBar"
                  _StyleDefs(64)  =   ":id=42,.parent=33"
                  _StyleDefs(65)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIyanIyanIyanIyanIyanIya"
                  _StyleDefs(66)  =   "bmp(1):id=1,nIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIya"
                  _StyleDefs(67)  =   "bmp(2):id=1,nIyanIyanAAAAJSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSm"
                  _StyleDefs(68)  =   "bmp(3):id=1,pZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpQAAAJyurZyurZyurZyurZyurZyurZyu"
                  _StyleDefs(69)  =   "bmp(4):id=1,rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyu"
                  _StyleDefs(70)  =   "bmp(5):id=1,rZyurQAAAKW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2"
                  _StyleDefs(71)  =   "bmp(6):id=1,taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2tQAAAK2+va2+va2+va2+va2+va2+va2+va2+"
                  _StyleDefs(72)  =   "bmp(7):id=1,va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
                  _StyleDefs(73)  =   "bmp(8):id=1,vQAAAK2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
                  _StyleDefs(74)  =   "bmp(9):id=1,va2+va2+va2+va2+va2+va2+va2+va2+va2+vQAAALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
                  _StyleDefs(75)  =   "bmp(10):id=1,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAA"
                  _StyleDefs(76)  =   "bmp(11):id=1,ALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
                  _StyleDefs(77)  =   "bmp(12):id=1,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
                  _StyleDefs(78)  =   "bmp(13):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3P"
                  _StyleDefs(79)  =   "bmp(14):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
                  _StyleDefs(80)  =   "bmp(15):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
                  _StyleDefs(81)  =   "bmp(16):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAM7X1s7X"
                  _StyleDefs(82)  =   "bmp(17):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
                  _StyleDefs(83)  =   "bmp(18):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1gAAAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
                  _StyleDefs(84)  =   "bmp(19):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1gAAANbj59bj59bj"
                  _StyleDefs(85)  =   "bmp(20):id=1,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
                  _StyleDefs(86)  =   "bmp(21):id=1,59bj59bj59bj59bj59bj5wAAANbj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
                  _StyleDefs(87)  =   "bmp(22):id=1,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj5wAAAN7r797r797r797r"
                  _StyleDefs(88)  =   "bmp(23):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
                  _StyleDefs(89)  =   "bmp(24):id=1,797r797r797r797r7wAAAN7r797r797r797r797r797r797r797r797r797r797r797r797r797r"
                  _StyleDefs(90)  =   "bmp(25):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r7wAAAN7r797r797r797r797r"
                  _StyleDefs(91)  =   "bmp(26):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
                  _StyleDefs(92)  =   "bmp(27):id=1,797r797r797r7wAAAA=="
                  _StyleDefs(93)  =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIyanIyanIyanIyanIyanIya"
                  _StyleDefs(94)  =   "bmp(1):id=2,nIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIya"
                  _StyleDefs(95)  =   "bmp(2):id=2,nIyanIyanAAAAJSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSm"
                  _StyleDefs(96)  =   "bmp(3):id=2,pZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpQAAAJyurZyurZyurZyurZyurZyurZyu"
                  _StyleDefs(97)  =   "bmp(4):id=2,rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyu"
                  _StyleDefs(98)  =   "bmp(5):id=2,rZyurQAAAKW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2"
                  _StyleDefs(99)  =   "bmp(6):id=2,taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2tQAAAK2+va2+va2+va2+va2+va2+va2+va2+"
                  _StyleDefs(100) =   "bmp(7):id=2,va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
                  _StyleDefs(101) =   "bmp(8):id=2,vQAAAK2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
                  _StyleDefs(102) =   "bmp(9):id=2,va2+va2+va2+va2+va2+va2+va2+va2+va2+vQAAALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
                  _StyleDefs(103) =   "bmp(10):id=2,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAA"
                  _StyleDefs(104) =   "bmp(11):id=2,ALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
                  _StyleDefs(105) =   "bmp(12):id=2,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
                  _StyleDefs(106) =   "bmp(13):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3P"
                  _StyleDefs(107) =   "bmp(14):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
                  _StyleDefs(108) =   "bmp(15):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
                  _StyleDefs(109) =   "bmp(16):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAM7X1s7X"
                  _StyleDefs(110) =   "bmp(17):id=2,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
                  _StyleDefs(111) =   "bmp(18):id=2,1s7X1s7X1s7X1s7X1s7X1s7X1gAAAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
                  _StyleDefs(112) =   "bmp(19):id=2,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1gAAANbj59bj59bj"
                  _StyleDefs(113) =   "bmp(20):id=2,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
                  _StyleDefs(114) =   "bmp(21):id=2,59bj59bj59bj59bj59bj5wAAANbj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
                  _StyleDefs(115) =   "bmp(22):id=2,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj5wAAAN7r797r797r797r"
                  _StyleDefs(116) =   "bmp(23):id=2,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
                  _StyleDefs(117) =   "bmp(24):id=2,797r797r797r797r7wAAAN7r797r797r797r797r797r797r797r797r797r797r797r797r797r"
                  _StyleDefs(118) =   "bmp(25):id=2,797r797r797r797r797r797r797r797r797r797r797r797r797r7wAAAN7r797r797r797r797r"
                  _StyleDefs(119) =   "bmp(26):id=2,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
                  _StyleDefs(120) =   "bmp(27):id=2,797r797r797r7wAAAA=="
               End
            End
         End
         Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
            Height          =   10695
            Left            =   15
            TabIndex        =   34
            Top             =   330
            Width           =   20175
            _LayoutVersion  =   1
            _ExtentX        =   35586
            _ExtentY        =   18865
            _DataPath       =   ""
            Bands           =   "frmOrderProductColor_Edit.frx":19EE
            Begin C1SizerLibCtl.C1Elastic C1Elastic1 
               Height          =   9075
               Left            =   0
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   480
               Width           =   19935
               _cx             =   35163
               _cy             =   16007
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
               _GridInfo       =   $"frmOrderProductColor_Edit.frx":4774
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.PictureBox Picture1 
                  BorderStyle     =   0  'None
                  Height          =   8895
                  Left            =   90
                  ScaleHeight     =   8895
                  ScaleWidth      =   19755
                  TabIndex        =   36
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   19755
                  Begin XtremeSuiteControls.FlatEdit FlatEdit52 
                     Height          =   495
                     Left            =   16320
                     TabIndex        =   155
                     Top             =   840
                     Width           =   2535
                     _Version        =   1048578
                     _ExtentX        =   4471
                     _ExtentY        =   873
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit51 
                     Height          =   495
                     Left            =   16320
                     TabIndex        =   154
                     Top             =   120
                     Width           =   2535
                     _Version        =   1048578
                     _ExtentX        =   4471
                     _ExtentY        =   873
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic3 
                     Height          =   3255
                     Left            =   15240
                     TabIndex        =   149
                     TabStop         =   0   'False
                     Top             =   5040
                     Width           =   3855
                     _cx             =   6800
                     _cy             =   5741
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
                     _GridInfo       =   $"frmOrderProductColor_Edit.frx":47FB
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   9
                     Begin VB.PictureBox Picture5 
                        Height          =   3075
                        Left            =   90
                        ScaleHeight     =   3015
                        ScaleWidth      =   3615
                        TabIndex        =   150
                        TabStop         =   0   'False
                        Top             =   90
                        Width           =   3675
                     End
                  End
                  Begin MSComDlg.CommonDialog CommonDialog1 
                     Left            =   18480
                     Top             =   4440
                     _ExtentX        =   847
                     _ExtentY        =   847
                     _Version        =   393216
                  End
                  Begin VB.Frame Frame1 
                     Height          =   75
                     Left            =   180
                     TabIndex        =   39
                     Top             =   2760
                     Width           =   14955
                  End
                  Begin VB.Frame Frame2 
                     Height          =   75
                     Left            =   60
                     TabIndex        =   38
                     Top             =   7500
                     Width           =   14895
                  End
                  Begin VB.PictureBox Picture2 
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     ForeColor       =   &H80000008&
                     Height          =   375
                     Left            =   5400
                     ScaleHeight     =   345
                     ScaleWidth      =   1905
                     TabIndex        =   37
                     TabStop         =   0   'False
                     Top             =   660
                     Width           =   1935
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit33 
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   40
                     Top             =   2280
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit22 
                     Height          =   375
                     Left            =   7620
                     TabIndex        =   41
                     Top             =   4095
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
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
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit18 
                     Height          =   375
                     Left            =   7620
                     TabIndex        =   42
                     Top             =   2985
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
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
                  End
                  Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
                     Height          =   555
                     Left            =   9960
                     TabIndex        =   43
                     Top             =   8520
                     Width           =   2055
                     _ExtentX        =   3625
                     _ExtentY        =   979
                     _LayoutType     =   0
                     _RowHeight      =   -2147483647
                     _WasPersistedAsPixels=   0
                     Columns(0)._VlistStyle=   0
                     Columns(0)._MaxComboItems=   5
                     Columns(0).DataField=   ""
                     Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                     Columns.Count   =   1
                     Splits(0)._UserFlags=   0
                     Splits(0).RecordSelectors=   0   'False
                     Splits(0).RecordSelectorWidth=   953
                     Splits(0)._SavedRecordSelectors=   0   'False
                     Splits(0).ScrollBars=   2
                     Splits(0).AllowColSelect=   0   'False
                     Splits(0).DividerColor=   15790320
                     Splits(0).SpringMode=   0   'False
                     Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                     Splits(0)._ColumnProps(0)=   "Columns.Count=1"
                     Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
                     Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                     Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                     Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
                     Splits.Count    =   1
                     PrintInfos(0)._StateFlags=   3
                     PrintInfos(0).Name=   "piInternal 0"
                     PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
                     PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
                     PrintInfos(0).PageHeaderHeight=   0
                     PrintInfos(0).PageFooterHeight=   0
                     PrintInfos.Count=   1
                     AllowUpdate     =   0   'False
                     DefColWidth     =   0
                     HeadLines       =   1
                     FootLines       =   1
                     MultipleLines   =   0
                     CellTipsWidth   =   0
                     DeadAreaBackColor=   15790320
                     RowDividerColor =   15790320
                     RowSubDividerColor=   15790320
                     DirectionAfterEnter=   1
                     MaxRows         =   250000
                     ViewColumnCaptionWidth=   0
                     ViewColumnWidth =   0
                     _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
                     _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                     _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                     _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                     _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=114,.bold=0,.fontsize=825,.italic=0"
                     _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
                     _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
                     _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
                     _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                     _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
                     _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
                     _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
                     _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
                     _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                     _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
                     _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
                     _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
                     _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
                     _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
                     _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
                     _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
                     _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
                     _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                     _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
                     _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                     _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                     _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
                     _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
                     _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
                     _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
                     _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
                     _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
                     _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
                     _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
                     _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
                     _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
                     _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
                     _StyleDefs(40)  =   "Named:id=33:Normal"
                     _StyleDefs(41)  =   ":id=33,.parent=0"
                     _StyleDefs(42)  =   "Named:id=34:Heading"
                     _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                     _StyleDefs(44)  =   ":id=34,.wraptext=-1"
                     _StyleDefs(45)  =   "Named:id=35:Footing"
                     _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                     _StyleDefs(47)  =   "Named:id=36:Selected"
                     _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                     _StyleDefs(49)  =   "Named:id=37:Caption"
                     _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
                     _StyleDefs(51)  =   "Named:id=38:HighlightRow"
                     _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                     _StyleDefs(53)  =   "Named:id=39:EvenRow"
                     _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                     _StyleDefs(55)  =   "Named:id=40:OddRow"
                     _StyleDefs(56)  =   ":id=40,.parent=33"
                     _StyleDefs(57)  =   "Named:id=41:RecordSelector"
                     _StyleDefs(58)  =   ":id=41,.parent=34"
                     _StyleDefs(59)  =   "Named:id=42:FilterBar"
                     _StyleDefs(60)  =   ":id=42,.parent=33"
                  End
                  Begin MSComCtl2.DTPicker DTPicker1 
                     Height          =   375
                     Left            =   10800
                     TabIndex        =   44
                     Top             =   2985
                     Width           =   1575
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "宋体"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   225574913
                     CurrentDate     =   43041
                  End
                  Begin XtremeSuiteControls.PushButton PushButton1 
                     Height          =   375
                     Left            =   3480
                     TabIndex        =   45
                     Top             =   660
                     Visible         =   0   'False
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit2 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   46
                     Top             =   120
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
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
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit5 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   12720
                     TabIndex        =   47
                     Top             =   660
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
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
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit6 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   48
                     Top             =   1200
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
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
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit8 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   49
                     Top             =   2985
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
                  Begin XtremeSuiteControls.FlatEdit FlatEdit9 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   50
                     Top             =   2985
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
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
                  Begin XtremeSuiteControls.FlatEdit FlatEdit11 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   51
                     Top             =   660
                     Width           =   1815
                     _Version        =   1048578
                     _ExtentX        =   3201
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
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
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit15 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   52
                     Top             =   4095
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
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
                  Begin XtremeSuiteControls.FlatEdit FlatEdit19 
                     Bindings        =   "frmOrderProductColor_Edit.frx":4875
                     DataField       =   "B_CodeID"
                     BeginProperty DataFormat 
                        Type            =   1
                        Format          =   "0"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   2052
                        SubFormatType   =   1
                     EndProperty
                     DataMember      =   "0.00"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   53
                     Top             =   7620
                     Width           =   735
                     _Version        =   1048578
                     _ExtentX        =   1296
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
                  Begin XtremeSuiteControls.FlatEdit FlatEdit20 
                     DataField       =   "B_CodeID"
                     BeginProperty DataFormat 
                        Type            =   0
                        Format          =   "0"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   2052
                        SubFormatType   =   0
                     EndProperty
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   54
                     Top             =   7620
                     Width           =   735
                     _Version        =   1048578
                     _ExtentX        =   1296
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
                  Begin XtremeSuiteControls.FlatEdit FlatEdit21 
                     DataField       =   "B_CodeID"
                     BeginProperty DataFormat 
                        Type            =   0
                        Format          =   "0"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   2052
                        SubFormatType   =   0
                     EndProperty
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   9000
                     TabIndex        =   55
                     Top             =   7620
                     Width           =   735
                     _Version        =   1048578
                     _ExtentX        =   1296
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
                     Left            =   3480
                     TabIndex        =   56
                     Top             =   2985
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.PushButton PushButton3 
                     Height          =   375
                     Left            =   6360
                     TabIndex        =   57
                     Top             =   2985
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.PushButton PushButton5 
                     Height          =   375
                     Left            =   6360
                     TabIndex        =   58
                     Top             =   4095
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin MSComCtl2.DTPicker DTPicker2 
                     Height          =   375
                     Left            =   10800
                     TabIndex        =   59
                     Top             =   4095
                     Width           =   1575
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "宋体"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   225574913
                     CurrentDate     =   43041
                  End
                  Begin XtremeSuiteControls.ComboBox ComboBox4 
                     Height          =   345
                     Left            =   12780
                     TabIndex        =   60
                     Top             =   7635
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
                     Style           =   2
                     Text            =   "ComboBox1"
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit1 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   9000
                     TabIndex        =   61
                     Top             =   660
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
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
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit3 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   62
                     Top             =   120
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
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
                     Enabled         =   0   'False
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit4 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   9000
                     TabIndex        =   63
                     Top             =   180
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
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
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit10 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   12720
                     TabIndex        =   64
                     Top             =   120
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
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
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit13 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   435
                     Left            =   1920
                     TabIndex        =   65
                     Top             =   3480
                     Width           =   13275
                     _Version        =   1048578
                     _ExtentX        =   23416
                     _ExtentY        =   767
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
                     MultiLine       =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit17 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   435
                     Left            =   1920
                     TabIndex        =   66
                     Top             =   4650
                     Width           =   13275
                     _Version        =   1048578
                     _ExtentX        =   23416
                     _ExtentY        =   767
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
                     MultiLine       =   -1  'True
                  End
                  Begin XtremeSuiteControls.ComboBox ComboBox5 
                     Height          =   345
                     Left            =   9000
                     TabIndex        =   67
                     Top             =   1215
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
                     Style           =   2
                     Text            =   "ComboBox1"
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit14 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   68
                     Top             =   4095
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
                     Left            =   3480
                     TabIndex        =   69
                     Top             =   4095
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit12 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   14100
                     TabIndex        =   70
                     Top             =   2985
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   661
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
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit16 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   14100
                     TabIndex        =   71
                     Top             =   4095
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   661
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
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit23 
                     Height          =   375
                     Left            =   7620
                     TabIndex        =   72
                     Top             =   5265
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
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
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit24 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   73
                     Top             =   5265
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
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
                  Begin XtremeSuiteControls.PushButton PushButton6 
                     Height          =   375
                     Left            =   6360
                     TabIndex        =   74
                     Top             =   5265
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit25 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   75
                     Top             =   5265
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
                  Begin XtremeSuiteControls.PushButton PushButton7 
                     Height          =   375
                     Left            =   3480
                     TabIndex        =   76
                     Top             =   5265
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit26 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   435
                     Left            =   1920
                     TabIndex        =   77
                     Top             =   5820
                     Width           =   13275
                     _Version        =   1048578
                     _ExtentX        =   23416
                     _ExtentY        =   767
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
                     MultiLine       =   -1  'True
                  End
                  Begin MSComCtl2.DTPicker DTPicker3 
                     Height          =   375
                     Left            =   10800
                     TabIndex        =   78
                     Top             =   5265
                     Width           =   1575
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "宋体"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   225574913
                     CurrentDate     =   43041
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit27 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   14100
                     TabIndex        =   79
                     Top             =   5265
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   661
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
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit28 
                     Height          =   375
                     Left            =   7620
                     TabIndex        =   80
                     Top             =   6435
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
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
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit29 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   5460
                     TabIndex        =   81
                     Top             =   6435
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
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
                  Begin XtremeSuiteControls.PushButton PushButton8 
                     Height          =   375
                     Left            =   6420
                     TabIndex        =   82
                     Top             =   6420
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit30 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   83
                     Top             =   6435
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
                  Begin XtremeSuiteControls.PushButton PushButton9 
                     Height          =   375
                     Left            =   3480
                     TabIndex        =   84
                     Top             =   6435
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin MSComCtl2.DTPicker DTPicker4 
                     Height          =   375
                     Left            =   10860
                     TabIndex        =   85
                     Top             =   6435
                     Width           =   1575
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "宋体"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   225574913
                     CurrentDate     =   43041
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit31 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   14100
                     TabIndex        =   86
                     Top             =   6435
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   661
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
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit32 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   435
                     Left            =   1920
                     TabIndex        =   87
                     Top             =   6960
                     Width           =   13275
                     _Version        =   1048578
                     _ExtentX        =   23416
                     _ExtentY        =   767
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
                     MultiLine       =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit34 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   9000
                     TabIndex        =   88
                     Top             =   1740
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
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
                  Begin XtremeSuiteControls.FlatEdit FlatEdit35 
                     Height          =   375
                     Left            =   12720
                     TabIndex        =   89
                     Top             =   1740
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
                  End
                  Begin XtremeSuiteControls.ComboBox ComboBox1 
                     Height          =   390
                     Left            =   1920
                     TabIndex        =   90
                     Top             =   8190
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   688
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Style           =   2
                     Text            =   "ComboBox1"
                  End
                  Begin XtremeSuiteControls.ComboBox ComboBox2 
                     Height          =   390
                     Left            =   5400
                     TabIndex        =   91
                     Top             =   8190
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
                     _ExtentY        =   688
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Style           =   2
                     Text            =   "ComboBox2"
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit36 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   12720
                     TabIndex        =   92
                     Top             =   1200
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
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
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit37 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   93
                     Top             =   1740
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
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
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.PushButton PushButton12 
                     Height          =   435
                     Left            =   15420
                     TabIndex        =   146
                     Top             =   3840
                     Width           =   1755
                     _Version        =   1048578
                     _ExtentX        =   3096
                     _ExtentY        =   767
                     _StockProps     =   79
                     Caption         =   "预览图片"
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.PushButton PushButton13 
                     Height          =   435
                     Left            =   15420
                     TabIndex        =   147
                     Top             =   4440
                     Width           =   1755
                     _Version        =   1048578
                     _ExtentX        =   3096
                     _ExtentY        =   767
                     _StockProps     =   79
                     Caption         =   "上传图片"
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.PushButton PushButton14 
                     Height          =   435
                     Left            =   17280
                     TabIndex        =   148
                     Top             =   3840
                     Width           =   1755
                     _Version        =   1048578
                     _ExtentX        =   3096
                     _ExtentY        =   767
                     _StockProps     =   79
                     Caption         =   "删除图片"
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit53 
                     Height          =   855
                     Left            =   16320
                     TabIndex        =   156
                     Top             =   1560
                     Width           =   2535
                     _Version        =   1048578
                     _ExtentX        =   4471
                     _ExtentY        =   1508
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     MultiLine       =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit54 
                     Height          =   855
                     Left            =   16320
                     TabIndex        =   158
                     Top             =   2640
                     Width           =   2535
                     _Version        =   1048578
                     _ExtentX        =   4471
                     _ExtentY        =   1508
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     MultiLine       =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit7 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   162
                     Top             =   2280
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
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
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit55 
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   164
                     Top             =   1200
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit56 
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   165
                     Top             =   1680
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit58 
                     Height          =   375
                     Left            =   12720
                     TabIndex        =   166
                     Top             =   2280
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit57 
                     Height          =   375
                     Left            =   9000
                     TabIndex        =   167
                     Top             =   2280
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                  End
                  Begin XtremeSuiteControls.Label Label70 
                     Height          =   375
                     Left            =   7680
                     TabIndex        =   168
                     Top             =   2280
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "色布公斤："
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
                     Left            =   240
                     TabIndex        =   163
                     Top             =   2340
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "公  斤 数："
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
                  Begin XtremeSuiteControls.Label Label71 
                     Height          =   255
                     Left            =   11400
                     TabIndex        =   161
                     Top             =   2400
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "染厂折率："
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
                  Begin XtremeSuiteControls.Label Label69 
                     Height          =   255
                     Left            =   4080
                     TabIndex        =   160
                     Top             =   1800
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "条     重："
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
                  Begin XtremeSuiteControls.Label Label68 
                     Height          =   375
                     Left            =   4080
                     TabIndex        =   159
                     Top             =   1200
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "条     数："
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
                  Begin XtremeSuiteControls.Label Label67 
                     Height          =   255
                     Left            =   15240
                     TabIndex        =   157
                     Top             =   2880
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "送布地址："
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
                  Begin XtremeSuiteControls.Label Label66 
                     Height          =   255
                     Left            =   15000
                     TabIndex        =   153
                     Top             =   1680
                     Width           =   1335
                     _Version        =   1048578
                     _ExtentX        =   2355
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "白坯布地址:"
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
                  Begin XtremeSuiteControls.Label Label65 
                     Height          =   255
                     Left            =   15120
                     TabIndex        =   152
                     Top             =   240
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "坯布规格:"
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
                  Begin XtremeSuiteControls.Label Label64 
                     Height          =   375
                     Left            =   15000
                     TabIndex        =   151
                     Top             =   960
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "坯布厂电话："
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
                     Left            =   4080
                     TabIndex        =   142
                     Top             =   3045
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "染厂 跟单："
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
                  Begin XtremeSuiteControls.Label Label10 
                     Height          =   255
                     Left            =   240
                     TabIndex        =   141
                     Top             =   3045
                     Width           =   1335
                     _Version        =   1048578
                     _ExtentX        =   2355
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "染     厂："
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
                     Left            =   240
                     TabIndex        =   140
                     Top             =   1260
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "米     数："
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
                     Left            =   11280
                     TabIndex        =   139
                     Top             =   660
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "色     号："
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
                     Left            =   240
                     TabIndex        =   138
                     Top             =   720
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "颜     色："
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
                  Begin XtremeSuiteControls.Label Label3 
                     Height          =   255
                     Left            =   7680
                     TabIndex        =   137
                     Top             =   180
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "门   幅："
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
                     Left            =   4080
                     TabIndex        =   136
                     Top             =   180
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "品     名："
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
                     Left            =   240
                     TabIndex        =   135
                     Top             =   180
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "订  单 号："
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
                     Height          =   375
                     Left            =   9480
                     TabIndex        =   134
                     Top             =   2985
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "染厂交期："
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
                     Height          =   375
                     Left            =   240
                     TabIndex        =   133
                     Top             =   3510
                     Width           =   1695
                     _Version        =   1048578
                     _ExtentX        =   2990
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "染厂  备注："
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
                     Height          =   375
                     Left            =   4080
                     TabIndex        =   132
                     Top             =   4095
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "深加工跟单："
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
                     Height          =   375
                     Left            =   9480
                     TabIndex        =   131
                     Top             =   4095
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "深加工交期："
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
                  Begin XtremeSuiteControls.Label Label17 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   130
                     Top             =   4680
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "深加工备注："
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
                  Begin XtremeSuiteControls.Label Label19 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   129
                     Top             =   7620
                     Width           =   1455
                     _Version        =   1048578
                     _ExtentX        =   2566
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "纸      管："
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
                  Begin XtremeSuiteControls.Label lbl 
                     Height          =   375
                     Left            =   4080
                     TabIndex        =   128
                     Top             =   7620
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "袋    重："
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
                     Height          =   255
                     Left            =   7680
                     TabIndex        =   127
                     Top             =   7680
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "空   加："
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
                  Begin XtremeSuiteControls.Label Label21 
                     Height          =   255
                     Left            =   11280
                     TabIndex        =   126
                     Top             =   7680
                     Width           =   1335
                     _Version        =   1048578
                     _ExtentX        =   2355
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "包装  方式："
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
                     Left            =   11280
                     TabIndex        =   125
                     Top             =   180
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "克     重："
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
                  Begin XtremeSuiteControls.Label Label22 
                     Height          =   255
                     Left            =   4080
                     TabIndex        =   124
                     Top             =   720
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "颜色 标识:"
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
                     Left            =   7680
                     TabIndex        =   123
                     Top             =   720
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "花   型："
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
                     Left            =   2760
                     TabIndex        =   122
                     Top             =   7680
                     Width           =   495
                     _Version        =   1048578
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "公斤"
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
                     Left            =   6240
                     TabIndex        =   121
                     Top             =   7680
                     Width           =   495
                     _Version        =   1048578
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "公斤"
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
                  Begin XtremeSuiteControls.Label Label25 
                     Height          =   255
                     Left            =   9840
                     TabIndex        =   120
                     Top             =   7680
                     Width           =   495
                     _Version        =   1048578
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "公斤"
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
                     Height          =   1575
                     Left            =   14880
                     TabIndex        =   119
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   135
                     _Version        =   1048578
                     _ExtentX        =   238
                     _ExtentY        =   2778
                     _StockProps     =   79
                     ForeColor       =   255
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "宋体"
                        Size            =   15
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                  End
                  Begin XtremeSuiteControls.Label Label18 
                     Height          =   375
                     Left            =   7680
                     TabIndex        =   118
                     Top             =   1200
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "进度工艺："
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
                  Begin XtremeSuiteControls.Label Label14 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   117
                     Top             =   4095
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "深加工单位："
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
                  Begin XtremeSuiteControls.Label Label27 
                     Height          =   255
                     Left            =   12660
                     TabIndex        =   116
                     Top             =   3045
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "染厂加工费:"
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
                     Height          =   195
                     Left            =   12660
                     TabIndex        =   115
                     Top             =   4185
                     Width           =   1455
                     _Version        =   1048578
                     _ExtentX        =   2566
                     _ExtentY        =   344
                     _StockProps     =   79
                     Caption         =   "深加工加工费:"
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
                     Left            =   7080
                     TabIndex        =   114
                     Top             =   3045
                     Width           =   435
                     _Version        =   1048578
                     _ExtentX        =   767
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "电话:"
                  End
                  Begin XtremeSuiteControls.Label Label30 
                     Height          =   375
                     Left            =   7080
                     TabIndex        =   113
                     Top             =   4095
                     Width           =   495
                     _Version        =   1048578
                     _ExtentX        =   873
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "电话:"
                  End
                  Begin XtremeSuiteControls.Label Label31 
                     Height          =   375
                     Left            =   7080
                     TabIndex        =   112
                     Top             =   5265
                     Width           =   495
                     _Version        =   1048578
                     _ExtentX        =   873
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "电话:"
                  End
                  Begin XtremeSuiteControls.Label Label32 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   111
                     Top             =   5265
                     Width           =   1335
                     _Version        =   1048578
                     _ExtentX        =   2355
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "深加工单位2："
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
                     Left            =   4080
                     TabIndex        =   110
                     Top             =   5265
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "深加工跟单："
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
                     Height          =   375
                     Left            =   240
                     TabIndex        =   109
                     Top             =   5850
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "深加工备注："
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
                     Height          =   195
                     Left            =   12660
                     TabIndex        =   108
                     Top             =   5355
                     Width           =   1455
                     _Version        =   1048578
                     _ExtentX        =   2566
                     _ExtentY        =   344
                     _StockProps     =   79
                     Caption         =   "深加工加工费:"
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
                     Height          =   375
                     Left            =   9480
                     TabIndex        =   107
                     Top             =   5265
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "深加工交期："
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
                     Height          =   375
                     Left            =   7080
                     TabIndex        =   106
                     Top             =   6435
                     Width           =   495
                     _Version        =   1048578
                     _ExtentX        =   873
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "电话:"
                  End
                  Begin XtremeSuiteControls.Label Label38 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   105
                     Top             =   6435
                     Width           =   1335
                     _Version        =   1048578
                     _ExtentX        =   2355
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "深加工单位3："
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
                     Height          =   375
                     Left            =   4140
                     TabIndex        =   104
                     Top             =   6435
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "深加工跟单："
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
                  Begin XtremeSuiteControls.Label Label40 
                     Height          =   375
                     Left            =   9540
                     TabIndex        =   103
                     Top             =   6435
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "深加工交期："
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
                  Begin XtremeSuiteControls.Label Label41 
                     Height          =   195
                     Left            =   12660
                     TabIndex        =   102
                     Top             =   6525
                     Width           =   1455
                     _Version        =   1048578
                     _ExtentX        =   2566
                     _ExtentY        =   344
                     _StockProps     =   79
                     Caption         =   "深加工加工费:"
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
                  Begin XtremeSuiteControls.Label Label42 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   101
                     Top             =   7020
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "深加工备注："
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
                     Left            =   4080
                     TabIndex        =   100
                     Top             =   2280
                     Width           =   1275
                     _Version        =   1048578
                     _ExtentX        =   2249
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "裁剪折率:"
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
                  Begin XtremeSuiteControls.Label Label44 
                     Height          =   255
                     Left            =   7680
                     TabIndex        =   99
                     Top             =   1800
                     Width           =   1275
                     _Version        =   1048578
                     _ExtentX        =   2249
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "预计投坯量:"
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
                  Begin XtremeSuiteControls.Label Label45 
                     Height          =   255
                     Left            =   11280
                     TabIndex        =   98
                     Top             =   1800
                     Width           =   1275
                     _Version        =   1048578
                     _ExtentX        =   2249
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "实际投坯量:"
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
                  Begin VB.Label Label46 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "标签模板:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   240
                     TabIndex        =   97
                     Top             =   8220
                     Width           =   975
                  End
                  Begin XtremeSuiteControls.Label Label47 
                     Height          =   255
                     Left            =   4140
                     TabIndex        =   96
                     Top             =   8265
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "细码单:"
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
                  Begin XtremeSuiteControls.Label Label48 
                     Height          =   255
                     Left            =   11280
                     TabIndex        =   95
                     Top             =   1260
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "染厂颜色:"
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
                     Left            =   240
                     TabIndex        =   94
                     Top             =   1800
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "码     数："
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
            End
         End
      End
   End
End
Attribute VB_Name = "frmOrderProductColor_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public colorid As String
Public departid As String
Public departmentid As String
Public processid As String
Public processmentid As String
Public processid2 As String
Public processmentid2 As String
Public processid3 As String
Public processmentid3 As String
Public ordertocolorid As String '复制行将外连接的订单明细的主键
Public lGroupID As Long
Private num As Long '标识公斤数和米数
Public id As String
Public Valuation As String
Public itemid As String   '色布计划主键
Private theidColor As String
Private a As Long

Private szFile As String

Private client As String
Public colororderid As String
Public colorplanid As String
Private B_orderitemid As String
Private rss As RecordSet

Private cls1 As New clsPicture

Private Type RGB
        Red   As Byte
        Green   As Byte
        Blue   As Byte
End Type

'Private theOrderID As String
'
'Public Property Let OrderID(ByVal vData As String)
'    theOrderID = vData
'End Property
'
'Private Sub FillData()
'    If Len(theOrderID) <= 0 Then
'        Exit Sub
'    End If
'
'    SetComboListIndex theOrderID
'End Sub
'Private Sub SetComboListIndex(ByVal vData As String)
'    Dim i As Long
'    For i = 0 To ComboBox1.ListCount - 1
'        If ComboBox1.List(i) = vData Then
'            ComboBox1.ListIndex = i
'        End If
'    Next
'End Sub
Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "保存"
            save (0)
        Case "保存并复制本合同"
            save (1)
'            Saveandcopy_1
        Case "保存并复制本订单"
            save (2)
'            Saveandcopy
        Case "第一单"
            movefrist
        Case "上一单"
            moveshang
        Case "下一单"
            movenext
        Case "最后单"
            movelast
        Case "退出"
            Unload Me
    End Select
End Sub
    


Private Sub ComboBox5_Click()
'        If Len(ComboBox5.Text) > 0 Then
            Dim rs As New RecordSet
            Dim sql As String
            sql = "Select B_ProgressItem From G_ProgressCraft where B_Parent='" & ComboBox5.Text & "'"
            rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
'        End If
        TDBGrid1.DataSource = rs
        TDBGrid1.Columns("B_ProgressItem").Caption = "进度工艺明细"
        TDBGrid1.MarqueeStyle = dbgHighlightRow
End Sub


Private Sub FlatEdit11_Change()
        On Error Resume Next
        Dim rs As New RecordSet
        Dim sql As String
        sql = "select * from G_Color where B_SID='" & colorid & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

        Dim ys As RGB
        With ys
            .Red = rs!B_Red
            .Green = rs!B_Green
            .Blue = rs!B_Blue
        End With
        Picture2.BackColor = rs!B_hex
End Sub
'裁剪折率
Private Sub FlatEdit33_Change()
    Dim f As Long
    If Val(FlatEdit33.Text) <> 0 Then
        f = Val(FlatEdit7.Text) * (Val(FlatEdit33.Text) + 1)
        FlatEdit57.Text = f
    End If
End Sub
'公斤数
Private Sub FlatEdit7_Change()
    If Val(FlatEdit33.Text) <> 0 Then
        f = Val(FlatEdit7.Text) * (Val(FlatEdit33.Text) + 1)
        FlatEdit57.Text = f
    End If
End Sub
'条数
Private Sub FlatEdit55_Change()
    If Val(FlatEdit56.Text) <> 0 Then
        f = Val(FlatEdit56.Text) * Val(FlatEdit55.Text)
        FlatEdit7.Text = f   '公斤数
    End If
End Sub
'条重
Private Sub FlatEdit56_Change()
    If Val(FlatEdit56.Text) <> 0 Then
        f = Val(FlatEdit56.Text) * Val(FlatEdit55.Text)
        FlatEdit7.Text = f
    End If
End Sub
'染厂折率
Private Sub FlatEdit58_Change()
    If Val(FlatEdit57.Text) <> 0 Then
        f = Val(FlatEdit57.Text) * (Val(FlatEdit58.Text) + 1)
        FlatEdit34.Text = f
    End If
End Sub

Private Sub Form_Load()
    InitFrm
'    dingdanhao
    baozhuang
    process
'    FillData
    TDBGrid1.Visible = False
    DateTimePicker1.Value = Now
    Grid
    C1Tab1.CurrTab = 1
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
    With ActiveBar22
        .ClientAreaControl = C1Tab1
        .RecalcLayout
    End With
        With ActiveBar23
        .ClientAreaControl = C1Elastic2
        .RecalcLayout
    End With
    DTPicker1.Value = Now
    DTPicker2.Value = Now
    DTPicker3.Value = Now
    DTPicker4.Value = Now
    setValuation
    Sum
    biaoq
    xima
End Sub

'标签模板
Private Sub biaoq()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "SELECT *From G_JBCPSet Where B_Effective = 1 ORDER BY B_Order"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Do While Not rs.EOF
        ComboBox1.AddItem "" & rs!B_GroupName & ""
        rs.movenext
    Loop
End Sub
'细码
Private Sub xima()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "SELECT * From G_JRKXMDYS Where B_Effective = 1 ORDER BY B_Order"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Do While Not rs.EOF
        ComboBox2.AddItem "" & rs!B_GroupName & ""
        rs.movenext
    Loop
End Sub
Private Sub Sum()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_BillDetailColor a inner join G_BillColor b on a.B_ID=b.B_ID where b.B_BelongOrderID='" & id & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    a = rs.RecordCount
End Sub


Private Sub baozhuang()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "Select B_SID From G_PackWay Where 1=1"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

    Do While Not rs.EOF
        ComboBox4.AddItem rs!B_sid
        rs.movenext
    Loop
End Sub
Private Sub process()
    Dim rs As New RecordSet
    Dim sql As String
    sql = " select B_SID from G_ProgressCraftCT where 1=1"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

    Do While Not rs.EOF
        ComboBox5.AddItem rs!B_sid
        rs.movenext
    Loop
End Sub



'---------------------------------------设置按钮的触发事件
Private Sub PushButton1_Click()
      On Error Resume Next
      Dim sql As String
      Dim rs As New RecordSet
      Dim frm1 As New frmpopupColor
      frm1.Show vbModal
      FlatEdit11.Text = Trim(frm1.colorname)
      colorid = frm1.colorid

        sql = "select * from G_Color where B_SID='" & colorid & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

        Dim ys As RGB
        With ys
            .Red = rs!B_Red
            .Green = rs!B_Green
            .Blue = rs!B_Blue
        End With
        Picture2.BackColor = rs!B_hex
      Unload frm1
End Sub

'Private Sub PushButton1_FetchCellStyle(ByVal Condition As Integer, _
'    ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, _
'    ByVal CellStyle As TrueOleDBGrid80.StyleDi)
'    On Error Resume Next
'    Dim ys As RGB
''    CellStyle.BackColor = TDBGrid3.Columns("B_Hex").CellValue(Bookmark)
'    CellStyle.ForeColor = PushButton1.CellValue(Bookmark)
'End Sub

'色布计划预览图片
Private Sub PushButton12_Click()
    On Error GoTo IFERR
    
    With CommonDialog1
        .ShowOpen
   
        szFile = .FileName
    End With
     
    If Len(szFile) <= 0 Then
        Exit Sub
    End If
    cls1.InitCls szFile, Picture5
    
'    Picture3.Picture = LoadPicture(szFile)
    
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "图片格式不支持，请提供JPG或者BMP格式图片！" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub
'色布计划上传照片
Private Sub PushButton13_Click()
Dim sql As String
Dim rs As New RecordSet
'    If TDBGrid4.ApproxCount <= 0 Then
'        MsgBox "当前没有订单号不能上传", vbInformation, "提示"
'        Exit Sub
'    End If
    If szFile <> "" Then
'                 需要引用：Microsoft Scripting Runtime
        Dim fso As New FileSystemObject
    
        Dim lLength As Long
        Dim oFile As File
    
        Set oFile = fso.GetFile(szFile)
        
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
Private Sub saveImage()
  
'    If rsgrid4!B_ordercode = "" Then
''        saveImage = False
'        Exit Sub
'    End If
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_image_NEW where 1=0"
    rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
    
     If szFile = "" Then
     
     Else
        Dim rs1 As New RecordSet
        Dim sql1 As String
        sql1 = "select * from G_image_NEW where B_BDCItemID='" & itemid & "'"
        rs1.Open sql1, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
        If rs1.RecordCount > 0 Then
'            Dim sql2 As String
'            sql2 = "update G_Image set B_Picture='" & szFile & "' where  B_ID='" & theID & "' and B_itemid='" & rsgrid4!B_OrderCode & "'"
'            Gm.cnnToolImage.cnn.Execute sql2
            
            PicSaveToDB rs1!B_picture, szFile
            rs1.Update
            rs1.Close
            Set rs1 = Nothing
        Else
            rs.AddNew
'            rs!B_id = theid
            PicSaveToDB rs!B_picture, szFile
            rs!B_BDCItemID = itemid  '明细表的主键  一张图对应一个色布计划的主键
            rs.Update
            rs.Close
            Set rs = Nothing
        End If
        
    End If
End Sub
'色布计划删除照片
Private Sub PushButton14_Click()
    Dim sql As String
    If theid > 0 Then
        sql = "delete from G_image where B_BDCItemID='" & itemid & "'"
        Gm.cnnToolImage.cnn.Execute sql
    End If
    Picture5.Picture = Nothing
End Sub

Private Sub PushButton2_Click()
    Dim frm1 As New frmpopupDepart
      frm1.Show vbModal
      FlatEdit8.Text = Trim(frm1.departName)
     departid = frm1.departid
      Unload frm1
End Sub

Private Sub PushButton3_Click()
    Dim frm1 As New frmpopupDepartment
      frm1.Show vbModal
      FlatEdit9.Text = Trim(frm1.departmentName)
      FlatEdit18.Text = frm1.phone
     departmentid = frm1.departmentid
      Unload frm1
End Sub

Private Sub PushButton4_Click()
    Dim frm1 As New frmpopupProcess
    frm1.Show vbModal
    FlatEdit14.Text = Trim(frm1.processName)
    
    processid = frm1.processid
    Unload frm1
End Sub

Private Sub PushButton5_Click()
    Dim frm1 As New frmpopupProcessment
    frm1.Show vbModal
    FlatEdit15.Text = Trim(frm1.processmentName)
    FlatEdit22.Text = frm1.phone
    processmentid = frm1.processmentid
    Unload frm1
End Sub
Private Sub PushButton7_Click()
    Dim frm1 As New frmpopupProcess
    frm1.Show vbModal
    FlatEdit25.Text = Trim(frm1.processName)
    processid2 = frm1.processid
    Unload frm1
End Sub
Private Sub PushButton6_Click()
    Dim frm1 As New frmpopupProcessment
    frm1.Show vbModal
    FlatEdit24.Text = Trim(frm1.processmentName)
    FlatEdit23.Text = frm1.phone
    processmentid2 = frm1.processmentid
    Unload frm1
End Sub
Private Sub PushButton9_Click()
    Dim frm1 As New frmpopupProcess
    frm1.Show vbModal
    FlatEdit30.Text = Trim(frm1.processName)
    processid3 = frm1.processid
    Unload frm1
End Sub
Private Sub PushButton8_Click()
    Dim frm1 As New frmpopupProcessment
    frm1.Show vbModal
    FlatEdit29.Text = Trim(frm1.processmentName)
    FlatEdit28.Text = frm1.phone
    processmentid3 = frm1.processmentid
    Unload frm1
End Sub

'---------------------------------------设置控件的中只能输入数字和小数点


Private Sub setValuation()
    If Len(Valuation) > 0 Then
        If Len(Trim(Valuation)) = 2 Then
            FlatEdit7.Enabled = False
            FlatEdit7.BackColor = &HC0C0C0
            num = 1
        Else
            FlatEdit6.Enabled = False
            FlatEdit6.BackColor = &HC0C0C0
              num = 2
        End If
    End If
End Sub
Private Sub save(ByVal num As Long)

     If yanzhenColor(id) = False Then
        Exit Sub
    End If
    If Trim(FlatEdit3.Text) = "" Then
        MsgBox "订单号不能为空", vbInformation, "提示"
        Exit Sub
    End If
     If Trim(FlatEdit8.Text) = "" Then
        MsgBox "染厂不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(FlatEdit9.Text) = "" Then
        MsgBox "染厂跟单不能为空", vbInformation, "提示"
        Exit Sub
    End If
'    If Trim(FlatEdit12.Text) <= 0 Then
'        MsgBox "染厂加工费不能为空", vbInformation, "提示"
'        Exit Sub
'    End If

     If Trim(ComboBox5.Text) = "" Then
        MsgBox "进度工艺不能为空", vbInformation, "提示"
        Exit Sub
    End If
    
    If Trim(FlatEdit35.Text) = "" Then
        MsgBox "实际投坯量不能为空", vbInformation, "提示"
        Exit Sub
    End If

    If num = 0 Then
             SavedetailColor
    End If
    If num = 1 Then
             Saveandcopy_1
    End If
         If num = 2 Then
             Saveandcopy
    End If
End Sub

Private Sub SavedetailColor()
     If Len(itemid) > 0 Then
        savedetail_update
     Else
        Dim rs As New RecordSet
        Dim sql As String
        '2018年4月13日19:53:30
        sql = "select * from G_BillColor where B_belongorderid='" & id & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
'        sql = "select * from G_BilldetailColor where B_itemid='" & itemid & "'"
'        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
      
        If rs.RecordCount <= 0 Then
               savemain
        Else
                theidColor = rs!B_id
        End If
        
            savedetail
            
    End If
    validation (id)
    Me.Hide
End Sub

Private Sub savemain()
    Set clsBL = New clsBL
    Dim sql As String
            Dim rs As New RecordSet
            sql = "select * from G_DraftBillWhite where 1=1 "
            rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            rs.AddNew
            Dim a As String
            a = Format(Now, "YYYY-MM-DD")
            rs!B_datecreate = a
            rs.Update
            theidColor = rs!B_id
            
            Dim rs1 As New RecordSet
            Dim sql1 As String
            sql1 = "select * from G_BillColor where 1=1 "
            rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            rs1.AddNew
               rs1!B_id = theidColor
               Dim b As String
               b = Format(Now, "YYYY-MM-DD")
               rs1!B_datecreate = b
               rs1!B_BID = B_BID_CC
               rs1!B_ObjectID = B_ObjectID_CC
               rs1!B_BillType = B_BillType_CC
               rs1!B_Codeid = clsBL.GetFrameCodeDetail_01(B_ObjectID_CC)
               rs1!B_BelongOrderID = id
               rs1.Update
               Dim rs2 As New RecordSet
               Dim sql2 As String
               sql2 = "delete from G_DraftBillColor where B_ID='" & theidwhite & "'"
               Gm.cnnTool.cnn.Execute sql2
     
End Sub
Private Sub savedetail()
    On Error Resume Next
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_DraftBillDetailColor where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            '获取最新的一个条码的自增数字
    lIncr = GetNewBCIncr
    szBC13 = GetBC13(FillGetBC12(lIncr))
    rs.AddNew
    rs!B_datecreate = Now
    rs.Update
     itemid = rs!B_ItemID
     Dim rs1 As New RecordSet
     Dim sql1 As String
     
     sql1 = "select * from G_BillDetailColor where 1=1"
     rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
     
 rs1.AddNew
      rs1!B_GroupID = lGroupID
      rs1!B_orderitemid = IIf(IsNull(ordertocolorid), 0, ordertocolorid)
      rs1!B_ItemID = itemid
      rs1!B_id = theidColor
      rs1!B_ItemIDB = FlatEdit3.Text
      rs1!B_GoodsNameAlias = FlatEdit2.Text
      rs1!B_Width = FlatEdit4.Text
      rs1!B_weight = FlatEdit10.Text
      rs1!B_color = colorid
      rs1!B_orderColor = FlatEdit11.Text
      rs1!B_Producer = FlatEdit1.Text
      rs1!B_SeHao = FlatEdit5.Text
     
      rs1!B_meter = IIf(IsNull(FlatEdit6.Text), "", FlatEdit6.Text)
      rs1!B_kg = FlatEdit7.Text
      If Len(FlatEdit37.Text) > 0 Then
        rs1!B_BoxQty = IIf(IsNull(FlatEdit37.Text), 0, FlatEdit37.Text)
      End If
      rs1!B_depart = IIf(IsNull(departid), "", departid)
      rs1!B_department = IIf(IsNull(departmentid), "", departmentid)

      rs1!B_departdate = Format(DTPicker1.Value, "YYYY-MM-DD")
      
      rs1!B_flowCard = FlatEdit12.Text
      rs1!B_departdannote = FlatEdit13.Text
      rs1!B_processunit = IIf(IsNull(processid), "", processid)
      rs1!B_processdocumentary = IIf(IsNull(processmentid), "", processmentid)
      
      rs1!B_phone1 = FlatEdit22.Text
      rs1!B_processdate = Format(DTPicker2.Value, "YYYY-MM-DD")
      If Len(FlatEdit16.Text) > 0 Then
        rs1!B_processcost = IIf(IsNull(FlatEdit16.Text), "", FlatEdit16.Text)
      End If
      rs1!B_processnote = FlatEdit17.Text
      rs1!B_processunit2 = IIf(IsNull(processid2), "", processid2)
      rs1!B_processdocumentary2 = IIf(IsNull(processmentid2), "", processmentid2)
      rs1!B_phone2 = FlatEdit23.Text
      rs1!B_processdate2 = DTPicker3.Value
      If Len(FlatEdit27.Text) > 0 Then
        rs1!B_processCost2 = FlatEdit27.Text
      End If
      rs1!B_processnote2 = FlatEdit26.Text
      rs1!B_processunit3 = IIf(IsNull(processid3), "", processid3)
      rs1!B_processdocumentary3 = IIf(IsNull(processmentid3), "", processmentid3)
      rs1!B_processdate3 = DTPicker4.Value
      If Len(FlatEdit31.Text) > 0 Then
        rs1!B_processCost3 = FlatEdit31.Text
      End If
      rs1!B_processnote3 = FlatEdit32.Text
      rs1!B_Progressprocess = ComboBox5.Text
      If Len(FlatEdit19.Text) > 0 Then
        rs1!B_Paper = IIf(IsNull(FlatEdit19.Text), 0, FlatEdit19.Text)
      End If
      If Len(FlatEdit20.Text) > 0 Then
        rs1!B_pocket = FlatEdit20.Text
      End If
      rs1!B_Empty = FlatEdit21.Text
      rs1!B_packagstyle = ComboBox4.Text
      rs1!B_departCost = FlatEdit12.Text
      rs1!B_fold = FlatEdit33.Text
      rs1!B_Cast = FlatEdit34.Text
      rs1!B_PracticeCast = FlatEdit35.Text
      rs1!B_LabelTemplate = ComboBox1.Text
      rs1!B_DetailTemplate = ComboBox2.Text
      rs1!B_DepartColor = FlatEdit36.Text
'      rs1!B_Progressprocess = ComboBox5.Text
'      rs1!B_Paper = FlatEdit19.Text
'      rs1!B_pocket = FlatEdit20.Text
'      rs1!B_Empty = FlatEdit21.Text
'      rs1!B_packagstyle = ComboBox4.Text
'      rs1!B_departCost = FlatEdit12.Text
    
    rs1!B_PBGuige = FlatEdit51.Text    '坯布规格  坯布厂电话  坯布厂地址  送布地址
    rs1!B_PBPhone = FlatEdit52.Text
    rs1!B_PBDiZhi = FlatEdit53.Text
    rs1!B_SBDiZhi = FlatEdit54.Text
    
    rs1!B_TiaoShu = FlatEdit55.Text
    rs1!B_TiaoZhong = FlatEdit56.Text
    rs1!B_ColorQty = FlatEdit57.Text
    rs1!B_ColorZL = FlatEdit58.Text
    
    
     
     
      rs1.Update
      Dim sql2 As String
      sql2 = "delete from G_DraftBillDetailColor where B_itemid='" & itemid & "'"
      Gm.cnnTool.cnn.Execute sql2
End Sub

'--------------------------------------------------------------------------------
' Project    :       织造企业MIS系统
' Procedure  :       savedetail_update
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       YZ6XXTUV8AGVQLQ
' Date-Time  :       2018/12/3-16:02:49
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub savedetail_update()
        Dim sql As String
        Dim a As String
        Dim b As String
        Dim c As String
        Dim d As String
        a = Format(DTPicker1.Value, "YYYY-MM-dd")
        b = Format(DTPicker2.Value, "YYYY-MM-dd")
        c = Format(DTPicker3.Value, "YYYY-MM-dd")
        d = Format(DTPicker4.Value, "YYYY-MM-dd")
         sql = "exec usp_updateColordetail '" & itemid & "','" & FlatEdit3.Text & "','" & FlatEdit2.Text & "','" & FlatEdit4.Text & "','" & FlatEdit10.Text & "'"
         sql = sql & ",'" & colorid & "','" & FlatEdit1.Text & "','" & FlatEdit5.Text & "','" & FlatEdit6.Text & "','" & FlatEdit7.Text & "','" & FlatEdit37.Text & "','" & departid & "','" & departmentid & "','" & FlatEdit18.Text & "'"
         sql = sql & ",'" & a & "','" & FlatEdit13.Text & "','" & processid & "','" & processmentid & "','" & FlatEdit22.Text & "','" & b & "'"
         sql = sql & ",'" & FlatEdit16.Text & "','" & FlatEdit17.Text & "'"
         sql = sql & ",'" & processid2 & "','" & processmentid2 & "','" & FlatEdit23.Text & "','" & c & "','" & FlatEdit27.Text & "','" & FlatEdit26.Text & "'"
         sql = sql & ",'" & processid3 & "','" & processmentid3 & "','" & FlatEdit28.Text & "','" & d & "','" & FlatEdit31.Text & "','" & FlatEdit32.Text & "'"
         sql = sql & ",'" & ComboBox5.Text & "','" & FlatEdit19.Text & "','" & FlatEdit20.Text & "','" & FlatEdit21.Text & "','" & ComboBox4.Text & "','" & FlatEdit12.Text & "','" & FlatEdit33.Text & "','" & FlatEdit34.Text & "','" & FlatEdit35.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & FlatEdit36.Text & "'"
         sql = sql & ",'" & FlatEdit51.Text & "','" & FlatEdit52.Text & "','" & FlatEdit53.Text & "','" & FlatEdit54.Text & "'"
          sql = sql & ",'" & FlatEdit55.Text & "','" & FlatEdit56.Text & "','" & FlatEdit57.Text & "','" & FlatEdit58.Text & "'"
         Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
        
End Sub

Private Sub Saveandcopy()
          Dim sql As String
        Dim a As String
        Dim b As String
        Dim c As String
        Dim d As String
        a = Format(DTPicker1.Value, "YYYY-MM-dd")
        b = Format(DTPicker2.Value, "YYYY-MM-dd")
        c = Format(DTPicker3.Value, "YYYY-MM-dd")
        d = Format(DTPicker3.Value, "YYYY-MM-dd")
        
         sql = "exec usp_saveandCopy '" & id & "','" & FlatEdit2.Text & "','" & FlatEdit4.Text & "','" & FlatEdit10.Text & "','" & FlatEdit1.Text & "','" & FlatEdit5.Text & "'"
         sql = sql & ",'" & departid & "','" & departmentid & "','" & FlatEdit18.Text & "'"
         sql = sql & ",'" & a & "','" & FlatEdit13.Text & "','" & processid & "','" & processmentid & "','" & FlatEdit22.Text & "','" & b & "'"
         sql = sql & ",'" & FlatEdit16.Text & "','" & FlatEdit17.Text & "'"
         sql = sql & ",'" & processid2 & "','" & processmentid2 & "','" & FlatEdit23.Text & "','" & c & "','" & FlatEdit27.Text & "','" & FlatEdit26.Text & "'"
         sql = sql & ",'" & processid3 & "','" & processmentid3 & "','" & FlatEdit28.Text & "','" & d & "','" & FlatEdit31.Text & "','" & FlatEdit32.Text & "'"
         sql = sql & ",'" & ComboBox5.Text & "','" & FlatEdit19.Text & "','" & FlatEdit20.Text & "','" & FlatEdit21.Text & "','" & ComboBox4.Text & "','" & FlatEdit12.Text & "','" & FlatEdit33.Text & "','" & FlatEdit34.Text & "','" & FlatEdit35.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & FlatEdit3.Text & "','" & FlatEdit36.Text & "'"
         sql = sql & ",'" & FlatEdit51.Text & "','" & FlatEdit52.Text & "','" & FlatEdit53.Text & "','" & FlatEdit54.Text & "'"
          sql = sql & ",'" & FlatEdit55.Text & "','" & FlatEdit56.Text & "','" & FlatEdit57.Text & "','" & FlatEdit58.Text & "'"
         Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
        MsgBox "保存成功并复制成功", vbInformation, "提示"
        validation (id)
        Me.Hide
End Sub
Private Sub Saveandcopy_1()
          Dim sql As String
        Dim a As String
        Dim b As String
         Dim c As String
        Dim d As String
        a = Format(DTPicker1.Value, "YYYY-MM-dd")
        b = Format(DTPicker2.Value, "YYYY-MM-dd")
         c = Format(DTPicker3.Value, "YYYY-MM-dd")
        d = Format(DTPicker3.Value, "YYYY-MM-dd")
         sql = "exec usp_saveandCopy '" & id & "','" & FlatEdit2.Text & "','" & FlatEdit4.Text & "','" & FlatEdit10.Text & "','" & FlatEdit1.Text & "','" & FlatEdit5.Text & "'"
         sql = sql & ",'" & departid & "','" & departmentid & "','" & FlatEdit18.Text & "'"
         sql = sql & ",'" & a & "','" & FlatEdit13.Text & "','" & processid & "','" & processmentid & "','" & FlatEdit22.Text & "','" & b & "'"
         sql = sql & ",'" & FlatEdit16.Text & "','" & FlatEdit17.Text & "'"
         sql = sql & ",'" & processid2 & "','" & processmentid2 & "','" & FlatEdit23.Text & "','" & c & "','" & FlatEdit27.Text & "','" & FlatEdit26.Text & "'"
         sql = sql & ",'" & processid3 & "','" & processmentid3 & "','" & FlatEdit28.Text & "','" & d & "','" & FlatEdit31.Text & "','" & FlatEdit32.Text & "'"
         sql = sql & ",'" & ComboBox5.Text & "','" & FlatEdit19.Text & "','" & FlatEdit20.Text & "','" & FlatEdit21.Text & "','" & ComboBox4.Text & "','" & FlatEdit12.Text & "','" & FlatEdit33.Text & "','" & FlatEdit34.Text & "','" & FlatEdit35.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & "" & "','" & FlatEdit36.Text & "'"
         sql = sql & ",'" & FlatEdit51.Text & "','" & FlatEdit52.Text & "','" & FlatEdit53.Text & "','" & FlatEdit54.Text & "'"
          sql = sql & ",'" & FlatEdit55.Text & "','" & FlatEdit56.Text & "','" & FlatEdit57.Text & "','" & FlatEdit58.Text & "'"
         Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
        MsgBox "保存成功并复制成功", vbInformation, "提示"
        validation (id)
        Me.Hide
End Sub

Private Sub movefrist()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select top 1 * from G_BilldetailColor where B_ID=(select B_ID from G_BillColor where B_BelongOrderID='" & id & "') "
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     If rs.RecordCount <= 0 Then
            MsgBox "当前没有数据", vbInformation, "提示"
     Else
            Label26.Caption = 1
            itemid = rs!B_ItemID
            openbill
     End If
End Sub

Private Sub moveshang()
     Dim rs As New RecordSet
     Dim sql As String
     sql = "select top 1 * from G_BilldetailColor where B_ID=(select B_ID from G_BillColor where B_BelongOrderID='" & id & "') and B_itemid<'" & itemid & "'  Order by B_itemid desc"
      rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
      
      If rs.RecordCount <= 0 Then
        MsgBox "已经是第一单了", vbOKOnly + vbInformation, "提示"
     Else
        Label26.Caption = Val(Label26.Caption) - 1
        itemid = rs!B_ItemID
        openbill
     End If
End Sub

Private Sub movenext()
     Dim rs As New RecordSet
     Dim sql As String
     sql = "select top 1 * from G_BilldetailColor where B_ID=(select B_ID from G_BillColor where B_BelongOrderID='" & id & "') and B_itemid>'" & itemid & "'  Order by B_itemid asc"
      rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
      
      If rs.RecordCount <= 0 Then
        MsgBox "最后一单了", vbOKOnly + vbInformation, "提示"
     Else
        Label26.Caption = Val(Label26.Caption) + 1
        itemid = rs!B_ItemID
        openbill
     End If
End Sub

Private Sub movelast()
     Dim rs As New RecordSet
     Dim sql As String
     sql = "select top 1 * from G_BilldetailColor where B_ID=(select B_ID from G_BillColor where B_BelongOrderID='" & id & "')   Order by B_itemid desc"
      rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
      
      If rs.RecordCount <= 0 Then
        MsgBox "当前没有任何数据", vbOKOnly + vbInformation, "提示"
     Else
         Label26.Caption = a
       itemid = rs!B_ItemID
        openbill
     End If
End Sub

Private Sub openbill()
     Dim sql As String
    Dim rs As New RecordSet
    sql = "select *from G_Billdetailcolor where B_itemid='" & itemid & "' "
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    FlatEdit3.Text = IIf(IsNull(rs!B_ItemIDB), "", rs!B_ItemIDB)
    FlatEdit2.Text = IIf(IsNull(rs!B_GoodsNameAlias), "", rs!B_GoodsNameAlias)
    FlatEdit4.Text = IIf(IsNull(rs!B_Width), "", rs!B_Width)
    FlatEdit10.Text = IIf(IsNull(rs!B_weight), "", rs!B_weight)
    colorid = IIf(IsNull(rs!B_color), "", rs!B_color)
'    FlatEdit11.Text = GetColorName(rs!B_Color)
    FlatEdit11.Text = IIf(IsNull(rs!B_orderColor), "", rs!B_orderColor)
    FlatEdit1.Text = IIf(IsNull(rs!B_Producer), "", rs!B_Producer)
    FlatEdit5.Text = IIf(IsNull(rs!B_SeHao), "", rs!B_SeHao)
    FlatEdit37.Text = IIf(IsNull(rs!B_BoxQty), "", rs!B_BoxQty)
'    If Trim(rs!B_Meter) > 0 Then
        FlatEdit6.Text = rs!B_meter
'        FlatEdit6.Enabled = True
'         FlatEdit7.Enabled = False
'         FlatEdit7.BackColor = &HC0C0C0
'    End If
'    If Trim(rs!B_KG) > 0 Then
        FlatEdit7.Text = rs!B_kg
'        FlatEdit6.Enabled = False
'         FlatEdit7.Enabled = True
'         FlatEdit6.BackColor = &HC0C0C0
'    End If
'    frm1.FlatEdit6.Text = IIf(IsNull(rs!B_meter), "", rs!B_meter)
'    frm1.FlatEdit7.Text = IIf(IsNull(rs!B_KG), "", rs!B_KG)
    FlatEdit8.Text = getDepartName(IIf(IsNull(rs!B_depart), "", rs!B_depart))
    departid = IIf(IsNull(rs!B_depart), "", rs!B_depart)
    FlatEdit9.Text = getprocessName(IIf(IsNull(rs!B_department), "", rs!B_department))
    departmentid = IIf(IsNull(rs!B_department), "", rs!B_department)
    If IIf(IsNull(rs!B_departdate), "", rs!B_departdate) = "" Then
        DTPicker1.Value = Now
    Else
        DTPicker1.Value = rs!B_departdate
    End If
'    frm1.DTPicker1.Value = IIf(IsNull(rsdetailColor!B_departdate), "", rsdetailColor!B_departdate)
'    frm1.FlatEdit12.Text = IIf(IsNull(rs!B_flowCard), "", rs!B_flowCard)
    FlatEdit13.Text = IIf(IsNull(rs!B_departdannote), "", rs!B_departdannote)
    FlatEdit14.Text = getDepartName(IIf(IsNull(rs!B_processunit), "", rs!B_processunit))
    processid = IIf(IsNull(rs!B_processunit), "", rs!B_processunit)
    FlatEdit15.Text = getprocessName(IIf(IsNull(rs!B_processdocumentary), "", rs!B_processdocumentary))
    processmentid = IIf(IsNull(rs!B_processdocumentary), "", rs!B_processdocumentary)
   If IIf(IsNull(rs!B_processdate), "", rs!B_processdate) = "" Then
        DTPicker2.Value = Now
    Else
        DTPicker2.Value = rs!B_processdate
    End If
'    frm1.DTPicker2.Value = IIf(IsNull(rs!B_processdate), "", rs!B_processdate)
    FlatEdit17.Text = IIf(IsNull(rs!B_processnote), "", rs!B_processnote)
     If IIf(IsNull(rs!B_Progressprocess), "", rs!B_Progressprocess) = "" Then
            ComboBox5.Text = ""
     Else
        ComboBox5.Text = GetProgressCraftCT(IIf(IsNull(rs!B_Progressprocess), "", rs!B_Progressprocess))
     End If
'    frm1.ComboBox5.Text = GetProgressCraftCT(rs!B_Progressprocess)
    FlatEdit19.Text = IIf(IsNull(rs!B_Paper), "", rs!B_Paper)
    FlatEdit20.Text = IIf(IsNull(rs!B_pocket), "", rs!B_pocket)
    FlatEdit21.Text = IIf(IsNull(rs!B_Empty), "", rs!B_processnote)
    If IIf(IsNull(rs!B_packagstyle), "", rs!B_packagstyle) = "" Then
            ComboBox4.Text = ""
     Else
        ComboBox4.Text = GetB_packagstyle(rs!B_packagstyle)
     End If
End Sub

Private Function GetColorName(ByVal da As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select B_Name from G_Color where B_SID='" & da & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        GetColorName = rs!B_name
    Else
        GetColorName = ""
    End If
End Function

Private Function getDepartName(ByVal da As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select B_ClientName from G_ContactCompany where B_Clientid='" & da & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        getDepartName = rs!B_ClientName
    Else
        getDepartName = ""
    End If
End Function
Private Function getprocessName(ByVal da As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select B_Name from G_Employee where B_SID='" & da & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        getprocessName = rs!B_name
    Else
        getprocessName = ""
    End If
End Function

Private Function GetProgressCraftCT(ByVal da As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select B_SID from G_ProgressCraftCT where B_SID='" & da & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        GetProgressCraftCT = rs!B_sid
    Else
        GetProgressCraftCT = ""
    End If
End Function
Private Function GetB_packagstyle(ByVal da As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "Select B_SID From G_PackWay where B_SID='" & da & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        GetB_packagstyle = rs!B_sid
    Else
        GetB_packagstyle = ""
    End If
End Function

Private Sub FlatEdit19_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit20_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit21_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
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
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit16_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit33_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit35_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit19_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            FlatEdit20.SetFocus
    End Select
End Sub
Private Sub FlatEdit20_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            FlatEdit21.SetFocus
    End Select
End Sub
Private Sub FlatEdit21_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            ComboBox4.SetFocus
    End Select
End Sub
Private Sub FlatEdit33_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            FlatEdit35.SetFocus
    End Select
End Sub
Private Sub validation(ByVal id As String)
    Dim sql As String
    Dim rs As New RecordSet
    sql = "select * from G_BillDetailColor where B_ID in (select B_ID from G_BillColor where B_BelongOrderid='" & id & "')"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    Do While Not rs.EOF
        If IIf(IsNull(rs!B_Date), "", rs!B_Date) <> "" Then
                Exit Sub
        End If
        rs.movenext
    Loop
    rs.MoveFirst
    Do While Not rs.EOF
        If IIf(IsNull(rs!B_departdate), "", rs!B_departdate) = "" Then
            Exit Sub
        End If
        rs.movenext
    Loop
    rs.MoveFirst
    Do While Not rs.EOF
        rs!B_Date = Now
        rs.movenext
    Loop
End Sub
Public Function yanzhenColor(ByVal theid As String) As Boolean
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
    sql = "Select * From G_SystemUser where B_UserName='" & Gm.SysID.SystemUser & "'"
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs!B_SuperAdmin = 1 Then
        yanzhenColor = True
        Exit Function
    End If
    
    sql1 = "select distinct B_date from G_BilldetailColor where B_ID=(select B_ID from G_BillColor where B_belongorderid='" & theid & "')"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
'    sql2 = "select * from G_BillColor where B_Belongorderid='" & theid & "'"
 sql2 = "SELECT * FROM G_UserPro WHERE B_username='" & Gm.SysID.SystemUser & "' AND B_objectid='11S005'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
   If rs2.RecordCount > 0 Then
        If IIf(IsNull(rs2!B_new), 0, rs2!B_new) = 1 Then
            yanzhenColor = True
        Else
            yanzhenColor = False
            MsgBox "请设置权限", vbInformation, "提示"
            Exit Function
        End If
        If IIf(IsNull(rs1!B_Date), "", rs1!B_Date) <> "" Then
            If DateDiff("s", rs1!B_Date, Now) > 84600 Then
                yanzhenColor = False
                MsgBox "已经超过制作本单据的时间不能进行修改", vbInformation, "提示"
            Else
                yanzhenColor = True
            End If
        End If
    Else
        yanzhenColor = False
        MsgBox "你没有此权限", vbInformation, "提示"
    End If
End Function


Private Sub ActiveBar23_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "保存"
            savecg
        Case "退出"
            Unload Me
        Case "新增"
            AddNew
        Case "删除"
            de
    End Select
End Sub


Private Sub PushButton10_Click()
    Dim frm1 As New frmPopupClient_Edit
     frm1.a = "色布供应商"
     frm1.Show vbModal
    client = frm1.clientid
    FlatEdit45.Text = frm1.ClientName
    Unload frm1
End Sub

Private Sub Grid()
    On Error Resume Next
    
    Set rss = New RecordSet
    Dim sql1 As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_BilldetailColor where B_itemid='" & colororderid & "' "
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
   
'    sql1 = "select * from G_BilldetailColor where B_orderitemid='" & rs!B_orderitemid & "' and isnull(B_intype,0)=1"
'    rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    sql1 = "exec usp_SelectColororder '" & rs!B_orderitemid & "' "
    rss.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    TDBGrid2.DataSource = rss
    If rss.RecordCount > 0 Then
        rss.MoveFirst
    End If
    setgrid
    
    
    
    
    
End Sub
Private Sub setgrid()
    TDBGrid2.Columns("B_ItemIDB").Caption = "订单号"
    TDBGrid2.Columns("B_ClientName").Caption = "色布供应商"
    TDBGrid2.Columns("B_Width").Caption = "门幅"
    TDBGrid2.Columns("B_weight").Caption = "克重"
    TDBGrid2.Columns("B_GoodsNameAlias").Caption = "品名"
    TDBGrid2.Columns("B_Name").Caption = "颜色"
    TDBGrid2.Columns("B_SeHao").Caption = "色号"
    TDBGrid2.Columns("B_ps").Caption = "匹数"
    TDBGrid2.Columns("B_kg").Caption = "公斤"
    TDBGrid2.Columns("B_meter").Caption = "米数"
    TDBGrid2.Columns("B_qty").Caption = "码数"
    TDBGrid2.Columns("B_departdate").Caption = "交期"
    TDBGrid2.Columns("B_MemoDetail").Caption = "备注"
    TDBGrid2.Columns("B_hex").Caption = "色块"
    TDBGrid2.Columns("B_price").Caption = "单价"
    TDBGrid2.Columns("B_ItemIDB").width = 900
    TDBGrid2.Columns("B_Width").width = 900
    TDBGrid2.Columns("B_weight").width = 900
    TDBGrid2.Columns("B_SeHao").width = 900
    TDBGrid2.Columns("B_ps").width = 900
    TDBGrid2.Columns("B_kg").width = 900
    TDBGrid2.Columns("B_meter").width = 900
    TDBGrid2.Columns("B_qty").width = 900
    TDBGrid2.Columns("B_hex").width = 900
    TDBGrid2.Columns("B_price").width = 900
    
    TDBGrid2.Columns("B_Clientid").Visible = False
    TDBGrid2.Columns("B_Clientid").AllowSizing = False
    TDBGrid2.Columns("B_Clientid").Locked = True
        TDBGrid2.Columns("B_color").Visible = False
    TDBGrid2.Columns("B_color").AllowSizing = False
    TDBGrid2.Columns("B_color").Locked = True
      TDBGrid2.Columns("B_itemid").Visible = False
    TDBGrid2.Columns("B_itemid").AllowSizing = False
    TDBGrid2.Columns("B_itemid").Locked = True
    
    TDBGrid2.Columns("B_Hex").FetchStyle = True
    TDBGrid2.MarqueeStyle = dbgHighlightRow
    TDBGrid2.HoldFields
End Sub


Private Sub savecg()
    If Trim(client) = "" Then
        MsgBox "客户不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(FlatEdit41.Text) = "" Then
        MsgBox "门幅不能为空", vbInformation, "提示"
        Exit Sub
    End If
    
    If Trim(FlatEdit42.Text) = "" Then
        MsgBox "克重不能为空", vbInformation, "提示"
        Exit Sub
    End If
    
    If Trim(FlatEdit40.Text) = "" Then
        MsgBox "品名不能为空", vbInformation, "提示"
        Exit Sub
    End If
    
    If Trim(colorid) = "" And Trim(FlatEdit43.Text) = "" Then
        MsgBox "颜色或者色号不能为空", vbInformation, "提示"
        Exit Sub
    End If
    
    If Trim(FlatEdit49.Text) = "" And Trim(FlatEdit44.Text) = "" And Trim(FlatEdit46.Text) = "" And Trim(FlatEdit38.Text) = "" Then
        MsgBox "匹数,公斤数,米数,码数,任写其一", vbInformation, "提示"
        Exit Sub
    End If
        
    If Trim(FlatEdit50.Text) = "" Then
        MsgBox "单价不能为空", vbInformation, "提示"
        Exit Sub
    End If
    
    If B_orderitemid = "" Then
        saveALL
    Else
        saveupdate
    End If
    AddNew
End Sub
Private Sub saveALL()
    Dim idd As String
    Dim rs As New RecordSet
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim sql1 As String
    Dim sql2 As String
    Dim rs3 As New RecordSet
    Dim sql3 As String
    
    sql3 = "select * from G_BilldetailColor where B_itemid='" & colororderid & "' "
    rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    sql = "select * from G_draftBilldetailColor where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.AddNew
    rs!B_datecreate = Now
    rs.Update
    idd = rs!B_ItemID
    
    sql1 = "select * from G_BilldetailColor where 1=1"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs1.AddNew
    rs1!B_ItemID = idd
    rs1!B_ItemIDB = FlatEdit39.Text
    rs1!B_Clientid = client
    rs1!B_Width = FlatEdit41.Text
    rs1!B_weight = FlatEdit42.Text
    rs1!B_GoodsNameAlias = FlatEdit40.Text
    rs1!B_color = colorplanid
    rs1!B_orderColor = FlatEdit48.Text
    rs1!B_SeHao = FlatEdit43.Text
    rs1!B_price = FlatEdit50.Text
    rs1!B_ps = FlatEdit49.Text
    rs1!B_kg = FlatEdit44.Text
    If Len(Trim(FlatEdit46.Text)) > 0 Then
        rs1!B_meter = IIf(IsNull(FlatEdit46.Text), 0, FlatEdit46.Text)
    End If
    If Len(Trim(FlatEdit38.Text)) > 0 Then
        rs1!B_qty = IIf(IsNull(FlatEdit38.Text), 0, FlatEdit38.Text)
    End If
    rs1!B_departdate = DateTimePicker1.Value
    rs1!B_MemoDetail = FlatEdit47.Text
    rs1!B_intype = 1
    rs1!B_orderitemid = rs3!B_orderitemid
    rs1.Update
    
    sql2 = "delete from G_draftBilldetailColor where B_itemid='" & idd & "'"
    Gm.cnnTool.cnn.Execute sql2
'    MsgBox "保存成功", vbInformation, "提示"
    rss.requery
End Sub

Private Sub PushButton11_Click()
'    Dim frm1 As New frmpopupColor
'    frm1.Show vbModal
'    If frm1.bsaved = True Then
'        colorplanid = frm1.colorid
'        FlatEdit48.Text = frm1.colorname
'    End If
    On Error Resume Next
    Dim sql As String
    Dim rs As New RecordSet
    Dim frm1 As New frmpopupColor
    frm1.Show vbModal
    FlatEdit48.Text = Trim(frm1.colorname)
    colorplanid = frm1.colorid
    If frm1.bsaved = True Then
        FlatEdit48.Enabled = True
    Else
        If colorplanid = "" Then
             FlatEdit48.Enabled = False
        End If
    End If
    sql = "select * from G_Color where B_SID='" & colorplanid & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

        If rs!B_hex <> "" Then
        
            Picture4.BackColor = rs!B_hex
        Else
            Picture4.BackColor = &H8000000F
        End If
    
        
    Unload frm1
End Sub

Private Sub saveupdate()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "exec usp_OrderColor '" & B_orderitemid & "','" & FlatEdit39.Text & "','" & client & "','" & FlatEdit41.Text & "','" & FlatEdit42.Text & "','" & FlatEdit40.Text & "',"
    sql = sql & "'" & colorid & "','" & FlatEdit43.Text & "','" & FlatEdit49.Text & "','" & FlatEdit44.Text & "',"
    sql = sql & "'" & FlatEdit46.Text & "','" & FlatEdit38.Text & "','" & DateTimePicker1.Value & "','" & FlatEdit47.Text & "','" & FlatEdit48.Text & "','" & FlatEdit50.Text & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset
'    MsgBox "保存完成", vbInformation, "提示"
    rss.requery
End Sub

'Private Function clientname(ByVal client As String)
'    Dim rs As New RecordSet
'    Dim sql As String
'    sql = "select * from G_ContactCompany where B_clientid='" & client & "'"
'    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
'    If rs.RecordCount > 0 Then
'        clientname = rs!B_ClientName
'    Else
'        clientname = ""
'    End If
'End Function
'
'Private Function colorname(ByVal client As String)
'    Dim rs As New RecordSet
'    Dim sql As String
'    sql = "select * from G_Color where B_sid='" & client & "'"
'    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
'    If rs.RecordCount > 0 Then
'        colorname = rs!B_name
'    Else
'        colorname = ""
'    End If
'End Function

Private Sub FlatEdit49_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub

Private Sub FlatEdit44_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit46_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub

Private Sub FlatEdit38_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit50_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub

Private Sub TDBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    B_orderitemid = rss!B_ItemID
    FlatEdit39.Text = rss!B_ItemIDB
    client = rss!B_Clientid
    FlatEdit45.Text = rss!B_ClientName
    FlatEdit41.Text = rss!B_Width
    FlatEdit42.Text = rss!B_weight
    FlatEdit40.Text = rss!B_GoodsNameAlias
    colorplanid = rss!B_color
    FlatEdit48.Text = rss!B_name
    FlatEdit43.Text = rss!B_SeHao
    FlatEdit49.Text = rss!B_ps
    FlatEdit44.Text = rss!B_kg
    FlatEdit46.Text = IIf(IsNull(rss!B_meter), "", rss!B_meter)
    FlatEdit38.Text = IIf(IsNull(rss!B_qty), "", rss!B_qty)
    DateTimePicker1.Value = rss!B_departdate
    FlatEdit47.Text = rss!B_MemoDetail
    Picture4.BackColor = rss!B_hex
    FlatEdit50.Text = rss!B_price
End Sub

Private Sub AddNew()
    Dim a As String
    a = FlatEdit39.Text
    B_orderitemid = ""
    colorplanid = ""
    client = ""
    
     FlatEdit39.Text = ""
    FlatEdit45.Text = ""
    FlatEdit41.Text = ""
    FlatEdit42.Text = ""
    FlatEdit40.Text = ""
    FlatEdit48.Text = ""
    FlatEdit43.Text = ""
    FlatEdit49.Text = ""
    FlatEdit44.Text = ""
    FlatEdit46.Text = ""
    FlatEdit38.Text = ""
    FlatEdit47.Text = ""
    FlatEdit50.Text = ""
    FlatEdit39.Text = a
End Sub

Private Sub de()
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim sql As String
    Dim rs As New RecordSet
    sql = "delete from G_BilldetailColor where B_itemid='" & rss!B_ItemID & "'" '
    Gm.cnnTool.cnn.Execute sql
    rss.requery
End Sub

Private Sub FlatEdit6_KeyPress(KeyAscii As Integer)
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
Private Sub FlatEdit37_KeyPress(KeyAscii As Integer)
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
Private Sub TDBGrid2_FetchCellStyle(ByVal Condition As Integer, _
    ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, _
    ByVal CellStyle As TrueOleDBGrid80.StyleDisp)

'    Dim ys As RGB
    On Error Resume Next
    Debug.Print TDBGrid2.Columns("B_Hex").CellValue(bookmark)
    CellStyle.BackColor = TDBGrid2.Columns("B_Hex").CellValue(bookmark)
     CellStyle.ForeColor = TDBGrid2.Columns("B_Hex").CellValue(bookmark)
End Sub
'从表G_BillDetailColor获取当前最新一个条码的自增数字
Private Function GetNewBCIncr() As Long
    Dim rs As New RecordSet
    strSQL = "select top 1 * from G_BillDetailColor order by B_BCIncr desc"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Dim lRturn As Long
    If rs.RecordCount <= 0 Then
        lRturn = 1
    Else
        lRturn = IIf(IsNull(rs!B_BCIncr), 0, rs!B_BCIncr) + 1
    End If
    
    rs.Close
    Set rs = Nothing
    
    GetNewBCIncr = lRturn
End Function
Private Function GetBC13(ByVal vBC12 As String) As String
    Dim szRturn As String
    szRturn = GetEAN13CheckOut(vBC12)
    
    GetBC13 = vBC12 & szRturn
End Function

'获取最新的一个13位条码
Private Function GetBC13Ex() As String
    Dim szIncr As String
    szIncr = GetNewBCIncr
    
    Dim szBC12 As String
    szBC12 = FillGetBC12(GetNewBCIncr)
    
    GetBC13Ex = GetBC13(szBC12)
End Function
'传入参数：任意长度的自增数字的字符串类型
'返回值：返回BC13条码的前面12位字符
Private Function FillGetBC12(ByVal vIncr As String) As String
    Dim cls1 As New clsString
    Dim szReturn As String
    
    szReturn = cls1.FillRepeat(vIncr, 11, "0", True)
    szReturn = COLORBC13FIRST & szReturn
    
    FillGetBC12 = szReturn
End Function

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

Private Sub ImageUI()
               If itemid = "" Then
                        Exit Sub
                End If
                Dim rs As New RecordSet
                Dim sql As String
                Dim clsFile01 As New clsFile
                Dim szPic As String
                
                sql = "select * from G_image_NEW where B_BDCItemID='" & itemid & "'"
                rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
                
                If rs.RecordCount > 0 Then
                    szPic = App.Path & "\temp\" & rs!B_id & rs!B_ItemID & ".JPG"
                    Debug.Print szPic
                    
                    clsFile01.DownloadPic rs!B_picture, szPic
                    cls1.InitCls szPic, Picture5
                    
                    'PicShow2Ctl rs!B_picture, Picture3
                Else
                    Picture3.Picture = Nothing
                End If

End Sub
