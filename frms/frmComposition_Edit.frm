VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmComposition_Edit 
   Caption         =   "白坯构成"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10995
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
   ScaleHeight     =   7815
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10995
      _LayoutVersion  =   1
      _ExtentX        =   19394
      _ExtentY        =   13785
      _DataPath       =   ""
      Bands           =   "frmComposition_Edit.frx":0000
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   7815
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   10995
         _cx             =   19394
         _cy             =   13785
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
         Align           =   5
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
         _GridInfo       =   $"frmComposition_Edit.frx":01C8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.Frame Frame1 
            Caption         =   "原料"
            Height          =   3705
            Left            =   90
            TabIndex        =   31
            Top             =   90
            Width           =   10815
            Begin XtremeSuiteControls.CheckBox CheckBox1 
               Height          =   255
               Left            =   4800
               TabIndex        =   32
               Top             =   2580
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PushButton5 
               Height          =   555
               Left            =   8400
               TabIndex        =   33
               Top             =   3000
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   979
               _StockProps     =   79
               Caption         =   "退出"
               UseVisualStyle  =   -1  'True
               ImageGap        =   15
               IconWidth       =   16
               Icon            =   "frmComposition_Edit.frx":024E
            End
            Begin XtremeSuiteControls.PushButton PushButton4 
               Height          =   555
               Left            =   6510
               TabIndex        =   34
               Top             =   3000
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   979
               _StockProps     =   79
               Caption         =   "保存"
               UseVisualStyle  =   -1  'True
               ImageGap        =   15
               IconWidth       =   16
               Icon            =   "frmComposition_Edit.frx":072E
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               Height          =   375
               Left            =   4740
               TabIndex        =   35
               Top             =   360
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               Enabled         =   0   'False
               BackColor       =   14737632
               Locked          =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit2 
               Height          =   375
               Left            =   4740
               TabIndex        =   36
               Top             =   1080
               Width           =   1395
               _Version        =   1048578
               _ExtentX        =   2461
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               Enabled         =   0   'False
            End
            Begin XtremeSuiteControls.PushButton PushButton2 
               Height          =   375
               Left            =   6120
               TabIndex        =   37
               Top             =   1080
               Width           =   375
               _Version        =   1048578
               _ExtentX        =   661
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit3 
               Height          =   375
               Left            =   1380
               TabIndex        =   38
               Top             =   1800
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.ComboBox ComboBox1 
               Height          =   315
               Left            =   1380
               TabIndex        =   39
               Top             =   390
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit4 
               Height          =   375
               Left            =   8220
               TabIndex        =   40
               Top             =   1080
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit6 
               Height          =   375
               Left            =   8220
               TabIndex        =   41
               Top             =   360
               Width           =   1515
               _Version        =   1048578
               _ExtentX        =   2672
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit7 
               Height          =   375
               Left            =   4740
               TabIndex        =   42
               Top             =   1800
               Width           =   1395
               _Version        =   1048578
               _ExtentX        =   2461
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               Enabled         =   0   'False
            End
            Begin XtremeSuiteControls.PushButton PushButton3 
               Height          =   615
               Left            =   1380
               TabIndex        =   43
               Top             =   960
               Width           =   1515
               _Version        =   1048578
               _ExtentX        =   2672
               _ExtentY        =   1085
               _StockProps     =   79
               Caption         =   "选择库存"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PushButton11 
               Height          =   375
               Left            =   6120
               TabIndex        =   44
               Top             =   1800
               Width           =   375
               _Version        =   1048578
               _ExtentX        =   661
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PushButton13 
               Height          =   555
               Left            =   4620
               TabIndex        =   45
               Top             =   3000
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   979
               _StockProps     =   79
               Caption         =   "保存并新增"
               UseVisualStyle  =   -1  'True
               ImageGap        =   15
               IconWidth       =   16
               Icon            =   "frmComposition_Edit.frx":0C0E
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit14 
               Height          =   375
               Left            =   1380
               TabIndex        =   46
               Top             =   2520
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit17 
               Height          =   375
               Left            =   8220
               TabIndex        =   47
               Top             =   1800
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.CheckBox CheckBox3 
               Height          =   375
               Left            =   9720
               TabIndex        =   58
               Top             =   360
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "不定重"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   255
               Left            =   3960
               TabIndex        =   57
               Top             =   420
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "订单号:"
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   315
               Left            =   3960
               TabIndex        =   56
               Top             =   1110
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "品    名:"
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label3 
               Height          =   255
               Left            =   480
               TabIndex        =   55
               Top             =   1860
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "原订单号:"
            End
            Begin XtremeSuiteControls.Label Label4 
               Height          =   255
               Left            =   7200
               TabIndex        =   54
               Top             =   1140
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "规   格:"
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Left            =   480
               TabIndex        =   53
               Top             =   420
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "入库类型:"
            End
            Begin XtremeSuiteControls.Label Label6 
               Height          =   195
               Left            =   7200
               TabIndex        =   52
               Top             =   450
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "公    斤:"
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label7 
               Height          =   255
               Left            =   3960
               TabIndex        =   51
               Top             =   1860
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "供应商:"
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label16 
               Height          =   255
               Left            =   480
               TabIndex        =   50
               Top             =   2580
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "备        注:"
            End
            Begin XtremeSuiteControls.Label Label19 
               Height          =   255
               Left            =   7200
               TabIndex        =   49
               Top             =   1860
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "单   价:"
            End
            Begin XtremeSuiteControls.Label Label20 
               Height          =   255
               Left            =   3960
               TabIndex        =   48
               Top             =   2580
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "补    单:"
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "白坯"
            Height          =   3870
            Left            =   90
            TabIndex        =   2
            Top             =   3855
            Width           =   10815
            Begin XtremeSuiteControls.CheckBox CheckBox2 
               Height          =   375
               Left            =   8280
               TabIndex        =   3
               Top             =   2640
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
               _ExtentY        =   661
               _StockProps     =   79
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit5 
               Height          =   375
               Left            =   1380
               TabIndex        =   4
               Top             =   1920
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.ComboBox ComboBox2 
               Height          =   315
               Left            =   1320
               TabIndex        =   5
               Top             =   420
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.PushButton PushButton6 
               Height          =   615
               Left            =   1320
               TabIndex        =   6
               Top             =   990
               Width           =   1515
               _Version        =   1048578
               _ExtentX        =   2672
               _ExtentY        =   1085
               _StockProps     =   79
               Caption         =   "选择库存"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit8 
               Height          =   375
               Left            =   4740
               TabIndex        =   7
               Top             =   390
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               Enabled         =   0   'False
               BackColor       =   14737632
               Locked          =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit9 
               Height          =   375
               Left            =   4740
               TabIndex        =   8
               Top             =   1140
               Width           =   1395
               _Version        =   1048578
               _ExtentX        =   2461
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               Enabled         =   0   'False
            End
            Begin XtremeSuiteControls.PushButton PushButton7 
               Height          =   375
               Left            =   6120
               TabIndex        =   9
               Top             =   1155
               Width           =   375
               _Version        =   1048578
               _ExtentX        =   661
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit10 
               Height          =   375
               Left            =   4740
               TabIndex        =   10
               Top             =   1920
               Width           =   1395
               _Version        =   1048578
               _ExtentX        =   2461
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               Enabled         =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit11 
               Height          =   375
               Left            =   8160
               TabIndex        =   11
               Top             =   1155
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit12 
               Height          =   375
               Left            =   8160
               TabIndex        =   12
               Top             =   390
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit13 
               Height          =   375
               Left            =   8160
               TabIndex        =   13
               Top             =   1920
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.PushButton PushButton12 
               Height          =   375
               Left            =   6120
               TabIndex        =   14
               Top             =   1920
               Width           =   375
               _Version        =   1048578
               _ExtentX        =   661
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PushButton9 
               Height          =   555
               Left            =   8400
               TabIndex        =   15
               Top             =   3240
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   979
               _StockProps     =   79
               Caption         =   "退出"
               UseVisualStyle  =   -1  'True
               ImageGap        =   15
               IconWidth       =   16
               Icon            =   "frmComposition_Edit.frx":10EE
            End
            Begin XtremeSuiteControls.PushButton PushButton10 
               Height          =   555
               Left            =   6480
               TabIndex        =   16
               Top             =   3240
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   979
               _StockProps     =   79
               Caption         =   "保存"
               UseVisualStyle  =   -1  'True
               ImageGap        =   15
               IconWidth       =   16
               Icon            =   "frmComposition_Edit.frx":15CE
            End
            Begin XtremeSuiteControls.PushButton PushButton14 
               Height          =   555
               Left            =   4560
               TabIndex        =   17
               Top             =   3240
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   979
               _StockProps     =   79
               Caption         =   "保存并新增"
               UseVisualStyle  =   -1  'True
               ImageGap        =   15
               IconWidth       =   16
               Icon            =   "frmComposition_Edit.frx":1AAE
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit15 
               Height          =   375
               Left            =   4740
               TabIndex        =   18
               Top             =   2640
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit16 
               Height          =   375
               Left            =   1380
               TabIndex        =   19
               Top             =   2640
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.Label Label8 
               Height          =   255
               Left            =   480
               TabIndex        =   30
               Top             =   450
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "入库类型:"
            End
            Begin XtremeSuiteControls.Label Label9 
               Height          =   255
               Left            =   480
               TabIndex        =   29
               Top             =   1980
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "原订单号:"
            End
            Begin XtremeSuiteControls.Label Label10 
               Height          =   255
               Left            =   3900
               TabIndex        =   28
               Top             =   1980
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "供应商:"
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label11 
               Height          =   315
               Left            =   3900
               TabIndex        =   27
               Top             =   1185
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "品    名:"
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label12 
               Height          =   255
               Left            =   3900
               TabIndex        =   26
               Top             =   450
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "订单号:"
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label13 
               Height          =   195
               Left            =   7200
               TabIndex        =   25
               Top             =   480
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "数    量:"
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label14 
               Height          =   255
               Left            =   7200
               TabIndex        =   24
               Top             =   1215
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "门    幅:"
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label15 
               Height          =   255
               Left            =   7200
               TabIndex        =   23
               Top             =   1980
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "克    重:"
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label17 
               Height          =   255
               Left            =   3840
               TabIndex        =   22
               Top             =   2700
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "备        注:"
            End
            Begin XtremeSuiteControls.Label Label18 
               Height          =   255
               Left            =   480
               TabIndex        =   21
               Top             =   2700
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "单        价:"
            End
            Begin XtremeSuiteControls.Label Label21 
               Height          =   255
               Left            =   7200
               TabIndex        =   20
               Top             =   2700
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "补    单:"
            End
         End
      End
   End
End
Attribute VB_Name = "frmComposition_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rss As RecordSet
Public itemid As String
Public Whiteitemid As String
Public id As String
'原料品名id
Public OriginalProduct As String
'原料客户id
Public Originalsuppliers As String
'白坯品名id
Public whiteProduct As String
'白坯客户id
Public producerid As String

Private rsgrid5 As RecordSet

Public Property Set Rs5(ByRef vData As RecordSet)
    Set rsgrid5 = vData

End Property

Private Sub CheckBox3_Click()
    If CheckBox3.Value = 1 Then
'        CheckBox3.Value = 1
        FlatEdit6.BackColor = &H8000000F
        FlatEdit6.Enabled = False
        FlatEdit6.Text = ""
    Else
'        CheckBox3.Value = 0
        FlatEdit6.BackColor = &H80000005
        FlatEdit6.Enabled = True
    End If
End Sub

'Private Compositionid As String
'Private Whiteid As String
'Private Clientid As String
'Private WhiteClientid As String

Private Sub ComboBox1_Click()
    If ComboBox1.Text = "采购" Then
        Original
       
    Else
        original_1
    End If
End Sub
Private Sub ComboBox2_Click()
    If ComboBox2.Text = "采购" Or ComboBox2.Text = "加工" Then
        white
    Else
        White_1
    End If
    
End Sub

Private Sub Form_Load()
    InitFrm
    Storageway
    WhiteStorageway
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
End Sub

Private Sub Storageway()
    ComboBox1.AddItem "采购"
    ComboBox1.AddItem "调拨"
    ComboBox1.ListIndex = 0
End Sub
Private Sub WhiteStorageway()
    ComboBox2.AddItem "采购"
    ComboBox2.AddItem "调拨"
    ComboBox2.AddItem "加工"
    ComboBox2.ListIndex = 0
End Sub
'选择订单号
Private Sub PushButton1_Click()
    Dim frm1 As New frmpopupWhiteDingdan
    frm1.id = id
    frm1.Show vbModal
    If frm1.bsaved = False Then
        Exit Sub
    End If
    FlatEdit1.Text = frm1.departmentid
    Unload frm1
End Sub



Private Sub PushButton14_Click()
           If yanzhenWhiteComposition(id, itemid) = False Then
             Exit Sub
        End If
       If ComboBox2.Text = "采购" Or ComboBox2.Text = "加工" Then
            If Trim(FlatEdit8.Text) = "" Then
                MsgBox "订单号不能为空", vbInformation, "提示"
                Exit Sub
            End If
        End If
         If Trim(FlatEdit9.Text) = "" Then
            MsgBox "品名不能为空", vbInformation, "提示"
            Exit Sub
        End If
         If Trim(FlatEdit10.Text) = "" Then
            MsgBox "供应商不能为空", vbInformation, "提示"
              Exit Sub
        End If
         If Trim(FlatEdit13.Text) = "" Then
            MsgBox "克重不能为空", vbInformation, "提示"
            Exit Sub
        End If
        If Trim(FlatEdit12.Text) = "" Then
            MsgBox "数量不能为空", vbInformation, "提示"
            Exit Sub
        End If
        If ComboBox2.Text = "调拨" Then
            If Trim(FlatEdit5.Text) = "" Then
               MsgBox "原订单号不能为空", vbInformation, "提示"
               Exit Sub
            End If
        End If
     SaveWhite
     createdate (id)
      rsgrid5.requery
        ComboBox2.Clear
'        FlatEdit8.Text = ""
        FlatEdit9.Text = ""
        FlatEdit10.Text = ""
        FlatEdit5.Text = ""
        FlatEdit13.Text = ""
        FlatEdit11.Text = ""
       WhiteStorageway
       Whiteitemid = ""
End Sub

'原料选择库存
Private Sub PushButton3_Click()
    Dim sql As String
    Dim rs As New RecordSet
    If ComboBox1.Text = "调拨" Then
        Dim frm1 As New frmpopupOriginal
        frm1.Show vbModal
        If frm1.bsaved = False Then
            Exit Sub
        End If
        
        FlatEdit4.Text = frm1.specifications
        FlatEdit3.Text = frm1.OrderCode
        FlatEdit2.Text = frm1.WhiteName
        FlatEdit7.Text = frm1.ClientName
        OriginalProduct = frm1.GoodsID
        Originalsuppliers = frm1.supplier
        
        If Originalsuppliers = "" Then
            sql = "select * from G_ContactCompany where B_ClientID='kx1'"
            rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            Originalsuppliers = rs!B_Clientid
            FlatEdit7.Text = rs!B_ClientName
        End If
        
        Unload frm1
    End If
End Sub

'选择库存
Private Sub PushButton6_Click()
      If ComboBox2.Text = "调拨" Then
          Dim frm1 As New frmpopupWhiteStock
          frm1.Show vbModal
          
          If frm1.theSaved = False Then
            Exit Sub
          End If

          FlatEdit9.Text = frm1.WhiteName
          FlatEdit11.Text = frm1.Widthid
          FlatEdit13.Text = frm1.UnitWeight
          FlatEdit5.Text = frm1.OrderCode
          FlatEdit10.Text = frm1.ProducerName
          whiteProduct = frm1.GoodsID
          producerid = frm1.producer
          Unload frm1
    End If
End Sub

'选择品名
Private Sub PushButton2_Click()
            Dim frm1 As New frmpopupComposition
            frm1.Show vbModal
            FlatEdit2.Text = frm1.CompositionName
            OriginalProduct = frm1.Compositionid
            Unload frm1
End Sub
'选择供应商
Private Sub PushButton11_Click()
            Dim frm1 As New frmPopupDanWei
            frm1.ContactType = "原料供应商"
            frm1.Show vbModal
            Originalsuppliers = frm1.Clientid
            FlatEdit7.Text = frm1.ClientName
            Unload frm1
End Sub
Private Sub PushButton12_Click()
            Dim frm1 As New frmPopupDanWei
            frm1.ContactType = "白坯加工商"
            frm1.Show vbModal
            producerid = frm1.Clientid
            FlatEdit10.Text = frm1.ClientName
            Unload frm1
End Sub

Private Sub PushButton4_Click()
        If yanzhenWhiteComposition(id, itemid) = False Then
             Exit Sub
        End If
        If ComboBox1.Text = "采购" Then
            If Trim(FlatEdit1.Text) = "" Then
                MsgBox "订单号不能为空", vbInformation, "提示"
                Exit Sub
            End If
        End If
         If Trim(FlatEdit2.Text) = "" Then
            MsgBox "品名不能为空", vbInformation, "提示"
            Exit Sub
        End If
        If ComboBox1.Text = "调拨" Then
             If Trim(FlatEdit3.Text) = "" Then
                MsgBox "原订单号不能为空", vbInformation, "提示"
                Exit Sub
            End If
        End If
        If ComboBox1.Text <> "调拨" Then
             If Trim(FlatEdit4.Text) = "" Then
                MsgBox "规格不能为空", vbInformation, "提示"
                Exit Sub
            End If
        End If
        
        If CheckBox3.Value = 0 Then
            If FlatEdit6.Text = "" Or Val(FlatEdit6.Text) <= 0 Then
                MsgBox "数量不能为空或要大于0", vbInformation, "提示"
                  Exit Sub
            End If
        Else
            FlatEdit6.Text = ""
        End If
        
        If Trim(FlatEdit7.Text) = "" Then
            MsgBox "供应商不能为空", vbInformation, "提示"
            Exit Sub
        End If
     
     Saveoriginal
     createdate (id)
     '刷新上级窗体的记录集
     rsgrid5.requery
     Me.Hide
End Sub

Private Sub PushButton5_Click()
    Unload Me
End Sub
Private Sub PushButton10_Click()
        If yanzhenWhiteComposition(id, itemid) = False Then
             Exit Sub
        End If
       If ComboBox2.Text = "采购" Or ComboBox2.Text = "加工" Then
            If Trim(FlatEdit8.Text) = "" Then
                MsgBox "订单号不能为空", vbInformation, "提示"
                Exit Sub
            End If
        End If
         If Trim(FlatEdit9.Text) = "" Then
            MsgBox "品名不能为空", vbInformation, "提示"
            Exit Sub
        End If
         If Trim(FlatEdit10.Text) = "" Then
            MsgBox "供应商不能为空", vbInformation, "提示"
              Exit Sub
        End If
        If ComboBox2.Text <> "调拨" Then
             If Trim(FlatEdit13.Text) = "" Then
                MsgBox "克重不能为空", vbInformation, "提示"
                Exit Sub
            End If
        End If
        If Trim(FlatEdit12.Text) = "" Or Val(FlatEdit12.Text) <= 0 Then
            MsgBox "数量不能为空或要大于0", vbInformation, "提示"
            Exit Sub
        End If
        If ComboBox2.Text = "调拨" Then
            If Trim(FlatEdit5.Text) = "" Then
               MsgBox "原订单号不能为空", vbInformation, "提示"
               Exit Sub
            End If
        End If
     SaveWhite
     createdate (id)
      rsgrid5.requery
      Me.Hide
End Sub
'白坯品名
Private Sub PushButton7_Click()
        Dim frm1 As New frmpopupWhite
        frm1.Show vbModal
        whiteProduct = frm1.Whiteid
        FlatEdit9.Text = frm1.WhiteName
        Unload frm1
End Sub

Private Sub PushButton8_Click()
    Dim frm1 As New frmpopupWhiteDingdan
    frm1.id = id
    frm1.Show vbModal
    If frm1.bsaved = False Then
        Exit Sub
    End If
    FlatEdit8.Text = frm1.departmentid
    Unload frm1
End Sub

Private Sub PushButton9_Click()
    Unload Me
End Sub
'原料中采购方式的切换
Private Sub Original()
'        FlatEdit2.Enabled = True
    FlatEdit3.Enabled = True
    FlatEdit4.Enabled = True
'    FlatEdit7.Enabled = True
    FlatEdit3.Enabled = False
    PushButton3.Enabled = False
    FlatEdit2.BackColor = &H80000005
    FlatEdit3.BackColor = &H80000005
    FlatEdit4.BackColor = &H80000005
    FlatEdit7.BackColor = &H80000005
    FlatEdit17.Locked = False
     FlatEdit3.Text = ""
     PushButton2.Enabled = True
     PushButton11.Enabled = True
End Sub
Private Sub original_1()
    FlatEdit2.Enabled = False
    FlatEdit3.Enabled = False
    FlatEdit4.Enabled = False
    FlatEdit7.Enabled = False
    PushButton3.Enabled = True
    FlatEdit2.BackColor = &H80000005
    FlatEdit3.BackColor = &H80000005
    FlatEdit4.BackColor = &H80000005
    FlatEdit7.BackColor = &H80000005
    FlatEdit17.Locked = True
    PushButton2.Enabled = False
    PushButton11.Enabled = False
End Sub
'白坯中采购方式的切换
Private Sub white()
'     FlatEdit5.Enabled = True
'    FlatEdit9.Enabled = True
'    FlatEdit10.Enabled = True
    FlatEdit11.Enabled = True
    FlatEdit13.Enabled = True
    FlatEdit5.Enabled = False
    PushButton6.Enabled = False
    FlatEdit5.BackColor = &H80000005
    FlatEdit9.BackColor = &H80000005
    FlatEdit10.BackColor = &H80000005
    FlatEdit11.BackColor = &H80000005
    FlatEdit13.BackColor = &H80000005
    FlatEdit16.Locked = False
    FlatEdit5.Text = ""
    PushButton7.Enabled = True
    PushButton12.Enabled = True
End Sub
Private Sub White_1()
    FlatEdit5.Enabled = False
    FlatEdit9.Enabled = False
    FlatEdit10.Enabled = False
    FlatEdit11.Enabled = False
    FlatEdit13.Enabled = False
    PushButton6.Enabled = True
    FlatEdit5.BackColor = &H80000005
    FlatEdit9.BackColor = &H80000005
    FlatEdit10.BackColor = &H80000005
    FlatEdit11.BackColor = &H80000005
    FlatEdit13.BackColor = &H80000005
    FlatEdit16.Locked = True
    PushButton7.Enabled = False
    PushButton12.Enabled = False
End Sub

Private Sub Saveoriginal()
        If Len(Trim(itemid)) > 0 Then
            Saveoriginal_update
        Else
            Saveoriginal_Edit
        End If
End Sub
Private Sub Saveoriginal_Edit()
       Set rss = New RecordSet
       Dim sql As String
       Dim sql1 As String
       Dim sql2 As String
       Dim a As String
       Dim b As String
       Dim c As String
       '草稿表中获取主键id
       sql = "select * from G_DraftWhiteComposition where 1=0"
       rss.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        rss.addnew
        rss!B_creatime = Now
        rss.Update
       itemid = rss!B_itemid
        a = ""
        b = ""
        c = ""
       '进行正式表数据添加
        If ComboBox1.Text = "采购" Then
            sql1 = "exec usp_Saveoriginal '" & itemid & "','" & Frame1.Caption & "','" & FlatEdit1.Text & "','" & id & "','" & OriginalProduct & "',"
            sql1 = sql1 & "'" & Originalsuppliers & "','" & FlatEdit4.Text & "','" & c & "','" & ComboBox1.Text & "','" & a & "','" & b & "','" & FlatEdit6.Text & "','" & Gm.SysID.SystemUser & "','" & FlatEdit14.Text & "','" & FlatEdit17.Text & "','" & CheckBox1.Value & "','" & CheckBox3.Value & "'"
            Gm.cnnTool.cnn.Execute sql1
       Else
'            sql1 = "exec usp_Saveoriginal '" & itemid & "','" & Frame1.Caption & "','" & a & "','" & id & "','" & OriginalProduct & "',"
'            sql1 = sql1 & "'" & b & "','" & FlatEdit4.Text & "','" & c & "','" & ComboBox1.Text & "','" & FlatEdit3.Text & "','" & Originalsuppliers & "','" & FlatEdit6.Text & "','" & Gm.SysID.SystemUser & "','" & FlatEdit14.Text & "','" & FlatEdit17.Text & "','" & CheckBox1.Value & "','" & CheckBox3.Value & "'"
'            Debug.Print sql1
'            Gm.cnnTool.cnn.Execute sql1
            sql1 = "exec usp_Saveoriginal '" & itemid & "','" & Frame1.Caption & "','" & FlatEdit1.Text & "','" & id & "','" & OriginalProduct & "',"
            sql1 = sql1 & "'" & b & "','" & FlatEdit4.Text & "','" & c & "','" & ComboBox1.Text & "','" & FlatEdit3.Text & "','" & Originalsuppliers & "','" & FlatEdit6.Text & "','" & Gm.SysID.SystemUser & "','" & FlatEdit14.Text & "','" & FlatEdit17.Text & "','" & CheckBox1.Value & "','" & CheckBox3.Value & "'"
            Debug.Print sql1
            Gm.cnnTool.cnn.Execute sql1
       End If
       '删除明细表
       sql2 = "delete from G_DraftWhiteComposition where B_itemid='" & itemid & "'"
       Gm.cnnTool.cnn.Execute sql2
'       MsgBox "保存成功", vbInformation, "提示"
End Sub
Private Sub PushButton13_Click()
          If yanzhenWhiteComposition(id, itemid) = False Then
             Exit Sub
        End If
        If ComboBox1.Text = "采购" Then
            If Trim(FlatEdit1.Text) = "" Then
                MsgBox "订单号不能为空", vbInformation, "提示"
                Exit Sub
            End If
        End If
         If Trim(FlatEdit2.Text) = "" Then
            MsgBox "品名不能为空", vbInformation, "提示"
            Exit Sub
        End If
        If ComboBox1.Text = "调拨" Then
             If Trim(FlatEdit3.Text) = "" Then
                MsgBox "原订单号不能为空", vbInformation, "提示"
                Exit Sub
            End If
        End If
         If Trim(FlatEdit4.Text) = "" Then
            MsgBox "规格不能为空", vbInformation, "提示"
            Exit Sub
        End If
       If CheckBox3.Value = 0 Then
            If Trim(FlatEdit6.Text) = "" Or FlatEdit6.Text <= 0 Then
                MsgBox "数量不能为空或要大于0", vbInformation, "提示"
                  Exit Sub
            End If
        Else
            FlatEdit6.Text = ""
        End If
        If Trim(FlatEdit7.Text) = "" Then
            MsgBox "供应商不能为空", vbInformation, "提示"
            Exit Sub
        End If
     
     Saveoriginal
     createdate (id)
     '刷新上级窗体的记录集
     rsgrid5.requery
        ComboBox1.Clear
'        FlatEdit1.Text = ""
        FlatEdit2.Text = ""
        FlatEdit3.Text = ""
        FlatEdit4.Text = ""
        FlatEdit6.Text = ""
        FlatEdit7.Text = ""
       Storageway
       itemid = ""
End Sub
Private Sub Saveoriginal_update()
        Dim sql As String
        Dim sql1 As String
        Dim a As String
        Dim b As String
        Dim c As String
        a = ""
        b = ""
        c = ""
        If ComboBox1.Text = "采购" Then
            sql = "exec usp_Saveoriginalupdate '" & itemid & "','" & Frame1.Caption & "','" & FlatEdit1.Text & "','" & OriginalProduct & "',"
            sql = sql & "'" & Originalsuppliers & "','" & FlatEdit4.Text & "','" & c & "','" & ComboBox1.Text & "','" & a & "','" & b & "','" & FlatEdit6.Text & "','" & FlatEdit14.Text & "','" & FlatEdit17.Text & "','" & CheckBox1.Value & "','" & CheckBox3.Value & "'"
            Debug.Print sql
            Gm.cnnTool.cnn.Execute sql
          
        Else
'            sql1 = "exec usp_Saveoriginalupdate '" & itemid & "','" & Frame1.Caption & "','" & a & "','" & OriginalProduct & "',"
'            sql1 = sql1 & "'" & b & "','" & FlatEdit4.Text & "','" & c & "','" & ComboBox1.Text & "','" & FlatEdit3.Text & "','" & Originalsuppliers & "','" & FlatEdit6.Text & "','" & FlatEdit14.Text & "','" & FlatEdit17.Text & "','" & CheckBox1.Value & "','" & CheckBox3.Value & "'"
'            Gm.cnnTool.cnn.Execute sql1

            sql1 = "exec usp_Saveoriginalupdate '" & itemid & "','" & Frame1.Caption & "','" & FlatEdit1.Text & "','" & OriginalProduct & "',"
            sql1 = sql1 & "'" & b & "','" & FlatEdit4.Text & "','" & c & "','" & ComboBox1.Text & "','" & FlatEdit3.Text & "','" & Originalsuppliers & "','" & FlatEdit6.Text & "','" & FlatEdit14.Text & "','" & FlatEdit17.Text & "','" & CheckBox1.Value & "','" & CheckBox3.Value & "'"
            Gm.cnnTool.cnn.Execute sql1
        End If
'        MsgBox "修改成功", vbInformation, "提示"
        rsgrid5.requery
End Sub
'保存白坯的
Private Sub SaveWhite()
        If Len(Trim(Whiteitemid)) > 0 Then
            SaveWhite_update
        Else
            SaveWhite_Edit
        End If
End Sub
Private Sub SaveWhite_Edit()
         Set rss = New RecordSet
       Dim sql As String
       Dim sql1 As String
       Dim sql2 As String
       Dim a As String
       Dim b As String
       '草稿表中获取主键id
       sql = "select * from G_DraftWhiteComposition where 1=0"
       rss.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        rss.addnew
        rss!B_creatime = Now
        rss.Update
       Whiteitemid = rss!B_itemid
        a = ""
        b = ""
       '进行正式表数据添加
        If ComboBox2.Text = "采购" Or ComboBox2.Text = "加工" Then
            sql1 = "exec usp_Saveoriginal '" & Whiteitemid & "','" & Frame2.Caption & "','" & FlatEdit8.Text & "','" & id & "','" & whiteProduct & "',"
            sql1 = sql1 & "'" & producerid & "','" & FlatEdit11.Text & "','" & FlatEdit13.Text & "','" & ComboBox2.Text & "','" & a & "','" & b & "','" & FlatEdit12.Text & "','" & Gm.SysID.SystemUser & "','" & FlatEdit15.Text & "','" & FlatEdit16.Text & "','" & CheckBox2.Value & "',''"
            Debug.Print sql1
            Gm.cnnTool.cnn.Execute sql1
       Else
'            sql1 = "exec usp_Saveoriginal '" & Whiteitemid & "','" & Frame2.Caption & "','" & a & "','" & id & "','" & whiteProduct & "',"
'            sql1 = sql1 & "'" & b & "','" & FlatEdit11.Text & "','" & FlatEdit13.Text & "','" & ComboBox2.Text & "','" & FlatEdit5.Text & "','" & producerid & "','" & FlatEdit12.Text & "','" & Gm.SysID.SystemUser & "','" & FlatEdit15.Text & "','" & FlatEdit16.Text & "','" & CheckBox2.Value & "',''"
'            Gm.cnnTool.cnn.Execute sql1
            sql1 = "exec usp_Saveoriginal '" & Whiteitemid & "','" & Frame2.Caption & "','" & FlatEdit8.Text & "','" & id & "','" & whiteProduct & "',"
            sql1 = sql1 & "'" & b & "','" & FlatEdit11.Text & "','" & FlatEdit13.Text & "','" & ComboBox2.Text & "','" & FlatEdit5.Text & "','" & producerid & "','" & FlatEdit12.Text & "','" & Gm.SysID.SystemUser & "','" & FlatEdit15.Text & "','" & FlatEdit16.Text & "','" & CheckBox2.Value & "',''"
            Gm.cnnTool.cnn.Execute sql1
            
       End If
       '删除明细表
       sql2 = "delete from G_DraftWhiteComposition where B_itemid='" & Whiteitemid & "'"
       Gm.cnnTool.cnn.Execute sql2
'       MsgBox "保存成功", vbInformation, "提示"
End Sub
Private Sub SaveWhite_update()
        Dim sql As String
        Dim sql1 As String
        Dim a As String
        Dim b As String
        a = ""
        b = ""
       If ComboBox2.Text = "采购" Or ComboBox2.Text = "加工" Then
            sql = "exec usp_Saveoriginalupdate '" & Whiteitemid & "','" & Frame2.Caption & "','" & FlatEdit8.Text & "','" & whiteProduct & "',"
            sql = sql & "'" & producerid & "','" & FlatEdit11.Text & "','" & FlatEdit13.Text & "','" & ComboBox2.Text & "','" & a & "','" & b & "','" & FlatEdit12.Text & "','" & FlatEdit15.Text & "','" & FlatEdit16.Text & "','" & CheckBox2.Value & "',''"
            Gm.cnnTool.cnn.Execute sql
       Else
'            sql1 = "exec usp_Saveoriginalupdate '" & Whiteitemid & "','" & Frame2.Caption & "','" & a & "','" & id & "','" & whiteProduct & "',"
'            sql1 = sql1 & "'" & b & "','" & FlatEdit11.Text & "','" & FlatEdit13.Text & "','" & ComboBox2.Text & "','" & FlatEdit5.Text & "','" & producerid & "','" & FlatEdit12.Text & "','" & FlatEdit15.Text & "','" & FlatEdit16.Text & "','" & CheckBox2.Value & "',''"
'            Gm.cnnTool.cnn.Execute sql1

            sql1 = "exec usp_Saveoriginalupdate '" & Whiteitemid & "','" & Frame2.Caption & "','" & FlatEdit8.Text & "','" & id & "','" & whiteProduct & "',"
            sql1 = sql1 & "'" & b & "','" & FlatEdit11.Text & "','" & FlatEdit13.Text & "','" & ComboBox2.Text & "','" & FlatEdit5.Text & "','" & producerid & "','" & FlatEdit12.Text & "','" & FlatEdit15.Text & "','" & FlatEdit16.Text & "','" & CheckBox2.Value & "',''"
            Gm.cnnTool.cnn.Execute sql1
       End If
'        MsgBox "修改成功", vbInformation, "提示"
        rsgrid5.requery
End Sub

Private Sub FlatEdit6_KeyPress(KeyAscii As Integer)
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
Private Sub FlatEdit17_KeyPress(KeyAscii As Integer)
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

'当订单号合同中订单号都有开始计算时间
Private Sub createdate(ByVal id As String)
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
    sql = "select distinct B_itemidb from G_WhiteComposition where B_id='" & id & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    sql1 = "select distinct B_orderCode from G_Billdetailorder where B_id='" & id & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    sql2 = "select *  from G_WhiteComposition where B_id='" & id & "'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount = rs1.RecordCount Then
        Do While Not rs2.EOF
            rs2!B_Date = Now
            rs2.movenext
        Loop
    End If
End Sub
Public Function yanzhenWhiteComposition(ByVal theid As String, ByVal itemid As String) As Boolean
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
    Dim sql3 As String
    Dim rs3 As New RecordSet
    
    sql = "Select * From G_SystemUser where B_UserName='" & Gm.SysID.SystemUser & "'"
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs!B_SuperAdmin = 1 Then
        yanzhenWhiteComposition = True
        Exit Function
    End If
    
    sql1 = "select * from G_WhiteComposition where B_ID='" & theid & "'and B_itemid='" & itemid & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    
    sql2 = "select distinct B_UserName from G_WhiteComposition where B_ID='" & theid & "'and B_itemid='" & itemid & "'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    Debug.Print sql2
    If rs2.RecordCount > 0 Then
        If rs2!B_UserName = Gm.SysID.SystemUser Then
            yanzhenWhiteComposition = True
        Else
            yanzhenWhiteComposition = False
            MsgBox "不是本制单人不能修改", vbInformation, "提示"
            Exit Function
        End If
    Else
        yanzhenWhiteComposition = True
    End If
    If rs1.RecordCount > 0 Then
        If IIf(IsNull(rs1!B_Date), "", rs1!B_Date) <> "" Then
            sql3 = "SELECT B_value FROM G_Config_OneInt WHERE B_groupname='织造系统_合同构成修改时间'"
            rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            
            If DateDiff("s", rs1!B_Date, Now) > rs3!B_Value Then
                yanzhenWhiteComposition = False
                MsgBox "已经超过制作本单据的时间不能进行修改", vbInformation, "提示"
            Else
                yanzhenWhiteComposition = True
            End If
        End If
    End If
End Function


