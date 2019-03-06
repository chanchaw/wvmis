VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmOrder_Edit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置明细"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOrder_Edit.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   11475
   StartUpPosition =   2  '屏幕中心
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7515
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11475
      _LayoutVersion  =   1
      _ExtentX        =   20241
      _ExtentY        =   13256
      _DataPath       =   ""
      Bands           =   "frmOrder_Edit.frx":038A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6435
         Left            =   480
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   840
         Width           =   10395
         _cx             =   18336
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
         _GridInfo       =   $"frmOrder_Edit.frx":1136
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   6255
            Left            =   90
            ScaleHeight     =   6255
            ScaleWidth      =   10215
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   90
            Width           =   10215
            Begin VB.PictureBox Picture2 
               Height          =   375
               Left            =   4740
               ScaleHeight     =   315
               ScaleWidth      =   1515
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   2520
               Width           =   1575
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit12 
               Height          =   375
               Left            =   4740
               TabIndex        =   29
               Top             =   4080
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
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   6300
               TabIndex        =   23
               Top             =   2160
               Width           =   375
               _Version        =   1048578
               _ExtentX        =   661
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1560
               TabIndex        =   0
               Top             =   1260
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
               Left            =   4740
               TabIndex        =   15
               Top             =   1260
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit3 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   7920
               TabIndex        =   16
               Top             =   1260
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
               Left            =   1560
               TabIndex        =   17
               Top             =   2160
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
               Left            =   7920
               TabIndex        =   18
               Top             =   2160
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit6 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1560
               TabIndex        =   19
               Top             =   3120
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit8 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1560
               TabIndex        =   20
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit10 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   7920
               TabIndex        =   21
               Top             =   4980
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit11 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   4680
               TabIndex        =   22
               Top             =   2160
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit7 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   7920
               TabIndex        =   24
               Top             =   3120
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
               Left            =   7920
               TabIndex        =   25
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
               Height          =   390
               Left            =   4740
               TabIndex        =   27
               Top             =   3105
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit13 
               Height          =   975
               Left            =   1560
               TabIndex        =   31
               Top             =   5760
               Width           =   5115
               _Version        =   1048578
               _ExtentX        =   9022
               _ExtentY        =   1720
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               MultiLine       =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit14 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1560
               TabIndex        =   34
               Top             =   4980
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit15 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   4740
               TabIndex        =   35
               Top             =   4980
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
               Left            =   3180
               TabIndex        =   37
               Top             =   600
               Width           =   375
               _Version        =   1048578
               _ExtentX        =   661
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit16 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1560
               TabIndex        =   38
               Top             =   600
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
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
               Left            =   4740
               TabIndex        =   41
               Top             =   600
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
            Begin XtremeSuiteControls.Label Label19 
               Height          =   255
               Left            =   3720
               TabIndex        =   42
               Top             =   660
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "款号："
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
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "上级页面中选定客户后“源订单号”才有数据显示"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   4200
               TabIndex        =   40
               Top             =   45
               Width           =   4980
            End
            Begin XtremeSuiteControls.Label Label17 
               Height          =   255
               Left            =   360
               TabIndex        =   39
               Top             =   660
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
            Begin XtremeSuiteControls.Label Label16 
               Height          =   255
               Left            =   3720
               TabIndex        =   33
               Top             =   5040
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "中样费:"
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
               Left            =   360
               TabIndex        =   32
               Top             =   5040
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "制版费:"
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
            Begin XtremeSuiteControls.Label Label14 
               Height          =   315
               Left            =   360
               TabIndex        =   30
               Top             =   5730
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "后道工序:"
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
            Begin XtremeSuiteControls.Label Label13 
               Height          =   255
               Left            =   3720
               TabIndex        =   28
               Top             =   4140
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
            Begin XtremeSuiteControls.Label Label12 
               Height          =   255
               Left            =   3720
               TabIndex        =   26
               Top             =   3120
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "计价单位:"
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
            Begin XtremeSuiteControls.Label Label11 
               Height          =   255
               Left            =   7080
               TabIndex        =   14
               Top             =   5040
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "金额："
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
               Left            =   7080
               TabIndex        =   13
               Top             =   4140
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
            Begin XtremeSuiteControls.Label Label9 
               Height          =   255
               Left            =   360
               TabIndex        =   12
               Top             =   4140
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
            Begin XtremeSuiteControls.Label Label8 
               Height          =   255
               Left            =   7080
               TabIndex        =   11
               Top             =   3180
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "米数："
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
               Left            =   360
               TabIndex        =   10
               Top             =   3120
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "色号/花号："
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
            Begin XtremeSuiteControls.Label Label6 
               Height          =   255
               Left            =   7080
               TabIndex        =   9
               Top             =   2220
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "花型："
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
               Left            =   3720
               TabIndex        =   8
               Top             =   2220
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "颜色："
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
            Begin XtremeSuiteControls.Label Label4 
               Height          =   255
               Left            =   360
               TabIndex        =   7
               Top             =   2220
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "克重："
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
               Left            =   7080
               TabIndex        =   6
               Top             =   1320
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
            Begin XtremeSuiteControls.Label Label2 
               Height          =   255
               Left            =   3720
               TabIndex        =   5
               Top             =   1320
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
            Begin XtremeSuiteControls.Label Label1 
               Height          =   255
               Left            =   360
               TabIndex        =   4
               Top             =   1320
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
         End
      End
   End
End
Attribute VB_Name = "frmOrder_Edit"
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
'验证身份和时间
Private clspI As New clspI
Public client As String


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
    If Trim(FlatEdit2.Text) = "" Then
        MsgBox "品名不能为", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(FlatEdit3.Text) = "" Then
        MsgBox "门幅不能为空", vbInformation, "提示"
        Exit Sub
    End If
'     If Trim(FlatEdit4.Text) = "" Then
'        MsgBox "克重不能为空", vbInformation, "提示"
'        Exit Sub
'    End If
     If Trim(FlatEdit11.Text) = "" Then
        MsgBox "颜色不能为空", vbInformation, "提示"
        Exit Sub
    End If
'     If Trim(FlatEdit5.Text) = "" Then
'        MsgBox "花型不能为空", vbInformation, "提示"
'        Exit Sub
'    End If
'     If Trim(FlatEdit6.Text) = "" Then
'        MsgBox "花号不能为空", vbInformation, "提示"
'        Exit Sub
'    End If
 
     If Trim(ComboBox1.Text) = "" Then
            MsgBox "计价单位不能为空", vbInformation, "提示"
            Exit Sub
    End If
 
 
    If Val(FlatEdit7.Text) <= 0 And Val(FlatEdit8.Text) <= 0 And Val(FlatEdit12.Text) <= 0 Then
            MsgBox "三个中有一个不能为空", vbInformation, "提示"
            Exit Sub
    End If
  

'    If Trim(FlatEdit8.Text) = "" Then
'            MsgBox "公斤数不能为空", vbInformation, "提示"
'            Exit Sub
'    End If
'    If Trim(FlatEdit9.Text) = "" Then
'            MsgBox "码数不能为空", vbInformation, "提示"
'            Exit Sub
'    End If
'
     If Trim(FlatEdit9.Text) = "" Then
        MsgBox "单价不能为空", vbInformation, "提示"
        Exit Sub
    End If
    
    savedetail
    
End Sub


Private Sub ComboBox1_Click()
    If ComboBox1.Text = "米数" Then
        FlatEdit10.Text = Val(FlatEdit7.Text) * Val(FlatEdit9.Text)
    End If
    If ComboBox1.Text = "公斤数" Then
        FlatEdit10.Text = Val(FlatEdit8.Text) * Val(FlatEdit9.Text)
    End If
    If ComboBox1.Text = "码数" Then
        FlatEdit10.Text = Val(FlatEdit12.Text) * Val(FlatEdit9.Text)
    End If
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

Private Sub Form_Load()
    InitFrm
    num = 0
    bool = False
'    setValuation
    Valuation
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
End Sub

Private Sub Valuation()
    ComboBox1.AddItem "米数"
    ComboBox1.AddItem "公斤数"
    ComboBox1.AddItem "码数"
End Sub







'Private Sub setValuation()
'    If Len(Valuation) > 0 Then
'        If Len(Trim(Valuation)) = 2 Then
'            FlatEdit8.Enabled = False
'            FlatEdit8.BackColor = &HC0C0C0
'            num = 1
'        Else
'            FlatEdit7.Enabled = False
'            FlatEdit7.BackColor = &HC0C0C0
'            num = 2
'        End If
'    End If
'End Sub


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
    If ComboBox1.Text = "米数" Then
        FlatEdit10.Text = Val(FlatEdit7.Text) * Val(FlatEdit9)
     End If
        FlatEdit12.Text = Format(Val(FlatEdit7.Text) / 0.9144, "0.00")
   
End Sub
Private Sub FlatEdit8_Change()
    If ComboBox1.Text = "公斤数" Then
        FlatEdit10.Text = Val(FlatEdit8.Text) * Val(FlatEdit9)
    End If
End Sub
Private Sub FlatEdit9_Change()
    If ComboBox1.Text = "米数" Then
        FlatEdit10.Text = Val(FlatEdit7.Text) * Val(FlatEdit9.Text)
    End If
     If ComboBox1.Text = "公斤数" Then
        FlatEdit10.Text = Val(FlatEdit8.Text) * Val(FlatEdit9.Text)
    End If
    If ComboBox1.Text = "码数" Then
        FlatEdit10.Text = Val(FlatEdit12.Text) * Val(FlatEdit9.Text)
    End If
End Sub
Private Sub FlatEdit12_Change()
    If ComboBox1.Text = "码数" Then
        FlatEdit10.Text = Val(FlatEdit8.Text) * Val(FlatEdit12.Text)
    End If
End Sub
Private Sub FlatEdit14_Change()
    If ComboBox1.Text = "米数" Then
        FlatEdit10.Text = Val(FlatEdit7.Text) * Val(FlatEdit9.Text)
        FlatEdit10.Text = Val(FlatEdit14.Text) + Val(FlatEdit15.Text) + Val(FlatEdit7.Text) * Val(FlatEdit9.Text)
    End If
     If ComboBox1.Text = "公斤数" Then
        FlatEdit10.Text = Val(FlatEdit8.Text) * Val(FlatEdit9.Text)
        FlatEdit10.Text = Val(FlatEdit14.Text) + Val(FlatEdit15.Text) + Val(FlatEdit8.Text) * Val(FlatEdit9.Text)
    End If
    If ComboBox1.Text = "码数" Then
        FlatEdit10.Text = Val(FlatEdit12.Text) * Val(FlatEdit9.Text)
        FlatEdit10.Text = Val(FlatEdit14.Text) + Val(FlatEdit15.Text) + Val(FlatEdit12.Text) * Val(FlatEdit9.Text)
    End If
        
End Sub
Private Sub FlatEdit15_Change()
    If ComboBox1.Text = "米数" Then
        FlatEdit10.Text = Val(FlatEdit7.Text) * Val(FlatEdit9.Text)
        FlatEdit10.Text = Val(FlatEdit15.Text) + Val(FlatEdit14.Text) + Val(FlatEdit7.Text) * Val(FlatEdit9.Text)
    End If
     If ComboBox1.Text = "公斤数" Then
        FlatEdit10.Text = Val(FlatEdit8.Text) * Val(FlatEdit9.Text)
        FlatEdit10.Text = Val(FlatEdit15.Text) + Val(FlatEdit14.Text) + Val(FlatEdit8.Text) * Val(FlatEdit9.Text)
    End If
    If ComboBox1.Text = "码数" Then
        FlatEdit10.Text = Val(FlatEdit12.Text) * Val(FlatEdit9.Text)
        FlatEdit10.Text = Val(FlatEdit15.Text) + Val(FlatEdit14.Text) + Val(FlatEdit12.Text) * Val(FlatEdit9.Text)
    End If
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
            sql3 = "exec usp_savedetail '" & id & "','" & FlatEdit1.Text & "','" & FlatEdit2.Text & "','" & FlatEdit3.Text & "','" & FlatEdit4.Text & "','" & colorid & "','" & FlatEdit5.Text & "','" & FlatEdit6.Text & "','" & FlatEdit7.Text & "','" & FlatEdit8.Text & "','" & FlatEdit12.Text & "','" & FlatEdit9.Text & "','" & FlatEdit10.Text & "','" & FlatEdit13.Text & "','" & ComboBox1.Text & "','" & FlatEdit14.Text & "','" & FlatEdit15.Text & "','" & FlatEdit11.Text & "','" & FlatEdit16.Text & "','" & FlatEdit17.Text & "'"
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
         sql2 = "exec usp_savedbilletailupdate '" & itemid & "','" & FlatEdit1.Text & "','" & FlatEdit2.Text & "','" & FlatEdit3.Text & "','" & FlatEdit4.Text & "','" & colorid & "','" & FlatEdit5.Text & "','" & FlatEdit6.Text & "','" & FlatEdit7.Text & "','" & FlatEdit8.Text & "','" & FlatEdit12.Text & "','" & FlatEdit9.Text & "','" & FlatEdit10.Text & "','" & FlatEdit13.Text & "','" & ComboBox1.Text & "','" & FlatEdit14.Text & "','" & FlatEdit15.Text & "','" & FlatEdit11.Text & "','" & FlatEdit16.Text & "','" & FlatEdit17.Text & "'"
         Debug.Print sql2
        Gm.cnnTool.cnn.Execute sql2
        
    Else
        Dim sql As String
        sql = "exec usp_savedetailupdate '" & itemid & "','" & FlatEdit1.Text & "','" & FlatEdit2.Text & "','" & FlatEdit3.Text & "','" & FlatEdit4.Text & "','" & colorid & "','" & FlatEdit5.Text & "','" & FlatEdit6.Text & "','" & FlatEdit7.Text & "','" & FlatEdit8.Text & "','" & FlatEdit12.Text & "','" & FlatEdit9.Text & "','" & FlatEdit10.Text & "','" & FlatEdit13.Text & "','" & ComboBox1.Text & "','" & FlatEdit14.Text & "','" & FlatEdit15.Text & "','" & FlatEdit11.Text & "','" & FlatEdit16.Text & "','" & FlatEdit17.Text & "'"
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
Private Sub FlatEdit14_KeyPress(KeyAscii As Integer)
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
Private Sub FlatEdit15_KeyPress(KeyAscii As Integer)
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
    sql1 = "exec usp_savedetail '" & id & "','" & FlatEdit1.Text & "','" & FlatEdit2.Text & "','" & FlatEdit3.Text & "','" & FlatEdit4.Text & "','" & colorid & "','" & FlatEdit5.Text & "','" & FlatEdit6.Text & "','" & FlatEdit7.Text & "','" & FlatEdit8.Text & "','" & FlatEdit12.Text & "','" & FlatEdit9.Text & "','" & FlatEdit10.Text & "','" & FlatEdit13.Text & "','" & ComboBox1.Text & "','" & FlatEdit14.Text & "','" & FlatEdit15.Text & "','" & FlatEdit11.Text & "','" & FlatEdit16.Text & "'"
    Debug.Print sql1
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    Dim sql2 As String
    sql2 = "insert into G_BillDetailOrder (B_ID,B_itemid,B_OrderCode,B_GoodsID,B_Width,B_Weight,B_ColorID,B_HX,B_PatternCode,B_Meter,B_KG,B_Qty,B_Price,B_Sum,B_MemoDetail,B_BackMaterial,B_PlateMake,B_Sample,B_color,B_SourceOrderCode)  select B_ID,B_itemid,B_OrderCode,B_GoodsID,B_Width,B_Weight,B_ColorID,B_HX,B_PatternCode,B_Meter,B_KG,B_Qty,B_Price,B_Sum,B_MemoDetail,B_BackMaterial,B_PlateMake,B_Sample,B_color,B_SourceOrderCode from G_DraftBillDetailOrder  where B_ID='" & id & "'"
    Gm.cnnTool.cnn.Execute sql2
    Dim sql3 As String
    sql3 = "delete from G_DraftBillDetailOrder where B_ID='" & id & "'"
    Gm.cnnTool.cnn.Execute sql3
End Sub
Private Sub FlatEdit1_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            FlatEdit2.SetFocus
    End Select
End Sub
Private Sub FlatEdit2_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            FlatEdit3.SetFocus
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
        If Len(FlatEdit11.Text) <= 0 Then
            PushButton1_Click
        Else
            FlatEdit5.SetFocus
        End If
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
Private Sub FlatEdit12_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
             FlatEdit9.SetFocus
    End Select
End Sub

Private Sub FlatEdit9_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
     
    Select Case KeyCode
        Case 13
'        If MsgBox("是否保存数据", vbYesNoCancel + vbDefaultButton2 + vbInformation, "提示") = vbYes Then
'            save
'        if
'            Unload Me
'        End If
        Dim szReturn As VbMsgBoxResult
        szReturn = MsgBox("是否要保存？", vbYesNoCancel + vbDefaultButton1, "提示")
        Select Case szReturn
        
           Case vbYes
               save
           Case vbNo
               Unload Me
           Case vbCancel
               
        End Select

    End Select
End Sub

Private Sub PushButton2_Click()
    If client = "" Then
        MsgBox "先将主表选择一个客户", vbInformation, "提示"
        Exit Sub
        
    End If
    Dim frm1 As New frmPopupitemidb
    frm1.clientid = client
    frm1.Show vbModal
    FlatEdit16.Text = frm1.ordercode
    
    Unload frm1
End Sub
