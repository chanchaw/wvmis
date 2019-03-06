VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmDingDanReport_Edit 
   BackColor       =   &H00FFFFFF&
   Caption         =   "合同综合报表"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDingDanReport_Edit.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   16545
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7365
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16545
      _LayoutVersion  =   1
      _ExtentX        =   29184
      _ExtentY        =   12991
      _DataPath       =   ""
      Bands           =   "frmDingDanReport_Edit.frx":058A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6735
         Left            =   300
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   420
         Width           =   15855
         _cx             =   27966
         _cy             =   11880
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
         _GridInfo       =   $"frmDingDanReport_Edit.frx":22FE
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   2025
            Left            =   30
            ScaleHeight     =   2025
            ScaleWidth      =   15795
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   30
            Width           =   15795
            Begin XtremeSuiteControls.CheckBox CheckBox1 
               Height          =   255
               Left            =   10320
               TabIndex        =   34
               Top             =   240
               Width           =   255
               _Version        =   1048578
               _ExtentX        =   450
               _ExtentY        =   450
               _StockProps     =   79
               UseVisualStyle  =   -1  'True
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   315
               Left            =   8580
               TabIndex        =   16
               Top             =   210
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   224722945
               CurrentDate     =   43058
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   315
               Left            =   4920
               TabIndex        =   14
               Top             =   210
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   224722945
               CurrentDate     =   43058
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   1320
               TabIndex        =   12
               Top             =   210
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   224722945
               CurrentDate     =   43058
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   2700
               TabIndex        =   4
               Top             =   840
               Width           =   315
               _Version        =   1048578
               _ExtentX        =   556
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               BackColor       =   -2147483633
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               Height          =   375
               Left            =   1320
               TabIndex        =   5
               Top             =   840
               Width           =   1395
               _Version        =   1048578
               _ExtentX        =   2461
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BackColor       =   14737632
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit2 
               Height          =   375
               Left            =   8580
               TabIndex        =   6
               Top             =   840
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.ComboBox ComboBox1 
               Height          =   300
               Left            =   4920
               TabIndex        =   7
               Top             =   870
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   529
               _StockProps     =   77
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
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
               Height          =   300
               Left            =   1320
               TabIndex        =   18
               Top             =   1507
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   529
               _StockProps     =   77
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.ComboBox ComboBox3 
               Height          =   300
               Left            =   4920
               TabIndex        =   21
               Top             =   1500
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   529
               _StockProps     =   77
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit3 
               Height          =   375
               Left            =   8580
               TabIndex        =   23
               Top             =   1470
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.ComboBox ComboBox4 
               Height          =   300
               Left            =   12180
               TabIndex        =   24
               Top             =   877
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   529
               _StockProps     =   77
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.ComboBox ComboBox5 
               Height          =   300
               Left            =   12180
               TabIndex        =   26
               Top             =   1507
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   529
               _StockProps     =   77
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.ComboBox ComboBox6 
               Height          =   300
               Left            =   12180
               TabIndex        =   28
               Top             =   217
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   529
               _StockProps     =   77
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit4 
               Height          =   375
               Left            =   15720
               TabIndex        =   30
               Top             =   180
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.ComboBox ComboBox7 
               Height          =   300
               Left            =   15720
               TabIndex        =   32
               Top             =   840
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   529
               _StockProps     =   77
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.Label Label14 
               Height          =   255
               Left            =   14760
               TabIndex        =   33
               Top             =   870
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "合同分类:"
               BackColor       =   -2147483633
            End
            Begin XtremeSuiteControls.Label Label13 
               Height          =   255
               Left            =   14760
               TabIndex        =   31
               Top             =   240
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "延续打印:"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label12 
               Height          =   255
               Left            =   11280
               TabIndex        =   29
               Top             =   240
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "计划完成:"
               BackColor       =   -2147483633
            End
            Begin XtremeSuiteControls.Label Label11 
               Height          =   255
               Left            =   11280
               TabIndex        =   27
               Top             =   1530
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "色布计划:"
               BackColor       =   -2147483633
            End
            Begin XtremeSuiteControls.Label Label10 
               Height          =   255
               Left            =   11280
               TabIndex        =   25
               Top             =   900
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "白坯计划:"
               BackColor       =   -2147483633
            End
            Begin XtremeSuiteControls.Label Label9 
               Height          =   255
               Left            =   7860
               TabIndex        =   22
               Top             =   1530
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "订单号:"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label8 
               Height          =   255
               Left            =   4020
               TabIndex        =   20
               Top             =   1530
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "作        废:"
               BackColor       =   -2147483633
            End
            Begin XtremeSuiteControls.Label Label6 
               Height          =   195
               Left            =   420
               TabIndex        =   19
               Top             =   1560
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "是否完工:"
               BackColor       =   -2147483633
            End
            Begin XtremeSuiteControls.Label Label5 
               Height          =   195
               Left            =   7860
               TabIndex        =   15
               Top             =   270
               Width           =   675
               _Version        =   1048578
               _ExtentX        =   1191
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "交    期:"
               BackColor       =   -2147483633
            End
            Begin XtremeSuiteControls.Label Label4 
               Height          =   195
               Left            =   4020
               TabIndex        =   13
               Top             =   270
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "终止日期:"
               BackColor       =   -2147483633
            End
            Begin XtremeSuiteControls.Label Label3 
               Height          =   255
               Left            =   420
               TabIndex        =   11
               Top             =   240
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "起始日期:"
               BackColor       =   -2147483633
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   315
               Left            =   420
               TabIndex        =   10
               Top             =   870
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "客    户:"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   255
               Left            =   7860
               TabIndex        =   9
               Top             =   900
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "合同号:"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label7 
               Height          =   255
               Left            =   4020
               TabIndex        =   8
               Top             =   900
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "审        核:"
               BackColor       =   -2147483633
            End
         End
         Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
            Height          =   2430
            Left            =   30
            TabIndex        =   2
            Top             =   2085
            Width           =   15795
            _ExtentX        =   27861
            _ExtentY        =   4286
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
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
         Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
            Height          =   2160
            Left            =   30
            TabIndex        =   17
            Top             =   4545
            Width           =   15795
            _ExtentX        =   27861
            _ExtentY        =   3810
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
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
End
Attribute VB_Name = "frmDingDanReport_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public mvarObjectID As String

Private client As String
Private rss As RecordSet
Private rss1 As RecordSet
Private a As Long
Private b As Long
Private s As Long
Private Col As Long
Private whi As Long
Private pla As Long '设置计划完成
Private ord As Long '合同分类


Public Property Let ObjectID(ByVal vData As String)
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
    ObjectID = mvarObjectID
End Property

Private Sub ComboBox6_Click()
    If ComboBox6.Text = "" Then
        ComboBox4.Enabled = True
        ComboBox5.Enabled = True
    Else
        ComboBox4.Enabled = False
        ComboBox5.Enabled = False
        ComboBox4.Text = "全部"
        ComboBox5.Text = "全部"
    End If
End Sub

Private Sub Form_Load()
    InitFrm
    client = ""
    Grid
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
    Audit
    cargoClear
    invalid
    white
    color
    plan
    order
    Debug.Print Month(Now)
    DTPicker1.Value = Year(Now) & -Month(Now) & "-01"
    DTPicker2.Value = Now
    DTPicker3.Value = Year(Now) & "-12" & "-31"
    
End Sub
Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "查询"
            Grid
        Case "退出"
            Unload Me
        Case "合同详细打印"
            printdetail
        Case "设置完成"
            setfinish
        Case "取消完成"
            cancelfinish
    End Select
End Sub
'审核绑定的数据
Private Sub Audit()
    ComboBox1.AddItem "未审核"
    ComboBox1.AddItem "已审核"
    ComboBox1.AddItem "全部"
    ComboBox1.Text = "已审核"
End Sub
'作废绑定的数据
Private Sub invalid()
    ComboBox3.AddItem "未作废"
    ComboBox3.AddItem "已作废"
    ComboBox3.AddItem "全部"
    ComboBox3.Text = "未作废"
End Sub
'货清绑定的数据
Private Sub cargoClear()
    ComboBox2.AddItem "已完工"
    ComboBox2.AddItem "未完工"
    ComboBox2.AddItem "全部"
    ComboBox2.Text = "未完工"
End Sub
'白坯计划绑定的数据
Private Sub white()
    ComboBox4.AddItem "已制作"
    ComboBox4.AddItem "未制作"
    ComboBox4.AddItem "全部"
    ComboBox4.Text = "全部"
End Sub
'色布计划绑定的数据
Private Sub color()
    ComboBox5.AddItem "已制作"
    ComboBox5.AddItem "未制作"
    ComboBox5.AddItem "全部"
    ComboBox5.Text = "全部"
End Sub

'色布计划绑定的数据
Private Sub plan()
    ComboBox6.AddItem ""
    ComboBox6.AddItem "计划未完成"
    ComboBox6.AddItem "计划已完成"
    ComboBox6.Text = "计划未完成"
End Sub
'色布计划绑定的数据
Private Sub order()
    ComboBox7.AddItem "成品合同"
    ComboBox7.AddItem "面料合同"
    ComboBox7.AddItem "全部"
    ComboBox7.Text = "全部"
End Sub

'绑定数据
Private Sub Grid()
    Gm.log4Runtime "开始查询"
    Dim sql As String
    Dim c As String
    Dim d As String
    Dim f As String
    Dim sql1 As String
    Dim rs As New RecordSet
    Set rss = New RecordSet
    Gm.log4Runtime "开始选择条件"
    DownChoose
    Gm.log4Runtime "选择条件结束"
    c = Format(DTPicker1.Value, "YYYY-MM-DD")
    d = Format(DTPicker2.Value, "YYYY-MM-DD")
    f = Format(DTPicker3.Value, "YYYY-MM-DD")
    If CheckBox1.Value = 0 Then
        f = ""
    End If
    sql1 = "select * from G_SystemUser where B_UserName='" & Gm.SysID.SystemUser & "'"
    rs.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If IIf(IsNull(rs!B_salesman), 0, rs!B_salesman) = 1 Then
        sql = "exec usp_DingDanReport_Edit '" & c & "','" & d & "','" & f & "','" & a & "','" & client & "','" & Trim(FlatEdit2.Text) & "','" & b & "','" & "" & "','" & s & "','" & Trim(FlatEdit3.Text) & "','" & whi & "','" & Col & "','" & pla & "','" & FlatEdit4.Text & "','" & ord & "'"
        Debug.Print sql
    Else
        sql = "exec usp_DingDanReport_Edit '" & c & "','" & d & "','" & f & "','" & a & "','" & client & "','" & Trim(FlatEdit2.Text) & "','" & b & "','" & Gm.SysID.SystemUser & "','" & s & "','" & Trim(FlatEdit3.Text) & "','" & whi & "','" & Col & "','" & pla & "','" & FlatEdit4.Text & "','" & ord & "'"
    End If
    
  
    rss.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid1.DataSource = rss
    Gm.log4Runtime "绑定数据到网格"
    setgrid
    Gm.log4Runtime "设置样式"
    sumall
    Gm.log4Runtime "合计"
    If rss.RecordCount > 0 Then
        rss.MoveFirst
    End If
End Sub

Private Sub setgrid()
    TDBGrid1.Columns("B_PactCode").Caption = "合同号"
    TDBGrid1.Columns("B_ClientName").Caption = "客户"
    TDBGrid1.Columns("B_Date").Caption = "单据日期"
    TDBGrid1.Columns("B_DeliveryDate").Caption = "交货日"
    TDBGrid1.Columns("B_PactQty").Caption = "合同数量"
    TDBGrid1.Columns("B_PactBoxQty").Caption = "合同金额"
    TDBGrid1.Columns("B_Memo").Caption = "备注"
    TDBGrid1.Columns("B_SPlan").Caption = "色布计划"
    TDBGrid1.Columns("B_BPlan").Caption = "白坯计划"
    TDBGrid1.Columns("B_userdes").Caption = "制作人"
    
    TDBGrid1.Columns("B_PactBoxQty").NumberFormat = "0.00"
    
    TDBGrid1.Columns("B_Memo").width = 5000
    TDBGrid1.Columns("B_PactQty").width = 1000
    TDBGrid1.Columns("B_PactBoxQty").width = 1000
    TDBGrid1.Columns("B_ClientName").width = 3000
    TDBGrid1.Columns("B_ID").width = 0
    TDBGrid1.Columns("B_ID").Visible = False
    TDBGrid1.Columns("B_ID").AllowSizing = False
    TDBGrid1.Columns("B_invalid").width = 0
    TDBGrid1.Columns("B_invalid").Visible = False
    TDBGrid1.Columns("B_invalid").AllowSizing = False
    TDBGrid1.MarqueeStyle = dbgHighlightRow
        TDBGrid1.Columns("B_contractlogo").width = 0
    TDBGrid1.Columns("B_contractlogo").Visible = False
    TDBGrid1.Columns("B_contractlogo").AllowSizing = False
    TDBGrid1.Columns("B_ticketClear").width = 0
    TDBGrid1.Columns("B_ticketClear").Visible = False
    TDBGrid1.Columns("B_ticketClear").AllowSizing = False
    
    
    TDBGrid1.HoldFields
'    If TDBGrid1.ApproxCount > 0 Then
'    Do While Not rss.EOF
'        If rss!B_invalid = 1 Then
'            TDBGrid1.BackColor = vbRed
'        End If
'        rss.movenext
'    Loop
'    rss.MoveFirst
'    End If
    TDBGrid1.Columns("B_SPlan").ValueItems.Presentation = dbgCheckBox
    TDBGrid1.Columns("B_BPlan").ValueItems.Presentation = dbgCheckBox
    TDBGrid1.FetchRowStyle = True
End Sub




Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, _
    ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    
    Dim lState As Long
    lState = Val(IIf(IsNull(TDBGrid1.Columns("B_invalid").CellValue(bookmark)), "0", TDBGrid1.Columns("B_invalid").CellValue(bookmark)))
    
    Dim lState1 As Long
    lState1 = Val(IIf(IsNull(TDBGrid1.Columns("B_ticketClear").CellValue(bookmark)), "0", TDBGrid1.Columns("B_ticketClear").CellValue(bookmark)))
    If lState = 1 Then
        RowStyle.BackColor = &HC0C0C0
    End If
    
    If lState1 = 1 Then
        RowStyle.ForeColor = &HFF&
    End If
End Sub
'下拉框选择
Private Sub DownChoose()

    If ComboBox1.Text = "未审核" Then
        a = 0
    End If
    If ComboBox1.Text = "已审核" Then
        a = 1
    End If
    If ComboBox1.Text = "全部" Then
        a = 2
    End If
    If ComboBox2.Text = "未完工" Then
        b = 0
    End If
    If ComboBox2.Text = "已完工" Then
        b = 1
    End If
    If ComboBox2.Text = "全部" Then
        b = 2
    End If
    If ComboBox3.Text = "未作废" Then
        s = 0
    End If
    If ComboBox3.Text = "已作废" Then
        s = 1
    End If
    If ComboBox3.Text = "全部" Then
        s = 2
    End If
     If ComboBox4.Text = "未制作" Then
        whi = 0
    End If
    If ComboBox4.Text = "已制作" Then
        whi = 1
    End If
    If ComboBox4.Text = "全部" Then
        whi = 2
    End If
     If ComboBox5.Text = "未制作" Then
        Col = 0
    End If
    If ComboBox5.Text = "已制作" Then
        Col = 1
    End If
    If ComboBox5.Text = "全部" Then
        Col = 2
    End If
    
    If ComboBox6.Text = "" Then
        pla = 2
    End If
    If ComboBox6.Text = "计划未完成" Then
        pla = 0
    End If
        If ComboBox6.Text = "计划已完成" Then
        pla = 1
    End If
    
    If ComboBox7.Text = "面料合同" Then
        ord = 0
    End If
    If ComboBox7.Text = "成品合同" Then
        ord = 1
    End If
    If ComboBox7.Text = "全部" Then
        ord = 2
    End If
  
End Sub
'获取客户id
Private Sub PushButton1_Click()
    Dim frm1 As New frmPopupClient
    frm1.Show vbModal
    FlatEdit1.Text = frm1.ClientName
    client = frm1.clientid
    Unload frm1
End Sub
'页脚小计
Private Sub sumall()
    Dim a As Long
    Dim b As Long
    If rss.RecordCount <= 0 Then
        Exit Sub
    End If
    a = 0
    b = 0
    rss.MoveFirst
    Do While Not rss.EOF
        a = a + IIf(IsNull(rss!B_PactQty), 0, rss!B_PactQty)
        b = b + IIf(IsNull(rss!B_PactBoxQty), 0, rss!B_PactBoxQty)
        rss.movenext
    Loop
    TDBGrid1.Columns("B_PactCode").FooterText = "合计"
    TDBGrid1.Columns("B_PactQty").FooterText = "" & a & ""
    TDBGrid1.Columns("B_PactBoxQty").FooterText = "" & b & ""
End Sub

Private Sub TDBGrid1_DblClick()
    Dim sql As String
    Dim rs As New RecordSet
    If rss.RecordCount > 0 Then
        sql = "select * from G_Billorder where B_ID='" & rss!B_id & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs.RecordCount <= 0 Then
            MsgBox "此订单不存在，可能已经删除", vbInformation, "提示"
            Exit Sub
        End If
        If rs!B_contractlogo = 0 Then
            
        
            Dim frm1 As New frmOrder
            frm1.theid = rss!B_id
            frm1.openbill
            frm1.Show
        End If
        If rs!B_contractlogo = 1 Then
            
        
            Dim frm2 As New frmOrderProduct
            frm2.theid = rss!B_id
            frm2.openbill
            frm2.Show
        End If
    End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim a As String
    Set rss1 = New RecordSet
    Dim sql1 As String
    If rss.RecordCount <= 0 Then
        a = ""
    Else
        a = rss!B_id
    End If
    
    sql1 = "exec usp_selectorder '" & a & "'"
    rss1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    TDBGrid2.DataSource = rss1
    
    
    setgrid2
    sumall2
    If rss1.RecordCount > 0 Then
        rss1.MoveFirst
    End If
    TDBGrid2.Columns("B_Hex").FetchStyle = True
End Sub
Private Sub setgrid2()
    TDBGrid2.Columns("B_orderCode").Caption = "订单号"
    TDBGrid2.Columns("B_GoodsID").Caption = "品名"
'    TDBGrid2.Columns("B_Name").Caption = "颜色"
    TDBGrid2.Columns("B_color").Caption = "基础颜色"
    TDBGrid2.Columns("B_Hex").Caption = "颜色标识"
    TDBGrid2.Columns("B_PatternCode").Caption = "花号"
    TDBGrid2.Columns("B_Pactqty").Caption = "订单数量"
    TDBGrid2.Columns("B_Price").Caption = "单价"
    TDBGrid2.Columns("B_Sum").Caption = "订单金额"
    TDBGrid2.Columns("B_Width").Caption = "门幅"
    TDBGrid2.Columns("B_Weight").Caption = "克重"
    TDBGrid2.Columns("B_SPaln").Caption = "色布计划"
    TDBGrid2.Columns("B_WPaln").Caption = "白坯计划"
    TDBGrid2.Columns("B_orderColor").Caption = "修改颜色名称"
    TDBGrid2.Columns("B_SeHao").Caption = "色布色号"
    TDBGrid2.Columns("B_intype").Caption = "类别"
    TDBGrid2.Columns("B_Positivefabric").Caption = "正面面料"
    TDBGrid2.Columns("B_Middlefabric").Caption = "中间面料"
    TDBGrid2.Columns("B_Backfabric").Caption = "背面面料"
    TDBGrid2.Columns("B_GoodManual").Caption = "手工缸号"
    TDBGrid2.Columns("B_Process").Caption = "工艺"
    TDBGrid2.Columns("B_Packaging").Caption = "包装"
    
    TDBGrid2.Columns("B_Price").NumberFormat = "0.00"
    TDBGrid2.Columns("B_Sum").NumberFormat = "0.00"
    
    TDBGrid2.Columns("B_orderCode").width = 1000
    TDBGrid2.Columns("B_SeHao").width = 1000
    TDBGrid2.Columns("B_orderColor").width = 1200
    TDBGrid2.Columns("B_color").width = 1000
    TDBGrid2.Columns("B_Hex").width = 1000
    TDBGrid2.Columns("B_Width").width = 1000
    TDBGrid2.Columns("B_Weight").width = 1000
    TDBGrid2.Columns("B_Pactqty").width = 1000
    TDBGrid2.Columns("B_Price").width = 1000
    TDBGrid2.Columns("B_Sum").width = 1500
    TDBGrid2.Columns("B_Positivefabric").width = 1000
    TDBGrid2.Columns("B_Middlefabric").width = 1000
    TDBGrid2.Columns("B_Backfabric").width = 1000
    TDBGrid2.Columns("B_GoodManual").width = 1000
    TDBGrid2.Columns("B_Process").width = 1000
    TDBGrid2.Columns("B_Packaging").width = 1000
    TDBGrid2.Columns("B_ID").width = 0
    TDBGrid2.Columns("B_ID").Visible = False
    TDBGrid2.Columns("B_ID").AllowSizing = False
    
    TDBGrid2.Columns("B_contractlogo").width = 0
    TDBGrid2.Columns("B_contractlogo").Visible = False
    TDBGrid2.Columns("B_contractlogo").AllowSizing = False
    
    
    TDBGrid2.Columns("B_SPaln").ValueItems.Presentation = dbgCheckBox
    TDBGrid2.Columns("B_WPaln").ValueItems.Presentation = dbgCheckBox
    TDBGrid2.MarqueeStyle = dbgHighlightRow
    TDBGrid2.HoldFields
End Sub


Private Sub sumall2()
    Dim a As Long
    Dim b As Long
    If rss1.RecordCount <= 0 Then
        Exit Sub
    End If
    a = 0
    b = 0
    rss1.MoveFirst
    Do While Not rss1.EOF
        a = a + IIf(IsNull(rss1!B_PactQty), 0, rss1!B_PactQty)
        b = b + IIf(IsNull(rss1!B_Sum), 0, rss1!B_Sum)
        rss1.movenext
    Loop
    TDBGrid2.Columns("B_orderCode").FooterText = "合计"
    TDBGrid2.Columns("B_PactQty").FooterText = "" & a & ""
    TDBGrid2.Columns("B_Sum").FooterText = "" & b & ""
End Sub
Private Sub TDBGrid2_DblClick()
    If rss1.RecordCount > 0 Then
        If rss1!B_contractlogo = 0 Then
            Dim frm1 As New frmOrder
            frm1.theid = rss1!B_id
            frm1.openbill
            frm1.Show
        End If
        If rss1!B_contractlogo = 1 Then
        
            Dim frm2 As New frmOrderProduct
            frm2.theid = rss1!B_id
            frm2.openbill
            frm2.Show
        End If
    End If
End Sub
Private Sub TDBGrid2_FetchCellStyle(ByVal Condition As Integer, _
    ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, _
    ByVal CellStyle As TrueOleDBGrid80.StyleDisp)
    On Error Resume Next
    
    CellStyle.BackColor = TDBGrid2.Columns("B_Hex").CellValue(bookmark)
    CellStyle.ForeColor = TDBGrid2.Columns("B_Hex").CellValue(bookmark)
End Sub

'打印
Private Sub printdetail()
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If

    Dim c As String
    Dim d As String
    Dim f As String
  
    DownChoose
    c = Format(DTPicker1.Value, "YYYY-MM-DD")
    d = Format(DTPicker2.Value, "YYYY-MM-DD")
    f = Format(DTPicker3.Value, "YYYY-MM-DD")
    If CheckBox1.Value = 0 Then
        f = ""
    End If
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    sql1 = "select * from G_SystemUser where B_UserName='" & Gm.SysID.SystemUser & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Dim frm1 As New frmModBLRPreviewOri
    
    
    
    If IIf(IsNull(rs1!B_salesman), 0, rs1!B_salesman) = 1 Then
        sql = "exec usp_DingDanReport_print '" & c & "','" & d & "','" & f & "','" & a & "','" & client & "','" & Trim(FlatEdit2.Text) & "','" & b & "','" & "" & "','" & s & "','" & Trim(FlatEdit3.Text) & "','" & Gm.SysID.SystemUserName & "','" & ord & "'"
        Debug.Print sql
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Else
        sql = "exec usp_DingDanReport_print '" & c & "','" & d & "','" & f & "','" & a & "','" & client & "','" & Trim(FlatEdit2.Text) & "','" & b & "','" & Gm.SysID.SystemUser & "','" & s & "','" & Trim(FlatEdit3.Text) & "','" & Gm.SysID.SystemUserName & "','" & ord & "'"

        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    End If
    
    frm1.obj = "11S057"
    Set frm1.RecordSet = rs.Clone
    frm1.ObjectID = "22B025"
    frm1.Show
End Sub

Private Sub PrintHang()
  
    Dim sql As String
    Dim rs As New RecordSet
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim a As String
    Dim b As String
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    b = Format(DTPicker2.Value, "YYYY-MM-DD")
    sql = "exec usp_dingdandetailreport '" & a & "','" & b & "','" & rss1!B_ItemID & "','" & Gm.SysID.SystemUserName & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockReadOnly
    
    Dim frm1 As New frmModBLRPreviewOri
    Set frm1.RecordSet = rs.Clone
    frm1.ObjectID = "22B025"
    frm1.Show
    
End Sub

Private Sub setfinish()
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim sql As String
    sql = "update G_BillOrder set B_ticketClear=1 where B_id='" & rss!B_id & "'"
    Gm.cnnTool.cnn.Execute sql
    Grid
    
End Sub
Private Sub cancelfinish()
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim sql As String
    sql = "update G_BillOrder set B_ticketClear=0 where B_id='" & rss!B_id & "'"
    Gm.cnnTool.cnn.Execute sql
    Grid
End Sub
