VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmProductClientReturn 
   Caption         =   "��Ʒ�ͻ��˻�"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18420
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
   ScaleHeight     =   7950
   ScaleWidth      =   18420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Tag             =   "19B104"
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7950
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18420
      _LayoutVersion  =   1
      _ExtentX        =   32491
      _ExtentY        =   14023
      _DataPath       =   ""
      Bands           =   "frmProductClientReturn.frx":0000
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5775
         Left            =   300
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   17415
         _cx             =   30718
         _cy             =   10186
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
         _GridInfo       =   $"frmProductClientReturn.frx":5258
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   1185
            Left            =   30
            ScaleHeight     =   1185
            ScaleWidth      =   17355
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   30
            Width           =   17355
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   375
               Left            =   4680
               TabIndex        =   3
               Top             =   180
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               Format          =   210108417
               CurrentDate     =   43106
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   8640
               TabIndex        =   4
               Top             =   720
               Width           =   315
               _Version        =   1048578
               _ExtentX        =   556
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               Height          =   375
               Left            =   7320
               TabIndex        =   5
               Top             =   720
               Width           =   1335
               _Version        =   1048578
               _ExtentX        =   2355
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               Left            =   7320
               TabIndex        =   6
               Top             =   180
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit3 
               Height          =   375
               Left            =   1200
               TabIndex        =   7
               Top             =   180
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit4 
               Height          =   375
               Left            =   1200
               TabIndex        =   8
               Top             =   720
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.ComboBox ComboBox1 
               Height          =   315
               Left            =   4680
               TabIndex        =   9
               Top             =   750
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.PushButton PushButton2 
               Height          =   375
               Left            =   11640
               TabIndex        =   10
               Top             =   720
               Width           =   315
               _Version        =   1048578
               _ExtentX        =   556
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit5 
               Height          =   375
               Left            =   10320
               TabIndex        =   11
               Top             =   720
               Width           =   1335
               _Version        =   1048578
               _ExtentX        =   2355
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit6 
               Height          =   375
               Left            =   10320
               TabIndex        =   12
               Top             =   180
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.ComboBox ComboBox2 
               Height          =   315
               Left            =   13680
               TabIndex        =   13
               Top             =   210
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit7 
               Height          =   375
               Left            =   13680
               TabIndex        =   14
               Top             =   720
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.ComboBox ComboBox3 
               Height          =   315
               Left            =   16800
               TabIndex        =   15
               Top             =   240
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit8 
               Height          =   375
               Left            =   16800
               TabIndex        =   16
               Top             =   720
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               Left            =   300
               TabIndex        =   28
               Top             =   240
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "���ݱ��:"
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   255
               Left            =   6840
               TabIndex        =   27
               Top             =   240
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "�˷�:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   315
               Left            =   6840
               TabIndex        =   26
               Top             =   750
               Width           =   495
               _Version        =   1048578
               _ExtentX        =   873
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "�˷�:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label3 
               Height          =   255
               Left            =   3480
               TabIndex        =   25
               Top             =   240
               Width           =   1095
               _Version        =   1048578
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "�� �� �� ��:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label4 
               Height          =   255
               Left            =   300
               TabIndex        =   24
               Top             =   780
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "�� �� ��:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Left            =   3480
               TabIndex        =   23
               Top             =   780
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "������㷽ʽ:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label6 
               Height          =   315
               Left            =   9600
               TabIndex        =   22
               Top             =   750
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "װж��:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               Left            =   9600
               TabIndex        =   21
               Top             =   240
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "װж��:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label9 
               Height          =   255
               Left            =   12480
               TabIndex        =   20
               Top             =   240
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "�˷ѽ��㷽ʽ:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label11 
               Height          =   255
               Left            =   15600
               TabIndex        =   19
               Top             =   270
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "�����Ƿ��Ѹ�:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label10 
               Height          =   255
               Left            =   12480
               TabIndex        =   18
               Top             =   780
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "�˷��绰:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               Left            =   15600
               TabIndex        =   17
               Top             =   780
               Width           =   1095
               _Version        =   1048578
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "��  ʻ  Ա:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
            Height          =   4500
            Left            =   30
            TabIndex        =   29
            Top             =   1245
            Width           =   17355
            _ExtentX        =   30612
            _ExtentY        =   7938
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
            _StyleDefs(8)   =   ":id=1,.fontname=����"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgpicMode=2,.bgbmp=1,.bold=0"
            _StyleDefs(11)  =   ":id=2,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(12)  =   ":id=2,.fontname=����"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgpicMode=2,.bgbmp=2,.bold=0"
            _StyleDefs(14)  =   ":id=3,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(15)  =   ":id=3,.fontname=����"
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
Attribute VB_Name = "frmProductClientReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const theObjectID As String = "12B016"  '�������ݶ�����
Public rsdetail As RecordSet
Private theBLTool As New clsAutoCreateBL
Private dingdan As String
Public Originalsuppliers As String
Public fh As String

Public id As String '�������������
Private printdetail As Boolean
Public mvarObjectID As String
Public bol As Boolean '��֤�Ƿ񱣴�

Public Property Let ObjectID(ByVal vData As String)
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
    ObjectID = mvarObjectID
End Property

Private Function GetCodeID() As String
    GetCodeID = theBLTool.GetFrameCodeDetail(theObjectID)
End Function




Private Sub Form_Load()
     DTPicker1.Value = Now
    InitFrm
    printdetail = False
    
        TDBGrid1.HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid1.Splits(0).HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid1.HighlightRowStyle.ForeColor = &H80000008
    TDBGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
    TDBGrid1.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
    FlatEdit3.Text = GetCodeID
'    id = ""
    setRs
    cob1
    cob2
    cob3
End Sub
'��������2 �Ľ��㷽ʽ
Private Sub cob2()
    ComboBox2.Clear
    Dim rs As RecordSet
    Set rs = New RecordSet
    Dim sql As String
    sql = "Select B_SID From G_BalanceCope Where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    Do While Not rs.EOF
        ComboBox2.AddItem rs!B_sid
        rs.movenext
    Loop
End Sub
Private Sub cob3()
    ComboBox3.Clear
    ComboBox3.AddItem "��"
    ComboBox3.AddItem "��"
    ComboBox3.AddItem ""
    ComboBox3.Text = "��"
End Sub

'�󶨲ݸ�����
Private Sub setRs()
    Set rsdetail = New RecordSet
    rsdetail.Fields.Append "B_hex", adVarChar, 100
    rsdetail.Fields.Append "B_colorname", adVarChar, 100
    rsdetail.Fields.Append "B_colorid", adVarChar, 100
    rsdetail.Fields.Append "B_departColor", adVarChar, 100
    rsdetail.Fields.Append "B_itemid", adVarChar, 100
    rsdetail.Fields.Append "B_LoadTime", adVarChar, 100
    rsdetail.Fields.Append "B_itemidb", adVarChar, 100
    rsdetail.Fields.Append "B_name", adVarChar, 100
'    rsDetail.Fields.Append "B_SID", adVarChar, 100
    rsdetail.Fields.Append "B_SeHao", adVarChar, 100
    rsdetail.Fields.Append "B_Hua", adVarChar, 100
    rsdetail.Fields.Append "B_width", adVarChar, 100
    rsdetail.Fields.Append "B_weight", adVarChar, 100
 
    rsdetail.Fields.Append "B_PIshu", adVarChar, 100
    rsdetail.Fields.Append "B_kg", adVarChar, 100
    rsdetail.Fields.Append "B_DeliveryGoods", adVarChar, 100
    rsdetail.Fields.Append "B_Deliveryaddress", adVarChar, 100
    rsdetail.Fields.Append "B_PactCode", adVarChar, 100
    rsdetail.Fields.Append "B_Client", adVarChar, 100
    rsdetail.Fields.Append "B_Clientid", adVarChar, 100
    rsdetail.Fields.Append "B_Waitfreight", adVarChar, 100
        rsdetail.Fields.Append "B_freight", adVarChar, 100
    rsdetail.Fields.Append "B_Prepaidfreight", adVarChar, 100
    rsdetail.Fields.Append "B_memo", adVarChar, 100
    rsdetail.Fields.Append "B_orderitemid", adVarChar, 100
    rsdetail.Open
    
    TDBGrid1.DataSource = rsdetail
    setrsDetail
End Sub
Private Sub setrsDetail()
    setGridShow
       TDBGrid1.Columns("B_PIshu").NumberFormat = "0.0"
   TDBGrid1.Columns("B_kg").NumberFormat = "0.0"
'   TDBGrid1.Columns("B_Meter").NumberFormat = "0.0"
'   TDBGrid1.Columns("B_Ma").NumberFormat = "0.0"
'
   TDBGrid1.Columns("B_itemidb").Locked = True
   TDBGrid1.Columns("B_name").Locked = True
   TDBGrid1.Columns("B_width").Locked = True
   TDBGrid1.Columns("B_weight").Locked = True
'   TDBGrid1.Columns("B_SeHao").Locked = True
   TDBGrid1.Columns("B_Hua").Locked = True
   TDBGrid1.Columns("B_PactCode").Locked = True
   TDBGrid1.Columns("B_Client").Locked = True
   TDBGrid1.Columns("B_hex").Locked = True
   TDBGrid1.Columns("B_colorname").Locked = True
   TDBGrid1.Columns("B_departColor").Locked = True

  
   TDBGrid1.Columns("B_Waitfreight").ValueItems.Presentation = dbgCheckBox
   TDBGrid1.Columns("B_Prepaidfreight").ValueItems.Presentation = dbgCheckBox
     TDBGrid1.Columns("B_itemid").Locked = True
   TDBGrid1.Columns("B_itemid").AllowSizing = False
   TDBGrid1.Columns("B_itemid").Visible = False
   TDBGrid1.Columns("B_LoadTime").Locked = True
   TDBGrid1.Columns("B_LoadTime").AllowSizing = False
   TDBGrid1.Columns("B_LoadTime").Visible = False
     TDBGrid1.Columns("B_Clientid").Locked = True
   TDBGrid1.Columns("B_Clientid").AllowSizing = False
   TDBGrid1.Columns("B_Clientid").Visible = False
   
   TDBGrid1.Columns("B_colorid").Locked = True
   TDBGrid1.Columns("B_colorid").AllowSizing = False
   TDBGrid1.Columns("B_colorid").Visible = False
     TDBGrid1.Columns("B_orderitemid").Visible = False
   TDBGrid1.Columns("B_orderitemid").Locked = True
   TDBGrid1.Columns("B_orderitemid").AllowSizing = False
   TDBGrid1.Columns("B_DeliveryGoods").Caption = "�˻���λ"
   TDBGrid1.Columns("B_Deliveryaddress").Caption = "�˻���ַ"
     TDBGrid1.Columns("B_departColor").Visible = False
   TDBGrid1.Columns("B_departColor").Locked = True
   TDBGrid1.Columns("B_departColor").AllowSizing = False
   
 TDBGrid1.Columns("B_Hex").FetchStyle = True
   TDBGrid1.HoldFields
   TDBGrid1.MarqueeStyle = dbgHighlightRow
End Sub
Private Sub setGridShow()
    Dim cls1 As New clsGridShow
    With cls1
        .ObjectID = "11S069"
        .InitClass TDBGrid1, 3
        .ShowGridFormat
    End With
End Sub
Private Sub setGridStyle()
    Dim i As Long
    Dim strSQL As String
    Dim dWidth As Integer
    Dim szFieldName As String
    
    For i = 0 To TDBGrid1.Columns.Count - 1
        If TDBGrid1.Columns(i).width > 0 Then
            If TDBGrid1.Columns(i).Visible = True Then
                szFieldName = TDBGrid1.Columns(i).DataField
                dWidth = TDBGrid1.Columns(i).width
                strSQL = "update G_BLSField set B_GridWidth='" & dWidth & "' where B_ObjectID='11S069' and B_FieldName='" & szFieldName & "'"
                Gm.cnnTool.cnn.Execute strSQL
            End If
        End If
    Next
End Sub
Private Sub PushButton1_Click()
     Dim frm1 As New frmPopupDanWei
        frm1.ContactType = "��������"
        frm1.Show vbModal
        Originalsuppliers = frm1.Clientid
        FlatEdit1.Text = frm1.ClientName
        Unload frm1
End Sub
Private Sub PushButton2_Click()
       Dim frm1 As New frmpopupEmploy
        frm1.ContactType = "ɫ������װж��"
        frm1.Show vbModal
        fh = frm1.Clientid
        FlatEdit5.Text = frm1.ClientName
        Unload frm1
End Sub

'�󶨽��㷽ʽ
Private Sub cob1()
   ComboBox1.Clear
    Dim sql As String
    Dim rs As New RecordSet
    sql = "Select B_SID From G_Balance Where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            ComboBox1.AddItem "" & rs!B_sid & ""
            rs.movenext
        Loop
    End If

End Sub

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "���沢��ӡ"
            saveandprint
        Case "����"
            save
        Case "�˳�"
            Unload Me
        Case "������ʽ"
            setGridStyle
        Case "������"
            addnew
        Case "ɾ����"
            DeleteHang
        Case "ѡ�������ϸ�뵥"
            choosema
        Case "��һ��"
            MoveFirst
        Case "��һ��"
            MovePrevious
        Case "��һ��"
            movenext
        Case "���"
            movelast
        Case "����"
            add
         Case "�޸�"
            upd
    End Select
End Sub
Private Sub add()
    saveAudit (1)
 FlatEdit1 = ""
    FlatEdit2 = ""
    FlatEdit4 = ""
    FlatEdit5 = ""
    FlatEdit6 = ""
     FlatEdit7 = ""
    FlatEdit8 = ""
    id = ""
    cob1
    cob2
    cob3
    fh = ""
    Originalsuppliers = ""
    FlatEdit3.Text = GetCodeID
    setRs
    Dim a As Long
    Dim b As Long
    a = 0
    b = 0
   
    TDBGrid1.Columns("B_pishu").FooterText = "" & a & ""
    TDBGrid1.Columns("B_kg").FooterText = "" & b & ""
End Sub
'������
Private Sub addnew()
    Dim rs As New RecordSet
    Dim itemid As String
    Dim itemidb As String
    
    Dim B_name As String
    
    Dim B_SeHao As String
    Dim B_Hua As String
    Dim B_Width As String
    Dim B_weight As String
    Dim B_PactCode As String
    Dim B_Client As String
    Dim B_Clientid As String
    Dim B_hex As String
    Dim B_colorname As String
    Dim B_sid As String
    Dim bool As Boolean
    
    Dim bool2 As Boolean
    Dim sql As String
'    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim rs2 As RecordSet
    
    
     bool2 = False
    Dim frm1 As New frmProductDelivery_Edit
        If Len(dingdan) > 0 Then
            frm1.FlatEdit1.Text = dingdan
            frm1.FlatEdit1.SelStart = Len(dingdan)
            frm1.grid
        End If
   
    frm1.Show vbModal
    If frm1.bsaved = True Then
'        itemid = frm1.itemid
'        itemidb = frm1.itemidb
'        B_name = frm1.B_GoodsNameAlias
'        B_SeHao = frm1.B_SeHao
'        B_Hua = frm1.B_Producer
'        B_Width = frm1.B_Width
'        B_weight = frm1.B_weight
'        B_PactCode = frm1.B_PactCode
'        B_Client = frm1.B_Client
'        B_ClientID = frm1.B_ClientID
'        B_hex = frm1.B_hex
'        B_colorname = frm1.B_name
'        B_sid = frm1.B_sid
        Dim tdbgRow As Variant
        For Each tdbgRow In frm1.TDBGrid1.SelBookmarks
            frm1.rss.Bookmark = tdbgRow
            Set rs = New RecordSet
            TDBGridAddOneRow frm1.rss!B_itemid, frm1.rss!B_logo
            

        Next
        dingdan = frm1.FlatEdit1.Text



        '2018��3��24�շ���---
        
        
'        dingdan = frm1.dingdan
'        bool = True
'        Set rs = frm1.rsDetail
'
'        If rs.RecordCount > 0 Then
'
'            If TDBGrid1.ApproxCount > 0 Then
'                Set rs2 = rsDetail.Clone
'                rsDetail.MoveFirst
'                Do While Not rsDetail.EOF
'                        rs.MoveFirst
'
'
'                        Do While Not rs.EOF
'                                rs2.Filter = "B_itemid='" & rs!B_itemid & "'"
'                                If rs2.RecordCount > 0 Then
'                                    bool = False
'
'                                End If
'                                rs2.Filter = ""
'                                If bool = True Then
'                                rsDetail.AddNew
'                                rsDetail!B_itemid = rs!B_itemid
'                                rsDetail!B_ItemIDB = rs!B_ItemIDB
'                                rsDetail!B_name = rs!B_GoodsNameAlias
'                                rsDetail!B_SeHao = rs!B_SeHao
'                                rsDetail!B_Hua = rs!B_Producer
'                                rsDetail!B_Width = rs!B_Width
'                                rsDetail!B_weight = rs!B_weight
'                                rsDetail!B_PactCode = rs!B_PactCode
'                                rsDetail!B_Client = rs!B_Client
'                                rsDetail!B_Clientid = rs!B_Clientid
'                                rsDetail!B_hex = rs!B_hex
'                                rsDetail!B_colorname = rs!B_name
'                                rsDetail!B_Colorid = rs!B_sid
'                                rsDetail!B_freight = 0
'                                rsDetail.Update
'                                End If
'                                rs.movenext
'                                bool = True
'                         Loop
'
'                     rsDetail.movenext
'                Loop
'                 rsDetail.MoveFirst
'            Else
'                rs.MoveFirst
'                Do While Not rs.EOF
'                    rsDetail.AddNew
'
'                    rsDetail!B_itemid = rs!B_itemid
'                    rsDetail!B_ItemIDB = rs!B_ItemIDB
'                    rsDetail!B_name = rs!B_GoodsNameAlias
'                    rsDetail!B_SeHao = rs!B_SeHao
'                    rsDetail!B_Hua = rs!B_Producer
'                    rsDetail!B_Width = rs!B_Width
'                    rsDetail!B_weight = rs!B_weight
'                    rsDetail!B_PactCode = rs!B_PactCode
'                    rsDetail!B_Client = rs!B_Client
'                    rsDetail!B_Clientid = rs!B_Clientid
'                    rsDetail!B_hex = rs!B_hex
'                    rsDetail!B_colorname = rs!B_name
'                    rsDetail!B_Colorid = rs!B_sid
'                    rsDetail!B_freight = 0
'                    rsDetail.Update
'                    rs.movenext
'                Loop
'                rsDetail.MoveFirst
'            End If
'        Else
'            itemid = frm1.itemid
'            itemidb = frm1.itemidb
'            B_name = frm1.B_GoodsNameAlias
'            B_SeHao = frm1.B_SeHao
'            B_Hua = frm1.B_Producer
'            B_Width = frm1.B_Width
'            B_weight = frm1.B_weight
'            B_PactCode = frm1.B_PactCode
'            B_Client = frm1.B_Client
'            B_Clientid = frm1.B_Clientid
'            B_hex = frm1.B_hex
'            B_colorname = frm1.B_name
'            B_sid = frm1.B_sid
'
''            Debug.Print rs.RecordCount
''            rs.MoveFirst
''            Do While Not rs.EOF
'            Dim rsss As New RecordSet
'            Set rsss = rsDetail.Clone
'            rsss.Filter = "B_itemid='" & itemid & "'"
'            If rsss.RecordCount > 0 Then
'                MsgBox "�����ظ�", vbInformation, "��ʾ"
''                rsDetail.Filter = ""
'                Exit Sub
'            End If
'            rsDetail.AddNew
'            rsDetail!B_itemid = itemid
'            rsDetail!B_ItemIDB = itemidb
'
'            rsDetail!B_name = B_name
'            rsDetail!B_SeHao = B_SeHao
'            rsDetail!B_Hua = B_Hua
'            rsDetail!B_Width = B_Width
'            rsDetail!B_weight = B_weight
'            rsDetail!B_PactCode = B_PactCode
'            rsDetail!B_Client = B_Client
'            rsDetail!B_Clientid = B_Clientid
'            rsDetail!B_hex = B_hex
'            rsDetail!B_colorname = B_colorname
'            rsDetail!B_Colorid = B_sid
'            rsDetail!B_freight = 0
'            rsDetail.Update
            
'                rs.movenext
'            Loop
'        End If
                   
    Else
        Exit Sub
    End If

    Unload frm1
    sumall
    TDBGrid1.SetFocus
    TDBGrid1.Col = 11
End Sub

Private Sub TDBGridAddOneRow(ByVal vItemID As Long, ByVal vlogo As String)
    Dim sql As String
    Dim rs As New RecordSet
    sql = "exec usp_ProductdeliveryChoose '" & vItemID & "', '" & vlogo & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            rsdetail.addnew
'            rsDetail!B_itemid = rs!B_itemid
            rsdetail!B_ItemIDB = rs!B_ItemIDB
            rsdetail!B_name = rs!B_GoodsNameAlias
            rsdetail!B_SeHao = IIf(IsNull(rs!B_SeHao), "", rs!B_SeHao)
            rsdetail!B_Hua = IIf(IsNull(rs!B_Producer), "", rs!B_Producer)
            rsdetail!B_Width = rs!B_Width
            rsdetail!B_weight = rs!B_weight
            rsdetail!B_PactCode = IIf(IsNull(rs!B_PactCode), "", rs!B_PactCode)
            rsdetail!B_Client = rs!B_ClientName
            rsdetail!B_Clientid = rs!B_Clientid
            rsdetail!B_hex = rs!B_hex
            rsdetail!B_colorname = rs!B_name
            rsdetail!B_colorid = rs!B_sid
            rsdetail!B_Freight = 0
            rsdetail!B_orderitemid = rs!B_itemid
            rsdetail!B_DepartColor = IIf(IsNull(rs!B_DepartColor), "", rs!B_DepartColor)
            rsdetail.Update
End Sub




'ѡ��ϸ�뵥
Private Sub choosema()

'    Dim B_ItemIDB As String
'    Dim B_GoodsNameAlias As String
'    Dim B_SeHao As String
'    Dim B_Producer As String
'    Dim B_Width As String
'    Dim B_weight As String
'    Dim B_PactCode As String
'    Dim B_ClientName As String
'    Dim sum1 As String
'    Dim sum2 As String
'    Dim B_Clientid As String
'    Dim B_DTRK As String
'
'     Dim B_hex As String
'    Dim B_colorname As String
'    Dim B_sid As String
'
'    Dim frm1 As New frmProductDelivery_Edit
'    frm1.Show vbModal
'     If frm1.bsaved = True Then
'        B_ItemIDB = frm1.B_ItemIDB
'        B_GoodsNameAlias = frm1.B_GoodsNameAlias
'        B_SeHao = frm1.B_SeHao
'        B_Producer = frm1.B_Producer
'        B_Width = frm1.B_Width
'        B_weight = frm1.B_weight
'        B_PactCode = frm1.B_PactCode
'        B_ClientName = frm1.B_ClientName
'        sum1 = frm1.sum1
'        sum2 = frm1.sum2
'        B_Clientid = frm1.B_Clientid
'        B_DTRK = frm1.B_DTRK
'        B_hex = frm1.B_hex
'        B_colorname = frm1.B_name
'        B_sid = frm1.B_sid
'
'        If TDBGrid1.ApproxCount > 0 Then
'            rsdetail.MoveFirst
'            Do While Not rsdetail.EOF
'                If rsdetail!B_LoadTime = B_DTRK Then
'                    MsgBox "���ظ���¼", vbInformation + vbOKOnly, "��ʾ"
'                    Exit Sub
'                End If
'                rsdetail.movenext
'            Loop
'        End If
'
'    Else
'        Exit Sub
'    End If
'    Unload frm1
'
'    rsdetail.addnew
'    rsdetail!B_LoadTime = B_DTRK
'    rsdetail!B_ItemIDB = B_ItemIDB
'    rsdetail!B_name = B_GoodsNameAlias
'    rsdetail!B_SeHao = B_SeHao
'    rsdetail!B_Hua = B_Producer
'    rsdetail!B_Width = B_Width
'    rsdetail!B_weight = B_weight
'    rsdetail!B_Freight = 0
'    rsdetail!B_PIShu = sum1
'    rsdetail!B_kg = sum2
'      rsdetail!B_PactCode = B_PactCode
'    rsdetail!B_Client = B_ClientName
'    rsdetail!B_Clientid = B_Clientid
'    rsdetail!B_hex = B_hex
'    rsdetail!B_colorname = B_colorname
'    rsdetail!B_colorid = B_sid
'    rsdetail.Update
'    sumall
End Sub

Private Sub TDBGrid1_ButtonClick(ByVal ColIndex As Integer)
    
If TDBGrid1.Columns("B_process").ColIndex = ColIndex Then
     Dim frm2 As New frmPopupDanWei
    frm2.ContactType = "Ⱦ��"
    frm2.Show vbModal
    rsdetail!B_processid = frm2.Clientid
    rsdetail!B_process = frm2.ClientName
    Unload frm2
    End If
End Sub
Private Sub TDBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If ColIndex = TDBGrid1.Columns("B_PIshu").ColIndex Then
        TDBGrid1.Columns("B_PIshu").Value = Abs(Val(TDBGrid1.Columns("B_PIshu").Value))
    End If
       If ColIndex = TDBGrid1.Columns("B_kg").ColIndex Then
        TDBGrid1.Columns("B_kg").Value = Abs(Val(TDBGrid1.Columns("B_kg").Value))
    End If
    If ColIndex = TDBGrid1.Columns("B_freight").ColIndex Then
        TDBGrid1.Columns("B_freight").Value = Abs(Val(TDBGrid1.Columns("B_freight").Value))
    End If
    
    sumall
End Sub
Private Sub sumall()
     Dim rs As New RecordSet
    Dim a As Long
    Dim b As Double
    Dim c As String
    If rsdetail.RecordCount <= 0 Then
        Exit Sub
    End If
    a = 0
    b = 0
    Set rs = rsdetail.Clone
    rs.MoveFirst
    Do While Not rs.EOF
        a = a + IIf(IsNull(rs!B_PIShu), 0, rs!B_PIShu)
        b = b + IIf(IsNull(rs!B_kg), 0, rs!B_kg)
        rs.movenext
    Loop
    c = Format(b, "0.0")
    TDBGrid1.Columns("B_hex").FooterText = "�ϼ�"
    TDBGrid1.Columns("B_PIshu").FooterText = "" & a & ""
    TDBGrid1.Columns("B_kg").FooterText = "" & c & ""
End Sub
'����ɾ��
Private Sub DeleteHang()
    If Gm.PI.JudgeDelete(Me.Tag) = False Then
        Exit Sub
    End If
    On Error Resume Next
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim sql As String
    Dim sql1 As String
    Dim rs As New RecordSet
    Dim rs2 As New RecordSet
    Dim sql2 As String
    Dim sql3 As String
    
    sql2 = "select * from G_billdetailProduct where B_ID='" & id & "'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs2.RecordCount = 1 Then
    
         sql = "select * from G_billdetailProduct where B_itemid='" & rsdetail!B_itemid & "'"
         rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
         If rs.RecordCount > 0 Then
            If MsgBox("�˵���ֻʣһ������,ɾ����ȫ��ɾ�����Ƿ�ɾ��", vbInformation + vbYesNo + vbDefaultButton2, "��ʾ") = vbYes Then
                          Dim c As New Clsfreight
                    If c.Freight(FlatEdit3.Text) = False Then
                        MsgBox "�˵����������˷�,�޷�ɾ��", vbInformation, "��ʾ"
                        Exit Sub
                    End If
                    sql1 = "delete from G_billdetailProduct where B_itemid='" & rsdetail!B_itemid & "'"
                    Gm.cnnTool.cnn.Execute sql1
                        sql3 = "delete from G_billProduct where B_id='" & id & "'"
                    Gm.cnnTool.cnn.Execute sql3
                        FlatEdit1 = ""
                        FlatEdit2 = ""
                        FlatEdit4 = ""
                        FlatEdit5 = ""
                        FlatEdit6 = ""
                          FlatEdit7 = ""
                        FlatEdit8 = ""
                        id = ""
                        cob1
                        cob2
                        cob3
                        fh = ""
                        Originalsuppliers = ""
                        FlatEdit3.Text = GetCodeID
                        setRs
            Else
                Exit Sub
            End If
        End If
    Else
        sql1 = "delete from G_billdetailProduct where B_itemid='" & rsdetail!B_itemid & "'"
        Gm.cnnTool.cnn.Execute sql1
    End If
    
    rsdetail.delete
    If TDBGrid1.ApproxCount > 0 Then
        rsdetail.MoveFirst
    End If
    
End Sub

Private Sub save()
    If Gm.PI.JudgeNew(Me.Tag) = False Then
        Exit Sub
    End If
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim i As Long
    i = 1
    Dim a As String
    Dim rs As New RecordSet
    Dim sql As String
    Dim sql1 As String
    Dim rs2 As New RecordSet
    Dim sql2 As String
    
    If TDBGrid1.ApproxCount <= 0 Then
        MsgBox "û�����ݲ��ܱ���", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    
    rsdetail.MoveFirst
    Do While Not rsdetail.EOF
        If IIf(IsNull(rsdetail!B_PIShu), "", rsdetail!B_PIShu) = "" Or rsdetail!B_PIShu = 0 Then
            MsgBox "��" & i & "�м�������Ϊ�ջ���Ϊ0", vbInformation, "��ʾ"
            Exit Sub
        End If
        If IIf(IsNull(rsdetail!B_kg), "", rsdetail!B_kg) = "" Or rsdetail!B_kg = 0 Then
            MsgBox "��" & i & "�а�������Ϊ�ջ���Ϊ0", vbInformation, "��ʾ"
            Exit Sub
        End If
        If Abs(Val(rsdetail!B_Waitfreight)) = 1 Then
            If Val(rsdetail!B_Freight) <= 0 Then
                MsgBox "��" & i & "�в����˷�Ϊ0", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
        rsdetail.movenext
        i = i + 1
    Loop
    
    sql2 = "select * from G_billProduct where B_id='" & id & "'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs2.RecordCount > 0 Then
            saveupdate
            printdetail = True
            Exit Sub
    End If
    
    
    
    

    sql = "select * from G_draftBillProduct where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.addnew
    rs!B_datecreate = Now
    rs.Update
    id = rs!B_id
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    Dim d As String
    
    If ComboBox3.Text = "��" Then
        d = 1
    End If
    If ComboBox3.Text = "��" Then
        d = 0
    End If
    If ComboBox3.Text = "" Then
        d = ""
    End If
    sql1 = "exec usp_InsertProductOrder  '" & id & "','" & FlatEdit3.Text & "','" & a & "','PRO',"
    sql1 = sql1 & "'12B016','PRO02 ','" & Gm.SysID.SystemUser & "','" & FlatEdit2.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit6.Text & "','" & FlatEdit4.Text & "','" & ComboBox1.Text & "','" & Originalsuppliers & "','" & fh & "','" & ComboBox2.Text & "','" & FlatEdit7.Text & "','" & d & "','" & FlatEdit8.Text & "'"
    Gm.cnnTool.cnn.Execute sql1
    savedetail
      '���е������
   setAudit (0)
'    setOk
    sql = "delete from G_draftBillProduct where B_itemid='" & id & "'"
 FlatEdit1 = ""
    FlatEdit2 = ""
    FlatEdit4 = ""
    FlatEdit5 = ""
    FlatEdit6 = ""
     FlatEdit7 = ""
    FlatEdit8 = ""
    id = ""
    cob1
    cob2
    cob3
    fh = ""
    Originalsuppliers = ""
    FlatEdit3.Text = GetCodeID
    setRs
    printdetail = True
    bol = True
End Sub

Private Sub savedetail()
    Dim rs As RecordSet
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim sql1 As String
    Dim item As String
    Dim sql2 As String
    rsdetail.MoveFirst
    Do While Not rsdetail.EOF
        Set rs = New RecordSet
        sql = "select * from G_draftBilldetailProduct where 1=1"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        rs.addnew
        rs!B_datecreate = Now
        rs.Update
        item = rs!B_itemid
        
        sql2 = "exec usp_insertProductDelivery '" & item & "','" & id & "','" & rsdetail!B_ItemIDB & "'"
        sql2 = sql2 & ",'" & rsdetail!B_name & "','" & rsdetail!B_SeHao & "','" & rsdetail!B_Hua & "','" & rsdetail!B_Width & "','" & rsdetail!B_weight & "','" & rsdetail!B_PIShu & "'"
        sql2 = sql2 & ",'" & rsdetail!B_kg & "','" & rsdetail!B_PactCode & "','" & rsdetail!B_Clientid & "','" & rsdetail!B_Waitfreight & "','" & rsdetail!B_memo & "','" & rsdetail!B_colorid & "','" & rsdetail!B_colorname & "','" & rsdetail!B_Freight & "','" & rsdetail!B_Prepaidfreight & "','" & rsdetail!B_orderitemid & "','" & rsdetail!B_DepartColor & "','" & rsdetail!B_Deliveryaddress & "','" & rsdetail!B_DeliveryGoods & "'"

        Gm.cnnTool.cnn.Execute sql2
        sql1 = "delete from G_draftBilldetailProduct where B_itemid='" & item & "'"
        Gm.cnnTool.cnn.Execute sql1
        rsdetail.movenext
    Loop
End Sub
'���б����޸�ʱ���������ݽ��б���
Private Sub saveadddetail()
 Dim rs As RecordSet
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim sql1 As String
    Dim item As String
    Dim sql2 As String
 
    Set rs = New RecordSet
    sql = "select * from G_draftBilldetailProduct where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.addnew
    rs!B_datecreate = Now
    rs.Update
    item = rs!B_itemid
    
    sql2 = "exec usp_insertProductDelivery '" & item & "','" & id & "','" & rsdetail!B_ItemIDB & "'"
    sql2 = sql2 & ",'" & rsdetail!B_name & "','" & rsdetail!B_SeHao & "','" & rsdetail!B_Hua & "','" & rsdetail!B_Width & "','" & rsdetail!B_weight & "','" & rsdetail!B_PIShu & "'"
    sql2 = sql2 & ",'" & rsdetail!B_kg & "','" & rsdetail!B_PactCode & "','" & rsdetail!B_Clientid & "','" & rsdetail!B_Waitfreight & "','" & rsdetail!B_memo & "','" & rsdetail!B_colorid & "','" & rsdetail!B_colorname & "','" & rsdetail!B_Freight & "','" & rsdetail!B_Prepaidfreight & "','" & rsdetail!B_orderitemid & "','" & rsdetail!B_DepartColor & "','" & rsdetail!B_Deliveryaddress & "','" & rsdetail!B_DeliveryGoods & "' "

    Gm.cnnTool.cnn.Execute sql2
    sql1 = "delete from G_draftBilldetailProduct where B_itemid='" & item & "'"
    Gm.cnnTool.cnn.Execute sql1
  
End Sub
Private Sub setOk()
    Dim sql As String
    Dim rs As RecordSet
    
    rsdetail.MoveFirst
    Do While Not rsdetail.EOF
        If Len(rsdetail!B_LoadTime) > 0 Then
            sql = "update G_JRKBill set B_logo=1 where B_dtrk='" & rsdetail!B_LoadTime & "'"
             Gm.cnnTool.cnn.Execute sql
        End If
        rsdetail.movenext
    Loop
    
End Sub

Private Sub saveandprint()
    If Gm.PI.JudgeNew(Me.Tag) = False Then
        Exit Sub
    End If
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim i As Long
    i = 1
    Dim a As String
    Dim rs As New RecordSet
    Dim sql As String
    Dim sql1 As String
    Dim rs2 As New RecordSet
    Dim sql2 As String
    
    If TDBGrid1.ApproxCount <= 0 Then
        MsgBox "û�����ݲ��ܱ���", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    
    rsdetail.MoveFirst
    Do While Not rsdetail.EOF
        If IIf(IsNull(rsdetail!B_PIShu), "", rsdetail!B_PIShu) = "" Or rsdetail!B_PIShu = 0 Then
            MsgBox "��" & i & "�м�������Ϊ�ջ���Ϊ0", vbInformation, "��ʾ"
            Exit Sub
        End If
        If IIf(IsNull(rsdetail!B_kg), "", rsdetail!B_kg) = "" Or rsdetail!B_kg = 0 Then
            MsgBox "��" & i & "�а�������Ϊ�ջ���Ϊ0", vbInformation, "��ʾ"
            Exit Sub
        End If
        If Abs(rsdetail!B_Waitfreight) = 1 Then
            If Val(rsdetail!B_Freight) <= 0 Then
                MsgBox "��" & i & "�в����˷�Ϊ0", vbInformation, "��ʾ"
                Exit Sub
            End If
        End If
        rsdetail.movenext
        i = i + 1
    Loop
    
    sql2 = "select * from G_billProduct where B_id='" & id & "'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs2.RecordCount > 0 Then
            saveupdate1
    Else
    

    sql = "select * from G_draftBillProduct where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.addnew
    rs!B_datecreate = Now
    rs.Update
    id = rs!B_id
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    Dim d As String
    
    If ComboBox3.Text = "��" Then
        d = 1
    End If
    If ComboBox3.Text = "��" Then
        d = 0
    End If
    If ComboBox3.Text = "" Then
        d = ""
    End If
    sql1 = "exec usp_InsertProductOrder  '" & id & "','" & FlatEdit3.Text & "','" & a & "','PRO',"
    sql1 = sql1 & "'12B016','PRO02 ','" & Gm.SysID.SystemUser & "','" & FlatEdit2.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit6.Text & "','" & FlatEdit4.Text & "','" & ComboBox1.Text & "','" & Originalsuppliers & "','" & fh & "','" & ComboBox2.Text & "' ,'" & FlatEdit7.Text & "','" & d & "','" & FlatEdit8.Text & "'"
    Gm.cnnTool.cnn.Execute sql1
    savedetail
'    setOk
    sql = "delete from G_draftBillProduct where B_itemid='" & id & "'"
 
    End If
    
        Dim rs3 As New RecordSet
        Dim sql3 As String
         sql3 = "exec usp_Productdeliveryprint '" & id & "','" & Gm.SysID.SystemUserName & "'"
         Debug.Print sql
        rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        Dim frm1 As New frmModBLRPreviewOriColor
        Set frm1.RecordSet = rs3.Clone
            
        frm1.ObjectID = "22B095"
        frm1.Show vbModal
'        FlatEdit1 = ""
'    FlatEdit2 = ""
'    FlatEdit4 = ""
'    FlatEdit5 = ""
'    FlatEdit6 = ""
'    id = ""
'    cob1
'    cob2
'    fh = ""
'    Originalsuppliers = ""
'    FlatEdit3.Text = GetCodeID
'    setRs
    add
    bol = True
End Sub

Private Sub MoveFirst()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select top 1 B_id,B_CodeID,B_Date,B_Freight,B_PlateNumber,B_PNumber,B_payment,B_Shipment,b.B_ClientName,B_Hand,c.B_Name,B_cope,B_Freighttelephone,B_CostPay,B_drivename"
    sql = sql & " from G_BillProduct a left outer join G_ContactCompany b on a.B_Shipment=b.B_ClientID"
    sql = sql & " left outer join G_Employee c on a.B_Hand=c.B_SID where B_ObjectID='12B016' and  B_BillType='PRO02'"
    sql = sql & "order by B_id"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount <= 0 Then
        MsgBox "�޶�����Ϣ", vbInformation + vbOKOnly, "��ʾ"
        Exit Sub
    End If
    id = rs!B_id
   FlatEdit3.Text = IIf(IsNull(rs!B_Codeid), "", rs!B_Codeid)
    DTPicker1.Value = rs!B_Date
    FlatEdit2.Text = IIf(IsNull(rs!B_Freight), "", rs!B_Freight)
    FlatEdit6.Text = IIf(IsNull(rs!B_PlateNumber), "", rs!B_PlateNumber)
    FlatEdit4.Text = IIf(IsNull(rs!B_PNumber), "", rs!B_PNumber)
    ComboBox1.Text = rs!B_payment
    FlatEdit1.Text = IIf(IsNull(rs!B_ClientName), "", rs!B_ClientName)
    Originalsuppliers = rs!B_Shipment
    FlatEdit5.Text = IIf(IsNull(rs!B_name), "", rs!B_name)
    fh = rs!B_Hand
     ComboBox2.Text = IIf(IsNull(rs!B_cope), "", rs!B_cope)
       FlatEdit7.Text = IIf(IsNull(rs!B_Freighttelephone), "", rs!B_Freighttelephone)
    If rs!B_costpay = 0 Then
        ComboBox3.Text = "��"
    End If
    If rs!B_costpay = 1 Then
        ComboBox3.Text = "��"
    End If
    If rs!B_costpay = "" Then
        ComboBox3.Text = ""
    End If
    FlatEdit8.Text = IIf(IsNull(rs!B_drivename), "", rs!B_drivename)
    openbill
End Sub

Private Sub MovePrevious()
    Dim rs As New RecordSet
    Dim sql As String
    If id = "" Then
        movelast
        Exit Sub
    End If
    sql = "select top 1 a.B_id,B_CodeID,B_Date,B_Freight,B_PlateNumber,B_PNumber,B_payment,B_Shipment,b.B_ClientName,B_Hand,c.B_Name,B_cope,B_Freighttelephone,B_CostPay,B_drivename"
    sql = sql & " from G_BillProduct a left outer join G_ContactCompany b on a.B_Shipment=b.B_ClientID"
    sql = sql & " left outer join G_Employee c on a.B_Hand=c.B_SID where B_ObjectID='12B016' and  B_BillType='PRO02' and B_ID<'" & id & "'"
    sql = sql & "order by B_id desc"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount <= 0 Then
        MsgBox "���ǵ�һ��", vbInformation + vbOKOnly, "��ʾ"
        Exit Sub
    End If
    id = rs!B_id
     FlatEdit3.Text = IIf(IsNull(rs!B_Codeid), "", rs!B_Codeid)
    DTPicker1.Value = rs!B_Date
    FlatEdit2.Text = IIf(IsNull(rs!B_Freight), "", rs!B_Freight)
    FlatEdit6.Text = IIf(IsNull(rs!B_PlateNumber), "", rs!B_PlateNumber)
    FlatEdit4.Text = IIf(IsNull(rs!B_PNumber), "", rs!B_PNumber)
    ComboBox1.Text = rs!B_payment
    FlatEdit1.Text = IIf(IsNull(rs!B_ClientName), "", rs!B_ClientName)
    Originalsuppliers = rs!B_Shipment
    FlatEdit5.Text = IIf(IsNull(rs!B_name), "", rs!B_name)
    fh = rs!B_Hand
     ComboBox2.Text = IIf(IsNull(rs!B_cope), "", rs!B_cope)
       FlatEdit7.Text = IIf(IsNull(rs!B_Freighttelephone), "", rs!B_Freighttelephone)
    If rs!B_costpay = 0 Then
        ComboBox3.Text = "��"
    End If
    If rs!B_costpay = 1 Then
        ComboBox3.Text = "��"
    End If
    If rs!B_costpay = "" Then
        ComboBox3.Text = ""
    End If
    FlatEdit8.Text = IIf(IsNull(rs!B_drivename), "", rs!B_drivename)
    openbill
End Sub
Private Sub movenext()
    Dim rs As New RecordSet
    Dim sql As String
    If id = "" Then
        movelast
        Exit Sub
    End If
    sql = "select top 1 B_id,B_CodeID,B_Date,B_Freight,B_PlateNumber,B_PNumber,B_payment,B_Shipment,b.B_ClientName,B_Hand,c.B_Name,B_cope,B_Freighttelephone,B_CostPay,B_drivename"
    sql = sql & " from G_BillProduct a left outer join G_ContactCompany b on a.B_Shipment=b.B_ClientID"
    sql = sql & " left outer join G_Employee c on a.B_Hand=c.B_SID where B_ObjectID='12B016' and  B_BillType='PRO02' and B_ID>'" & id & "'"
    sql = sql & "order by B_id "
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount <= 0 Then
         MsgBox "�������һ��", vbInformation + vbOKOnly, "��ʾ"
        Exit Sub
    End If
    id = rs!B_id
    FlatEdit3.Text = IIf(IsNull(rs!B_Codeid), "", rs!B_Codeid)
    DTPicker1.Value = rs!B_Date
    FlatEdit2.Text = IIf(IsNull(rs!B_Freight), "", rs!B_Freight)
    FlatEdit6.Text = IIf(IsNull(rs!B_PlateNumber), "", rs!B_PlateNumber)
    FlatEdit4.Text = IIf(IsNull(rs!B_PNumber), "", rs!B_PNumber)
    ComboBox1.Text = rs!B_payment
    FlatEdit1.Text = IIf(IsNull(rs!B_ClientName), "", rs!B_ClientName)
    Originalsuppliers = rs!B_Shipment
    FlatEdit5.Text = IIf(IsNull(rs!B_name), "", rs!B_name)
    fh = rs!B_Hand
     ComboBox2.Text = IIf(IsNull(rs!B_cope), "", rs!B_cope)
       FlatEdit7.Text = IIf(IsNull(rs!B_Freighttelephone), "", rs!B_Freighttelephone)
    If rs!B_costpay = 0 Then
        ComboBox3.Text = "��"
    End If
    If rs!B_costpay = 1 Then
        ComboBox3.Text = "��"
    End If
    If rs!B_costpay = "" Then
        ComboBox3.Text = ""
    End If
    FlatEdit8.Text = IIf(IsNull(rs!B_drivename), "", rs!B_drivename)
    openbill
End Sub


Private Sub movelast()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select top 1 B_id,B_CodeID,B_Date,B_Freight,B_PlateNumber,B_PNumber,B_payment,B_Shipment,b.B_ClientName,B_Hand,c.B_Name,B_cope,B_Freighttelephone,B_CostPay,B_drivename"
    sql = sql & " from G_BillProduct a left outer join G_ContactCompany b on a.B_Shipment=b.B_ClientID"
    sql = sql & " left outer join G_Employee c on a.B_Hand=c.B_SID where B_ObjectID='12B016' and  B_BillType='PRO02'"
    sql = sql & "order by B_id desc"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount <= 0 Then
        MsgBox "�޶�����Ϣ", vbInformation + vbOKOnly, "��ʾ"
        Exit Sub
    End If
    id = rs!B_id
    FlatEdit3.Text = IIf(IsNull(rs!B_Codeid), "", rs!B_Codeid)
    DTPicker1.Value = rs!B_Date
    FlatEdit2.Text = IIf(IsNull(rs!B_Freight), "", rs!B_Freight)
    FlatEdit6.Text = IIf(IsNull(rs!B_PlateNumber), "", rs!B_PlateNumber)
    FlatEdit4.Text = IIf(IsNull(rs!B_PNumber), "", rs!B_PNumber)
    ComboBox1.Text = rs!B_payment
    FlatEdit1.Text = IIf(IsNull(rs!B_ClientName), "", rs!B_ClientName)
    Originalsuppliers = IIf(IsNull(rs!B_Shipment), "", rs!B_Shipment)
    FlatEdit5.Text = IIf(IsNull(rs!B_name), "", rs!B_name)
    fh = rs!B_Hand
     ComboBox2.Text = IIf(IsNull(rs!B_cope), "", rs!B_cope)
       FlatEdit7.Text = IIf(IsNull(rs!B_Freighttelephone), "", rs!B_Freighttelephone)
    If rs!B_costpay = 0 Then
        ComboBox3.Text = "��"
    End If
    If rs!B_costpay = 1 Then
        ComboBox3.Text = "��"
    End If
    If rs!B_costpay = "" Then
        ComboBox3.Text = ""
    End If
    FlatEdit8.Text = IIf(IsNull(rs!B_drivename), "", rs!B_drivename)
    openbill
End Sub

Public Sub openbill()
   Dim sql As String
   Dim rs As New RecordSet
   sql = "select c.B_HEX,case when isnull(B_ordercolor,'')='' then c.B_name when isnull(B_ordercolor,'')<>'' then a.B_ordercolor end as B_name,"
   sql = sql & " B_sid, B_ItemID,B_ItemIDB,B_GoodsNameAlias,B_SeHao,B_Producer,B_width,B_weight,B_ps,B_KG,a.B_PactCode,a.B_Clientid,b.B_ClientName,B_Waitfreight,B_MemoDetail,B_freight,B_Prepaidfreight,B_departColor,B_Deliveryaddress,B_DeliveryGoods"
   sql = sql & " from G_BillDetailProduct a left outer join G_ContactCompany b on a.B_Clientid=b.B_ClientID left outer join G_color c on a.B_color=c.B_sid"
   sql = sql & " where B_ID='" & id & "'"

   rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount <= 0 Then
        Exit Sub
    End If
   setRs
   rs.MoveFirst
   Do While Not rs.EOF
       rsdetail.addnew
       rsdetail!B_hex = rs!B_hex
       rsdetail!B_colorname = rs!B_name
       rsdetail!B_colorid = rs!B_sid
       rsdetail!B_DepartColor = IIf(IsNull(rs!B_DepartColor), "", rs!B_DepartColor)
       rsdetail!B_itemid = rs!B_itemid
       rsdetail!B_ItemIDB = rs!B_ItemIDB
       rsdetail!B_name = rs!B_GoodsNameAlias
       rsdetail!B_SeHao = rs!B_SeHao
       rsdetail!B_Hua = IIf(IsNull(rs!B_Producer), "", rs!B_Producer)
       rsdetail!B_Width = IIf(IsNull(rs!B_Width), "", rs!B_Width)
       rsdetail!B_weight = IIf(IsNull(rs!B_weight), "", rs!B_weight)
    
       rsdetail!B_PIShu = rs!B_ps
       rsdetail!B_kg = rs!B_kg
       rsdetail!B_PactCode = IIf(IsNull(rs!B_PactCode), "", rs!B_PactCode)
    
       rsdetail!B_Client = IIf(IsNull(rs!B_ClientName), "", rs!B_ClientName)
       rsdetail!B_Clientid = rs!B_Clientid
       rsdetail!B_Waitfreight = IIf(IsNull(rs!B_Waitfreight), 0, rs!B_Waitfreight)
       rsdetail!B_Freight = IIf(IsNull(rs!B_Freight), "", rs!B_Freight)
       rsdetail!B_Prepaidfreight = IIf(IsNull(rs!B_Prepaidfreight), 0, rs!B_Prepaidfreight)
       
       rsdetail!B_DeliveryGoods = IIf(IsNull(rs!B_DeliveryGoods), 0, rs!B_DeliveryGoods)
       rsdetail!B_Deliveryaddress = IIf(IsNull(rs!B_Deliveryaddress), 0, rs!B_Deliveryaddress)
       
       rsdetail!B_memo = rs!B_MemoDetail
       rsdetail!B_Client = IIf(IsNull(rs!B_ClientName), 0, rs!B_ClientName)
       rsdetail!B_Clientid = rs!B_Clientid
       rsdetail.Update
       rs.movenext
   Loop
   tp
    If rs.RecordCount > 0 Then
        rsdetail.MoveFirst
   End If
   sumall
End Sub
'�޸ı����ӡ
Private Sub saveupdate1()
    Dim sql1 As String
    Dim a As String
     a = Format(DTPicker1.Value, "YYYY-MM-DD")
    Dim d As String
    
    If ComboBox3.Text = "��" Then
        d = 1
    End If
    If ComboBox3.Text = "��" Then
        d = 0
    End If
    If ComboBox3.Text = "" Then
        d = ""
    End If
    sql1 = "exec usp_ProductDeliveryupdate  '" & id & "','" & FlatEdit3.Text & "','" & a & "','PRO',"
    sql1 = sql1 & "'12B016','PRO02 ','" & Gm.SysID.SystemUser & "','" & FlatEdit2.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit6.Text & "','" & FlatEdit4.Text & "','" & ComboBox1.Text & "','" & Originalsuppliers & "','" & fh & "','" & ComboBox2.Text & "','" & FlatEdit7.Text & "','" & d & "','" & FlatEdit8.Text & "'"
    Gm.cnnTool.cnn.Execute sql1
    savedetailupdate
      '���е������
   setAudit (0)
End Sub

Private Sub saveupdate()
    Dim sql1 As String
    Dim a As String
     a = Format(DTPicker1.Value, "YYYY-MM-DD")
    Dim d As String
    
    If ComboBox3.Text = "��" Then
        d = 1
    End If
    If ComboBox3.Text = "��" Then
        d = 0
    End If
    If ComboBox3.Text = "" Then
        d = ""
    End If
    sql1 = "exec usp_ProductDeliveryupdate  '" & id & "','" & FlatEdit3.Text & "','" & a & "','PRO',"
    sql1 = sql1 & "'12B016','PRO02 ','" & Gm.SysID.SystemUser & "','" & FlatEdit2.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit6.Text & "','" & FlatEdit4.Text & "','" & ComboBox1.Text & "','" & Originalsuppliers & "','" & fh & "','" & ComboBox2.Text & "','" & FlatEdit7.Text & "','" & d & "','" & FlatEdit8.Text & "'"
    Gm.cnnTool.cnn.Execute sql1
    savedetailupdate
      '���е������
   setAudit (0)
   FlatEdit1 = ""
    FlatEdit2 = ""
    FlatEdit4 = ""
    FlatEdit5 = ""
    FlatEdit6 = ""
     FlatEdit7 = ""
    FlatEdit8 = ""
    id = ""
    cob1
    cob2
    cob3
    fh = ""
    Originalsuppliers = ""
    FlatEdit3.Text = GetCodeID
    setRs
End Sub

Private Sub savedetailupdate()
    Dim sql3 As String
    Dim rs3 As RecordSet
    Dim sql2 As String
    rsdetail.MoveFirst
    Do While Not rsdetail.EOF
        Set rs3 = New RecordSet
        sql3 = "select * from G_billdetailProduct where B_itemid ='" & rsdetail!B_itemid & "'"
        rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs3.RecordCount > 0 Then
           sql2 = "exec usp_ColordetailDeliveryupdate '" & rsdetail!B_itemid & "','" & id & "','" & rsdetail!B_ItemIDB & "'"
           sql2 = sql2 & ",'" & rsdetail!B_name & "','" & rsdetail!B_SeHao & "','" & rsdetail!B_Hua & "','" & rsdetail!B_Width & "','" & rsdetail!B_weight & "','" & rsdetail!B_PIShu & "'"
           sql2 = sql2 & ",'" & rsdetail!B_kg & "','" & rsdetail!B_PactCode & "','" & rsdetail!B_Client & "','" & rsdetail!B_Waitfreight & "','" & rsdetail!B_memo & "','" & rsdetail!B_colorid & "','" & rsdetail!B_colorname & "','" & rsdetail!B_Freight & "','" & rsdetail!B_Prepaidfreight & "','" & rsdetail!B_DepartColor & "','" & rsdetail!B_Deliveryaddress & "','" & rsdetail!B_DeliveryGoods & "'"
           
           Gm.cnnTool.cnn.Execute sql2
        Else
            saveadddetail
        End If
        rsdetail.movenext
    Loop
End Sub
Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, _
    ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, _
    ByVal CellStyle As TrueOleDBGrid80.StyleDisp)

   
    On Error Resume Next
    Debug.Print TDBGrid1.Columns("B_Hex").CellValue(Bookmark)
    CellStyle.BackColor = TDBGrid1.Columns("B_Hex").CellValue(Bookmark)
     CellStyle.ForeColor = TDBGrid1.Columns("B_Hex").CellValue(Bookmark)
End Sub


Private Sub FlatEdit2_KeyPress(KeyAscii As Integer)
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

Private Sub saveAudit(ByVal a As Long)
    If a = 0 Then
        FlatEdit2.Enabled = False
        FlatEdit4.Enabled = False
        FlatEdit6.Enabled = False
        FlatEdit7.Enabled = False
        FlatEdit8.Enabled = False
        DTPicker1.Enabled = False
        PushButton1.Enabled = False
        PushButton2.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        TDBGrid1.Enabled = False
        
        ActiveBar21.Bands("band1").Tools("����ͼƬ").Visible = True
        ActiveBar21.Bands("Band2").Tools("������").Enabled = False
        ActiveBar21.Bands("Band2").Tools("ɾ����").Enabled = False
        
        
    End If
    If a = 1 Then
        FlatEdit2.Enabled = True
        FlatEdit4.Enabled = True
        FlatEdit6.Enabled = True
        FlatEdit7.Enabled = True
        FlatEdit8.Enabled = True
        DTPicker1.Enabled = True
        PushButton1.Enabled = True
        PushButton2.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        TDBGrid1.Enabled = True
        ActiveBar21.Bands("band1").Tools("����ͼƬ").Visible = False
         ActiveBar21.Bands("Band2").Tools("������").Enabled = True
        ActiveBar21.Bands("Band2").Tools("ɾ����").Enabled = True
    End If
    
    ActiveBar21.RecalcLayout
End Sub
Private Sub upd()
    If id <> "" Then
        setAudit (1)
        tp
        
    End If
End Sub
Private Sub tp()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "select * from  G_Billproduct  where B_id='" & id & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If IIf(IsNull(rs!B_Audit), 0, rs!B_Audit) = 0 Then
        saveAudit (0)
    End If
    If IIf(IsNull(rs!B_Audit), 0, rs!B_Audit) = 1 Then
        saveAudit (1)
    End If
End Sub
Private Sub setAudit(ByVal a As Long)
    Dim sql As String
    Dim rs As New RecordSet
    sql = "update G_Billproduct set B_Audit='" & a & "' where B_Id='" & id & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
End Sub

