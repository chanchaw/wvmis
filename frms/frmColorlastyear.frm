VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmColorlastyear 
   Caption         =   "ɫ������ȿ��"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15990
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   15990
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15990
      _LayoutVersion  =   1
      _ExtentX        =   28205
      _ExtentY        =   13679
      _DataPath       =   ""
      Bands           =   "frmColorlastyear.frx":0000
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5775
         Left            =   300
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   15135
         _cx             =   26696
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
         _GridInfo       =   $"frmColorlastyear.frx":41B2
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   1185
            Left            =   30
            ScaleHeight     =   1185
            ScaleWidth      =   15075
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   30
            Width           =   15075
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   375
               Left            =   4680
               TabIndex        =   3
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               _Version        =   393216
               Format          =   241631233
               CurrentDate     =   43106
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   9360
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
               Left            =   7800
               TabIndex        =   5
               Top             =   720
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit2 
               Height          =   375
               Left            =   7800
               TabIndex        =   6
               Top             =   180
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
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
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
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
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
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
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.PushButton PushButton2 
               Height          =   375
               Left            =   12600
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
               Left            =   11040
               TabIndex        =   11
               Top             =   720
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit6 
               Height          =   375
               Left            =   11040
               TabIndex        =   12
               Top             =   180
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
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
               Left            =   14520
               TabIndex        =   13
               Top             =   240
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.Label Label9 
               Height          =   255
               Left            =   13320
               TabIndex        =   22
               Top             =   270
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
            Begin XtremeSuiteControls.Label Label8 
               Height          =   255
               Left            =   10320
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
            Begin XtremeSuiteControls.Label Label6 
               Height          =   315
               Left            =   10320
               TabIndex        =   20
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
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Left            =   3600
               TabIndex        =   19
               Top             =   780
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "���㷽ʽ:"
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
               TabIndex        =   18
               Top             =   780
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "���ƺ�:"
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
               Left            =   3600
               TabIndex        =   17
               Top             =   240
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "�������:"
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
               Left            =   7080
               TabIndex        =   16
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
            Begin XtremeSuiteControls.Label Label2 
               Height          =   255
               Left            =   7080
               TabIndex        =   15
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
            Begin XtremeSuiteControls.Label Label7 
               Height          =   255
               Left            =   300
               TabIndex        =   14
               Top             =   240
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "���ݱ��:"
            End
         End
         Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
            Height          =   4500
            Left            =   30
            TabIndex        =   23
            Top             =   1245
            Width           =   15075
            _ExtentX        =   26591
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
Attribute VB_Name = "frmColorlastyear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cls1 As clsGridShow
Public rsdetail As RecordSet
Private Const theObjectID As String = "12B008"  '�������ݶ�����
Private theBLTool As New clsAutoCreateBL
Public mvarObjectID As String
Public dingdan As String


Public id As String
Public fh As String
Public Originalsuppliers As String

Private printdetail As Boolean    '������ɽ��д�ӡ����֤

Public Property Let ObjectID(ByVal vData As String)
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
    ObjectID = mvarObjectID
End Property


Private Function GetCodeID() As String
    GetCodeID = theBLTool.GetFrameCodeDetail(theObjectID)
End Function

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
            AddNew
        Case "ɾ����"
            DeleteHang
            
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
    End Select
End Sub
Private Sub add()
    FlatEdit1 = ""
    FlatEdit2 = ""
    FlatEdit4 = ""
    FlatEdit5 = ""
    FlatEdit6 = ""
    id = ""
    cob1
    cob2
    fh = ""
    Originalsuppliers = ""
    FlatEdit3.Text = GetCodeID
    setRs
    Dim a As Long
    Dim b As Long
    a = 0
    b = 0
    TDBGrid1.Columns("B_hex").FooterText = "�ϼ�"
    TDBGrid1.Columns("B_Weight").FooterText = "" & a & ""
    TDBGrid1.Columns("B_PIshu").FooterText = "" & b & ""
End Sub


Private Sub Form_Load()
    InitFrm
    printdetail = False
    DTPicker1 = Now
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
    FlatEdit3.Text = GetCodeID
    setRs
    cob1
    cob2
    
     TDBGrid1.HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid1.Splits(0).HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid1.HighlightRowStyle.ForeColor = &H80000008
    TDBGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
    TDBGrid1.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell
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

'�󶨲ݸ�����
Private Sub setRs()
    Set rsdetail = New RecordSet
    rsdetail.Fields.Append "B_HEX", adVarChar, 100
    rsdetail.Fields.Append "B_itemidb", adVarChar, 100
    
    rsdetail.Fields.Append "B_depart", adVarChar, 100
    rsdetail.Fields.Append "B_ColorName", adVarChar, 100
    rsdetail.Fields.Append "B_specification", adVarChar, 100
    rsdetail.Fields.Append "B_specification1", adVarChar, 100
    rsdetail.Fields.Append "B_Color", adVarChar, 100
    rsdetail.Fields.Append "B_SeHao", adVarChar, 100
    rsdetail.Fields.Append "B_PIshu", adVarChar, 100
    rsdetail.Fields.Append "B_Weight", adVarChar, 100
    rsdetail.Fields.Append "B_Meter", adVarChar, 100
    rsdetail.Fields.Append "B_Ma", adVarChar, 100
    rsdetail.Fields.Append "B_ClientName", adVarChar, 100
    rsdetail.Fields.Append "B_DeliveryNote", adVarChar, 100
    rsdetail.Fields.Append "B_type", adVarChar, 100
    
    rsdetail.Fields.Append "B_departid", adVarChar, 100
    rsdetail.Fields.Append "B_colorid", adVarChar, 100
    rsdetail.Fields.Append "B_ClientID", adVarChar, 100
    rsdetail.Fields.Append "B_itemid", adVarChar, 100
    rsdetail.Fields.Append "B_Waitfreight", adVarChar, 100
    rsdetail.Fields.Append "B_process", adVarChar, 100
    rsdetail.Fields.Append "B_processid", adVarChar, 100
    rsdetail.Fields.Append "B_repair", adVarChar, 100
    rsdetail.Fields.Append "B_repaircost", adVarChar, 100
    rsdetail.Fields.Append "B_cang", adVarChar, 100
    rsdetail.Fields.Append "B_Memo", adVarChar, 100

    rsdetail.Open
    
    TDBGrid1.DataSource = rsdetail
    setrsDetail
End Sub
Private Sub setrsDetail()
    setGridShow
'    TDBGrid1.Columns("B_HEX").Caption = "��ɫ����"
'    TDBGrid1.Columns("B_itemidb").Caption = ""
'    TDBGrid1.Columns("B_depart").Caption = ""
'    TDBGrid1.Columns("B_ColorName").Caption = ""
'    TDBGrid1.Columns("B_specification").Caption = ""
'    TDBGrid1.Columns("B_Color").Caption = ""
'    TDBGrid1.Columns("B_SeHao").Caption = ""
'    TDBGrid1.Columns("B_PIshu").Caption = ""
'    TDBGrid1.Columns("B_Weight").Caption = ""
'    TDBGrid1.Columns("B_Meter").Caption = ""
'    TDBGrid1.Columns("B_Ma").Caption = ""
'    TDBGrid1.Columns("B_ClientName").Caption = ""
'    TDBGrid1.Columns("B_DeliveryNote").Caption = ""
'    TDBGrid1.Columns("B_type").Caption = ""
'    TDBGrid1.Columns("B_Memo").Caption = ""
    TDBGrid1.Columns("B_repair").ValueItems.Presentation = dbgCheckBox
   TDBGrid1.Columns("B_Waitfreight").ValueItems.Presentation = dbgCheckBox
   
   
   
   
   TDBGrid1.Columns("B_process").Button = True
   TDBGrid1.Columns("B_depart").Locked = True
   TDBGrid1.Columns("B_ClientName").Locked = True
   
   TDBGrid1.Columns("B_HEX").Locked = True
   TDBGrid1.Columns("B_itemidb").Locked = True
'   TDBGrid1.Columns("B_depart").Locked = True
'   TDBGrid1.Columns("B_Color").Locked = True
'   TDBGrid1.Columns("B_specification").Locked = True
'   TDBGrid1.Columns("B_Color").Locked = True
   TDBGrid1.Columns("B_SeHao").Locked = True
   TDBGrid1.Columns("B_depart").Button = True
   TDBGrid1.Columns("B_ClientName").Button = True
'   TDBGrid1.Columns("B_Color").Button = True
   
   TDBGrid1.Columns("B_departid").Visible = False
   TDBGrid1.Columns("B_departid").Locked = True
   TDBGrid1.Columns("B_departid").AllowSizing = False
      TDBGrid1.Columns("B_colorid").Visible = False
   TDBGrid1.Columns("B_colorid").Locked = True
   TDBGrid1.Columns("B_colorid").AllowSizing = False
     TDBGrid1.Columns("B_ClientID").Visible = False
   TDBGrid1.Columns("B_ClientID").Locked = True
   TDBGrid1.Columns("B_ClientID").AllowSizing = False
   
        TDBGrid1.Columns("B_processid").Visible = False
   TDBGrid1.Columns("B_processid").Locked = True
   TDBGrid1.Columns("B_processid").AllowSizing = False
   
  
   
'   TDBGrid1.Columns("B_Hex").FetchStyle = True
   TDBGrid1.HoldFields
   TDBGrid1.MarqueeStyle = dbgHighlightRow
End Sub
Private Sub setGridShow()
    Dim cls1 As New clsGridShow
    With cls1
        .ObjectID = "11S049"
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
                strSQL = "update G_BLSField set B_GridWidth='" & dWidth & "' where B_ObjectID='11S049' and B_FieldName='" & szFieldName & "'"
                Gm.cnnTool.cnn.Execute strSQL
            End If
        End If
    Next
End Sub

Private Sub PushButton1_Click()
     Dim frm1 As New frmPopupDanWei
        frm1.ContactType = "��������"
        frm1.Show vbModal
        Originalsuppliers = frm1.clientid
        FlatEdit1.Text = frm1.ClientName
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
'������
Private Sub AddNew()
    Dim itemid As String
    Dim departid As String  'Ⱦ��id
    Dim colorid As String
    Dim hex As String
    Dim gg As String
    Dim khid As String
    Dim khname As String
    Dim itemidb As String
    Dim B_colorname As String
    
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim frm1 As New frmColorlastyear_Edit
'        If Len(dingdan) > 0 Then
'            frm1.FlatEdit1.Text = dingdan
'        End If
   
    frm1.Show vbModal
    If frm1.bsaved = True Then
'        itemid = frm1.itemid
        itemidb = frm1.itemidb
'        departid = frm1.B_Clientid
'        colorid = frm1.B_sid
'        hex = frm1.B_hex
'        gg = frm1.B_gg
'        khid = frm1.khid
'        khname = frm1.khname
'        dingdan = frm1.dingdan
         B_colorname = frm1.B_name
    Else
        Exit Sub
    End If
    Unload frm1
    
'
'    sql = "select * from G_BIlldetailcolor where B_itemid='" & itemid & "'"
'    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    rsdetail.AddNew
'    rsDetail!B_hex = hex
    rsdetail!B_ItemIDB = itemidb
'    rsDetail!B_depart = departid
    rsdetail!B_colorname = B_colorname
'    rsDetail!B_Specification = gg
'    rsDetail!B_Color = colorid
'    rsDetail!B_SeHao = rs!B_SeHao
'    rsDetail!B_departid = rs!B_depart
'    rsDetail!B_Colorid = rs!B_Color
'    rsDetail!B_ClientName = khname
'    rsDetail!B_Clientid = khid
'    rsDetail!B_Waitfreight = 0
    rsdetail.Update
    sumall
'    rsDetail.requery
    TDBGrid1.SetFocus
    TDBGrid1.Col = 7
End Sub
'���ݱ���
Private Sub save()
    Dim rrs As New RecordSet
    Dim strSQL As String

    Dim a As String
    Dim i As Long
    i = 1
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    rsdetail.MoveFirst
    Do While Not rsdetail.EOF
     
        If IIf(IsNull(rsdetail!B_PIShu), "", rsdetail!B_PIShu) = "" Or rsdetail!B_PIShu = 0 Then
            MsgBox "��" & i & "��ƥ������Ϊ�ջ���Ϊ0", vbInformation, "��ʾ"
            Exit Sub
        End If
        If IIf(IsNull(rsdetail!B_weight), "", rsdetail!B_weight) = "" Or rsdetail!B_weight = 0 Then
            MsgBox "��" & i & "�й��ﲻ��Ϊ�ջ���Ϊ0", vbInformation, "��ʾ"
            Exit Sub
        End If
        If rsdetail!B_color = "" Then
            MsgBox "��" & i & "����ɫ����Ϊ��", vbInformation, "��ʾ"
            Exit Sub
        End If
        rsdetail.movenext
        i = i + 1
    Loop
    
    
    If id <> "" Then
         strSQL = "select * from G_Billcolor where B_ID='" & id & "'"
        rrs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rrs.RecordCount > 0 Then
            savetoupdate
            Exit Sub
        Else
            id = ""
        End If
    End If
    
    Dim rs As New RecordSet
    Dim sql As String
    Dim sql1 As String
    sql = "select * from G_draftBillColor where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.AddNew
    rs!B_datecreate = Now
    rs.Update
    id = rs!B_id
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    
    sql1 = "exec usp_InsertColorOrder  '" & id & "','" & FlatEdit3.Text & "','" & a & "','CLC',"
    sql1 = sql1 & "'12B008','COL02 ','" & Gm.SysID.SystemUser & "','" & FlatEdit2.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit6.Text & "','" & FlatEdit4.Text & "','" & ComboBox1.Text & "','" & Originalsuppliers & "','" & fh & "','" & ComboBox2.Text & "','','','','" & "" & "','" & "" & "'"
    Gm.cnnTool.cnn.Execute sql1
    
    savedetail    '����������ϸ����
    
    sql = "delete from G_draftBillColor where B_itemid='" & id & "'"
   FlatEdit1 = ""
    FlatEdit2 = ""
    FlatEdit4 = ""
    FlatEdit5 = ""
    FlatEdit6 = ""
    id = ""
    cob1
    cob2
    fh = ""
    Originalsuppliers = ""
    FlatEdit3.Text = GetCodeID
    setRs
    printdetail = True
     Dim c As Long
    Dim b As Long
    a = 0
    b = 0
    
    TDBGrid1.Columns("B_Weight").FooterText = "" & c & ""
    TDBGrid1.Columns("B_PIshu").FooterText = "" & b & ""
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
        sql = "select * from G_draftBilldetailcolor where 1=1"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        rs.AddNew
        rs!B_datecreate = Now
        rs.Update
        item = rs!B_ItemID
'        If IIf(IsNull(rsDetail!B_repair), 0, rsDetail!B_repair) = 0 Then
'            rsDetail!B_repaircost = ""
'        End If
        sql2 = "exec usp_ColorlastyearDetail '" & item & "','" & id & "','" & rsdetail!B_ItemIDB & "','" & rsdetail!B_departid & "'"
        sql2 = sql2 & ",'" & rsdetail!B_colorname & "','" & rsdetail!B_specification & "','" & rsdetail!B_specification1 & "','" & rsdetail!B_colorid & "','" & rsdetail!B_SeHao & "'"
        sql2 = sql2 & ",'" & rsdetail!B_PIShu & "','" & rsdetail!B_weight & "','" & rsdetail!B_meter & "','" & rsdetail!B_Ma & "'"
        sql2 = sql2 & ",'" & rsdetail!B_Clientid & "','" & rsdetail!B_DeliveryNote & "','" & rsdetail!B_type & "','" & rsdetail!B_Waitfreight & "','" & rsdetail!B_processid & "','" & rsdetail!B_memo & "','" & "" & "','" & rsdetail!B_repaircost & "','" & rsdetail!B_color & "','" & rsdetail!B_Cang & "'"
        Gm.cnnTool.cnn.Execute sql2
        sql1 = "delete from G_draftBilldetailcolor where B_itemid='" & item & "'"
        Gm.cnnTool.cnn.Execute sql1
        rsdetail.movenext
    Loop
End Sub
'���д����������е���������
Private Sub saveadddetail()
    Dim rs As RecordSet
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim sql1 As String
    Dim item As String
    Dim sql2 As String

    Set rs = New RecordSet
    sql = "select * from G_draftBilldetailcolor where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.AddNew
    rs!B_datecreate = Now
    rs.Update
    item = rs!B_ItemID
'    If IIf(IsNull(rsDetail!B_repair), 0, rsDetail!B_repair) = 0 Then
'        rsDetail!B_repaircost = ""
'    End If
    sql2 = "exec usp_ColorlastyearDetail '" & item & "','" & id & "','" & rsdetail!B_ItemIDB & "','" & rsdetail!B_departid & "'"
    sql2 = sql2 & ",'" & rsdetail!B_colorname & "','" & rsdetail!B_specification & "','" & rsdetail!B_specification1 & "','" & rsdetail!B_colorid & "','" & rsdetail!B_SeHao & "'"
    sql2 = sql2 & ",'" & rsdetail!B_PIShu & "','" & rsdetail!B_weight & "','" & rsdetail!B_meter & "','" & rsdetail!B_Ma & "'"
    sql2 = sql2 & ",'" & rsdetail!B_Clientid & "','" & rsdetail!B_DeliveryNote & "','" & rsdetail!B_type & "','" & rsdetail!B_Waitfreight & "','" & rsdetail!B_processid & "','" & rsdetail!B_memo & "','" & "" & "','" & rsdetail!B_repaircost & "','" & rsdetail!B_color & "','" & rsdetail!B_Cang & "'"
    Gm.cnnTool.cnn.Execute sql2
    sql1 = "delete from G_draftBilldetailcolor where B_itemid='" & item & "'"
    Gm.cnnTool.cnn.Execute sql1

End Sub
'����ɾ��
Private Sub DeleteHang()
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
    
    sql2 = "select * from G_billdetailcolor where B_ID='" & id & "'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs2.RecordCount = 1 Then
    
         sql = "select * from G_billdetailcolor where B_itemid='" & rsdetail!B_ItemID & "'"
         rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
         If rs.RecordCount > 0 Then
            If MsgBox("�˵���ֻʣһ������,ɾ����ȫ��ɾ�����Ƿ�ɾ��", vbInformation + vbYesNo + vbDefaultButton2, "��ʾ") = vbYes Then
                    sql1 = "delete from G_billdetailcolor where B_itemid='" & rsdetail!B_ItemID & "'"
                    Gm.cnnTool.cnn.Execute sql1
                        sql3 = "delete from G_billcolor where B_id='" & id & "'"
                    Gm.cnnTool.cnn.Execute sql3
                        FlatEdit1 = ""
                        FlatEdit2 = ""
                        FlatEdit4 = ""
                        FlatEdit5 = ""
                        FlatEdit6 = ""
                        id = ""
                        cob1
                        cob2
                        fh = ""
                        Originalsuppliers = ""
                        FlatEdit3.Text = GetCodeID
                        setRs
            Else
                Exit Sub
            End If
        End If
    Else
        sql1 = "delete from G_billdetailcolor where B_itemid='" & rsdetail!B_ItemID & "'"
        Gm.cnnTool.cnn.Execute sql1
    End If
    
    rsdetail.delete
    If TDBGrid1.ApproxCount > 0 Then
        rsdetail.MoveFirst
    End If
    
End Sub



Private Sub PushButton2_Click()
       Dim frm1 As New frmpopupEmploy
        frm1.ContactType = "ɫ������װж��"
        frm1.Show vbModal
        fh = frm1.clientid
        FlatEdit5.Text = frm1.ClientName
        Unload frm1
End Sub

Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, _
    ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, _
    ByVal CellStyle As TrueOleDBGrid80.StyleDisp)

'    Dim ys As RGB
    On Error Resume Next
    Debug.Print TDBGrid1.Columns("B_Hex").CellValue(bookmark)
    CellStyle.BackColor = TDBGrid1.Columns("B_Hex").CellValue(bookmark)
     CellStyle.ForeColor = TDBGrid1.Columns("B_Hex").CellValue(bookmark)
End Sub
Private Sub TDBGrid1_BeforeColUpdate(ByVal colIndex As Integer, OldValue As Variant, Cancel As Integer)
    If colIndex = TDBGrid1.Columns("B_PIshu").colIndex Then
        TDBGrid1.Columns("B_PIshu").Value = Abs(Val(TDBGrid1.Columns("B_PIshu").Value))
    End If
       If colIndex = TDBGrid1.Columns("B_Weight").colIndex Then
        TDBGrid1.Columns("B_Weight").Value = Abs(Val(TDBGrid1.Columns("B_Weight").Value))
    End If
       If colIndex = TDBGrid1.Columns("B_Meter").colIndex Then
        TDBGrid1.Columns("B_Meter").Value = Abs(Val(TDBGrid1.Columns("B_Meter").Value))
    End If
       If colIndex = TDBGrid1.Columns("B_Ma").colIndex Then
        TDBGrid1.Columns("B_Ma").Value = Abs(Val(TDBGrid1.Columns("B_Ma").Value))
    End If
    
    If colIndex = TDBGrid1.Columns("B_repaircost").colIndex Then
        TDBGrid1.Columns("B_repaircost").Value = Abs(Val(TDBGrid1.Columns("B_repaircost").Value))
    End If
    
    
    
    sumall
End Sub
Private Sub saveandprint()
      Dim rrs As New RecordSet
    Dim strSQL As String


    Dim a As String
    Dim i As Long
    i = 1
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    rsdetail.MoveFirst
    Do While Not rsdetail.EOF
     
        If IIf(IsNull(rsdetail!B_PIShu), "", rsdetail!B_PIShu) = "" Or rsdetail!B_PIShu = 0 Then
            MsgBox "��" & i & "��ƥ������Ϊ�ջ���Ϊ0", vbInformation, "��ʾ"
            Exit Sub
        End If
        If IIf(IsNull(rsdetail!B_weight), "", rsdetail!B_weight) = "" Or rsdetail!B_weight = 0 Then
            MsgBox "��" & i & "�й��ﲻ��Ϊ�ջ���Ϊ0", vbInformation, "��ʾ"
            Exit Sub
        End If
        If rsdetail!B_color = "" Then
            MsgBox "��" & i & "����ɫ����Ϊ��", vbInformation, "��ʾ"
            Exit Sub
        End If
        rsdetail.movenext
        i = i + 1
    Loop
    
    strSQL = "select * from G_Billcolor where B_ID='" & id & "'"
    rrs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rrs.RecordCount > 0 Then
        savetoupdate1
        
    Else

        Dim rs As New RecordSet
        Dim sql As String
        Dim sql1 As String
        sql = "select * from G_draftBillColor where 1=1"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        rs.AddNew
        rs!B_datecreate = Now
        rs.Update
        id = rs!B_id
        a = Format(DTPicker1.Value, "YYYY-MM-DD")
        
        sql1 = "exec usp_InsertColorOrder  '" & id & "','" & FlatEdit3.Text & "','" & a & "','CLC',"
        sql1 = sql1 & "'12B008','COL02 ','" & Gm.SysID.SystemUser & "','" & FlatEdit2.Text & "'"
        sql1 = sql1 & ",'" & FlatEdit6.Text & "','" & FlatEdit4.Text & "','" & ComboBox1.Text & "','" & Originalsuppliers & "','" & fh & "','" & ComboBox2.Text & "','','','','" & "" & "','" & "" & "'"
        Gm.cnnTool.cnn.Execute sql1
        savedetail
        sql = "delete from G_draftBillColor where B_itemid='" & id & "'"
    End If
   
        Dim rs3 As New RecordSet
        Dim sql3 As String
        sql3 = "exec usp_Colorlastyearreport '" & id & "','" & Gm.SysID.SystemUserName & "'"
        rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        Dim frm1 As New frmModBLRPreviewOri
        Set frm1.RecordSet = rs3.Clone
            
        frm1.ObjectID = "22B066"
        frm1.Show
         FlatEdit1 = ""
    FlatEdit2 = ""
    FlatEdit4 = ""
    FlatEdit5 = ""
    FlatEdit6 = ""
    id = ""
    cob1
    cob2
    fh = ""
    Originalsuppliers = ""
    FlatEdit3.Text = GetCodeID
    setRs
    Dim c As Long
    Dim b As Long
    a = 0
    b = 0
    
    TDBGrid1.Columns("B_Weight").FooterText = "" & c & ""
    TDBGrid1.Columns("B_PIshu").FooterText = "" & b & ""
End Sub
Private Sub sumall()
     Dim rs As New RecordSet
    Dim a As Double
    Dim b As Long
    Dim c As Double
    Dim d As Double
    Dim e As String
    Dim f As String
    Dim g As String
    If rsdetail.RecordCount <= 0 Then
        Exit Sub
    End If
    a = 0
    b = 0
    c = 0
    d = 0
    Set rs = rsdetail.Clone
    rs.MoveFirst
    Do While Not rs.EOF
        a = a + IIf(IsNull(Val(rs!B_weight)), 0, Val(rs!B_weight))
        b = b + IIf(IsNull(rs!B_PIShu), 0, rs!B_PIShu)
        c = c + IIf(IsNull(Val(rs!B_meter)), 0, Val(rs!B_meter))
        d = d + IIf(IsNull(rs!B_Ma), 0, rs!B_Ma)
        rs.movenext
    Loop
    e = Format(a, "0.0")
    f = Format(c, "0.0")
    g = Format(d, "0.0")
    TDBGrid1.Columns("B_hex").FooterText = "�ϼ�"
    TDBGrid1.Columns("B_Weight").FooterText = "" & e & ""
    TDBGrid1.Columns("B_PIshu").FooterText = "" & b & ""
    TDBGrid1.Columns("B_Meter").FooterText = "" & f & ""
    TDBGrid1.Columns("B_Ma").FooterText = "" & g & ""
End Sub
'��ӡʱ�����޸�
Private Sub savetoupdate1()
    Dim sql3 As String
    Dim rs3 As RecordSet
    Dim sql2 As String
    Dim sql1 As String
    Dim a As String
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    sql1 = "exec usp_Colororder_Update  '" & id & "','" & FlatEdit3.Text & "','" & a & "','WHT',"
    sql1 = sql1 & "'12B008','COL02 ','" & Gm.SysID.SystemUser & "','" & FlatEdit2.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit6.Text & "','" & FlatEdit4.Text & "','" & ComboBox1.Text & "','" & Originalsuppliers & "','" & fh & "','" & ComboBox2.Text & "','','',''"
    Gm.cnnTool.cnn.Execute sql1
    rsdetail.MoveFirst
    Do While Not rsdetail.EOF
'
'        If IIf(IsNull(rsDetail!B_repair), 0, rsDetail!B_repair) = 0 Then
'            rsDetail!B_repaircost = ""
'        End If
        Set rs3 = New RecordSet
        sql3 = "select * from G_billdetailcolor where B_itemid ='" & rsdetail!B_ItemID & "'"
        rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs3.RecordCount > 0 Then
            sql2 = "exec usp_Colorlastyeardetail_Update '" & rsdetail!B_ItemID & "','" & id & "','" & rsdetail!B_ItemIDB & "','" & rsdetail!B_departid & "'"
            sql2 = sql2 & ",'" & rsdetail!B_colorname & "','" & rsdetail!B_specification & "','" & rsdetail!B_specification1 & "','" & rsdetail!B_colorid & "','" & rsdetail!B_SeHao & "'"
            sql2 = sql2 & ",'" & rsdetail!B_PIShu & "','" & rsdetail!B_weight & "','" & rsdetail!B_meter & "','" & rsdetail!B_Ma & "'"
            sql2 = sql2 & ",'" & rsdetail!B_Clientid & "','" & rsdetail!B_DeliveryNote & "','" & rsdetail!B_type & "','" & rsdetail!B_processid & "','" & rsdetail!B_memo & "','" & "" & "','" & rsdetail!B_repaircost & "','" & rsdetail!B_color & "','" & rsdetail!B_Cang & "'"
            Gm.cnnTool.cnn.Execute sql2
        Else
            saveadddetail
        End If
        rsdetail.movenext
    Loop
End Sub


Private Sub savetoupdate()
    Dim sql3 As String
    Dim rs3 As RecordSet
    Dim sql2 As String
    Dim sql1 As String
    Dim a As String
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    sql1 = "exec usp_Colororder_Update  '" & id & "','" & FlatEdit3.Text & "','" & a & "','WHT',"
    sql1 = sql1 & "'12B008','COL02 ','" & Gm.SysID.SystemUser & "','" & FlatEdit2.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit6.Text & "','" & FlatEdit4.Text & "','" & ComboBox1.Text & "','" & Originalsuppliers & "','" & fh & "','" & ComboBox2.Text & "','','',''"
    Gm.cnnTool.cnn.Execute sql1
    rsdetail.MoveFirst
    Do While Not rsdetail.EOF
'        If IIf(IsNull(rsDetail!B_repair), 0, rsDetail!B_repair) = 0 Then
'            rsDetail!B_repaircost = ""
'        End If
        Set rs3 = New RecordSet
        sql3 = "select * from G_billdetailcolor where B_itemid ='" & rsdetail!B_ItemID & "'"
        rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs3.RecordCount > 0 Then
            sql2 = "exec usp_Colorlastyeardetail_Update '" & rsdetail!B_ItemID & "','" & id & "','" & rsdetail!B_ItemIDB & "','" & rsdetail!B_departid & "'"
             sql2 = sql2 & ",'" & rsdetail!B_colorname & "','" & rsdetail!B_specification & "','" & rsdetail!B_specification1 & "','" & rsdetail!B_colorid & "','" & rsdetail!B_SeHao & "'"
             sql2 = sql2 & ",'" & rsdetail!B_PIShu & "','" & rsdetail!B_weight & "','" & rsdetail!B_meter & "','" & rsdetail!B_Ma & "'"
             sql2 = sql2 & ",'" & rsdetail!B_Clientid & "','" & rsdetail!B_DeliveryNote & "','" & rsdetail!B_type & "','" & rsdetail!B_processid & "','" & rsdetail!B_memo & "','" & "" & "','" & rsdetail!B_repaircost & "','" & rsdetail!B_color & "','" & rsdetail!B_Cang & "'"
             Gm.cnnTool.cnn.Execute sql2
        Else
            saveadddetail
        End If
        rsdetail.movenext
    Loop
     FlatEdit1 = ""
    FlatEdit2 = ""
    FlatEdit4 = ""
    FlatEdit5 = ""
    FlatEdit6 = ""
    id = ""
    cob1
    cob2
    fh = ""
    Originalsuppliers = ""
    FlatEdit3.Text = GetCodeID
    setRs
End Sub
Private Sub TDBGrid1_ButtonClick(ByVal colIndex As Integer)
    
If TDBGrid1.Columns("B_depart").colIndex = colIndex Then
     Dim frm2 As New frmPopupDanWei
    frm2.ContactType = "Ⱦ��"
    frm2.Show vbModal
    rsdetail!B_departid = frm2.clientid
    rsdetail!B_depart = frm2.ClientName
    Unload frm2
    End If
    
    If TDBGrid1.Columns("B_ClientName").colIndex = colIndex Then
     Dim frm1 As New frmPopupDanWei
    frm1.ContactType = "�ͻ�"
    frm1.Show vbModal
    rsdetail!B_Clientid = frm1.clientid
    rsdetail!B_ClientName = frm1.ClientName
    Unload frm1
    End If
    
    If TDBGrid1.Columns("B_color").colIndex = colIndex Then
     Dim frm3 As New frmpopupColor
    frm3.Show vbModal
    rsdetail!B_colorid = frm3.colorid
    rsdetail!B_color = frm3.colorname
    Unload frm3
    End If
End Sub


Private Sub MoveFirst()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select top 1 B_id,B_CodeID,B_Date,B_Freight,B_PlateNumber,B_PNumber,B_payment,B_Shipment,b.B_ClientName,B_Hand,c.B_Name,B_cope"
    sql = sql & " from G_Billcolor a left outer join G_ContactCompany b on a.B_Shipment=b.B_ClientID"
    sql = sql & " left outer join G_Employee c on a.B_Hand=c.B_SID where B_ObjectID='12B008' and  B_BillType='COL02'"
    sql = sql & "order by B_id"
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
    openbill
End Sub

Private Sub MovePrevious()
    Dim rs As New RecordSet
    Dim sql As String
    If id = "" Then
        movelast
        Exit Sub
    End If
    sql = "select top 1 a.B_id,B_CodeID,B_Date,B_Freight,B_PlateNumber,B_PNumber,B_payment,B_Shipment,b.B_ClientName,B_Hand,c.B_Name,B_cope"
    sql = sql & " from G_Billcolor a left outer join G_ContactCompany b on a.B_Shipment=b.B_ClientID"
    sql = sql & " left outer join G_Employee c on a.B_Hand=c.B_SID where B_ObjectID='12B008' and  B_BillType='COL02' and B_ID<'" & id & "'"
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
    openbill
End Sub
Private Sub movenext()
    Dim rs As New RecordSet
    Dim sql As String
    If id = "" Then
        movelast
        Exit Sub
    End If
    sql = "select top 1 B_id,B_CodeID,B_Date,B_Freight,B_PlateNumber,B_PNumber,B_payment,B_Shipment,b.B_ClientName,B_Hand,c.B_Name,B_cope"
    sql = sql & " from G_Billcolor a left outer join G_ContactCompany b on a.B_Shipment=b.B_ClientID"
    sql = sql & " left outer join G_Employee c on a.B_Hand=c.B_SID where B_ObjectID='12B008' and  B_BillType='COL02' and B_ID>'" & id & "'"
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
    openbill
End Sub


Private Sub movelast()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select top 1 B_id,B_CodeID,B_Date,B_Freight,B_PlateNumber,B_PNumber,B_payment,B_Shipment,b.B_ClientName,B_Hand,c.B_Name,B_cope"
    sql = sql & " from G_Billcolor a left outer join G_ContactCompany b on a.B_Shipment=b.B_ClientID"
    sql = sql & " left outer join G_Employee c on a.B_Hand=c.B_SID where B_ObjectID='12B008' and  B_BillType='COL02'"
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
    openbill
End Sub

Public Sub openbill()
   Dim sql As String
   Dim rs As New RecordSet

   sql = "exec usp_colorlastyearopenbill '" & id & "'"
   rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
   Debug.Print rs.RecordCount
   setRs
   Do While Not rs.EOF
        rsdetail.AddNew
'            rsdetail!B_hex = IIf(IsNull(rs!B_hex), "", rs!B_hex)
            rsdetail!B_ItemIDB = rs!B_ItemIDB
            
            rsdetail!B_depart = IIf(IsNull(rs!B_ClientName), "", rs!B_ClientName)
            rsdetail!B_colorname = rs!B_GoodsNameAlias
            rsdetail!B_specification = IIf(IsNull(rs!B_Width), "", rs!B_Width)
            rsdetail!B_specification1 = IIf(IsNull(rs!B_weight), "", rs!B_weight)
            rsdetail!B_color = IIf(IsNull(rs!B_name), "", rs!B_name)
            rsdetail!B_SeHao = rs!B_SeHao
            rsdetail!B_PIShu = rs!B_ps
            rsdetail!B_weight = rs!B_kg
            rsdetail!B_meter = rs!B_meter
            rsdetail!B_Ma = rs!B_BoxQty
            rsdetail!B_ClientName = IIf(IsNull(rs!B_Client), "", rs!B_Client)
            rsdetail!B_DeliveryNote = rs!B_TransfersID
            rsdetail!B_type = rs!B_type
            
            rsdetail!B_departid = IIf(IsNull(rs!B_depart), "", rs!B_depart)
            rsdetail!B_colorid = IIf(IsNull(rs!B_color), "", rs!B_color)
            rsdetail!B_Clientid = rs!B_Clientid
            rsdetail!B_ItemID = rs!B_ItemID
            rsdetail!B_memo = rs!B_MemoDetail
            rsdetail!B_process = IIf(IsNull(rs!B_processid), "", rs!B_processid)
            rsdetail!B_processid = IIf(IsNull(rs!B_process), "", rs!B_process)
            rsdetail!B_repaircost = IIf(IsNull(rs!B_returnprocesscost), "", rs!B_returnprocesscost)
            rsdetail!B_Cang = IIf(IsNull(rs!B_Cang), "", rs!B_Cang)
       rsdetail.Update
       rs.movenext
   Loop
   sumall
End Sub


