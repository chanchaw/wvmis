VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmOriginaltransfers 
   Caption         =   "ԭ�ϵ���"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16200
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
   ScaleHeight     =   7455
   ScaleWidth      =   16200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16200
      _LayoutVersion  =   1
      _ExtentX        =   28575
      _ExtentY        =   13150
      _DataPath       =   ""
      Bands           =   "frmOriginaltransfers.frx":0000
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5775
         Left            =   300
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   14775
         _cx             =   26061
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
         _GridInfo       =   $"frmOriginaltransfers.frx":51E4
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   1185
            Left            =   30
            ScaleHeight     =   1185
            ScaleWidth      =   14715
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   30
            Width           =   14715
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   375
               Left            =   4680
               TabIndex        =   3
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               _Version        =   393216
               Format          =   211419137
               CurrentDate     =   43106
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   9240
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
               Width           =   1455
               _Version        =   1048578
               _ExtentX        =   2566
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
               Left            =   12480
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
               Width           =   1455
               _Version        =   1048578
               _ExtentX        =   2566
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
               Left            =   14400
               TabIndex        =   22
               Top             =   210
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit8 
               Height          =   375
               Left            =   14400
               TabIndex        =   24
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
            Begin XtremeSuiteControls.Label Label12 
               Height          =   255
               Left            =   13200
               TabIndex        =   25
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
            Begin XtremeSuiteControls.Label Label9 
               Height          =   255
               Left            =   13200
               TabIndex        =   23
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
            Begin XtremeSuiteControls.Label Label8 
               Height          =   255
               Left            =   10320
               TabIndex        =   20
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
               TabIndex        =   19
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
               Left            =   3480
               TabIndex        =   18
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
            Begin XtremeSuiteControls.Label Label4 
               Height          =   255
               Left            =   300
               TabIndex        =   17
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
            Begin XtremeSuiteControls.Label Label3 
               Height          =   255
               Left            =   3480
               TabIndex        =   16
               Top             =   240
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
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
            Begin XtremeSuiteControls.Label Label1 
               Height          =   315
               Left            =   7080
               TabIndex        =   15
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
               TabIndex        =   14
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
               TabIndex        =   13
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
            TabIndex        =   21
            Top             =   1245
            Width           =   14715
            _ExtentX        =   25956
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
Attribute VB_Name = "frmOriginaltransfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cls1 As clsGridShow
Public rsdetail As RecordSet
Private Const theObjectID As String = "12B004"  '�������ݶ�����
Private theBLTool As New clsAutoCreateBL
Public mvarObjectID As String
Public dingdan As String
'Private client As String

Public id As String
Public fh As String
Public Originalsuppliers As String

Private printdetail As Boolean    '������ɽ��д�ӡ����֤
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
            addwhite
'            addnew
'            rsdetail.addnew
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
    FlatEdit8 = ""
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
    TDBGrid1.Columns("B_itemidb").FooterText = "�ϼ�"
    TDBGrid1.Columns("B_qty").FooterText = "" & a & ""
    TDBGrid1.Columns("B_sum").FooterText = "" & b & ""
End Sub


Private Sub Form_Load()
    InitFrm
    printdetail = False
    DTPicker1 = Now
    cob2
    TDBGrid1.HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid1.Splits(0).HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid1.HighlightRowStyle.ForeColor = &H80000008
    TDBGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
    TDBGrid1.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell
    
     Dim sql As String
    Dim rs As New RecordSet
    sql = "select * from G_Config_OneInt where B_groupName='ԭ�ϵ���¼�뷽ʽ'"
    Debug.Print sql
   rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
   If rs!B_Value = 0 Then
        TDBGrid1.Columns("B_itemidb").Locked = True
        TDBGrid1.Columns("B_specification").Locked = True
       TDBGrid1.Columns("B_process").Locked = True
       TDBGrid1.Columns("B_transfers").Locked = True
   End If
   bol = False
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
    FlatEdit3.Text = GetCodeID
    setRs
    cob1
    DTPicker1.Value = Now
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
    rsdetail.Fields.Append "B_itemid", adVarChar, 100
    rsdetail.Fields.Append "B_itemidb", adVarChar, 100
    rsdetail.Fields.Append "B_process", adVarChar, 100
    rsdetail.Fields.Append "B_ColorName", adVarChar, 100
    rsdetail.Fields.Append "B_specification", adVarChar, 100
    rsdetail.Fields.Append "B_boxqty", adVarChar, 100
    rsdetail.Fields.Append "B_qty", adVarChar, 100
    rsdetail.Fields.Append "B_price", adVarChar, 100
    rsdetail.Fields.Append "B_Sum", adVarChar, 100
    rsdetail.Fields.Append "B_Cang", adVarChar, 100
    rsdetail.Fields.Append "B_transfers", adVarChar, 100
    rsdetail.Fields.Append "B_DeliveryNote", adVarChar, 100
    rsdetail.Fields.Append "B_Memo", adVarChar, 100
    
    rsdetail.Fields.Append "B_Colorid", adVarChar, 100
    rsdetail.Fields.Append "B_processid", adVarChar, 100
    rsdetail.Fields.Append "B_transfersid", adVarChar, 100
    rsdetail.Open
    
    TDBGrid1.DataSource = rsdetail
    setrsDetail
End Sub
Private Sub setrsDetail()
    setGridShow
    TDBGrid1.Columns("B_itemid").Locked = True
'   TDBGrid1.Columns("B_itemidb").Locked = True
'   TDBGrid1.Columns("B_suppliers").Locked = True
   TDBGrid1.Columns("B_ColorName").Locked = True
'   TDBGrid1.Columns("B_specification").Locked = True
'   TDBGrid1.Columns("B_qty").Locked = True
'   TDBGrid1.Columns("B_price").Locked = True
   TDBGrid1.Columns("B_Sum").Locked = True
'   TDBGrid1.Columns("B_Cang").Locked = True
'   TDBGrid1.Columns("B_Memo").Locked = True
   TDBGrid1.Columns("B_ColorName").Button = True
   
   TDBGrid1.Columns("B_transfersid").Visible = False
   TDBGrid1.Columns("B_transfersid").Locked = True
   TDBGrid1.Columns("B_transfersid").AllowSizing = False
      TDBGrid1.Columns("B_processid").Visible = False
   TDBGrid1.Columns("B_processid").Locked = True
   TDBGrid1.Columns("B_processid").AllowSizing = False
   TDBGrid1.Columns("B_colorid").Visible = False
   TDBGrid1.Columns("B_colorid").Locked = True
   TDBGrid1.Columns("B_colorid").AllowSizing = False
   TDBGrid1.Columns("B_process").Button = True
   TDBGrid1.Columns("B_transfers").Button = True
'     TDBGrid1.Columns("B_itemidb").Visible = False
'   TDBGrid1.Columns("B_itemidb").Locked = True
'   TDBGrid1.Columns("B_itemidb").AllowSizing = False
   TDBGrid1.Columns("B_qty").NumberFormat = "0.00"
   TDBGrid1.HoldFields
   TDBGrid1.MarqueeStyle = dbgHighlightRow
End Sub
Private Sub setGridShow()
    Dim cls1 As New clsGridShow
    With cls1
        .ObjectID = "11S036"
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
                strSQL = "update G_BLSField set B_GridWidth='" & dWidth & "' where B_ObjectID='11S036' and B_FieldName='" & szFieldName & "'"
                Gm.cnnTool.cnn.Execute strSQL
            End If
        End If
    Next
End Sub



Private Sub PushButton1_Click()
     Dim frm1 As New frmPopupDanWei
        frm1.Caption = "�˷�"
        frm1.ContactType = "��������"
        frm1.Show vbModal
        Originalsuppliers = frm1.Clientid
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
Private Sub addwhite()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "select * from G_Config_OneInt where B_groupName='ԭ�ϵ���¼�뷽ʽ'"
    Debug.Print sql
   rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs!B_Value = 0 Then
        addnew
    Else
        rsdetail.addnew
    End If
End Sub
'������
Private Sub addnew()
    Dim itemidb As String
    Dim supplierid As String
    Dim supperliername As String
    Dim B_sid As String
    Dim B_name As String
    Dim B_gg As String
    Dim B_qty As String
    Dim B_price As String
    Dim B_memo As String
    
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim frm1 As New frmOriginalOrderInset_Edit
        If Len(dingdan) > 0 Then
            frm1.FlatEdit1.Text = dingdan
        End If
    frm1.Show vbModal
    If frm1.bsaved = True Then
'        itemidb = frm1.itemidb
''        supplierid = frm1.supplierid
''        supperliername = frm1.supperliername
'        B_sid = frm1.B_sid
'        B_name = frm1.B_name
'        B_gg = frm1.B_gg
'        B_qty = frm1.B_qty
'        B_price = 0
'        B_memo = frm1.B_memo
        Dim tdbgRow As Variant
        For Each tdbgRow In frm1.TDBGrid1.SelBookmarks
            frm1.rss.Bookmark = tdbgRow
                rsdetail.addnew
                rsdetail!B_ItemIDB = frm1.rss!B_OrderCode
                rsdetail!B_colorname = frm1.rss!B_name
                rsdetail!B_colorid = frm1.rss!B_sid
                rsdetail!B_specification = frm1.rss!B_specification
'                rsDetail!B_qty = frm1.rss!sum2
'                rsDetail!B_price = 0
'                rsDetail!B_memo = frm1.rss!B_MemoDetail
'                rsDetail!B_sum = Val(B_qty) * Val(B_price)
                rsdetail.Update
        Next
    Else
        Exit Sub
    End If
    Unload frm1
'    rsDetail.AddNew
'        rsDetail!B_ItemIDB = itemidb
''        rsDetail!B_suppliers = supperliername
''        rsDetail!B_suppliersid = supplierid
'        rsDetail!B_colorname = B_name
'        rsDetail!B_Colorid = B_sid
'        rsDetail!B_Specification = B_gg
'        rsDetail!B_qty = B_qty
'        rsDetail!B_price = B_price
'        rsDetail!B_memo = B_memo
'        rsDetail!B_sum = Val(B_qty) * Val(B_price)
'    rsDetail.Update
    sumall
End Sub
'���ݱ���
Private Sub save()
    Dim rrs As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet

    
    Dim i As Long
    i = 1
    Dim a As String
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    rsdetail.MoveFirst
    Do While Not rsdetail.EOF
     
        If rsdetail!B_price <= 0 Or rsdetail!B_price = "" Then
            MsgBox "��" & i & "�в���Ϊ0����Ϊ��", vbInformation, "��ʾ"
            Exit Sub
        End If
        If rsdetail!B_process = "" Then
            MsgBox "��" & i & "�мӹ�������Ϊ��", vbInformation, "��ʾ"
            Exit Sub
        End If
        If rsdetail!B_transfers = "" Then
            MsgBox "��" & i & "�е����ӹ�������Ϊ��", vbInformation, "��ʾ"
            Exit Sub
        End If
        If rsdetail!B_colorid = "" Then
            MsgBox "��" & i & "��ԭ�ϲ���Ϊ��", vbInformation, "��ʾ"
            Exit Sub
        End If
     
        rsdetail.movenext
        i = i + 1
    Loop
    sql2 = "select * from G_billyarn where B_id='" & id & "'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs2.RecordCount > 0 Then
            savetoupdate
            printdetail = True
            Exit Sub
    End If
    
    
    
    Dim rs As New RecordSet
    Dim sql As String
    Dim sql1 As String
    sql = "select * from G_draftBillyarn where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.addnew
    rs!B_datecreate = Now
    rs.Update
    id = rs!B_id
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    
    sql1 = "exec usp_Originalorderinsert  '" & id & "','" & FlatEdit3.Text & "','" & a & "',"
    sql1 = sql1 & "'12B004','YARN09','" & Gm.SysID.SystemUser & "','" & FlatEdit2.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit6.Text & "','" & FlatEdit4.Text & "','" & ComboBox1.Text & "','" & Originalsuppliers & "','" & fh & "','" & ComboBox2.Text & "','" & FlatEdit8.Text & "'"
    Gm.cnnTool.cnn.Execute sql1
    savedetail
      '���е������
   setAudit (0)
    sql = "delete from G_draftBillyarn where B_itemid='" & id & "'"
     FlatEdit1 = ""
    FlatEdit2 = ""
    FlatEdit4 = ""
    FlatEdit5 = ""
    FlatEdit6 = ""
    FlatEdit8 = ""
    id = ""
    cob1
    cob2
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
        sql = "select * from G_draftBilldetailyarn where 1=1"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        rs.addnew
        rs!B_datecreate = Now
        rs.Update
        item = rs!B_itemid
        
        sql2 = "exec usp_Originaltransfers '" & item & "','" & id & "','" & rsdetail!B_ItemIDB & "','" & rsdetail!B_processid & "'"
        sql2 = sql2 & ",'" & rsdetail!B_colorid & "','" & rsdetail!B_specification & "','" & rsdetail!B_qty & "','" & rsdetail!B_price & "'"
        sql2 = sql2 & ",'" & rsdetail!B_Sum & "','" & rsdetail!B_Cang & "','" & rsdetail!B_memo & "','" & rsdetail!B_TransfersID & "','" & rsdetail!B_BoxQty & "','" & rsdetail!B_DeliveryNote & "'"
        
        Gm.cnnTool.cnn.Execute sql2
        sql1 = "delete from G_draftBilldetailyarn where B_itemid='" & item & "'"
        Gm.cnnTool.cnn.Execute sql1
        rsdetail.movenext
    Loop
End Sub
'���б����޸�ʱ���������ݽ��б���
Private Sub saveadddetail()
     Dim rs As New RecordSet
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim sql1 As String
    Dim item As String
    Dim sql2 As String

    sql = "select * from G_draftBilldetailyarn where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.addnew
    rs!B_datecreate = Now
    rs.Update
    item = rs!B_itemid
    
    sql2 = "exec usp_Originaltransfers '" & item & "','" & id & "','" & rsdetail!B_ItemIDB & "','" & rsdetail!B_processid & "'"
    sql2 = sql2 & ",'" & rsdetail!B_colorid & "','" & rsdetail!B_specification & "','" & rsdetail!B_qty & "','" & rsdetail!B_price & "'"
    sql2 = sql2 & ",'" & rsdetail!B_Sum & "','" & rsdetail!B_Cang & "','" & rsdetail!B_memo & "','" & rsdetail!B_TransfersID & "','" & rsdetail!B_BoxQty & "','" & rsdetail!B_DeliveryNote & "'"
    
    Gm.cnnTool.cnn.Execute sql2
    sql1 = "delete from G_draftBilldetailyarn where B_itemid='" & item & "'"
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
    
    sql2 = "select * from G_billdetailyarn where B_ID='" & id & "'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs2.RecordCount = 1 Then
    
         sql = "select * from G_billdetailyarn where B_itemid='" & rsdetail!B_itemid & "'"
         rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
         If rs.RecordCount > 0 Then
            If MsgBox("�˵���ֻʣһ������,ɾ����ȫ��ɾ�����Ƿ�ɾ��", vbInformation + vbYesNo + vbDefaultButton2, "��ʾ") = vbYes Then
                          Dim c As New Clsfreight
                    If c.Freight(FlatEdit3.Text) = False Then
                        MsgBox "�˵����������˷�,�޷�ɾ��", vbInformation, "��ʾ"
                        Exit Sub
                    End If
                    sql1 = "delete from G_billdetailyarn where B_itemid='" & rsdetail!B_itemid & "'"
                    Gm.cnnTool.cnn.Execute sql1
                        sql3 = "delete from G_billyarn where B_id='" & id & "'"
                    Gm.cnnTool.cnn.Execute sql3
                        FlatEdit1 = ""
                        FlatEdit2 = ""
                        FlatEdit4 = ""
                        FlatEdit5 = ""
                        FlatEdit6 = ""
                        FlatEdit8 = ""
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
        sql1 = "delete from G_billdetailyarn where B_itemid='" & rsdetail!B_itemid & "'"
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
        fh = frm1.Clientid
        FlatEdit5.Text = frm1.ClientName
        Unload frm1
End Sub

Private Sub PushButton3_Click()
  
    Dim frm1 As New frmOriginalorderselect
    If TDBGrid1.ApproxCount > 0 Then
        frm1.FlatEdit1.Text = rsdetail!B_ItemIDB
    End If
    frm1.Show vbModal
    
End Sub

Private Sub TDBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If ColIndex = TDBGrid1.Columns("B_price").ColIndex Then
        TDBGrid1.Columns("B_price").Value = Abs(Val(TDBGrid1.Columns("B_price").Value))
        TDBGrid1.Columns("B_sum").Value = Abs(Val(TDBGrid1.Columns("B_qty").Value) * Val(TDBGrid1.Columns("B_price").Value))
    End If

     If ColIndex = TDBGrid1.Columns("B_qty").ColIndex Then
        TDBGrid1.Columns("B_qty").Value = Abs(Val(TDBGrid1.Columns("B_qty").Value))
        TDBGrid1.Columns("B_sum").Value = Abs(Val(TDBGrid1.Columns("B_qty").Value) * Val(TDBGrid1.Columns("B_price").Value))
    End If
    If ColIndex = TDBGrid1.Columns("B_boxqty").ColIndex Then
        TDBGrid1.Columns("B_boxqty").Value = Val(TDBGrid1.Columns("B_boxqty").Value)
    End If
    sumall
End Sub
Private Sub saveandprint()
       Dim rrs As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet

    
    Dim i As Long
    i = 1
    Dim a As String
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    rsdetail.MoveFirst
    Do While Not rsdetail.EOF
     
        If rsdetail!B_price <= 0 Or rsdetail!B_price = "" Then
            MsgBox "��" & i & "�в���Ϊ0����Ϊ��", vbInformation, "��ʾ"
            Exit Sub
        End If
         If rsdetail!B_process = "" Then
            MsgBox "��" & i & "�мӹ�������Ϊ��", vbInformation, "��ʾ"
            Exit Sub
        End If
          If rsdetail!B_transfers = "" Then
            MsgBox "��" & i & "�е����ӹ�������Ϊ��", vbInformation, "��ʾ"
            Exit Sub
        End If
        If rsdetail!B_colorid = "" Then
            MsgBox "��" & i & "��ԭ�ϲ���Ϊ��", vbInformation, "��ʾ"
            Exit Sub
        End If
        
        rsdetail.movenext
        i = i + 1
    Loop
    sql2 = "select * from G_billyarn where B_id='" & id & "'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs2.RecordCount > 0 Then
            savetoupdate1
            printdetail = True
    Else
    
 
    Dim rs As New RecordSet
    Dim sql As String
    Dim sql1 As String
    sql = "select * from G_draftBillyarn where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.addnew
    rs!B_datecreate = Now
    rs.Update
    id = rs!B_id
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    
    sql1 = "exec usp_Originalorderinsert  '" & id & "','" & FlatEdit3.Text & "','" & a & "',"
    sql1 = sql1 & "'12B004','YARN09','" & Gm.SysID.SystemUser & "','" & FlatEdit2.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit6.Text & "','" & FlatEdit4.Text & "','" & ComboBox1.Text & "','" & Originalsuppliers & "','" & fh & "','" & ComboBox2.Text & "','" & FlatEdit8.Text & "'"
    Gm.cnnTool.cnn.Execute sql1
    savedetail
    sql = "delete from G_draftBillyarn where B_itemid='" & id & "'"
    End If
    
    
    
    If printdetail = False Then
        Exit Sub
    End If
        Dim rs3 As New RecordSet
        Dim sql3 As String
        sql3 = "exec usp_Originaltransfersprint '" & id & "','" & Gm.SysID.SystemUserName & "'"
        rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        Dim frm1 As New frmModBLRPreviewOri
        Set frm1.RecordSet = rs3.Clone
            
        frm1.ObjectID = "22B047"
        frm1.Show
    FlatEdit1 = ""
    FlatEdit2 = ""
    FlatEdit4 = ""
    FlatEdit5 = ""
    FlatEdit6 = ""
    FlatEdit8 = ""
    id = ""
    cob1
    cob2
    fh = ""
    Originalsuppliers = ""
    FlatEdit3.Text = GetCodeID
    setRs
    bol = True
End Sub
Private Sub TDBGrid1_ButtonClick(ByVal ColIndex As Integer)
 If TDBGrid1.Columns("B_process").ColIndex = ColIndex Then
     Dim frm1 As New frmPopupDanWei
    frm1.Caption = "�ӹ���"
    frm1.ContactType = "�����ӹ���"
    frm1.Show vbModal
    rsdetail!B_processid = frm1.Clientid
    rsdetail!B_process = frm1.ClientName
    Unload frm1
    End If
     If TDBGrid1.Columns("B_transfers").ColIndex = ColIndex Then
     Dim frm2 As New frmPopupDanWei
     frm2.Caption = "�����ӹ���"
    frm2.ContactType = "�����ӹ���"
    frm2.Show vbModal
    rsdetail!B_TransfersID = frm2.Clientid
    rsdetail!B_transfers = frm2.ClientName
    Unload frm2
    End If
    
    
     If TDBGrid1.Columns("B_colorname").ColIndex = ColIndex Then
    Dim frm3 As New frmpopupComposition
    frm3.Show vbModal
    rsdetail!B_colorid = frm3.Compositionid
    rsdetail!B_colorname = frm3.CompositionName
    Unload frm3
    End If
End Sub

Private Sub sumall()
     Dim rs As New RecordSet
    Dim a As Double
    Dim b As Double
    Dim c As String
    Dim d As String
    If rsdetail.RecordCount <= 0 Then
        Exit Sub
    End If
    a = 0
    b = 0
    Set rs = rsdetail.Clone
    rs.MoveFirst
    Do While Not rs.EOF
        a = a + IIf(IsNull(rs!B_qty), 0, rs!B_qty)
        b = b + IIf(IsNull(rs!B_Sum), 0, rs!B_Sum)
        rs.movenext
    Loop
    c = Format(a, "0.00")
    d = Format(b, "0.00")
    TDBGrid1.Columns("B_itemidb").FooterText = "�ϼ�"
    TDBGrid1.Columns("B_qty").FooterText = "" & c & ""
    TDBGrid1.Columns("B_sum").FooterText = "" & d & ""
End Sub
'�����޸ı���ʹ�ӡ
Private Sub savetoupdate1()
    Dim sql2 As String
    Dim sql1 As String
    Dim a As String
    Dim sql3 As String
    Dim rs3 As RecordSet
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    sql1 = "exec usp_Originalorderinsert_update  '" & id & "','" & FlatEdit3.Text & "','" & a & "',"
    sql1 = sql1 & "'12B004','YARN09 ','" & Gm.SysID.SystemUser & "','" & FlatEdit2.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit6.Text & "','" & FlatEdit4.Text & "','" & ComboBox1.Text & "','" & Originalsuppliers & "','" & fh & "','" & ComboBox2.Text & "','" & FlatEdit8.Text & "'"
    Gm.cnnTool.cnn.Execute sql1
    rsdetail.MoveFirst
    Do While Not rsdetail.EOF
        Set rs3 = New RecordSet
        sql3 = "select * from G_billdetailyarn where B_itemid ='" & rsdetail!B_itemid & "'"
        rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs3.RecordCount > 0 Then
            sql2 = "exec usp_Originaltransfers_update '" & rsdetail!B_itemid & "','" & id & "','" & rsdetail!B_ItemIDB & "','" & rsdetail!B_processid & "'"
            sql2 = sql2 & ",'" & rsdetail!B_colorid & "','" & rsdetail!B_specification & "','" & rsdetail!B_qty & "','" & rsdetail!B_price & "'"
            sql2 = sql2 & ",'" & rsdetail!B_Sum & "','" & rsdetail!B_Cang & "','" & rsdetail!B_memo & "','" & rsdetail!B_TransfersID & "','" & rsdetail!B_BoxQty & "','" & rsdetail!B_DeliveryNote & "'"
            
            Gm.cnnTool.cnn.Execute sql2
        Else
            saveadddetail
        End If
        rsdetail.movenext
    Loop
      '���е������
   setAudit (0)
End Sub

Private Sub savetoupdate()
    Dim sql2 As String
    Dim sql1 As String
    Dim a As String
    Dim sql3 As String
    Dim rs3 As RecordSet
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    sql1 = "exec usp_Originalorderinsert_update  '" & id & "','" & FlatEdit3.Text & "','" & a & "',"
    sql1 = sql1 & "'12B004','YARN09 ','" & Gm.SysID.SystemUser & "','" & FlatEdit2.Text & "'"
    sql1 = sql1 & ",'" & FlatEdit6.Text & "','" & FlatEdit4.Text & "','" & ComboBox1.Text & "','" & Originalsuppliers & "','" & fh & "','" & ComboBox2.Text & "','" & FlatEdit8.Text & "'"
    Gm.cnnTool.cnn.Execute sql1
    rsdetail.MoveFirst
    Do While Not rsdetail.EOF
        Set rs3 = New RecordSet
        sql3 = "select * from G_billdetailyarn where B_itemid ='" & rsdetail!B_itemid & "'"
        rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs3.RecordCount > 0 Then
        
            sql2 = "exec usp_Originaltransfers_update '" & rsdetail!B_itemid & "','" & id & "','" & rsdetail!B_ItemIDB & "','" & rsdetail!B_processid & "'"
            sql2 = sql2 & ",'" & rsdetail!B_colorid & "','" & rsdetail!B_specification & "','" & rsdetail!B_qty & "','" & rsdetail!B_price & "'"
            sql2 = sql2 & ",'" & rsdetail!B_Sum & "','" & rsdetail!B_Cang & "','" & rsdetail!B_memo & "','" & rsdetail!B_TransfersID & "','" & rsdetail!B_BoxQty & "','" & rsdetail!B_DeliveryNote & "'"
            
            Gm.cnnTool.cnn.Execute sql2
        Else
            saveadddetail
        End If
        rsdetail.movenext
    Loop
      '���е������
   setAudit (0)
    
     FlatEdit1 = ""
    FlatEdit2 = ""
    FlatEdit4 = ""
    FlatEdit5 = ""
    FlatEdit6 = ""
    FlatEdit8 = ""
    id = ""
    cob1
    cob2
    fh = ""
    Originalsuppliers = ""
    FlatEdit3.Text = GetCodeID
    setRs
End Sub

Private Sub MoveFirst()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select top 1 B_id,B_CodeID,B_Date,B_Freightnum,B_HandNum,B_CarNum,B_payment,B_Shipment,b.B_ClientName,B_HandMent,c.B_Name,B_cope,B_drivename "
    sql = sql & " from G_BillYarn a left outer join G_ContactCompany b on a.B_Shipment=b.B_ClientID left outer join G_Employee c on a.B_HandMent=c.B_SID"
    sql = sql & " where B_ObjectID='12B004' and  B_BillType='YARN09'"
    sql = sql & "order by B_id"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount <= 0 Then
        MsgBox "�޶�����Ϣ", vbInformation + vbOKOnly, "��ʾ"
        Exit Sub
    End If
    id = rs!B_id
   FlatEdit3.Text = IIf(IsNull(rs!B_Codeid), "", rs!B_Codeid)
    DTPicker1.Value = rs!B_Date
    FlatEdit2.Text = IIf(IsNull(rs!B_Freightnum), "", rs!B_Freightnum)
    FlatEdit6.Text = IIf(IsNull(rs!B_HandNum), "", rs!B_HandNum)
    FlatEdit4.Text = IIf(IsNull(rs!B_CarNum), "", rs!B_CarNum)
    ComboBox1.Text = rs!B_payment
    FlatEdit1.Text = IIf(IsNull(rs!B_ClientName), "", rs!B_ClientName)
    Originalsuppliers = rs!B_Shipment
    FlatEdit5.Text = IIf(IsNull(rs!B_name), "", rs!B_name)
    fh = rs!B_HandMent
    ComboBox2.Text = IIf(IsNull(rs!B_cope), "", rs!B_cope)
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
     sql = "select top 1 B_id,B_CodeID,B_Date,B_Freightnum,B_HandNum,B_CarNum,B_payment,B_Shipment,b.B_ClientName,B_HandMent,c.B_Name,B_cope,B_drivename "
    sql = sql & " from G_BillYarn a left outer join G_ContactCompany b on a.B_Shipment=b.B_ClientID left outer join G_Employee c on a.B_HandMent=c.B_SID"
    sql = sql & " where B_ObjectID='12B004' and  B_BillType='YARN09' and B_ID<'" & id & "'"
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
    FlatEdit2.Text = IIf(IsNull(rs!B_Freightnum), "", rs!B_Freightnum)
    FlatEdit6.Text = IIf(IsNull(rs!B_HandNum), "", rs!B_HandNum)
    FlatEdit4.Text = IIf(IsNull(rs!B_CarNum), "", rs!B_CarNum)
    ComboBox1.Text = rs!B_payment
    FlatEdit1.Text = IIf(IsNull(rs!B_ClientName), "", rs!B_ClientName)
    Originalsuppliers = rs!B_Shipment
    FlatEdit5.Text = IIf(IsNull(rs!B_name), "", rs!B_name)
    fh = rs!B_HandMent
    ComboBox2.Text = IIf(IsNull(rs!B_cope), "", rs!B_cope)
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
      sql = "select top 1 B_id,B_CodeID,B_Date,B_Freightnum,B_HandNum,B_CarNum,B_payment,B_Shipment,b.B_ClientName,B_HandMent,c.B_Name,B_cope,B_drivename "
    sql = sql & " from G_BillYarn a left outer join G_ContactCompany b on a.B_Shipment=b.B_ClientID left outer join G_Employee c on a.B_HandMent=c.B_SID"
    sql = sql & " where B_ObjectID='12B004' and  B_BillType='YARN09' and B_ID>'" & id & "'"
    sql = sql & "order by B_id"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount <= 0 Then
         MsgBox "�������һ��", vbInformation + vbOKOnly, "��ʾ"
        Exit Sub
    End If
    id = rs!B_id
   FlatEdit3.Text = IIf(IsNull(rs!B_Codeid), "", rs!B_Codeid)
    DTPicker1.Value = rs!B_Date
    FlatEdit2.Text = IIf(IsNull(rs!B_Freightnum), "", rs!B_Freightnum)
    FlatEdit6.Text = IIf(IsNull(rs!B_HandNum), "", rs!B_HandNum)
    FlatEdit4.Text = IIf(IsNull(rs!B_CarNum), "", rs!B_CarNum)
    ComboBox1.Text = rs!B_payment
    FlatEdit1.Text = IIf(IsNull(rs!B_ClientName), "", rs!B_ClientName)
    Originalsuppliers = rs!B_Shipment
    FlatEdit5.Text = IIf(IsNull(rs!B_name), "", rs!B_name)
    fh = rs!B_HandMent
    ComboBox2.Text = IIf(IsNull(rs!B_cope), "", rs!B_cope)
    FlatEdit8.Text = IIf(IsNull(rs!B_drivename), "", rs!B_drivename)
    openbill
End Sub


Private Sub movelast()
    Dim rs As New RecordSet
    Dim sql As String
   sql = "select top 1 B_id,B_CodeID,B_Date,B_Freightnum,B_HandNum,B_CarNum,B_payment,B_Shipment,b.B_ClientName,B_HandMent,c.B_Name,B_cope,B_drivename "
    sql = sql & " from G_BillYarn a left outer join G_ContactCompany b on a.B_Shipment=b.B_ClientID left outer join G_Employee c on a.B_HandMent=c.B_SID"
    sql = sql & " where B_ObjectID='12B004' and  B_BillType='YARN09' "
    sql = sql & "order by B_id desc"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount <= 0 Then
        MsgBox "�޶�����Ϣ", vbInformation + vbOKOnly, "��ʾ"
        Exit Sub
    End If
    id = rs!B_id
   FlatEdit3.Text = IIf(IsNull(rs!B_Codeid), "", rs!B_Codeid)
    DTPicker1.Value = rs!B_Date
    FlatEdit2.Text = IIf(IsNull(rs!B_Freightnum), "", rs!B_Freightnum)
    FlatEdit6.Text = IIf(IsNull(rs!B_HandNum), "", rs!B_HandNum)
    FlatEdit4.Text = IIf(IsNull(rs!B_CarNum), "", rs!B_CarNum)
    ComboBox1.Text = rs!B_payment
    FlatEdit1.Text = IIf(IsNull(rs!B_ClientName), "", rs!B_ClientName)
    Originalsuppliers = rs!B_Shipment
    FlatEdit5.Text = IIf(IsNull(rs!B_name), "", rs!B_name)
    fh = rs!B_HandMent
    ComboBox2.Text = IIf(IsNull(rs!B_cope), "", rs!B_cope)
    FlatEdit8.Text = IIf(IsNull(rs!B_drivename), "", rs!B_drivename)
    openbill
End Sub

Public Sub openbill()
   Dim sql As String
   Dim rs As New RecordSet
   sql = "select B_ItemID,B_ID,B_OrderCode,B_producer,b.B_ClientName,B_GoodsID,c.B_Name,a.B_Specification,B_Qty,a.B_Price,B_Sum,B_LotNumber,B_MemoDetail,B_transfers,d.B_clientname as B_client,B_boxqty,B_DeliveryNote"
   sql = sql & "  from G_BilldetailYarn a left outer join G_ContactCompany b on a.B_producer=b.B_ClientID left outer join G_Yarn c on a.B_GoodsID=c.B_SID left outer join G_ContactCompany d on a.B_transfers=d.B_Clientid"
   sql = sql & " where B_ID='" & id & "'"
   
   rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
   Debug.Print
   setRs
   Do While Not rs.EOF
        rsdetail.addnew
       rsdetail!B_itemid = rs!B_itemid
         rsdetail!B_ItemIDB = rs!B_OrderCode
         rsdetail!B_process = IIf(IsNull(rs!B_ClientName), "", rs!B_ClientName)
         rsdetail!B_colorname = IIf(IsNull(rs!B_name), "", rs!B_name)
         rsdetail!B_specification = rs!B_specification
         rsdetail!B_BoxQty = IIf(IsNull(rs!B_BoxQty), "", rs!B_BoxQty)
         rsdetail!B_qty = rs!B_qty
         rsdetail!B_price = rs!B_price
         rsdetail!B_Sum = rs!B_Sum
         rsdetail!B_Cang = rs!B_LotNumber
         rsdetail!B_memo = rs!B_MemoDetail
         rsdetail!B_processid = rs!B_Producer
         rsdetail!B_colorid = rs!B_GoodsID
         rsdetail!B_transfers = IIf(IsNull(rs!B_Client), "", rs!B_Client)
         rsdetail!B_TransfersID = IIf(IsNull(rs!B_transfers), "", rs!B_transfers)
         rsdetail!B_DeliveryNote = IIf(IsNull(rs!B_DeliveryNote), "", rs!B_DeliveryNote)
       rsdetail.Update
       rs.movenext
   Loop
   tp
   If rs.RecordCount > 0 Then
    rsdetail.MoveFirst
   End If
   sumall
End Sub
Private Sub saveAudit(ByVal a As Long)
    If a = 0 Then
        FlatEdit2.Enabled = False
        FlatEdit4.Enabled = False
        FlatEdit6.Enabled = False
        FlatEdit8.Enabled = False
        DTPicker1.Enabled = False
        PushButton1.Enabled = False
        PushButton2.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        TDBGrid1.Enabled = False
        
        ActiveBar21.Bands("band1").Tools("����ͼƬ").Visible = True
        ActiveBar21.Bands("Band2").Tools("������").Enabled = False
        ActiveBar21.Bands("Band2").Tools("ɾ����").Enabled = False
        
        
    End If
    If a = 1 Then
        FlatEdit2.Enabled = True
        FlatEdit4.Enabled = True
        FlatEdit6.Enabled = True
        FlatEdit8.Enabled = True
        DTPicker1.Enabled = True
        PushButton1.Enabled = True
        PushButton2.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
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
    sql = "select * from  G_BillYarn  where B_id='" & id & "'"
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
    sql = "update G_BillYarn set B_Audit='" & a & "' where B_Id='" & id & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
End Sub




