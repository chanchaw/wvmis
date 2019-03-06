VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{0E8071F7-A7DD-47AC-95A9-365FFFE096DF}#1.0#0"; "UCButton.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmColorDJRK 
   Caption         =   "色布打卷手动入库"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12990
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
   ScaleHeight     =   7710
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   10950
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20250
      _LayoutVersion  =   1
      _ExtentX        =   35719
      _ExtentY        =   19315
      _DataPath       =   ""
      Bands           =   "frmColorDJRK.frx":0000
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   7575
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   12255
         _cx             =   21616
         _cy             =   13361
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
         GridRows        =   6
         GridCols        =   4
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmColorDJRK.frx":01C8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.Frame Frame4 
            Height          =   3210
            Left            =   8640
            TabIndex        =   8
            Top             =   4275
            Width           =   3525
            Begin VB.TextBox Text11 
               Height          =   375
               Left            =   1440
               TabIndex        =   36
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox Text10 
               Height          =   375
               Left            =   1440
               TabIndex        =   35
               Top             =   720
               Width           =   1815
            End
            Begin TA_UCButton.UCButton UCButton5 
               Height          =   615
               Left            =   2040
               TabIndex        =   25
               Top             =   2160
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   1085
               Caption         =   "保  存"
            End
            Begin TA_UCButton.UCButton UCButton4 
               Height          =   615
               Left            =   480
               TabIndex        =   24
               Top             =   2160
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   1085
               Caption         =   "退  出"
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "米  数："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   360
               TabIndex        =   23
               Top             =   1320
               Width           =   960
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "公  斤："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   360
               TabIndex        =   22
               Top             =   720
               Width           =   960
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "计米方式"
            Height          =   3210
            Left            =   5145
            TabIndex        =   7
            Top             =   4275
            Width           =   3435
            Begin VB.TextBox Text12 
               Height          =   375
               Left            =   2400
               TabIndex        =   37
               Top             =   2160
               Width           =   975
            End
            Begin VB.OptionButton Option3 
               Caption         =   "公式计米"
               Height          =   255
               Left            =   480
               TabIndex        =   20
               Top             =   2160
               Width           =   1215
            End
            Begin VB.OptionButton Option2 
               Caption         =   "手动输入"
               Height          =   495
               Left            =   480
               TabIndex        =   19
               Top             =   1320
               Width           =   2055
            End
            Begin VB.OptionButton Option1 
               Caption         =   "不计米"
               Height          =   195
               Left            =   480
               TabIndex        =   18
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "系数："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1680
               TabIndex        =   21
               Top             =   2160
               Width           =   720
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "订单信息"
            Height          =   3570
            Left            =   5145
            TabIndex        =   6
            Top             =   645
            Width           =   7020
            Begin VB.TextBox Text15 
               Height          =   375
               Left            =   7800
               TabIndex        =   44
               Top             =   2160
               Width           =   2055
            End
            Begin VB.TextBox Text14 
               Height          =   375
               Left            =   7800
               TabIndex        =   43
               Top             =   1560
               Width           =   2055
            End
            Begin VB.TextBox Text13 
               Height          =   375
               Left            =   7800
               TabIndex        =   42
               Top             =   960
               Width           =   2055
            End
            Begin VB.TextBox Text9 
               Height          =   375
               Left            =   4560
               TabIndex        =   34
               Top             =   2760
               Width           =   2055
            End
            Begin VB.TextBox Text8 
               Height          =   375
               Left            =   4560
               TabIndex        =   33
               Top             =   2160
               Width           =   2055
            End
            Begin VB.TextBox Text7 
               Height          =   375
               Left            =   4560
               TabIndex        =   32
               Top             =   1560
               Width           =   2055
            End
            Begin VB.TextBox Text6 
               Height          =   375
               Left            =   4560
               TabIndex        =   31
               Top             =   960
               Width           =   2055
            End
            Begin VB.TextBox Text5 
               Height          =   375
               Left            =   1200
               TabIndex        =   30
               Top             =   2760
               Width           =   2055
            End
            Begin VB.TextBox Text4 
               Height          =   375
               Left            =   1200
               TabIndex        =   29
               Top             =   2160
               Width           =   2055
            End
            Begin VB.TextBox Text3 
               Height          =   375
               Left            =   1200
               TabIndex        =   28
               Top             =   1560
               Width           =   2055
            End
            Begin VB.TextBox Text2 
               Height          =   375
               Left            =   1200
               TabIndex        =   27
               Top             =   960
               Width           =   2055
            End
            Begin VB.TextBox Text1 
               Height          =   375
               Left            =   1200
               TabIndex        =   26
               Top             =   360
               Width           =   2055
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "空 加 重："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   3
               Left            =   6720
               TabIndex        =   41
               Top             =   2160
               Width           =   1200
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "袋     重："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   6720
               TabIndex        =   40
               Top             =   1560
               Width           =   1320
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "纸 管 重："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   6720
               TabIndex        =   39
               Top             =   960
               Width           =   1200
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "订单米数："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   3360
               TabIndex        =   17
               Top             =   2760
               Width           =   1200
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "订单码数："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   3360
               TabIndex        =   16
               Top             =   2160
               Width           =   1200
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门    幅："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   3360
               TabIndex        =   15
               Top             =   1560
               Width           =   1200
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "订 单 号："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   3360
               TabIndex        =   14
               Top             =   960
               Width           =   1200
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "订单公斤："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   13
               Top             =   2760
               Width           =   1200
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "克   重："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   12
               Top             =   2160
               Width           =   1080
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "品   名："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   11
               Top             =   1560
               Width           =   1080
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "客   户："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   10
               Top             =   960
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "录入卡号："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   9
               Top             =   360
               Width           =   1200
            End
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   90
            ScaleHeight     =   495
            ScaleWidth      =   12075
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   90
            Width           =   12075
            Begin TA_UCButton.UCButton UCButton3 
               Height          =   495
               Left            =   7560
               TabIndex        =   5
               Top             =   0
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   873
               Caption         =   "获取订单信息"
            End
            Begin TA_UCButton.UCButton UCButton2 
               Height          =   495
               Left            =   3000
               TabIndex        =   4
               Top             =   0
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   873
               Caption         =   "删除当前"
            End
            Begin TA_UCButton.UCButton UCButton1 
               Height          =   495
               Left            =   360
               TabIndex        =   3
               Top             =   0
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   873
               Caption         =   "重排序号"
            End
         End
         Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
            Height          =   6840
            Left            =   90
            TabIndex        =   38
            Top             =   645
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   12065
            _LayoutType     =   0
            _RowHeight      =   -2147483647
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
            Splits(0).DividerColor=   15790320
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2805"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131601"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2805"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=131601"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            BorderStyle     =   0
            ColumnFooters   =   -1  'True
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.alignment=2,.valignment=2,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=1425,.italic=0,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(8)   =   ":id=1,.fontname=宋体"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1425,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(12)  =   ":id=2,.fontname=宋体"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1425,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(15)  =   ":id=3,.fontname=宋体"
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
         End
      End
   End
End
Attribute VB_Name = "frmColorDJRK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private GridRS As New RecordSet
Private m_Number As Long   '用来做序号和匹号

Private Sub Form_Load()
    InitFrm

End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
    
    Me.Show
    Text1.SetFocus  '获取焦点
    Option1.Value = True
     m_Number = 1
    setRs
    sumall
End Sub

Private Sub setRs()
    Set GridRS = New RecordSet

    GridRS.Fields.Append "rowindex", adVarChar, 100   '序号
    GridRS.Fields.Append "B_PH2", adVarChar, 100
    GridRS.Fields.Append "B_GJ", adVarChar, 100
    GridRS.Fields.Append "B_MS", adVarChar, 100
     GridRS.Open
    TDBGrid1.DataSource = GridRS
    
    Grid
    
End Sub

Private Sub Grid()
 
     TDBGrid1.Columns("rowindex").Caption = "序号"
    TDBGrid1.Columns("B_PH2").Caption = "匹号"
    TDBGrid1.Columns("B_GJ").Caption = "公斤"
    TDBGrid1.Columns("B_MS").Caption = "米数"
    
    
    TDBGrid1.Columns("rowindex").width = 1200
    TDBGrid1.Columns("B_PH2").width = 1200
    TDBGrid1.Columns("B_GJ").width = 1800
    TDBGrid1.Columns("B_MS").width = 1800
    
 
    TDBGrid1.Columns("B_GJ").NumberFormat = "0.0"
    TDBGrid1.Columns("B_MS").NumberFormat = "0.0"
    
    TDBGrid1.HeadLines = 1.4
    TDBGrid1.HoldFields
    TDBGrid1.MarqueeStyle = dbgHighlightRow
End Sub



'输入卡号之后点击Enter键
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
     GetRs
    End If
End Sub
'获取客户订单信息
Private Sub GetRs()
     Dim sql As String
     Dim rs As New RecordSet
      sql = "SELECT a.B_Clientid,c.B_ItemIDB,c.B_GoodsNameAlias,c.B_width,c.B_weight,isnull(c.B_meter,0)AS B_meter,isnull(c.B_KG,0)AS B_KG,isnull(c.B_Qty,0)AS B_Qty,isnull(c.B_Paper,0)as B_Paper,isnull(c.B_pocket,0)as B_pocket,isnull(c.B_Empty,0)as B_Empty "
      sql = sql & "FROM G_BillOrder a "
      sql = sql & " LEFT OUTER JOIN G_Billcolor b "
      sql = sql & " on a.B_ID= b.B_belongorderid "
      sql = sql & " LEFT OUTER JOIN G_Billdetailcolor c "
      sql = sql & " on b.B_id=c.B_ID "
      sql = sql & " where c.B_BCIncr='" & Val(Text1.Text) & "'"
      
      Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
        If rs.RecordCount > 0 Then
            Text2.Text = getClientName(rs!B_Clientid) '获取订单客户名
            Text3.Text = rs!B_GoodsNameAlias              '品名
            Text4.Text = rs!B_weight                                 '克重
            Text5.Text = rs!B_kg                                         '订单公斤
            Text6.Text = rs!B_ItemIDB                               '订单号
            Text7.Text = rs!B_Width                                  '门幅
            Text8.Text = rs!B_qty                                       '订单码数
            Text9.Text = rs!B_meter                                  '订单米数
            Text13.Text = rs!B_Paper                                '纸管重
            Text14.Text = rs!B_pocket                              '袋重
            Text15.Text = rs!B_Empty                                '空加重
        Else
          MsgBox "该卡号不存在，请重新输入！", vbInformation, "提示"
          Text1.SetFocus
        End If
         Text10.SetFocus '确定之后把焦点传到text10上

End Sub


Private Function getClientName(ByVal clientid As String) As String
     Dim rs As New RecordSet
     Dim sql As String
     sql = "Select B_ClientName From G_ContactCompany Where B_ClientID='" & clientid & "'"
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     If rs.RecordCount > 0 Then
        getClientName = rs!B_ClientName
     Else
        getClientName = ""
     End If
End Function

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
 
     If GridRS.RecordCount > 0 Then
       GridRS.movelast
       m_Number = GridRS!rowIndex + 1
    End If

    If Option1.Value = True Then
            GridRS.AddNew
            GridRS!rowIndex = m_Number
            GridRS!B_PH2 = m_Number
            GridRS!B_GJ = Text10.Text
            GridRS!B_MS = Text11.Text
            GridRS.Update
            
            sumall
          Text10.Text = ""
     ElseIf Option3 = True Then
            If Val(Text12.Text) > 0 Then
                 GridRS.AddNew
                 GridRS!rowIndex = m_Number
                 GridRS!B_PH2 = m_Number
                 GridRS!B_GJ = Text10.Text
                 GridRS!B_MS = Val(Text10.Text) * Val(Text12.Text)
                 GridRS.Update
                 
                 sumall
                 Text10.Text = ""
                 Text10.SetFocus
          Else
                MsgBox "系数为空，请先输入计米系数！", vbInformation, "提示"
          End If
    Else
      Text11.SetFocus
     End If
   
End If
End Sub
    
Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    
    If GridRS.RecordCount > 0 Then
       GridRS.movelast
       m_Number = GridRS!rowIndex + 1
    End If
    
    If Option2.Value = True Then
        GridRS.AddNew
        GridRS!rowIndex = m_Number
        GridRS!B_PH2 = m_Number
        GridRS!B_GJ = Val(Text10.Text)
        GridRS!B_MS = Val(Text11.Text)
        GridRS.Update
        
        sumall
        Text10.Text = ""
        Text11.Text = ""
        
   End If

        Text10.SetFocus
 
 End If
End Sub


'重新按顺序排列序号和匹号
Private Sub UCButton1_Click()
   GetNumber
End Sub

Private Sub UCButton2_Click()
         TDBGrid1.delete
End Sub
'获取客户订单信息
Private Sub UCButton3_Click()
    GetRs
End Sub
'退出
Private Sub UCButton4_Click()
    Unload Me
End Sub
'重新按顺序排列序号和匹号
Private Sub GetNumber()
 Dim a As Long
If GridRS.RecordCount > 0 Then
  GridRS.MoveFirst
   For a = 1 To GridRS.RecordCount
     
        GridRS("rowIndex") = a
        GridRS("B_PH2") = a
    
        GridRS.movenext
   Next
End If
End Sub
'合计
Private Sub sumall()
    Dim rss As New RecordSet
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim a As Double
    Dim m As Double
    Dim b As String
    Dim n As String
    a = 0
    m = 0
    
    Set rss = GridRS.Clone
    rss.MoveFirst
    Do While Not rss.EOF
        a = a + Val(IIf(IsNull(rss!B_GJ), 0, rss!B_GJ))
       m = m + Val(IIf(IsNull(rss!B_MS), 0, rss!B_MS))
        rss.movenext
    Loop
    b = Format(a, "0.0")
    n = Format(m, "0.0")
    TDBGrid1.Columns("rowIndex").FooterText = "合计"
    TDBGrid1.Columns("B_GJ").FooterText = "" & b & ""
    TDBGrid1.Columns("B_MS").FooterText = "" & n & ""

End Sub

'保存
Private Sub UCButton5_Click()
Dim sql As String
Dim rs As New RecordSet
Dim strSQL As String
Dim strRS As New RecordSet
Dim g_CN As String
Dim g_CUN As String
Dim g_IP As String
Dim m_Date As String
Dim m_BDCID As String
Dim m_EDP As String

sql = "SELECT* FROM G_Billdetailcolor  WHERE B_BCIncr='" & Val(Text1.Text) & "'"
rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

m_Date = Format(Now, "YYYY-MM-DD HH:MM:SS")
g_CN = Gm.HardWareID.CN
g_CUN = Gm.HardWareID.CUN
g_IP = Gm.HardWareID.iP
m_BDCID = rs!B_ItemID
m_EDP = "001"

GridRS.MoveFirst
Do While Not GridRS.EOF

        strSQL = ""
        strSQL = "exec dbo.[usp_SaveG_JRKBill_DJRK] '" & m_BDCID & "','未设置',"
        strSQL = strSQL & " '" & GridRS("B_GJ") & "','" & GridRS("B_MS") & "','" & m_Date & "',"
        strSQL = strSQL & " '" & g_CUN & "','" & g_CN & "','" & g_IP & "','" & Val(Text1.Text) & "',"
        strSQL = strSQL & " '" & Val(Text7.Text) & "','" & GridRS("B_PH2") & "','" & m_EDP & "','" & m_Date & "',"
        strSQL = strSQL & " '" & Val(Text1.Text) & "','" & GridRS("B_PH2") & "','" & Val(Text13.Text) & "','" & Val(Text14.Text) & "',"
        strSQL = strSQL & " '" & Val(Text15.Text) & "','" & m_Date & "','" & m_Date & "'"
       Debug.Print strSQL
       Gm.cnnTool.cnn.Execute strSQL

     GridRS.movenext

Loop
        clean    '清空所有text中的数据
        InitFrm
End Sub


Private Sub clean()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""

End Sub





