VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.2#0"; "Codejock.CommandBars.v16.2.4.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStorageQueryWhite 
   Caption         =   "白坯仓库综合管理"
   ClientHeight    =   7860
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStorageQueryWhite.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   7860
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11910
      _cx             =   21008
      _cy             =   13864
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
      BorderWidth     =   2
      ChildSpacing    =   0
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
      _GridInfo       =   $"frmStorageQueryWhite.frx":058A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   7800
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   11850
         _cx             =   20902
         _cy             =   13758
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
         BackColor       =   13557726
         ForeColor       =   -2147483630
         FrontTabColor   =   3263743
         BackTabColor    =   8355711
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "白坯库存|领料记录"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   6
         Position        =   5
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   1000
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Picture(0)      =   "frmStorageQueryWhite.frx":0610
         Picture(1)      =   "frmStorageQueryWhite.frx":09AA
         Flags(1)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   7770
            Left            =   1020
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   15
            Width           =   10815
            _cx             =   19076
            _cy             =   13705
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
            ChildSpacing    =   0
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
            _GridInfo       =   $"frmStorageQueryWhite.frx":0F44
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.PictureBox Picture2 
               Height          =   5130
               Left            =   6930
               ScaleHeight     =   5070
               ScaleWidth      =   3795
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   30
               Width           =   3855
               Begin VB.ComboBox Combo1 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Left            =   360
                  Style           =   2  'Dropdown List
                  TabIndex        =   15
                  Top             =   900
                  Width           =   2895
               End
               Begin XtremeCommandBars.BackstageButton BackstageButton1 
                  Height          =   555
                  Index           =   0
                  Left            =   360
                  TabIndex        =   16
                  Top             =   2460
                  Width           =   1335
                  _Version        =   1048578
                  _ExtentX        =   2355
                  _ExtentY        =   979
                  _StockProps     =   79
                  Caption         =   "生成单据"
               End
               Begin XtremeCommandBars.BackstageButton BackstageButton1 
                  Height          =   555
                  Index           =   1
                  Left            =   1980
                  TabIndex        =   17
                  Top             =   1680
                  Width           =   1335
                  _Version        =   1048578
                  _ExtentX        =   2355
                  _ExtentY        =   979
                  _StockProps     =   79
                  Caption         =   "关闭窗体"
               End
               Begin XtremeCommandBars.BackstageButton BackstageButton1 
                  Height          =   555
                  Index           =   2
                  Left            =   360
                  TabIndex        =   18
                  Top             =   1680
                  Width           =   1335
                  _Version        =   1048578
                  _ExtentX        =   2355
                  _ExtentY        =   979
                  _StockProps     =   79
                  Caption         =   "刷新库存"
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "单据类型:"
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   12
                     Charset         =   134
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   3
                  Left            =   360
                  TabIndex        =   19
                  Top             =   540
                  Width           =   1155
               End
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H8000000C&
               BorderStyle     =   0  'None
               Height          =   450
               Index           =   1
               Left            =   30
               ScaleHeight     =   450
               ScaleWidth      =   6900
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   30
               Width           =   6900
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "白坯库存表:"
                  ForeColor       =   &H8000000E&
                  Height          =   195
                  Index           =   1
                  Left            =   180
                  TabIndex        =   13
                  Top             =   120
                  Width           =   960
               End
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H8000000C&
               BorderStyle     =   0  'None
               Height          =   495
               Index           =   0
               Left            =   30
               ScaleHeight     =   495
               ScaleWidth      =   6900
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   5160
               Width           =   6900
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "待生成数据:"
                  ForeColor       =   &H8000000E&
                  Height          =   195
                  Index           =   0
                  Left            =   180
                  TabIndex        =   11
                  Top             =   120
                  Width           =   960
               End
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H8000000C&
               BorderStyle     =   0  'None
               Height          =   450
               Index           =   2
               Left            =   6930
               ScaleHeight     =   450
               ScaleWidth      =   3855
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   30
               Width           =   3855
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "设置查询参数:"
                  ForeColor       =   &H8000000E&
                  Height          =   195
                  Index           =   2
                  Left            =   180
                  TabIndex        =   9
                  Top             =   120
                  Width           =   1140
               End
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H8000000C&
               BorderStyle     =   0  'None
               Height          =   495
               Index           =   3
               Left            =   6930
               ScaleHeight     =   495
               ScaleWidth      =   3855
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   5160
               Width           =   3855
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "选择生成单据类型:"
                  ForeColor       =   &H8000000E&
                  Height          =   195
                  Index           =   9
                  Left            =   180
                  TabIndex        =   7
                  Top             =   60
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Picture3 
               BorderStyle     =   0  'None
               Height          =   2085
               Left            =   7305
               ScaleHeight     =   2085
               ScaleWidth      =   3480
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   5655
               Width           =   3480
               Begin XtremeCommandBars.BackstageButton BackstageButton1 
                  Height          =   555
                  Index           =   6
                  Left            =   360
                  TabIndex        =   4
                  Top             =   360
                  Width           =   1335
                  _Version        =   1048578
                  _ExtentX        =   2355
                  _ExtentY        =   979
                  _StockProps     =   79
                  Caption         =   "删除当前行"
               End
               Begin XtremeCommandBars.BackstageButton BackstageButton1 
                  Height          =   555
                  Index           =   7
                  Left            =   360
                  TabIndex        =   5
                  Top             =   1140
                  Width           =   1335
                  _Version        =   1048578
                  _ExtentX        =   2355
                  _ExtentY        =   979
                  _StockProps     =   79
                  Caption         =   "清空网格数据"
               End
            End
            Begin TrueOleDBGrid80.TDBGrid TDBGInv 
               Height          =   4680
               Left            =   30
               TabIndex        =   20
               Top             =   480
               Width           =   6900
               _ExtentX        =   12171
               _ExtentY        =   8255
               _LayoutType     =   0
               _RowHeight      =   23
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
               Splits(0).ShowCollapseExpandIcons=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).AllowColSelect=   0   'False
               Splits(0).AllowColMove=   -1  'True
               Splits(0).AllowRowSelect=   0   'False
               Splits(0).DividerColor=   13160660
               Splits(0).FilterBar=   -1  'True
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=2"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=273"
               Splits(0)._ColumnProps(5)=   "Column(0).WrapText=1"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=3281"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
               Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=273"
               Splits(0)._ColumnProps(11)=   "Column(1).WrapText=1"
               Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
               PrintInfos(0).PageFooterFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               Appearance      =   0
               ColumnFooters   =   -1  'True
               DefColWidth     =   0
               HeadLines       =   1.75
               FootLines       =   1.1
               PictureCurrentRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
               PictureCurrentRow(0)=   "bHQAAA4DAABCTQ4DAAAAAAAANgAAACgAAAARAAAADgAAAAEAGAAAAAAA2AIAAAAAAAAAAAAAAAAA"
               PictureCurrentRow(1)=   "AAAAAACltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrUA"
               PictureCurrentRow(2)=   "lKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKalAKW2taW2"
               PictureCurrentRow(3)=   "taW2taW2taW2taW2taW2taW2taW2taW2tZSmpaW2taW2taW2taW2taW2taW2tQCtvr2tvr2tvr2t"
               PictureCurrentRow(4)=   "vr2tvr2tvr2tvr2tvr2tvr2tvr0YGBicrq2tvr2tvr2tvr2tvr2tvr0ArcfGrcfGrcfGrcfGrcfG"
               PictureCurrentRow(5)=   "rcfGrcfGrcfGrcfGrcfGAAAAGBgYnK6trcfGrcfGrcfGrcfGALXHxrXHxrXHxrXHxrXHxrXHxrXH"
               PictureCurrentRow(6)=   "xrXHxrXHxrXHxgAAAAAAABgYGKW2tbXHxrXHxrXHxgC9z869z869z869z869z869z869z869z869"
               PictureCurrentRow(7)=   "z869z84AAAAAAAAAAAAYGBitvrW9z869z84Avc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/O"
               PictureCurrentRow(8)=   "AAAAAAAAAAAAAAAAKTAxvc/Ovc/OAMbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1gAAAAAA"
               PictureCurrentRow(9)=   "AAAAABAQEKW2tcbX1sbX1gDO19bO19bO19bO19bO19bO19bO19bO19bO19bO19YAAAAAAAAQEBC1"
               PictureCurrentRow(10)=   "trXO19bO19bO19YA1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufnAAAAEBAQtb7G1ufn1ufn"
               PictureCurrentRow(11)=   "1ufn1ufnAN7n797n797n797n797n797n797n797n797n797n7xAQEL3Hxt7n797n797n797n797n"
               PictureCurrentRow(12)=   "7wDe7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7++9z87e7+/e7+/e7+/e7+/e7+/e7+8A5+/3"
               PictureCurrentRow(13)=   "5+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/3AA=="
               PictureCurrentRow.vt=   9
               PictureStandardRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
               PictureStandardRow(0)=   "bHQAAN4DAABCTd4DAAAAAAAANgAAACgAAAARAAAAEgAAAAEAGAAAAAAAqAMAAAAAAAAAAAAAAAAA"
               PictureStandardRow(1)=   "AAAAAACMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpwA"
               PictureStandardRow(2)=   "lKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKalAJyurZyu"
               PictureStandardRow(3)=   "rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurQCltrWltrWltrWl"
               PictureStandardRow(4)=   "trWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrUArb69rb69rb69rb69rb69"
               PictureStandardRow(5)=   "rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69AK2+va2+va2+va2+va2+va2+va2+"
               PictureStandardRow(6)=   "va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+vQC1x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8a1"
               PictureStandardRow(7)=   "x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8YAtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfG"
               PictureStandardRow(8)=   "tcfGtcfGtcfGtcfGtcfGtcfGtcfGAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
               PictureStandardRow(9)=   "zr3Pzr3Pzr3Pzr3Pzr3PzgC9z869z869z869z869z869z869z869z869z869z869z869z869z869"
               PictureStandardRow(10)=   "z869z869z869z84Avc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/O"
               PictureStandardRow(11)=   "vc/Ovc/OAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
               PictureStandardRow(12)=   "1gDO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19YA1ufn"
               PictureStandardRow(13)=   "1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufnANbn59bn59bn"
               PictureStandardRow(14)=   "59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wDe7+/e7+/e7+/e7+/e"
               PictureStandardRow(15)=   "7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+8A3u/v3u/v3u/v3u/v3u/v3u/v"
               PictureStandardRow(16)=   "3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/vAN7v797v797v797v797v797v797v797v"
               PictureStandardRow(17)=   "797v797v797v797v797v797v797v797v797v7wA="
               PictureStandardRow.vt=   9
               MultipleLines   =   0
               EmptyRows       =   -1  'True
               CellTipsWidth   =   0
               MultiSelect     =   2
               GroupByCaption  =   "分组示意图"
               DeadAreaBackColor=   16252927
               RowDividerColor =   13160660
               RowSubDividerColor=   13160660
               DirectionAfterEnter=   1
               MaxRows         =   250000
               ViewColumnCaptionWidth=   0
               ViewColumnWidth =   6900.095
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(5)   =   ":id=0,.fontname=宋体"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF7FFFF&,.bold=0,.fontsize=900"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(8)   =   ":id=1,.fontname=宋体"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgpicMode=2,.bold=0,.fontsize=900"
               _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(12)  =   ":id=2,.fontname=宋体"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
               _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(15)  =   ":id=3,.fontname=宋体"
               _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H31CFFF&"
               _StyleDefs(18)  =   ":id=6,.fgcolor=&H80000008&"
               _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
               _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgpicMode=2,.appearance=1"
               _StyleDefs(24)  =   ":id=11,.bgbmp=3"
               _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(26)  =   "Splits(0).Style:id=55,.parent=1,.alignment=2,.valignment=2,.wraptext=-1"
               _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
               _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=56,.parent=2"
               _StyleDefs(29)  =   "Splits(0).FooterStyle:id=57,.parent=3"
               _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=58,.parent=5"
               _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=60,.parent=6"
               _StyleDefs(32)  =   "Splits(0).EditorStyle:id=59,.parent=7"
               _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=61,.parent=8"
               _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=62,.parent=9,.bgcolor=&HFFFFFF&"
               _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
               _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
               _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
               _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=82,.parent=55"
               _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=56,.alignment=0"
               _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=57"
               _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=59"
               _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=86,.parent=55"
               _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=56,.alignment=0"
               _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=57"
               _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=59"
               _StyleDefs(46)  =   "Named:id=33:Normal"
               _StyleDefs(47)  =   ":id=33,.parent=0"
               _StyleDefs(48)  =   "Named:id=34:Heading"
               _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(50)  =   ":id=34,.wraptext=-1,.bgpicMode=1,.bgbmp=1"
               _StyleDefs(51)  =   "Named:id=35:Footing"
               _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(53)  =   ":id=35,.bgpicMode=1,.bgbmp=2"
               _StyleDefs(54)  =   "Named:id=36:Selected"
               _StyleDefs(55)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(56)  =   "Named:id=37:Caption"
               _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(58)  =   "Named:id=38:HighlightRow"
               _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(60)  =   "Named:id=39:EvenRow"
               _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(62)  =   "Named:id=40:OddRow"
               _StyleDefs(63)  =   ":id=40,.parent=33"
               _StyleDefs(64)  =   "Named:id=41:RecordSelector"
               _StyleDefs(65)  =   ":id=41,.parent=34,.bgcolor=&HCEDFDE&,.bgpicMode=0,.borderColor=&H80000005&"
               _StyleDefs(66)  =   "Named:id=42:FilterBar"
               _StyleDefs(67)  =   ":id=42,.parent=33"
               _StyleDefs(68)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
               _StyleDefs(69)  =   "bmp(1):id=1,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
               _StyleDefs(70)  =   "bmp(2):id=1,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
               _StyleDefs(71)  =   "bmp(3):id=1,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(72)  =   "bmp(4):id=1,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(73)  =   "bmp(5):id=1,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
               _StyleDefs(74)  =   "bmp(6):id=1,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
               _StyleDefs(75)  =   "bmp(7):id=1,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(76)  =   "bmp(8):id=1,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(77)  =   "bmp(9):id=1,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(78)  =   "bmp(10):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
               _StyleDefs(79)  =   "bmp(11):id=1,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(80)  =   "bmp(12):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(81)  =   "bmp(13):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
               _StyleDefs(82)  =   "bmp(14):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(83)  =   "bmp(15):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(84)  =   "bmp(16):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
               _StyleDefs(85)  =   "bmp(17):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(86)  =   "bmp(18):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(87)  =   "bmp(19):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
               _StyleDefs(88)  =   "bmp(20):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(89)  =   "bmp(21):id=1,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(90)  =   "bmp(22):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
               _StyleDefs(91)  =   "bmp(23):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(92)  =   "bmp(24):id=1,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(93)  =   "bmp(25):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
               _StyleDefs(94)  =   "bmp(26):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(95)  =   "bmp(27):id=1,797v797v797v7wAAAA=="
               _StyleDefs(96)  =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
               _StyleDefs(97)  =   "bmp(1):id=2,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
               _StyleDefs(98)  =   "bmp(2):id=2,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
               _StyleDefs(99)  =   "bmp(3):id=2,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(100) =   "bmp(4):id=2,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(101) =   "bmp(5):id=2,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
               _StyleDefs(102) =   "bmp(6):id=2,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
               _StyleDefs(103) =   "bmp(7):id=2,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(104) =   "bmp(8):id=2,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(105) =   "bmp(9):id=2,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(106) =   "bmp(10):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
               _StyleDefs(107) =   "bmp(11):id=2,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(108) =   "bmp(12):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(109) =   "bmp(13):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
               _StyleDefs(110) =   "bmp(14):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(111) =   "bmp(15):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(112) =   "bmp(16):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
               _StyleDefs(113) =   "bmp(17):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(114) =   "bmp(18):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(115) =   "bmp(19):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
               _StyleDefs(116) =   "bmp(20):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(117) =   "bmp(21):id=2,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(118) =   "bmp(22):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
               _StyleDefs(119) =   "bmp(23):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(120) =   "bmp(24):id=2,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(121) =   "bmp(25):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
               _StyleDefs(122) =   "bmp(26):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(123) =   "bmp(27):id=2,797v797v797v7wAAAA=="
               _StyleDefs(124) =   "bmp(0):id=3,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
               _StyleDefs(125) =   "bmp(1):id=3,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
               _StyleDefs(126) =   "bmp(2):id=3,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
               _StyleDefs(127) =   "bmp(3):id=3,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(128) =   "bmp(4):id=3,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(129) =   "bmp(5):id=3,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
               _StyleDefs(130) =   "bmp(6):id=3,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
               _StyleDefs(131) =   "bmp(7):id=3,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(132) =   "bmp(8):id=3,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(133) =   "bmp(9):id=3,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(134) =   "bmp(10):id=3,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
               _StyleDefs(135) =   "bmp(11):id=3,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(136) =   "bmp(12):id=3,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(137) =   "bmp(13):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
               _StyleDefs(138) =   "bmp(14):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(139) =   "bmp(15):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(140) =   "bmp(16):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
               _StyleDefs(141) =   "bmp(17):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(142) =   "bmp(18):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(143) =   "bmp(19):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
               _StyleDefs(144) =   "bmp(20):id=3,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(145) =   "bmp(21):id=3,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(146) =   "bmp(22):id=3,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
               _StyleDefs(147) =   "bmp(23):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(148) =   "bmp(24):id=3,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(149) =   "bmp(25):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
               _StyleDefs(150) =   "bmp(26):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(151) =   "bmp(27):id=3,797v797v797v7wAAAA=="
            End
            Begin TrueOleDBGrid80.TDBGrid TDBGWait01 
               Height          =   2580
               Left            =   30
               TabIndex        =   21
               Top             =   5160
               Width           =   6900
               _ExtentX        =   12171
               _ExtentY        =   4551
               _LayoutType     =   0
               _RowHeight      =   23
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
               Splits(0).ShowCollapseExpandIcons=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).AllowColSelect=   0   'False
               Splits(0).AllowColMove=   -1  'True
               Splits(0).AllowRowSelect=   0   'False
               Splits(0).DividerColor=   13160660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=2"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=273"
               Splits(0)._ColumnProps(5)=   "Column(0).WrapText=1"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=3281"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
               Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=273"
               Splits(0)._ColumnProps(11)=   "Column(1).WrapText=1"
               Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
               PrintInfos(0).PageFooterFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               Appearance      =   0
               ColumnFooters   =   -1  'True
               DefColWidth     =   0
               HeadLines       =   1.75
               FootLines       =   1.1
               PictureCurrentRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
               PictureCurrentRow(0)=   "bHQAAA4DAABCTQ4DAAAAAAAANgAAACgAAAARAAAADgAAAAEAGAAAAAAA2AIAAAAAAAAAAAAAAAAA"
               PictureCurrentRow(1)=   "AAAAAACltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrUA"
               PictureCurrentRow(2)=   "lKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKalAKW2taW2"
               PictureCurrentRow(3)=   "taW2taW2taW2taW2taW2taW2taW2taW2tZSmpaW2taW2taW2taW2taW2taW2tQCtvr2tvr2tvr2t"
               PictureCurrentRow(4)=   "vr2tvr2tvr2tvr2tvr2tvr2tvr0YGBicrq2tvr2tvr2tvr2tvr2tvr0ArcfGrcfGrcfGrcfGrcfG"
               PictureCurrentRow(5)=   "rcfGrcfGrcfGrcfGrcfGAAAAGBgYnK6trcfGrcfGrcfGrcfGALXHxrXHxrXHxrXHxrXHxrXHxrXH"
               PictureCurrentRow(6)=   "xrXHxrXHxrXHxgAAAAAAABgYGKW2tbXHxrXHxrXHxgC9z869z869z869z869z869z869z869z869"
               PictureCurrentRow(7)=   "z869z84AAAAAAAAAAAAYGBitvrW9z869z84Avc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/O"
               PictureCurrentRow(8)=   "AAAAAAAAAAAAAAAAKTAxvc/Ovc/OAMbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1gAAAAAA"
               PictureCurrentRow(9)=   "AAAAABAQEKW2tcbX1sbX1gDO19bO19bO19bO19bO19bO19bO19bO19bO19bO19YAAAAAAAAQEBC1"
               PictureCurrentRow(10)=   "trXO19bO19bO19YA1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufnAAAAEBAQtb7G1ufn1ufn"
               PictureCurrentRow(11)=   "1ufn1ufnAN7n797n797n797n797n797n797n797n797n797n7xAQEL3Hxt7n797n797n797n797n"
               PictureCurrentRow(12)=   "7wDe7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7++9z87e7+/e7+/e7+/e7+/e7+/e7+8A5+/3"
               PictureCurrentRow(13)=   "5+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/3AA=="
               PictureCurrentRow.vt=   9
               PictureStandardRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
               PictureStandardRow(0)=   "bHQAAN4DAABCTd4DAAAAAAAANgAAACgAAAARAAAAEgAAAAEAGAAAAAAAqAMAAAAAAAAAAAAAAAAA"
               PictureStandardRow(1)=   "AAAAAACMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpwA"
               PictureStandardRow(2)=   "lKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKalAJyurZyu"
               PictureStandardRow(3)=   "rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurQCltrWltrWltrWl"
               PictureStandardRow(4)=   "trWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrUArb69rb69rb69rb69rb69"
               PictureStandardRow(5)=   "rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69AK2+va2+va2+va2+va2+va2+va2+"
               PictureStandardRow(6)=   "va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+vQC1x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8a1"
               PictureStandardRow(7)=   "x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8YAtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfG"
               PictureStandardRow(8)=   "tcfGtcfGtcfGtcfGtcfGtcfGtcfGAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
               PictureStandardRow(9)=   "zr3Pzr3Pzr3Pzr3Pzr3PzgC9z869z869z869z869z869z869z869z869z869z869z869z869z869"
               PictureStandardRow(10)=   "z869z869z869z84Avc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/O"
               PictureStandardRow(11)=   "vc/Ovc/OAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
               PictureStandardRow(12)=   "1gDO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19YA1ufn"
               PictureStandardRow(13)=   "1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufnANbn59bn59bn"
               PictureStandardRow(14)=   "59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wDe7+/e7+/e7+/e7+/e"
               PictureStandardRow(15)=   "7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+8A3u/v3u/v3u/v3u/v3u/v3u/v"
               PictureStandardRow(16)=   "3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/vAN7v797v797v797v797v797v797v797v"
               PictureStandardRow(17)=   "797v797v797v797v797v797v797v797v797v7wA="
               PictureStandardRow.vt=   9
               MultipleLines   =   0
               EmptyRows       =   -1  'True
               CellTipsWidth   =   0
               MultiSelect     =   2
               GroupByCaption  =   "分组示意图"
               DeadAreaBackColor=   16252927
               RowDividerColor =   13160660
               RowSubDividerColor=   13160660
               DirectionAfterEnter=   1
               MaxRows         =   250000
               ViewColumnCaptionWidth=   0
               ViewColumnWidth =   6900.095
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(5)   =   ":id=0,.fontname=宋体"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF7FFFF&,.bold=0,.fontsize=900"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(8)   =   ":id=1,.fontname=宋体"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgpicMode=2,.bold=0,.fontsize=900"
               _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(12)  =   ":id=2,.fontname=宋体"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
               _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(15)  =   ":id=3,.fontname=宋体"
               _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H31CFFF&"
               _StyleDefs(18)  =   ":id=6,.fgcolor=&H80000008&"
               _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgpicMode=2"
               _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgpicMode=2,.appearance=1"
               _StyleDefs(24)  =   ":id=11,.bgbmp=3"
               _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(26)  =   "Splits(0).Style:id=55,.parent=1,.alignment=2,.valignment=2,.wraptext=-1"
               _StyleDefs(27)  =   ":id=55,.bgpicMode=2"
               _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
               _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=56,.parent=2"
               _StyleDefs(30)  =   "Splits(0).FooterStyle:id=57,.parent=3"
               _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=58,.parent=5"
               _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=60,.parent=6"
               _StyleDefs(33)  =   "Splits(0).EditorStyle:id=59,.parent=7"
               _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=61,.parent=8"
               _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=62,.parent=9,.bgcolor=&HFFFFFF&"
               _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
               _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
               _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
               _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=82,.parent=55"
               _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=56,.alignment=0"
               _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=57"
               _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=59"
               _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=86,.parent=55"
               _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=56,.alignment=0"
               _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=57"
               _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=59"
               _StyleDefs(47)  =   "Named:id=33:Normal"
               _StyleDefs(48)  =   ":id=33,.parent=0"
               _StyleDefs(49)  =   "Named:id=34:Heading"
               _StyleDefs(50)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(51)  =   ":id=34,.wraptext=-1,.bgpicMode=1,.bgbmp=1"
               _StyleDefs(52)  =   "Named:id=35:Footing"
               _StyleDefs(53)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(54)  =   ":id=35,.bgpicMode=1,.bgbmp=2"
               _StyleDefs(55)  =   "Named:id=36:Selected"
               _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(57)  =   "Named:id=37:Caption"
               _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(59)  =   "Named:id=38:HighlightRow"
               _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(61)  =   "Named:id=39:EvenRow"
               _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(63)  =   "Named:id=40:OddRow"
               _StyleDefs(64)  =   ":id=40,.parent=33"
               _StyleDefs(65)  =   "Named:id=41:RecordSelector"
               _StyleDefs(66)  =   ":id=41,.parent=34,.bgcolor=&HCEDFDE&,.bgpicMode=0,.borderColor=&H80000005&"
               _StyleDefs(67)  =   "Named:id=42:FilterBar"
               _StyleDefs(68)  =   ":id=42,.parent=33"
               _StyleDefs(69)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
               _StyleDefs(70)  =   "bmp(1):id=1,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
               _StyleDefs(71)  =   "bmp(2):id=1,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
               _StyleDefs(72)  =   "bmp(3):id=1,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(73)  =   "bmp(4):id=1,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(74)  =   "bmp(5):id=1,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
               _StyleDefs(75)  =   "bmp(6):id=1,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
               _StyleDefs(76)  =   "bmp(7):id=1,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(77)  =   "bmp(8):id=1,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(78)  =   "bmp(9):id=1,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(79)  =   "bmp(10):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
               _StyleDefs(80)  =   "bmp(11):id=1,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(81)  =   "bmp(12):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(82)  =   "bmp(13):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
               _StyleDefs(83)  =   "bmp(14):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(84)  =   "bmp(15):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(85)  =   "bmp(16):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
               _StyleDefs(86)  =   "bmp(17):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(87)  =   "bmp(18):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(88)  =   "bmp(19):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
               _StyleDefs(89)  =   "bmp(20):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(90)  =   "bmp(21):id=1,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(91)  =   "bmp(22):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
               _StyleDefs(92)  =   "bmp(23):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(93)  =   "bmp(24):id=1,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(94)  =   "bmp(25):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
               _StyleDefs(95)  =   "bmp(26):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(96)  =   "bmp(27):id=1,797v797v797v7wAAAA=="
               _StyleDefs(97)  =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
               _StyleDefs(98)  =   "bmp(1):id=2,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
               _StyleDefs(99)  =   "bmp(2):id=2,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
               _StyleDefs(100) =   "bmp(3):id=2,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(101) =   "bmp(4):id=2,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(102) =   "bmp(5):id=2,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
               _StyleDefs(103) =   "bmp(6):id=2,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
               _StyleDefs(104) =   "bmp(7):id=2,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(105) =   "bmp(8):id=2,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(106) =   "bmp(9):id=2,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(107) =   "bmp(10):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
               _StyleDefs(108) =   "bmp(11):id=2,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(109) =   "bmp(12):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(110) =   "bmp(13):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
               _StyleDefs(111) =   "bmp(14):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(112) =   "bmp(15):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(113) =   "bmp(16):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
               _StyleDefs(114) =   "bmp(17):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(115) =   "bmp(18):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(116) =   "bmp(19):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
               _StyleDefs(117) =   "bmp(20):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(118) =   "bmp(21):id=2,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(119) =   "bmp(22):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
               _StyleDefs(120) =   "bmp(23):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(121) =   "bmp(24):id=2,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(122) =   "bmp(25):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
               _StyleDefs(123) =   "bmp(26):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(124) =   "bmp(27):id=2,797v797v797v7wAAAA=="
               _StyleDefs(125) =   "bmp(0):id=3,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
               _StyleDefs(126) =   "bmp(1):id=3,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
               _StyleDefs(127) =   "bmp(2):id=3,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
               _StyleDefs(128) =   "bmp(3):id=3,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(129) =   "bmp(4):id=3,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(130) =   "bmp(5):id=3,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
               _StyleDefs(131) =   "bmp(6):id=3,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
               _StyleDefs(132) =   "bmp(7):id=3,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(133) =   "bmp(8):id=3,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(134) =   "bmp(9):id=3,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(135) =   "bmp(10):id=3,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
               _StyleDefs(136) =   "bmp(11):id=3,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(137) =   "bmp(12):id=3,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(138) =   "bmp(13):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
               _StyleDefs(139) =   "bmp(14):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(140) =   "bmp(15):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(141) =   "bmp(16):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
               _StyleDefs(142) =   "bmp(17):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(143) =   "bmp(18):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(144) =   "bmp(19):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
               _StyleDefs(145) =   "bmp(20):id=3,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(146) =   "bmp(21):id=3,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(147) =   "bmp(22):id=3,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
               _StyleDefs(148) =   "bmp(23):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(149) =   "bmp(24):id=3,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(150) =   "bmp(25):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
               _StyleDefs(151) =   "bmp(26):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(152) =   "bmp(27):id=3,797v797v797v7wAAAA=="
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   7770
            Left            =   13470
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   15
            Width           =   10815
            _cx             =   19076
            _cy             =   13705
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
            ChildSpacing    =   0
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
            _GridInfo       =   $"frmStorageQueryWhite.frx":0FDD
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   5130
               Left            =   8085
               ScaleHeight     =   5130
               ScaleWidth      =   2700
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   30
               Width           =   2700
               Begin VB.Frame Frame2 
                  Caption         =   "生成单据"
                  Height          =   1995
                  Left            =   240
                  TabIndex        =   40
                  Top             =   3300
                  Width           =   2955
                  Begin VB.ComboBox Combo2 
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   405
                     Left            =   360
                     Style           =   2  'Dropdown List
                     TabIndex        =   41
                     Top             =   720
                     Width           =   2295
                  End
                  Begin XtremeCommandBars.BackstageButton BackstageButton1 
                     Height          =   555
                     Index           =   3
                     Left            =   360
                     TabIndex        =   42
                     Top             =   1320
                     Width           =   1335
                     _Version        =   1048578
                     _ExtentX        =   2355
                     _ExtentY        =   979
                     _StockProps     =   79
                     Caption         =   "生成单据"
                  End
                  Begin VB.Label Label1 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "单据类型:"
                     BeginProperty Font 
                        Name            =   "宋体"
                        Size            =   12
                        Charset         =   134
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Index           =   4
                     Left            =   360
                     TabIndex        =   43
                     Top             =   360
                     Width           =   1155
                  End
               End
               Begin VB.Frame Frame1 
                  Caption         =   "领料日期设定"
                  Height          =   1635
                  Left            =   240
                  TabIndex        =   35
                  Top             =   180
                  Width           =   2955
                  Begin MSComCtl2.DTPicker DTPSpendSDate 
                     Height          =   375
                     Left            =   1260
                     TabIndex        =   36
                     Top             =   420
                     Width           =   1455
                     _ExtentX        =   2566
                     _ExtentY        =   661
                     _Version        =   393216
                     Format          =   199622657
                     CurrentDate     =   42769
                  End
                  Begin MSComCtl2.DTPicker DTPSpendEDate 
                     Height          =   375
                     Left            =   1260
                     TabIndex        =   37
                     Top             =   900
                     Width           =   1455
                     _ExtentX        =   2566
                     _ExtentY        =   661
                     _Version        =   393216
                     Format          =   199622657
                     CurrentDate     =   42769
                  End
                  Begin VB.Label Label2 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "起始日期："
                     Height          =   195
                     Index           =   0
                     Left            =   240
                     TabIndex        =   39
                     Top             =   480
                     Width           =   900
                  End
                  Begin VB.Label Label2 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "终止日期："
                     Height          =   195
                     Index           =   1
                     Left            =   240
                     TabIndex        =   38
                     Top             =   960
                     Width           =   900
                  End
               End
               Begin XtremeCommandBars.BackstageButton BackstageButton1 
                  Height          =   555
                  Index           =   4
                  Left            =   1860
                  TabIndex        =   44
                  Top             =   2100
                  Width           =   1335
                  _Version        =   1048578
                  _ExtentX        =   2355
                  _ExtentY        =   979
                  _StockProps     =   79
                  Caption         =   "关闭窗体"
               End
               Begin XtremeCommandBars.BackstageButton BackstageButton1 
                  Height          =   555
                  Index           =   5
                  Left            =   240
                  TabIndex        =   45
                  Top             =   2100
                  Width           =   1335
                  _Version        =   1048578
                  _ExtentX        =   2355
                  _ExtentY        =   979
                  _StockProps     =   79
                  Caption         =   "刷新领料"
               End
            End
            Begin VB.PictureBox Picture4 
               BorderStyle     =   0  'None
               Height          =   2580
               Left            =   8085
               ScaleHeight     =   2580
               ScaleWidth      =   2700
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   5160
               Width           =   2700
               Begin XtremeCommandBars.BackstageButton BackstageButton1 
                  Height          =   555
                  Index           =   8
                  Left            =   420
                  TabIndex        =   32
                  Top             =   300
                  Width           =   1335
                  _Version        =   1048578
                  _ExtentX        =   2355
                  _ExtentY        =   979
                  _StockProps     =   79
                  Caption         =   "删除当前行"
               End
               Begin XtremeCommandBars.BackstageButton BackstageButton1 
                  Height          =   555
                  Index           =   9
                  Left            =   420
                  TabIndex        =   33
                  Top             =   1080
                  Width           =   1335
                  _Version        =   1048578
                  _ExtentX        =   2355
                  _ExtentY        =   979
                  _StockProps     =   79
                  Caption         =   "清空网格数据"
               End
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H8000000C&
               BorderStyle     =   0  'None
               Height          =   2580
               Index           =   4
               Left            =   8085
               ScaleHeight     =   2580
               ScaleWidth      =   2700
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   5160
               Width           =   2700
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "选择生成单据类型:"
                  ForeColor       =   &H8000000E&
                  Height          =   195
                  Index           =   10
                  Left            =   180
                  TabIndex        =   30
                  Top             =   120
                  Width           =   1500
               End
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H8000000C&
               BorderStyle     =   0  'None
               Height          =   1275
               Index           =   5
               Left            =   8085
               ScaleHeight     =   1275
               ScaleWidth      =   2700
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   30
               Width           =   2700
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "设置查询参数:"
                  ForeColor       =   &H8000000E&
                  Height          =   195
                  Index           =   17
                  Left            =   180
                  TabIndex        =   28
                  Top             =   120
                  Width           =   1140
               End
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H8000000C&
               BorderStyle     =   0  'None
               Height          =   1290
               Index           =   6
               Left            =   30
               ScaleHeight     =   1290
               ScaleWidth      =   10755
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   5160
               Width           =   10755
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "待生成数据:"
                  ForeColor       =   &H8000000E&
                  Height          =   195
                  Index           =   18
                  Left            =   180
                  TabIndex        =   26
                  Top             =   120
                  Width           =   960
               End
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H8000000C&
               BorderStyle     =   0  'None
               Height          =   1275
               Index           =   7
               Left            =   30
               ScaleHeight     =   1275
               ScaleWidth      =   10755
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   30
               Width           =   10755
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "客户订单领料记录:"
                  ForeColor       =   &H8000000E&
                  Height          =   195
                  Index           =   19
                  Left            =   180
                  TabIndex        =   24
                  Top             =   120
                  Width           =   1500
               End
            End
            Begin TrueOleDBGrid80.TDBGrid TDBGSpend 
               Height          =   3855
               Left            =   30
               TabIndex        =   46
               Top             =   1305
               Width           =   8055
               _ExtentX        =   14208
               _ExtentY        =   6800
               _LayoutType     =   0
               _RowHeight      =   23
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
               Splits(0).ShowCollapseExpandIcons=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).AllowColSelect=   0   'False
               Splits(0).AllowColMove=   -1  'True
               Splits(0).AllowRowSelect=   0   'False
               Splits(0).DividerColor=   13160660
               Splits(0).FilterBar=   -1  'True
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=2"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=273"
               Splits(0)._ColumnProps(5)=   "Column(0).WrapText=1"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=3281"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
               Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=273"
               Splits(0)._ColumnProps(11)=   "Column(1).WrapText=1"
               Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
               PrintInfos(0).PageFooterFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               Appearance      =   0
               ColumnFooters   =   -1  'True
               DefColWidth     =   0
               HeadLines       =   1.75
               FootLines       =   1.1
               PictureCurrentRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
               PictureCurrentRow(0)=   "bHQAAA4DAABCTQ4DAAAAAAAANgAAACgAAAARAAAADgAAAAEAGAAAAAAA2AIAAAAAAAAAAAAAAAAA"
               PictureCurrentRow(1)=   "AAAAAACltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrUA"
               PictureCurrentRow(2)=   "lKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKalAKW2taW2"
               PictureCurrentRow(3)=   "taW2taW2taW2taW2taW2taW2taW2taW2tZSmpaW2taW2taW2taW2taW2taW2tQCtvr2tvr2tvr2t"
               PictureCurrentRow(4)=   "vr2tvr2tvr2tvr2tvr2tvr2tvr0YGBicrq2tvr2tvr2tvr2tvr2tvr0ArcfGrcfGrcfGrcfGrcfG"
               PictureCurrentRow(5)=   "rcfGrcfGrcfGrcfGrcfGAAAAGBgYnK6trcfGrcfGrcfGrcfGALXHxrXHxrXHxrXHxrXHxrXHxrXH"
               PictureCurrentRow(6)=   "xrXHxrXHxrXHxgAAAAAAABgYGKW2tbXHxrXHxrXHxgC9z869z869z869z869z869z869z869z869"
               PictureCurrentRow(7)=   "z869z84AAAAAAAAAAAAYGBitvrW9z869z84Avc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/O"
               PictureCurrentRow(8)=   "AAAAAAAAAAAAAAAAKTAxvc/Ovc/OAMbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1gAAAAAA"
               PictureCurrentRow(9)=   "AAAAABAQEKW2tcbX1sbX1gDO19bO19bO19bO19bO19bO19bO19bO19bO19bO19YAAAAAAAAQEBC1"
               PictureCurrentRow(10)=   "trXO19bO19bO19YA1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufnAAAAEBAQtb7G1ufn1ufn"
               PictureCurrentRow(11)=   "1ufn1ufnAN7n797n797n797n797n797n797n797n797n797n7xAQEL3Hxt7n797n797n797n797n"
               PictureCurrentRow(12)=   "7wDe7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7++9z87e7+/e7+/e7+/e7+/e7+/e7+8A5+/3"
               PictureCurrentRow(13)=   "5+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/3AA=="
               PictureCurrentRow.vt=   9
               PictureStandardRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
               PictureStandardRow(0)=   "bHQAAN4DAABCTd4DAAAAAAAANgAAACgAAAARAAAAEgAAAAEAGAAAAAAAqAMAAAAAAAAAAAAAAAAA"
               PictureStandardRow(1)=   "AAAAAACMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpwA"
               PictureStandardRow(2)=   "lKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKalAJyurZyu"
               PictureStandardRow(3)=   "rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurQCltrWltrWltrWl"
               PictureStandardRow(4)=   "trWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrUArb69rb69rb69rb69rb69"
               PictureStandardRow(5)=   "rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69AK2+va2+va2+va2+va2+va2+va2+"
               PictureStandardRow(6)=   "va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+vQC1x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8a1"
               PictureStandardRow(7)=   "x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8YAtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfG"
               PictureStandardRow(8)=   "tcfGtcfGtcfGtcfGtcfGtcfGtcfGAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
               PictureStandardRow(9)=   "zr3Pzr3Pzr3Pzr3Pzr3PzgC9z869z869z869z869z869z869z869z869z869z869z869z869z869"
               PictureStandardRow(10)=   "z869z869z869z84Avc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/O"
               PictureStandardRow(11)=   "vc/Ovc/OAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
               PictureStandardRow(12)=   "1gDO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19YA1ufn"
               PictureStandardRow(13)=   "1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufnANbn59bn59bn"
               PictureStandardRow(14)=   "59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wDe7+/e7+/e7+/e7+/e"
               PictureStandardRow(15)=   "7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+8A3u/v3u/v3u/v3u/v3u/v3u/v"
               PictureStandardRow(16)=   "3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/vAN7v797v797v797v797v797v797v797v"
               PictureStandardRow(17)=   "797v797v797v797v797v797v797v797v797v7wA="
               PictureStandardRow.vt=   9
               MultipleLines   =   0
               EmptyRows       =   -1  'True
               CellTipsWidth   =   0
               MultiSelect     =   2
               GroupByCaption  =   "分组示意图"
               DeadAreaBackColor=   16252927
               RowDividerColor =   13160660
               RowSubDividerColor=   13160660
               DirectionAfterEnter=   1
               MaxRows         =   250000
               ViewColumnCaptionWidth=   0
               ViewColumnWidth =   6900.095
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(5)   =   ":id=0,.fontname=宋体"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF7FFFF&,.bold=0,.fontsize=900"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(8)   =   ":id=1,.fontname=宋体"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgpicMode=2,.bold=0,.fontsize=900"
               _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(12)  =   ":id=2,.fontname=宋体"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
               _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(15)  =   ":id=3,.fontname=宋体"
               _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H31CFFF&"
               _StyleDefs(18)  =   ":id=6,.fgcolor=&H80000008&"
               _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
               _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgpicMode=2,.appearance=1"
               _StyleDefs(24)  =   ":id=11,.bgbmp=3"
               _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(26)  =   "Splits(0).Style:id=55,.parent=1,.alignment=2,.valignment=2,.wraptext=-1"
               _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
               _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=56,.parent=2"
               _StyleDefs(29)  =   "Splits(0).FooterStyle:id=57,.parent=3"
               _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=58,.parent=5"
               _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=60,.parent=6"
               _StyleDefs(32)  =   "Splits(0).EditorStyle:id=59,.parent=7"
               _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=61,.parent=8"
               _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=62,.parent=9,.bgcolor=&HFFFFFF&"
               _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
               _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
               _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
               _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=82,.parent=55"
               _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=56,.alignment=0"
               _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=57"
               _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=59"
               _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=86,.parent=55"
               _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=56,.alignment=0"
               _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=57"
               _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=59"
               _StyleDefs(46)  =   "Named:id=33:Normal"
               _StyleDefs(47)  =   ":id=33,.parent=0"
               _StyleDefs(48)  =   "Named:id=34:Heading"
               _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(50)  =   ":id=34,.wraptext=-1,.bgpicMode=1,.bgbmp=1"
               _StyleDefs(51)  =   "Named:id=35:Footing"
               _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(53)  =   ":id=35,.bgpicMode=1,.bgbmp=2"
               _StyleDefs(54)  =   "Named:id=36:Selected"
               _StyleDefs(55)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(56)  =   "Named:id=37:Caption"
               _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(58)  =   "Named:id=38:HighlightRow"
               _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(60)  =   "Named:id=39:EvenRow"
               _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(62)  =   "Named:id=40:OddRow"
               _StyleDefs(63)  =   ":id=40,.parent=33"
               _StyleDefs(64)  =   "Named:id=41:RecordSelector"
               _StyleDefs(65)  =   ":id=41,.parent=34,.bgcolor=&HCEDFDE&,.bgpicMode=0,.borderColor=&H80000005&"
               _StyleDefs(66)  =   "Named:id=42:FilterBar"
               _StyleDefs(67)  =   ":id=42,.parent=33"
               _StyleDefs(68)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
               _StyleDefs(69)  =   "bmp(1):id=1,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
               _StyleDefs(70)  =   "bmp(2):id=1,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
               _StyleDefs(71)  =   "bmp(3):id=1,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(72)  =   "bmp(4):id=1,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(73)  =   "bmp(5):id=1,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
               _StyleDefs(74)  =   "bmp(6):id=1,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
               _StyleDefs(75)  =   "bmp(7):id=1,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(76)  =   "bmp(8):id=1,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(77)  =   "bmp(9):id=1,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(78)  =   "bmp(10):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
               _StyleDefs(79)  =   "bmp(11):id=1,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(80)  =   "bmp(12):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(81)  =   "bmp(13):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
               _StyleDefs(82)  =   "bmp(14):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(83)  =   "bmp(15):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(84)  =   "bmp(16):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
               _StyleDefs(85)  =   "bmp(17):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(86)  =   "bmp(18):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(87)  =   "bmp(19):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
               _StyleDefs(88)  =   "bmp(20):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(89)  =   "bmp(21):id=1,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(90)  =   "bmp(22):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
               _StyleDefs(91)  =   "bmp(23):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(92)  =   "bmp(24):id=1,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(93)  =   "bmp(25):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
               _StyleDefs(94)  =   "bmp(26):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(95)  =   "bmp(27):id=1,797v797v797v7wAAAA=="
               _StyleDefs(96)  =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
               _StyleDefs(97)  =   "bmp(1):id=2,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
               _StyleDefs(98)  =   "bmp(2):id=2,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
               _StyleDefs(99)  =   "bmp(3):id=2,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(100) =   "bmp(4):id=2,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(101) =   "bmp(5):id=2,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
               _StyleDefs(102) =   "bmp(6):id=2,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
               _StyleDefs(103) =   "bmp(7):id=2,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(104) =   "bmp(8):id=2,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(105) =   "bmp(9):id=2,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(106) =   "bmp(10):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
               _StyleDefs(107) =   "bmp(11):id=2,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(108) =   "bmp(12):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(109) =   "bmp(13):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
               _StyleDefs(110) =   "bmp(14):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(111) =   "bmp(15):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(112) =   "bmp(16):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
               _StyleDefs(113) =   "bmp(17):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(114) =   "bmp(18):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(115) =   "bmp(19):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
               _StyleDefs(116) =   "bmp(20):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(117) =   "bmp(21):id=2,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(118) =   "bmp(22):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
               _StyleDefs(119) =   "bmp(23):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(120) =   "bmp(24):id=2,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(121) =   "bmp(25):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
               _StyleDefs(122) =   "bmp(26):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(123) =   "bmp(27):id=2,797v797v797v7wAAAA=="
               _StyleDefs(124) =   "bmp(0):id=3,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
               _StyleDefs(125) =   "bmp(1):id=3,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
               _StyleDefs(126) =   "bmp(2):id=3,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
               _StyleDefs(127) =   "bmp(3):id=3,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(128) =   "bmp(4):id=3,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(129) =   "bmp(5):id=3,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
               _StyleDefs(130) =   "bmp(6):id=3,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
               _StyleDefs(131) =   "bmp(7):id=3,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(132) =   "bmp(8):id=3,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(133) =   "bmp(9):id=3,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(134) =   "bmp(10):id=3,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
               _StyleDefs(135) =   "bmp(11):id=3,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(136) =   "bmp(12):id=3,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(137) =   "bmp(13):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
               _StyleDefs(138) =   "bmp(14):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(139) =   "bmp(15):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(140) =   "bmp(16):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
               _StyleDefs(141) =   "bmp(17):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(142) =   "bmp(18):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(143) =   "bmp(19):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
               _StyleDefs(144) =   "bmp(20):id=3,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(145) =   "bmp(21):id=3,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(146) =   "bmp(22):id=3,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
               _StyleDefs(147) =   "bmp(23):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(148) =   "bmp(24):id=3,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(149) =   "bmp(25):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
               _StyleDefs(150) =   "bmp(26):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(151) =   "bmp(27):id=3,797v797v797v7wAAAA=="
            End
            Begin TrueOleDBGrid80.TDBGrid TDBGWait02 
               Height          =   2580
               Left            =   30
               TabIndex        =   47
               Top             =   5160
               Width           =   8055
               _ExtentX        =   14208
               _ExtentY        =   4551
               _LayoutType     =   0
               _RowHeight      =   23
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
               Splits(0).ShowCollapseExpandIcons=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).AllowColSelect=   0   'False
               Splits(0).AllowColMove=   -1  'True
               Splits(0).AllowRowSelect=   0   'False
               Splits(0).DividerColor=   13160660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=2"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=273"
               Splits(0)._ColumnProps(5)=   "Column(0).WrapText=1"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=3281"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
               Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=273"
               Splits(0)._ColumnProps(11)=   "Column(1).WrapText=1"
               Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
               PrintInfos(0).PageFooterFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               Appearance      =   0
               ColumnFooters   =   -1  'True
               DefColWidth     =   0
               HeadLines       =   1.75
               FootLines       =   1.1
               PictureCurrentRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
               PictureCurrentRow(0)=   "bHQAAA4DAABCTQ4DAAAAAAAANgAAACgAAAARAAAADgAAAAEAGAAAAAAA2AIAAAAAAAAAAAAAAAAA"
               PictureCurrentRow(1)=   "AAAAAACltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrUA"
               PictureCurrentRow(2)=   "lKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKalAKW2taW2"
               PictureCurrentRow(3)=   "taW2taW2taW2taW2taW2taW2taW2taW2tZSmpaW2taW2taW2taW2taW2taW2tQCtvr2tvr2tvr2t"
               PictureCurrentRow(4)=   "vr2tvr2tvr2tvr2tvr2tvr2tvr0YGBicrq2tvr2tvr2tvr2tvr2tvr0ArcfGrcfGrcfGrcfGrcfG"
               PictureCurrentRow(5)=   "rcfGrcfGrcfGrcfGrcfGAAAAGBgYnK6trcfGrcfGrcfGrcfGALXHxrXHxrXHxrXHxrXHxrXHxrXH"
               PictureCurrentRow(6)=   "xrXHxrXHxrXHxgAAAAAAABgYGKW2tbXHxrXHxrXHxgC9z869z869z869z869z869z869z869z869"
               PictureCurrentRow(7)=   "z869z84AAAAAAAAAAAAYGBitvrW9z869z84Avc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/O"
               PictureCurrentRow(8)=   "AAAAAAAAAAAAAAAAKTAxvc/Ovc/OAMbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1gAAAAAA"
               PictureCurrentRow(9)=   "AAAAABAQEKW2tcbX1sbX1gDO19bO19bO19bO19bO19bO19bO19bO19bO19bO19YAAAAAAAAQEBC1"
               PictureCurrentRow(10)=   "trXO19bO19bO19YA1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufnAAAAEBAQtb7G1ufn1ufn"
               PictureCurrentRow(11)=   "1ufn1ufnAN7n797n797n797n797n797n797n797n797n797n7xAQEL3Hxt7n797n797n797n797n"
               PictureCurrentRow(12)=   "7wDe7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7++9z87e7+/e7+/e7+/e7+/e7+/e7+8A5+/3"
               PictureCurrentRow(13)=   "5+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/3AA=="
               PictureCurrentRow.vt=   9
               PictureStandardRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
               PictureStandardRow(0)=   "bHQAAN4DAABCTd4DAAAAAAAANgAAACgAAAARAAAAEgAAAAEAGAAAAAAAqAMAAAAAAAAAAAAAAAAA"
               PictureStandardRow(1)=   "AAAAAACMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpwA"
               PictureStandardRow(2)=   "lKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKalAJyurZyu"
               PictureStandardRow(3)=   "rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurQCltrWltrWltrWl"
               PictureStandardRow(4)=   "trWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrUArb69rb69rb69rb69rb69"
               PictureStandardRow(5)=   "rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69AK2+va2+va2+va2+va2+va2+va2+"
               PictureStandardRow(6)=   "va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+vQC1x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8a1"
               PictureStandardRow(7)=   "x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8YAtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfG"
               PictureStandardRow(8)=   "tcfGtcfGtcfGtcfGtcfGtcfGtcfGAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
               PictureStandardRow(9)=   "zr3Pzr3Pzr3Pzr3Pzr3PzgC9z869z869z869z869z869z869z869z869z869z869z869z869z869"
               PictureStandardRow(10)=   "z869z869z869z84Avc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/O"
               PictureStandardRow(11)=   "vc/Ovc/OAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
               PictureStandardRow(12)=   "1gDO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19YA1ufn"
               PictureStandardRow(13)=   "1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufnANbn59bn59bn"
               PictureStandardRow(14)=   "59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wDe7+/e7+/e7+/e7+/e"
               PictureStandardRow(15)=   "7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+8A3u/v3u/v3u/v3u/v3u/v3u/v"
               PictureStandardRow(16)=   "3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/vAN7v797v797v797v797v797v797v797v"
               PictureStandardRow(17)=   "797v797v797v797v797v797v797v797v797v7wA="
               PictureStandardRow.vt=   9
               MultipleLines   =   0
               EmptyRows       =   -1  'True
               CellTipsWidth   =   0
               MultiSelect     =   2
               GroupByCaption  =   "分组示意图"
               DeadAreaBackColor=   16252927
               RowDividerColor =   13160660
               RowSubDividerColor=   13160660
               DirectionAfterEnter=   1
               MaxRows         =   250000
               ViewColumnCaptionWidth=   0
               ViewColumnWidth =   6900.095
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(5)   =   ":id=0,.fontname=宋体"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF7FFFF&,.bold=0,.fontsize=900"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(8)   =   ":id=1,.fontname=宋体"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgpicMode=2,.bold=0,.fontsize=900"
               _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(12)  =   ":id=2,.fontname=宋体"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
               _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=134"
               _StyleDefs(15)  =   ":id=3,.fontname=宋体"
               _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H31CFFF&"
               _StyleDefs(18)  =   ":id=6,.fgcolor=&H80000008&"
               _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
               _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgpicMode=2,.appearance=1"
               _StyleDefs(24)  =   ":id=11,.bgbmp=3"
               _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(26)  =   "Splits(0).Style:id=55,.parent=1,.alignment=2,.valignment=2,.wraptext=-1"
               _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
               _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=56,.parent=2"
               _StyleDefs(29)  =   "Splits(0).FooterStyle:id=57,.parent=3"
               _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=58,.parent=5"
               _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=60,.parent=6"
               _StyleDefs(32)  =   "Splits(0).EditorStyle:id=59,.parent=7"
               _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=61,.parent=8"
               _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=62,.parent=9,.bgcolor=&HFFFFFF&"
               _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
               _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
               _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
               _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=82,.parent=55"
               _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=56,.alignment=0"
               _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=57"
               _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=59"
               _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=86,.parent=55"
               _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=56,.alignment=0"
               _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=57"
               _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=59"
               _StyleDefs(46)  =   "Named:id=33:Normal"
               _StyleDefs(47)  =   ":id=33,.parent=0"
               _StyleDefs(48)  =   "Named:id=34:Heading"
               _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(50)  =   ":id=34,.wraptext=-1,.bgpicMode=1,.bgbmp=1"
               _StyleDefs(51)  =   "Named:id=35:Footing"
               _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(53)  =   ":id=35,.bgpicMode=1,.bgbmp=2"
               _StyleDefs(54)  =   "Named:id=36:Selected"
               _StyleDefs(55)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(56)  =   "Named:id=37:Caption"
               _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(58)  =   "Named:id=38:HighlightRow"
               _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(60)  =   "Named:id=39:EvenRow"
               _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(62)  =   "Named:id=40:OddRow"
               _StyleDefs(63)  =   ":id=40,.parent=33"
               _StyleDefs(64)  =   "Named:id=41:RecordSelector"
               _StyleDefs(65)  =   ":id=41,.parent=34,.bgcolor=&HCEDFDE&,.bgpicMode=0,.borderColor=&H80000005&"
               _StyleDefs(66)  =   "Named:id=42:FilterBar"
               _StyleDefs(67)  =   ":id=42,.parent=33"
               _StyleDefs(68)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
               _StyleDefs(69)  =   "bmp(1):id=1,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
               _StyleDefs(70)  =   "bmp(2):id=1,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
               _StyleDefs(71)  =   "bmp(3):id=1,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(72)  =   "bmp(4):id=1,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(73)  =   "bmp(5):id=1,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
               _StyleDefs(74)  =   "bmp(6):id=1,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
               _StyleDefs(75)  =   "bmp(7):id=1,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(76)  =   "bmp(8):id=1,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(77)  =   "bmp(9):id=1,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(78)  =   "bmp(10):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
               _StyleDefs(79)  =   "bmp(11):id=1,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(80)  =   "bmp(12):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(81)  =   "bmp(13):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
               _StyleDefs(82)  =   "bmp(14):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(83)  =   "bmp(15):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(84)  =   "bmp(16):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
               _StyleDefs(85)  =   "bmp(17):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(86)  =   "bmp(18):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(87)  =   "bmp(19):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
               _StyleDefs(88)  =   "bmp(20):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(89)  =   "bmp(21):id=1,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(90)  =   "bmp(22):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
               _StyleDefs(91)  =   "bmp(23):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(92)  =   "bmp(24):id=1,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(93)  =   "bmp(25):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
               _StyleDefs(94)  =   "bmp(26):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(95)  =   "bmp(27):id=1,797v797v797v7wAAAA=="
               _StyleDefs(96)  =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
               _StyleDefs(97)  =   "bmp(1):id=2,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
               _StyleDefs(98)  =   "bmp(2):id=2,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
               _StyleDefs(99)  =   "bmp(3):id=2,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(100) =   "bmp(4):id=2,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(101) =   "bmp(5):id=2,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
               _StyleDefs(102) =   "bmp(6):id=2,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
               _StyleDefs(103) =   "bmp(7):id=2,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(104) =   "bmp(8):id=2,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(105) =   "bmp(9):id=2,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(106) =   "bmp(10):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
               _StyleDefs(107) =   "bmp(11):id=2,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(108) =   "bmp(12):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(109) =   "bmp(13):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
               _StyleDefs(110) =   "bmp(14):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(111) =   "bmp(15):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(112) =   "bmp(16):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
               _StyleDefs(113) =   "bmp(17):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(114) =   "bmp(18):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(115) =   "bmp(19):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
               _StyleDefs(116) =   "bmp(20):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(117) =   "bmp(21):id=2,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(118) =   "bmp(22):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
               _StyleDefs(119) =   "bmp(23):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(120) =   "bmp(24):id=2,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(121) =   "bmp(25):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
               _StyleDefs(122) =   "bmp(26):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(123) =   "bmp(27):id=2,797v797v797v7wAAAA=="
               _StyleDefs(124) =   "bmp(0):id=3,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
               _StyleDefs(125) =   "bmp(1):id=3,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
               _StyleDefs(126) =   "bmp(2):id=3,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
               _StyleDefs(127) =   "bmp(3):id=3,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(128) =   "bmp(4):id=3,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
               _StyleDefs(129) =   "bmp(5):id=3,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
               _StyleDefs(130) =   "bmp(6):id=3,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
               _StyleDefs(131) =   "bmp(7):id=3,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(132) =   "bmp(8):id=3,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
               _StyleDefs(133) =   "bmp(9):id=3,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(134) =   "bmp(10):id=3,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
               _StyleDefs(135) =   "bmp(11):id=3,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
               _StyleDefs(136) =   "bmp(12):id=3,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(137) =   "bmp(13):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
               _StyleDefs(138) =   "bmp(14):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(139) =   "bmp(15):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
               _StyleDefs(140) =   "bmp(16):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
               _StyleDefs(141) =   "bmp(17):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(142) =   "bmp(18):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
               _StyleDefs(143) =   "bmp(19):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
               _StyleDefs(144) =   "bmp(20):id=3,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(145) =   "bmp(21):id=3,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
               _StyleDefs(146) =   "bmp(22):id=3,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
               _StyleDefs(147) =   "bmp(23):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(148) =   "bmp(24):id=3,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(149) =   "bmp(25):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
               _StyleDefs(150) =   "bmp(26):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
               _StyleDefs(151) =   "bmp(27):id=3,797v797v797v7wAAAA=="
            End
         End
      End
   End
End
Attribute VB_Name = "frmStorageQueryWhite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'使用说明：
'本模板只适用于带有批次的库存模式。
'A_Storage：当前仓库综合管理针对的仓库的名称。
'           两个下拉的ComboBox从G_StorageBillType中查询获取属于该仓库的单据
'           同时还要根据不同的标签页的类型
'           第一个标签页只获取G_StorageBillType的字段B_CreateType=CF -- 可被正向生成的单据
'           第二个标签页只获取G_StorageBillType的字段B_CreateType=CF -- 可被逆向生成的单据
            
'strSQLInv：第一个标签页显示对应仓库的库存的存储过程语句。查询接口都设置为空 -- 查询所有库存
'strSQLInvSpend：第二个标签页显示的领料库存的存储过程语句。查询接口只有起始和终止日期。从UI获取
'A_BLObjectID：仓库的单据。根据其编号获取表示数量的字段。用于锁定待编辑网格中其他所有的字段
'A_BLObjectIDCreated：生成的目标单据的单据编号。原本的设计是一个仓库只有一个单据
'                     即入库单，其中有单据类型的选择（只有采购入库、生成入库等的选择）
'                     而生成的单据需要有该仓库下的所有单据类型，所以又做了一个和单据A_BLObjectID
'                     完全一致的单据，只有单据类型上不同。
'A_rsQuantity：表示数量的字段的集合
'A_rsInv：库存记录集
'A_rsInvSpend：领料库存记录集
'A_rsWait01：第一个待编辑网格的记录集
'A_rsWait02：第二个待编辑网格的记录集
'其他要求：
'库存存储过程（即strSQLInv）中的数量字段不可使用别名，必须和单据中的一致
'因为单据中设定了该名称的字段为表示数量的字段
'
'领料库存的存储过程存在3种数量字段：1.领料数量 2.退料数量 3.库存领料数量
'其中领料数量的字段不可使用别名，后面两种类型的数量字段可随意


'下面是相关的常量
Private Const A_Storage As String = "白坯仓库"
Private Const strSQLInv As String = "exec dbo.[usp_InvWhite] '',''" '库存存储过程
Private Const strSQLInvSpend As String = "" '领料库存
Private Const A_BLObjectID As String = "12B005" '根据该单据对象的设置，将非数量的字段锁定为不可修改
Private Const A_BLObjectIDCreated As String = "12B006"   '被生成的单据对象编号
Private Const A_GridStyleObjectID As String = "13B009"  '设置网格显示用到的对象
Private Const A_InvKeyField As String = "B_GoodsID" '库存表中的主键字段
Private A_rsQuantity As RecordSet  '数量记录集

Private A_rsInv As RecordSet  '库存记录集
Private A_rsInvSpend As RecordSet  '领料库存
Private A_rsWait01 As RecordSet  '待编辑数据记录集
Private A_rsWait02 As RecordSet  '待编辑数据记录集


Private strSQL As String
Public mvarObjectID As String

Private A_BillTypeCombo As New cls_Link_Data_Ctl   '库存可生成的单据类型
Private A_BillTypeComboRe As New cls_Link_Data_Ctl  '领料退料之类的单据类型

Public Property Let ObjectID(ByVal vData As String)
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
    ObjectID = mvarObjectID
End Property

Public Sub LoadObject()

End Sub

'获取库存
Private Sub GetInv()
    Dim cls1 As New clsGridShow
    Set A_rsInv = New RecordSet
    A_rsInv.Open strSQLInv, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    TDBGInv.DataSource = A_rsInv
    
    With cls1
        .ObjectID = A_GridStyleObjectID
        .InitClass TDBGInv, 5
        .ShowGridFormat
    End With
    
    
    '初始化库存网格下面的网格
    InitWait01
    
    
    '网格控件汇总
    GetStat A_rsInv, TDBGInv, "B_Qty"
    GetStat A_rsInv, TDBGInv, "B_BoxQty"
    
    '锁定数量列不可定位进入光标
    LockInvCol
End Sub


'获取领料库存
Private Sub GetInvSpend()
    Dim cls1 As New clsGridShow
    Set A_rsInvSpend = New RecordSet
    
    
    strSQL = "exec dbo." & strSQLInvSpend & " '" & DTPSpendSDate & "','" & DTPSpendEDate & "'"
    A_rsInvSpend.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    Debug.Print strSQL
    
    TDBGSpend.DataSource = A_rsInvSpend
    
    With cls1
        .ObjectID = "110212"
        .InitClass TDBGSpend, 3
        .ShowGridFormat
    End With
    
    
    '初始化库存网格下面的网格
    InitWait02
End Sub

Private Sub InitDate()
    Dim szEDate As String
    Dim szSDate As String
    
    szEDate = Format(Now, "YYYY-MM-DD")
    szSDate = DateAdd("d", -7, szEDate)
    szSDate = Format(szSDate, "YYYY-MM-DD")
    
    DTPSpendSDate = szSDate
    DTPSpendEDate = szEDate
End Sub


'初始化窗体
Private Sub InitFrm()
    C1Tab1.CurrTab = 0
    
    '初始化日期控件
    InitDate
    
    '初始化单据类型
    InitBillType
    
    '获取库存
    GetInv
    
End Sub

Private Sub BackstageButton1_Click(Index As Integer)
    Select Case Index
        Case 0   '从库存生成单据
            CreateBill_FromInv
        Case 1, 4
            Unload Me
        Case 2   '刷新库存
            GetInv
            
        Case 5 '刷新领料库存
            GetInvSpend
        Case 3  '生成领料退料
            CreateBill_FromInvSpend
            
            
        Case 6   '库存 - 删除当前行
            DelItem_Inv
        Case 7   '库存 - 清空网格数据
            ClearItem_Inv
            
        Case 8   '第二个待编辑网格 - 删除当前行
            DelItem_InvSpend
        Case 9   '第二个待编辑网格 - 清空网格数据
            ClearItem_InvSpend
    End Select
End Sub

Private Sub Form_Load()
    InitFrm
End Sub

Private Sub InitWait01()
    '拷贝A_RS1的字段框架，并且绑定到下面的网格中
    FillRSFrame A_rsInv, A_rsWait01
    TDBGWait01.DataSource = A_rsWait01
    
    Dim cls1 As New clsGridShow
    With cls1
        .ObjectID = A_GridStyleObjectID
        .InitClass TDBGWait01, 5
        .ShowGridFormat
    End With
    
    LockedTDBG TDBGWait01
End Sub

'领料退料之类的单据类型
Private Sub InitWait02()
    '拷贝A_RS1的字段框架，并且绑定到下面的网格中
    FillRSFrameExcept A_rsInvSpend, A_rsWait02, "B_Qty02,B_BoxQty02,B_Qty01,B_BoxQty01", ","
    TDBGWait02.DataSource = A_rsWait02
    
    Dim cls1 As New clsGridShow
    With cls1
        .ObjectID = "110213"
        .InitClass TDBGWait02, 3
        .ShowGridFormat
    End With
    
    LockedTDBG TDBGWait02
    
End Sub

'将所有非数量的字段都锁定为不可编辑
Private Sub LockedTDBG(ByRef vTDBG As TDBGrid)
    Dim rs As New RecordSet
    Dim szFields As String
    
    strSQL = "SELECT * FROM G_BLField AS gb"
    strSQL = strSQL & " WHERE gb.B_ObjectID='" & A_BLObjectID & "'"
    strSQL = strSQL & " AND abs(isnull(gb.B_IsQuantity,0))=1"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    
    '数量字段记录集
    Set A_rsQuantity = New RecordSet
    A_rsQuantity.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    szFields = ""
    Do While Not rs.EOF
        szFields = szFields & rs!B_FieldName & ","
        rs.movenext
    Loop
    
    If Right$(szFields, 1) = "," Then
        szFields = Left$(szFields, Len(szFields) - 1)
    End If
    
    Dim cls1 As New clsGridShow
    cls1.SetColLockedExcept vTDBG, szFields, ","
    
    Set cls1 = Nothing
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Select2Wait01()
    Dim i As Long
    A_rsWait01.AddNew
    For i = 0 To A_rsWait01.Fields.Count - 1
        A_rsWait01.Fields(i).Value = A_rsInv.Fields(i).Value
    Next
End Sub

Private Sub Select2Wait02()
    Dim i As Long
    A_rsWait02.AddNew
    For i = 0 To A_rsWait02.Fields.Count - 1
        A_rsWait02.Fields(i).Value = A_rsInvSpend.Fields(i).Value
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Gm.CacheFrms.DelFrm mvarObjectID
End Sub

Private Sub TDBGInv_DblClick()
    Select2Wait01
End Sub

Private Sub CreateBill_FromInv()
    If JudgeCreateInv = False Then
        Exit Sub
    End If

    CreateBL A_BillTypeCombo.GetData, A_rsWait01
End Sub




'生成除   采购入库单   之外的所有单据
Private Sub CreateBL(ByVal vBillType As String, ByRef vRsWait As RecordSet)
    Dim clsAutoCreateBL1 As clsAutoCreateBL
    Dim rsBill As RecordSet
    Dim rsDetail As RecordSet
    Dim szTemp As String
    Dim rsWait As New RecordSet
    Set rsWait = vRsWait.Clone
    
    
    Set clsAutoCreateBL1 = New clsAutoCreateBL
    clsAutoCreateBL1.InitCls A_BLObjectIDCreated
    
    
    '写草稿主表数据
    Set rsBill = New RecordSet
    rsBill.Fields.Append "B_BillType", adVarChar, 100
    rsBill.Fields.Append "B_UserName", adVarChar, 100
    rsBill.Fields.Append "B_ContactCom", adVarChar, 100
    
    rsBill.Open
    rsBill.AddNew
    rsBill!B_BillType = vBillType
    rsBill!B_UserName = Gm.SysID.SystemUser
    rsBill!B_ContactCom = rsWait!B_producer
    
    clsAutoCreateBL1.CreateOneBillByRs rsBill
    rsBill.Close
    Set rsBill = Nothing
    
    
    '写草稿明细表数据
    Set rsDetail = New RecordSet
    rsDetail.Fields.Append "B_GoodsID", adVarChar, 100
    rsDetail.Fields.Append "B_Width", adVarChar, 100
    rsDetail.Fields.Append "B_UnitWeight", adVarChar, 100
    rsDetail.Fields.Append "B_ItemIDB", adVarChar, 100
    A_rsQuantity.MoveFirst
    Do While Not A_rsQuantity.EOF
        rsDetail.Fields.Append A_rsQuantity!B_FieldName, adDouble
        A_rsQuantity.movenext
    Loop
    
    rsDetail.Open
    
    
    rsWait.MoveFirst
    Do While Not rsWait.EOF
        rsDetail.AddNew
        rsDetail!B_GoodsID = rsWait!B_GoodsID
        rsDetail!B_width = rsWait!B_width
        rsDetail!B_UnitWeight = rsWait!B_UnitWeight
        rsDetail!B_ItemIDB = rsWait!B_ItemIDB
        
        A_rsQuantity.MoveFirst
        Do While Not A_rsQuantity.EOF
            szTemp = Trim(A_rsQuantity!B_FieldName)
            rsDetail(szTemp) = rsWait(szTemp)
            A_rsQuantity.movenext
        Loop
        
        rsWait.movenext
    Loop
    

    clsAutoCreateBL1.CreateDetailsByRs rsDetail
    rsDetail.Close
    Set rsDetail = Nothing
    
    '生成之后不登帐，打开草稿表
    '2. 再登帐为正式表
    'clsAutoCreateBL1.Update
    
    'MsgBox "生成完毕！", vbOKOnly + vbInformation, "提示"
    
    clsAutoCreateBL1.openBill
    
    Set clsAutoCreateBL1 = Nothing
End Sub

'初始化BillType的ComboBox
Private Sub InitBillType()
    '除了  采购入库单   之外的所有其他单据
    strSQL = "SELECT * FROM G_StorageBillType AS gsbt"
    strSQL = strSQL & " WHERE gsbt.B_StorageID='" & A_Storage & "'"
    strSQL = strSQL & " AND gsbt.B_CreateType='" & CREATEFROMINVETORY & "'"

    A_BillTypeCombo.InitCls Combo1, strSQL, "B_BillName", "B_SID", "", "单据类型未初始化！"
    A_BillTypeCombo.ShowData
    
    
    
    '领料退料之类的单据类型
    Set A_BillTypeComboRe = New cls_Link_Data_Ctl
    strSQL = "SELECT * FROM G_StorageBillType AS gsbt"
    strSQL = strSQL & " WHERE gsbt.B_StorageID='" & A_Storage & "'"
    strSQL = strSQL & " AND gsbt.B_CreateType='" & CREATEREVERSE & "'"
    A_BillTypeComboRe.InitCls Combo2, strSQL, "B_BillName", "B_ID", "", "单据类型未初始化！"
    A_BillTypeComboRe.ShowData
End Sub

  
Private Sub ExeTDBGridFilterChange(ByRef vTDBGrid As TDBGrid, ByRef vRs As RecordSet)
    On Error GoTo IFERR
    Dim Col As Integer
    Col = vTDBGrid.Col
       
    vTDBGrid.HoldFields
    vRs.Filter = GetTDBGridFilterString(vTDBGrid)
    vTDBGrid.Col = Col
    vTDBGrid.EditActive = True
       
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "错误发生于对网格控件进行过滤中" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub
  
Private Function GetTDBGridFilterString(ByRef vTDBGrid As TDBGrid) As String
    On Error Resume Next
    Dim tmp As String
    Dim n As Integer
    Dim Col As TrueOleDBGrid80.Column
    Dim cols As TrueOleDBGrid80.Columns
       
    Set cols = vTDBGrid.Columns
       
    For Each Col In cols
        If Trim(Col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
            Select Case Col.DataWidth
                Case 23, 6, 11
                    tmp = tmp & Col.DataField & " =" & Col.FilterText
                Case Else
                    tmp = tmp & Col.DataField & " LIKE '%" & Col.FilterText & "%'"
            End Select
        End If
    Next Col
                      
    GetTDBGridFilterString = tmp
End Function

Private Sub TDBGInv_FilterChange()
    ExeTDBGridFilterChange TDBGInv, A_rsInv
End Sub

Private Sub TDBGSpend_DblClick()
    Select2Wait02
End Sub

Private Sub TDBGSpend_FilterChange()
    ExeTDBGridFilterChange TDBGSpend, A_rsInvSpend
End Sub

'生成领料退料类型的单据
Private Sub CreateBill_FromInvSpend()
    If JudgeCreateInvSpend = False Then
        Exit Sub
    End If
    CreateBL A_BillTypeComboRe.GetData, A_rsWait02
End Sub

'在库存页面中，判断是否可生成
'判断生成的数量不可大于库存数量
Private Function JudgeCreateInv() As Boolean
    JudgeCreateInv = True
    
    Dim lDJZ As String
    lDJZ = A_BillTypeCombo.GetData
    If Len(lDJZ) <= 0 Then
        JudgeCreateInv = False
        MsgBox "没有选择单据类型！", vbOKOnly + vbInformation, "提示"
        Combo1.SetFocus
        Exit Function
    End If
    
    
    If A_rsWait01.State <> adStateOpen Then
        JudgeCreateInv = False
        Exit Function
    End If
    
    If A_rsWait01.RecordCount <= 0 Then
        JudgeCreateInv = False
        Exit Function
    End If
    
    Dim rs1 As New RecordSet
    Dim rs2 As New RecordSet
    Dim szFieldName As String
    Dim dSum As Double
    Dim dNumericOne As Double
    
    Set rs1 = A_rsInv.Clone
    Set rs2 = A_rsWait01.Clone
    
    Do While Not rs2.EOF
        'rs1.Filter = " B_ItemID='" & rs2!B_ItemID & "'"
        rs1.Filter = " " & A_InvKeyField & "='" & rs2(Trim$(A_InvKeyField)) & "'"
        
        A_rsQuantity.MoveFirst
        dSum = 0
        Do While Not A_rsQuantity.EOF
            szFieldName = Trim$(A_rsQuantity!B_FieldName)
            dNumericOne = IIf(IsNull(rs2(szFieldName)), 0, rs2(szFieldName))
            If dNumericOne > rs1(szFieldName) Then
                JudgeCreateInv = False
                MsgBox "生成的数量不可大于库存数量！", vbOKOnly + vbInformation, "提示"
                Exit Function
            End If
            
            '多个数量相加等待本循环外判断：多个数量不可全部为0
            dSum = dSum + dNumericOne
            A_rsQuantity.movenext
        Loop
        
        '多个数量不可全部为0
        If dSum <= 0 Then
            JudgeCreateInv = False
            MsgBox "多个数量至少设置一个！", vbOKOnly + vbInformation, "提示"
            Exit Function
        End If
            
        rs2.movenext
    Loop
    
    rs1.Close
    Set rs1 = Nothing
    
    rs2.Close
    Set rs2 = Nothing
End Function


Private Function JudgeCreateInvSpend() As Boolean
    On Error Resume Next
    JudgeCreateInvSpend = True
    
    Dim lDJZ As Long
    lDJZ = A_BillTypeComboRe.GetData
    If lDJZ <= 0 Then
        JudgeCreateInvSpend = False
        MsgBox "没有选择单据类型！", vbOKOnly + vbInformation, "提示"
        Combo2.SetFocus
        Exit Function
    End If
    
    If A_rsWait02.State <> adStateOpen Then
        JudgeCreateInvSpend = False
        Exit Function
    End If
    
    If A_rsWait02.RecordCount <= 0 Then
        JudgeCreateInvSpend = False
        Exit Function
    End If
    
    Dim rs1 As New RecordSet
    Dim rs2 As New RecordSet
    Dim szFieldName As String
    Dim dSum As Double
    Dim dNumericOne As Double
    
    Set rs1 = A_rsInvSpend.Clone
    Set rs2 = A_rsWait02.Clone
    
    Do While Not rs2.EOF
        rs1.Filter = " B_ItemID='" & rs2!B_itemid & "'"
        
        
        A_rsQuantity.MoveFirst
        dSum = 0 '初始化合计变量，等待在数量循环中统计多个数量的合计数
        Do While Not A_rsQuantity.EOF
            szFieldName = Trim$(A_rsQuantity!B_FieldName)
            dNumericOne = IIf(IsNull(rs2(szFieldName)), 0, rs2(szFieldName))
            If dNumericOne > rs1(szFieldName) Then
                JudgeCreateInvSpend = False
                MsgBox "生成的数量不可大于库存数量！", vbOKOnly + vbInformation, "提示"
                Exit Function
            End If
            
            '多个数量相加等待本循环外判断：多个数量不可全部为0
            dSum = dSum + dNumericOne
            A_rsQuantity.movenext
        Loop
        
        '多个数量不可全部为0
        If dSum <= 0 Then
            JudgeCreateInvSpend = False
            MsgBox "多个数量至少设置一个！", vbOKOnly + vbInformation, "提示"
            Exit Function
        End If
        
        
        rs2.movenext
    Loop
    
    rs1.Close
    Set rs1 = Nothing
    
    rs2.Close
    Set rs2 = Nothing
End Function


Private Sub DelItem_Inv()
    If A_rsWait01.RecordCount > 0 Then
        A_rsWait01.delete
        A_rsWait01.MoveFirst
    End If
End Sub

Private Sub ClearItem_Inv()
    If A_rsWait01.RecordCount > 0 Then
        A_rsWait01.MoveFirst
        Do While Not A_rsWait01.EOF
            A_rsWait01.delete
            A_rsWait01.movenext
        Loop
    End If
End Sub


Private Sub DelItem_InvSpend()
    If A_rsWait02.RecordCount > 0 Then
        A_rsWait02.delete
    End If
End Sub

Private Sub ClearItem_InvSpend()
    If A_rsWait02.RecordCount > 0 Then
        A_rsWait02.MoveFirst
        Do While Not A_rsWait02.EOF
            A_rsWait02.delete
            A_rsWait02.movenext
        Loop
    End If
End Sub


'参数解释：
'设定记录集和绑定网格以及需要统计的字段后，函数自动计算该字段在网格中的位置进行统计
Private Sub GetStat(ByRef rs As ADODB.RecordSet, ByRef TDBGrid1 As TDBGrid, ByVal m_Field As String)
    On Error Resume Next
    Dim i As Long
    Dim j As Long
    Dim sum As Double
    
    
    sum = 0
    i = 0
    j = 0
    For j = 0 To TDBGrid1.Columns.Count - 1
        If TDBGrid1.Columns(j).DataField = m_Field Then
            j = TDBGrid1.Columns(j).ColIndex
            Exit For
        End If
    Next
    
    TDBGrid1.Columns(j).FooterText = ""
    
    rs.MoveFirst
    For i = 0 To rs.RecordCount - 1
        sum = sum + rs.Fields(m_Field)
        rs.movenext
    Next
    
    TDBGrid1.Columns(j).FooterAlignment = dbgRight
    TDBGrid1.Columns(j).FooterText = sum
    TDBGrid1.Columns(j).FooterText = Format(TDBGrid1.Columns(j).FooterText, "###0.0")
End Sub


Private Sub LockInvCol()
    TDBGInv.Columns("B_Qty").AllowFocus = False
    TDBGInv.Columns("B_BoxQty").AllowFocus = False
    TDBGInv.Columns("B_Sum").AllowFocus = False
End Sub


