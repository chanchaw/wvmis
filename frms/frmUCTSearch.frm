VERSION 5.00
Object = "{0E8071F7-A7DD-47AC-95A9-365FFFE096DF}#1.0#0"; "UCButton.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form frmUCTSearch 
   BackColor       =   &H00CEDFDE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "过滤查询"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   3420
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8910
      _cx             =   15716
      _cy             =   6033
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      Version         =   800
      BackColor       =   13557726
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   8
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
      GridRows        =   3
      GridCols        =   3
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmUCTSearch.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00CEDFDE&
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   120
         ScaleHeight     =   570
         ScaleWidth      =   8670
         TabIndex        =   16
         Top             =   2730
         Width           =   8670
         Begin TA_UCButton.UCButton cmdMoveOut 
            Height          =   375
            Left            =   3900
            TabIndex        =   26
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "删除<<"
         End
         Begin TA_UCButton.UCButton cmdMoveIn 
            Height          =   375
            Left            =   2700
            TabIndex        =   25
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "增加>>"
         End
         Begin TA_UCButton.UCButton cmdCancel 
            Height          =   375
            Left            =   7260
            TabIndex        =   24
            Top             =   120
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   661
            Caption         =   "取消  "
            Icon            =   "frmUCTSearch.frx":0068
            IconMask        =   "frmUCTSearch.frx":02FE
            CaptionAlignment=   1
         End
         Begin TA_UCButton.UCButton cmdOK 
            Height          =   375
            Left            =   5760
            TabIndex        =   23
            Top             =   120
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   661
            Caption         =   "确定  "
            Icon            =   "frmUCTSearch.frx":0594
            IconMask        =   "frmUCTSearch.frx":092E
            CaptionAlignment=   1
         End
         Begin TA_UCButton.UCButton cmdReset 
            Height          =   375
            Left            =   60
            TabIndex        =   22
            Top             =   120
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   661
            Caption         =   "重置"
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00CEDFDE&
         BorderStyle     =   0  'None
         Height          =   390
         Index           =   2
         Left            =   5115
         ScaleHeight     =   390
         ScaleWidth      =   3675
         TabIndex        =   14
         Top             =   120
         Width           =   3675
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "已选择查询项目"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00CEDFDE&
         BorderStyle     =   0  'None
         Height          =   390
         Index           =   1
         Left            =   2130
         ScaleHeight     =   390
         ScaleWidth      =   2925
         TabIndex        =   12
         Top             =   120
         Width           =   2925
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "查询条件"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00CEDFDE&
         BorderStyle     =   0  'None
         Height          =   390
         Index           =   0
         Left            =   120
         ScaleHeight     =   390
         ScaleWidth      =   1950
         TabIndex        =   10
         Top             =   120
         Width           =   1950
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "查询项目"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   1215
         End
      End
      Begin TrueOleDBList80.TDBList TDBList2 
         Height          =   2100
         Left            =   5115
         TabIndex        =   17
         Top             =   570
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   3704
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
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).ScrollBars=   0
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   0
         BorderStyle     =   0
         MatchEntry      =   0
         RightToLeft     =   0   'False
         MatchCompare    =   -6
         MatchCol        =   0
         ColumnHeaders   =   0   'False
         ColumnFooters   =   0   'False
         DataMode        =   0
         MultiSelect     =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   0
         Caption         =   ""
         ExposeCellMode  =   0
         LayoutName      =   ""
         LayoutFileName  =   ""
         LayoutUrl       =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         ListField       =   ""
         BoundColumn     =   ""
         IntegralHeight  =   0   'False
         CellTipsWidth   =   0
         CellTipsDelay   =   1000
         RowMember       =   ""
         MouseIcon       =   0
         MouseIcon.vt    =   3
         MousePointer    =   0
         MatchEntryTimeout=   2000
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DataView        =   0
         GroupByCaption  =   "Drag a column header here to group by that column"
         ScrollTrack     =   0   'False
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         AddItemSeparator=   ";"
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=134"
         _StyleDefs(5)   =   ":id=0,.fontname=宋体"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=134"
         _StyleDefs(8)   =   ":id=1,.fontname=宋体"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H31CFFF&"
         _StyleDefs(28)  =   ":id=19,.fgcolor=&H80000008&"
         _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(41)  =   "Named:id=33:Normal"
         _StyleDefs(42)  =   ":id=33,.parent=0"
         _StyleDefs(43)  =   "Named:id=34:Heading"
         _StyleDefs(44)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(45)  =   ":id=34,.wraptext=-1"
         _StyleDefs(46)  =   "Named:id=35:Footing"
         _StyleDefs(47)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   "Named:id=36:Selected"
         _StyleDefs(49)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(50)  =   "Named:id=37:Caption"
         _StyleDefs(51)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(52)  =   "Named:id=38:HighlightRow"
         _StyleDefs(53)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(54)  =   "Named:id=39:EvenRow"
         _StyleDefs(55)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(56)  =   "Named:id=40:OddRow"
         _StyleDefs(57)  =   ":id=40,.parent=33"
         _StyleDefs(58)  =   "Named:id=41:RecordSelector"
         _StyleDefs(59)  =   ":id=41,.parent=34"
         _StyleDefs(60)  =   "Named:id=42:FilterBar"
         _StyleDefs(61)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBList80.TDBList TDBList1 
         Height          =   2100
         Left            =   120
         TabIndex        =   21
         Top             =   570
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   3704
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
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   0
         BorderStyle     =   0
         MatchEntry      =   0
         RightToLeft     =   0   'False
         MatchCompare    =   -6
         MatchCol        =   0
         ColumnHeaders   =   0   'False
         ColumnFooters   =   0   'False
         DataMode        =   0
         MultiSelect     =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   0
         Caption         =   ""
         ExposeCellMode  =   0
         LayoutName      =   ""
         LayoutFileName  =   ""
         LayoutUrl       =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         ListField       =   ""
         BoundColumn     =   ""
         IntegralHeight  =   0   'False
         CellTipsWidth   =   0
         CellTipsDelay   =   1000
         RowMember       =   ""
         MouseIcon       =   0
         MouseIcon.vt    =   3
         MousePointer    =   0
         MatchEntryTimeout=   2000
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DataView        =   0
         GroupByCaption  =   "Drag a column header here to group by that column"
         ScrollTrack     =   0   'False
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         AddItemSeparator=   ";"
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=134"
         _StyleDefs(5)   =   ":id=0,.fontname=宋体"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=宋体"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H31CFFF&"
         _StyleDefs(28)  =   ":id=19,.fgcolor=&H80000008&"
         _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(41)  =   "Named:id=33:Normal"
         _StyleDefs(42)  =   ":id=33,.parent=0"
         _StyleDefs(43)  =   "Named:id=34:Heading"
         _StyleDefs(44)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(45)  =   ":id=34,.wraptext=-1"
         _StyleDefs(46)  =   "Named:id=35:Footing"
         _StyleDefs(47)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   "Named:id=36:Selected"
         _StyleDefs(49)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(50)  =   "Named:id=37:Caption"
         _StyleDefs(51)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(52)  =   "Named:id=38:HighlightRow"
         _StyleDefs(53)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(54)  =   "Named:id=39:EvenRow"
         _StyleDefs(55)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(56)  =   "Named:id=40:OddRow"
         _StyleDefs(57)  =   ":id=40,.parent=33"
         _StyleDefs(58)  =   "Named:id=41:RecordSelector"
         _StyleDefs(59)  =   ":id=41,.parent=34"
         _StyleDefs(60)  =   "Named:id=42:FilterBar"
         _StyleDefs(61)  =   ":id=42,.parent=33"
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2100
         Index           =   1
         Left            =   2130
         ScaleHeight     =   2100
         ScaleWidth      =   2925
         TabIndex        =   1
         Top             =   570
         Visible         =   0   'False
         Width           =   2925
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   180
            TabIndex        =   19
            Top             =   660
            Width           =   2475
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "请输入："
            Height          =   255
            Left            =   180
            TabIndex        =   20
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2100
         Index           =   0
         Left            =   2130
         ScaleHeight     =   2100
         ScaleWidth      =   2925
         TabIndex        =   8
         Top             =   570
         Visible         =   0   'False
         Width           =   2925
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   180
            TabIndex        =   18
            Top             =   660
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "请输入："
            Height          =   255
            Left            =   180
            TabIndex        =   9
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2100
         Index           =   2
         Left            =   2130
         ScaleHeight     =   2100
         ScaleWidth      =   2925
         TabIndex        =   2
         Top             =   570
         Visible         =   0   'False
         Width           =   2925
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   780
            TabIndex        =   3
            Top             =   1260
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            Format          =   269287425
            CurrentDate     =   38665
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   780
            TabIndex        =   4
            Top             =   720
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            Format          =   269287425
            CurrentDate     =   38665
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "和"
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "介于"
            Height          =   195
            Left            =   180
            TabIndex        =   6
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "日期："
            Height          =   255
            Left            =   180
            TabIndex        =   5
            Top             =   300
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frmUCTSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'----调用参数
Public rs        As New Recordset
Public ObjectID  As String
Public FieldType As Integer
'----返回值
Public strResult As String
Public OK        As Boolean
'----模块内变量
Dim rsF          As New Recordset
Dim rsSelected   As New Recordset
Dim rsCN         As New Recordset
Dim rsItem       As New Recordset
Dim curRow       As Integer
Dim TmpRs        As New Recordset

Private Sub GetSearchField()
    Dim strSQL As String
    Dim i      As Integer
    
    On Error Resume Next
    rsF.Fields.Append "B_FieldName", adVarChar, 40, adFldIsNullable '
    rsF.Fields.Append "B_CnName", adVarChar, 40, adFldIsNullable
    rsF.Fields.Append "B_Type", adInteger, 5, adFldIsNullable '
    rsF.Open
    
    For i = 0 To rs.Fields.Count - 1
        rsF.AddNew
        rsF("B_FieldName") = rs.Fields(i).name
        
        Select Case rs.Fields(i).Type
            Case adVarChar, adChar, adLongVarChar
                rsF("B_Type") = 1      ' 1--String型
            Case adInteger, adSingle, adSmallInt, adDouble, adDecimal, adBigInt
                rsF("B_Type") = 2      ' 2--数值型
            Case adDate, adDBDate, adDBTime, adDBTimeStamp
                rsF("B_Type") = 3      ' 3--日期型
        End Select
        rsF.Update
    Next
    
    '----根据字段类型打开对应表
    Select Case FieldType
        Case 1
            strSQL = "Select * From G_FieldSystem"
        Case 2
            strSQL = "Select * From G_FieldUser"
        Case 3
            strSQL = "Select * From G_BLSField Where B_ObjectID='" & Trim(ObjectID) & "'"
        Case 4
            strSQL = "Select * From G_BLField  Where B_ObjectID='" & Trim(ObjectID) & "'"
        Case 5
            strSQL = "Select * From G_BLRField  Where B_ObjectID='" & Trim(ObjectID) & "'"
    End Select
    rsCN.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
        
    
    rsF.MoveFirst
    Do While Not rsF.EOF
        rsCN.Filter = "B_FieldName='" & rsF("B_FieldName") & "'"
        
        If Not IsNull(rsCN("B_CnName")) Then
            rsF("B_CnName") = rsCN("B_CnName")
            rsF.Update
        End If
        
        rsF.MoveNext
    Loop
    
End Sub

Private Sub cmdReset_Click()
    strResult = ""
    OK = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    WhichPictureShow
End Sub
Private Sub Form_Load()
    
    AnimateForm Me
    
    GetSearchField
    FillTdbList1
    InitSelectedItem
    DTPicker1.Value = Date - 30
    DTPicker2.Value = Date

    TDBList1.Bookmark = 1
    TDBList1.row = 0

End Sub

Private Sub TDBlist1_Click()
    WhichPictureShow
    
End Sub

'----移入需查询项目
Private Sub cmdMoveIn_Click()
    On Error GoTo IFERR
    If TDBList1.row = -1 Then
        MsgBox "请选择查询项目！", vbInformation, "选择项目"
        Exit Sub
    End If
    
    If CheckExistItem Then Exit Sub
    
    rsSelected.AddNew
   
    rsF.Filter = "B_CnName='" & TDBList1.Text & "'"

    Select Case rsF("B_Type")
    
        Case 1      '-- 字符串
            If Len(Text1) = 0 Then
                MsgBox "请输入查询条件！", vbInformation, "查询条件"
                rsSelected.CancelUpdate
                Exit Sub
            End If
            
            rsSelected("ItemValue") = Text1.Text

        Case 2      '---数值
             If Len(Text2) = 0 Then
                MsgBox "请输入查询条件！", vbInformation, "查询条件"
                rsSelected.CancelUpdate
                Exit Sub
            End If
            
            rsSelected("ItemValue") = Text2.Text
            
        Case 3      '--日期
            rsSelected("ItemValue") = DTPicker1.Value & "到" & DTPicker2.Value

    End Select
    
    rsSelected("ItemName") = TDBList1.Text
    
    rsSelected.Update
    TDBList2.Refresh
    
    Exit Sub
    
IFERR:
    rsSelected.CancelUpdate

End Sub

'----移出不需查询的项目
Private Sub cmdMoveOut_Click()
    On Error GoTo IFERR

    If rsSelected.RecordCount < 1 Then Exit Sub
        rsSelected.MoveFirst
        curRow = TDBList2.row
        rsSelected.Move TDBList2.row, 0
        
        If rsSelected.BOF Or rsSelected.EOF Then
            MsgBox "请选择要删除的已选择查询项目！", vbInformation
            Exit Sub
        End If
        rsSelected.Delete
    
    Exit Sub
    
IFERR:
    MsgBox Err.Description
End Sub

Private Sub InitSelectedItem()
    On Error Resume Next
    rsSelected.Fields.Append "ItemName", adVarChar, 20, adFldIsNullable
    rsSelected.Fields.Append "ItemValue", adVarChar, 200, adFldIsNullable
    
    rsSelected.Open
    
    Set TDBList2.DataSource = rsSelected
    Set TDBList2.RowSource = rsSelected
    TDBList2.Refresh
    
End Sub

Private Sub cmdOK_Click()
    On Error GoTo IFERR
    Dim strDate()  As String
    
    If Not rsSelected Is Nothing Then
        If rsSelected.BOF And rsSelected.EOF Then
            If TDBList1.row <> -1 And (Len(Text1) <> 0 Or Len(Text2) <> 0 Or Picture3(2).Visible = True) Then
                cmdMoveIn_Click
                
            End If
        End If
        rsSelected.MoveFirst
        
    End If
    
    Do While Not rsSelected.EOF
                 
        rsF.Filter = "B_CnName ='" & rsSelected("ItemName") & "'"
        Select Case rsF("B_Type")
        
            Case 3      '--日期
                strDate = Split(rsSelected("ItemValue"), "到")
                'strResult = strResult & " And  " & rsF("B_FieldName") & " Between  '" & strDate(0) & "' And '" & strDate(1) & "'"
                strResult = strResult & " And " & rsF("B_FieldName") & " >='" & strDate(0) & "'"
                strResult = strResult & " And " & rsF("B_FieldName") & " <='" & strDate(1) & "'"
            Case 2      '---数值
                strResult = strResult & " And  " & rsF("B_FieldName") & "=" & rsSelected("ItemValue")
            
            Case 1      '--字符串
                strResult = strResult & " And  " & rsF("B_FieldName") & " Like '%" & rsSelected("ItemValue") & "%'"
        
        End Select
        
        rsSelected.MoveNext
    Loop
    
    strResult = Trim(strResult)
    strResult = Trim(Right(strResult, Len(strResult) - 3))
    
    OK = True
    Me.Hide
    Exit Sub
    
IFERR:

    Me.Hide
End Sub

Private Function CheckExistItem() As Boolean

    If rsSelected.RecordCount < 1 Then Exit Function
    rsSelected.MoveFirst
    
    Do Until rsSelected.EOF
        If rsSelected(0) = TDBList1.Text Then
            MsgBox "此查询项目已选择", vbInformation
            CheckExistItem = True
            Exit Do
        End If
        
        rsSelected.MoveNext
    Loop
End Function

Private Sub FillTdbList1()
    On Error Resume Next
    TmpRs.Fields.Append "ItemName", adVarChar, 20, adFldIsNullable
    TmpRs.Open
    
    rsF.MoveFirst
    Do While Not rsF.EOF
        If Len(rsF("B_CnName")) > 0 Then
            TmpRs.AddNew
            TmpRs(0) = rsF("B_CnName")
            TmpRs.Update
        End If
        rsF.MoveNext
    Loop

    Set TDBList1.DataSource = TmpRs
    Set TDBList1.RowSource = TmpRs
    TDBList1.Refresh
    TDBList1.Bookmark = 1
End Sub

Private Sub TDBList2_Click()
    On Error Resume Next
    rsSelected.MoveFirst
    curRow = TDBList2.row
    rsSelected.Move TDBList2.row, 0
End Sub

Private Sub WhichPictureShow()
    On Error Resume Next
    rsF.Filter = "B_CnName='" & TDBList1.Text & "'"

    Select Case rsF("B_Type")

        Case 1       '--字符串
            Picture3(0).Visible = True
            Picture3(1).Visible = False
            Picture3(2).Visible = False
            Text1.SetFocus
        Case 2       '--数值
            Picture3(0).Visible = False
            Picture3(1).Visible = True
            Picture3(2).Visible = False
            Text2.SetFocus
        Case 3       '--日期
            Picture3(0).Visible = False
            Picture3(1).Visible = False
            Picture3(2).Visible = True
            DTPicker1.SetFocus
    End Select
End Sub

Private Sub Text1_Keydown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            cmdMoveIn_Click
            cmdOK.SetFocus
            
        Case vbKeyRight
            cmdMoveIn_Click
        
        Case vbKeyUp, vbKeyDown
            TDBList1.SetFocus
            
    End Select
    
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdMoveIn_Click
            cmdOK.SetFocus
            
        Case vbKeyRight
            cmdMoveIn_Click
            
        Case vbKeyUp, vbKeyDown
            TDBList1.SetFocus
            
    End Select
End Sub

Private Sub TDBList2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then
        cmdMoveOut_Click
    End If
End Sub
Private Sub TDBList1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyRight Then
        WhichPictureShow
    End If
End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyRight Then
        cmdMoveIn_Click
    End If
End Sub

Private Sub DTPicker2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyRight Then
        cmdMoveIn_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not rsF Is Nothing Then
        If rsF.State = adStateOpen Then rsF.Close
        Set rsF = Nothing
    End If
    
    If Not rsSelected Is Nothing Then
        If rsSelected.State = adStateOpen Then rsSelected.Close
        Set rsSelected = Nothing
    End If
    
    If Not rsCN Is Nothing Then
        If rsCN.State = adStateOpen Then rsCN.Close
        Set rsCN = Nothing
    End If
    
    If Not rsItem Is Nothing Then
        If rsItem.State = adStateOpen Then rsItem.Close
        Set rsItem = Nothing
    End If
    
End Sub


