VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{332B766E-0D0F-451B-B35F-358EC95AC208}#1.0#0"; "UCCommonCtls.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.2#0"; "Codejock.CommandBars.v16.2.4.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmStorageRKCKNewFP 
   Caption         =   "��Ʒ����"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStorageRKCKNewFP.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   15000
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   8715
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15000
      _LayoutVersion  =   1
      _ExtentX        =   26458
      _ExtentY        =   15372
      _DataPath       =   ""
      Bands           =   "frmStorageRKCKNewFP.frx":058A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   9435
         Left            =   600
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   60
         Width           =   14175
         _cx             =   25003
         _cy             =   16642
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
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
         ChildSpacing    =   1
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
         _GridInfo       =   $"frmStorageRKCKNewFP.frx":0F5A
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.Frame Frame1 
            Caption         =   "�ѹ�ѡ������"
            Height          =   2100
            Left            =   10050
            TabIndex        =   54
            Top             =   3810
            Width           =   2250
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   195
               Index           =   2
               Left            =   1560
               TabIndex        =   61
               Top             =   1590
               Width           =   360
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Index           =   2
               Left            =   420
               TabIndex        =   60
               Top             =   1500
               Width           =   660
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ƥ"
               Height          =   195
               Index           =   0
               Left            =   1560
               TabIndex        =   58
               Top             =   450
               Width           =   180
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   195
               Index           =   1
               Left            =   1560
               TabIndex        =   57
               Top             =   1050
               Width           =   360
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ƥ��"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Index           =   0
               Left            =   420
               TabIndex        =   56
               Top             =   360
               Width           =   660
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Index           =   1
               Left            =   420
               TabIndex        =   55
               Top             =   960
               Width           =   660
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "ѡ�����"
            Height          =   2100
            Left            =   12315
            TabIndex        =   50
            Top             =   3810
            Width           =   1830
            Begin XtremeCommandBars.BackstageButton ccButton2 
               Height          =   495
               Index           =   0
               Left            =   360
               TabIndex        =   51
               Top             =   300
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "ȫѡ"
            End
            Begin XtremeCommandBars.BackstageButton ccButton2 
               Height          =   495
               Index           =   1
               Left            =   360
               TabIndex        =   52
               Top             =   900
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "ȫ��ѡ"
            End
            Begin XtremeCommandBars.BackstageButton ccButton2 
               Height          =   495
               Index           =   2
               Left            =   360
               TabIndex        =   53
               Top             =   1500
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "��ѡ"
            End
         End
         Begin C1SizerLibCtl.C1Tab C1Tab2 
            Height          =   3765
            Left            =   10050
            TabIndex        =   3
            Top             =   30
            Width           =   4095
            _cx             =   7223
            _cy             =   6641
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
            Appearance      =   2
            MousePointer    =   0
            Version         =   800
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FrontTabColor   =   -2147483633
            BackTabColor    =   15465210
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "���ų���|�ͻ�����"
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   3
            Position        =   0
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   0   'False
            ShowFocusRect   =   0   'False
            TabsPerPage     =   2
            BorderWidth     =   0
            BoldCurrent     =   -1  'True
            DogEars         =   -1  'True
            MultiRow        =   0   'False
            MultiRowOffset  =   200
            CaptionStyle    =   0
            TabHeight       =   300
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   3375
               Left            =   4740
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   345
               Width           =   4005
               _cx             =   7064
               _cy             =   5953
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
               _GridInfo       =   $"frmStorageRKCKNewFP.frx":0FDE
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.PictureBox Picture1 
                  BorderStyle     =   0  'None
                  Height          =   3315
                  Left            =   30
                  ScaleHeight     =   3315
                  ScaleWidth      =   3945
                  TabIndex        =   5
                  TabStop         =   0   'False
                  Top             =   30
                  Width           =   3945
                  Begin TA_UCCommonCtls.UCListBox UCListBox1 
                     Height          =   435
                     Left            =   420
                     TabIndex        =   6
                     Top             =   180
                     Width           =   2655
                     _ExtentX        =   4683
                     _ExtentY        =   767
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   9
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Caption         =   "�ͻ�"
                  End
                  Begin TA_UCCommonCtls.UCListBox UCListBox2 
                     Height          =   435
                     Left            =   420
                     TabIndex        =   7
                     Top             =   720
                     Width           =   2655
                     _ExtentX        =   4683
                     _ExtentY        =   767
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   9
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Caption         =   "������"
                  End
                  Begin XtremeCommandBars.BackstageButton ccButton5 
                     Height          =   495
                     Index           =   0
                     Left            =   480
                     TabIndex        =   8
                     Top             =   1560
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   873
                     _StockProps     =   79
                     Caption         =   "��ȡ���"
                  End
                  Begin XtremeCommandBars.BackstageButton ccButton5 
                     Height          =   495
                     Index           =   1
                     Left            =   2100
                     TabIndex        =   9
                     Top             =   1560
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   873
                     _StockProps     =   79
                     Caption         =   "���ɷ�����"
                  End
                  Begin XtremeCommandBars.BackstageButton ccButton5 
                     Height          =   495
                     Index           =   2
                     Left            =   2100
                     TabIndex        =   10
                     Top             =   2220
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   873
                     _StockProps     =   79
                     Caption         =   "�˳�"
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic4 
               Height          =   3375
               Left            =   45
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   345
               Width           =   4005
               _cx             =   7064
               _cy             =   5953
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
               _GridInfo       =   $"frmStorageRKCKNewFP.frx":1058
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.PictureBox Picture5 
                  BorderStyle     =   0  'None
                  Height          =   3315
                  Left            =   30
                  ScaleHeight     =   3315
                  ScaleWidth      =   3945
                  TabIndex        =   12
                  TabStop         =   0   'False
                  Top             =   30
                  Width           =   3945
                  Begin VB.TextBox Text1 
                     Height          =   375
                     Index           =   0
                     Left            =   840
                     TabIndex        =   0
                     Top             =   180
                     Width           =   1395
                  End
                  Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
                     Bindings        =   "frmStorageRKCKNewFP.frx":10D2
                     Height          =   2580
                     Left            =   240
                     TabIndex        =   13
                     Top             =   660
                     Width           =   2010
                     _ExtentX        =   3545
                     _ExtentY        =   4551
                     _LayoutType     =   0
                     _RowHeight      =   17
                     _WasPersistedAsPixels=   0
                     Columns(0)._VlistStyle=   0
                     Columns(0)._MaxComboItems=   5
                     Columns(0).DataField=   ""
                     Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                     Columns.Count   =   1
                     Splits(0)._UserFlags=   0
                     Splits(0).ExtendRightColumn=   -1  'True
                     Splits(0).AllowFocus=   0   'False
                     Splits(0).RecordSelectors=   0   'False
                     Splits(0).RecordSelectorWidth=   503
                     Splits(0)._SavedRecordSelectors=   0   'False
                     Splits(0).ScrollBars=   2
                     Splits(0).DividerColor=   13160660
                     Splits(0).SpringMode=   0   'False
                     Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                     Splits(0)._ColumnProps(0)=   "Columns.Count=1"
                     Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
                     Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                     Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
                     Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
                     Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                     Splits.Count    =   1
                     PrintInfos(0)._StateFlags=   3
                     PrintInfos(0).Name=   "piInternal 0"
                     PrintInfos(0).PageHeaderFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=����"
                     PrintInfos(0).PageFooterFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=����"
                     PrintInfos(0).PageHeaderHeight=   0
                     PrintInfos(0).PageFooterHeight=   0
                     PrintInfos.Count=   1
                     AllowUpdate     =   0   'False
                     Appearance      =   0
                     ColumnHeaders   =   0   'False
                     DefColWidth     =   0
                     HeadLines       =   1.2
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
                     DeadAreaBackColor=   16252927
                     RowDividerColor =   13160660
                     RowSubDividerColor=   13160660
                     DirectionAfterEnter=   1
                     MaxRows         =   250000
                     ViewColumnCaptionWidth=   0
                     ViewColumnWidth =   0
                     _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
                     _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                     _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                     _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                     _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
                     _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=134"
                     _StyleDefs(5)   =   ":id=0,.fontname=����"
                     _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF7FFFF&,.bold=0,.fontsize=900"
                     _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=134"
                     _StyleDefs(8)   =   ":id=1,.fontname=����"
                     _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                     _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
                     _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=134"
                     _StyleDefs(12)  =   ":id=2,.fontname=����"
                     _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
                     _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=134"
                     _StyleDefs(15)  =   ":id=3,.fontname=����"
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
                     _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1,.alignment=2,.valignment=2"
                     _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                     _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
                     _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                     _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                     _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
                     _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
                     _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
                     _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
                     _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
                     _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
                     _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
                     _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
                     _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
                     _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
                     _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
                     _StyleDefs(42)  =   "Named:id=33:Normal"
                     _StyleDefs(43)  =   ":id=33,.parent=0"
                     _StyleDefs(44)  =   "Named:id=34:Heading"
                     _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                     _StyleDefs(46)  =   ":id=34,.wraptext=-1,.bgpicMode=1,.bgbmp=1"
                     _StyleDefs(47)  =   "Named:id=35:Footing"
                     _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                     _StyleDefs(49)  =   ":id=35,.bgpicMode=1,.bgbmp=2"
                     _StyleDefs(50)  =   "Named:id=36:Selected"
                     _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                     _StyleDefs(52)  =   "Named:id=37:Caption"
                     _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
                     _StyleDefs(54)  =   "Named:id=38:HighlightRow"
                     _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                     _StyleDefs(56)  =   "Named:id=39:EvenRow"
                     _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                     _StyleDefs(58)  =   "Named:id=40:OddRow"
                     _StyleDefs(59)  =   ":id=40,.parent=33"
                     _StyleDefs(60)  =   "Named:id=41:RecordSelector"
                     _StyleDefs(61)  =   ":id=41,.parent=34,.bgcolor=&HCEDFDE&,.bgpicMode=0,.borderColor=&H80000005&"
                     _StyleDefs(62)  =   "Named:id=42:FilterBar"
                     _StyleDefs(63)  =   ":id=42,.parent=33"
                     _StyleDefs(64)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
                     _StyleDefs(65)  =   "bmp(1):id=1,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
                     _StyleDefs(66)  =   "bmp(2):id=1,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
                     _StyleDefs(67)  =   "bmp(3):id=1,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
                     _StyleDefs(68)  =   "bmp(4):id=1,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
                     _StyleDefs(69)  =   "bmp(5):id=1,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
                     _StyleDefs(70)  =   "bmp(6):id=1,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
                     _StyleDefs(71)  =   "bmp(7):id=1,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                     _StyleDefs(72)  =   "bmp(8):id=1,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                     _StyleDefs(73)  =   "bmp(9):id=1,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                     _StyleDefs(74)  =   "bmp(10):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
                     _StyleDefs(75)  =   "bmp(11):id=1,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                     _StyleDefs(76)  =   "bmp(12):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                     _StyleDefs(77)  =   "bmp(13):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
                     _StyleDefs(78)  =   "bmp(14):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                     _StyleDefs(79)  =   "bmp(15):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                     _StyleDefs(80)  =   "bmp(16):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
                     _StyleDefs(81)  =   "bmp(17):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                     _StyleDefs(82)  =   "bmp(18):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                     _StyleDefs(83)  =   "bmp(19):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
                     _StyleDefs(84)  =   "bmp(20):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                     _StyleDefs(85)  =   "bmp(21):id=1,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                     _StyleDefs(86)  =   "bmp(22):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
                     _StyleDefs(87)  =   "bmp(23):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                     _StyleDefs(88)  =   "bmp(24):id=1,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                     _StyleDefs(89)  =   "bmp(25):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
                     _StyleDefs(90)  =   "bmp(26):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                     _StyleDefs(91)  =   "bmp(27):id=1,797v797v797v7wAAAA=="
                     _StyleDefs(92)  =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
                     _StyleDefs(93)  =   "bmp(1):id=2,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
                     _StyleDefs(94)  =   "bmp(2):id=2,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
                     _StyleDefs(95)  =   "bmp(3):id=2,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
                     _StyleDefs(96)  =   "bmp(4):id=2,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
                     _StyleDefs(97)  =   "bmp(5):id=2,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
                     _StyleDefs(98)  =   "bmp(6):id=2,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
                     _StyleDefs(99)  =   "bmp(7):id=2,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                     _StyleDefs(100) =   "bmp(8):id=2,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                     _StyleDefs(101) =   "bmp(9):id=2,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                     _StyleDefs(102) =   "bmp(10):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
                     _StyleDefs(103) =   "bmp(11):id=2,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                     _StyleDefs(104) =   "bmp(12):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                     _StyleDefs(105) =   "bmp(13):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
                     _StyleDefs(106) =   "bmp(14):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                     _StyleDefs(107) =   "bmp(15):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                     _StyleDefs(108) =   "bmp(16):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
                     _StyleDefs(109) =   "bmp(17):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                     _StyleDefs(110) =   "bmp(18):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                     _StyleDefs(111) =   "bmp(19):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
                     _StyleDefs(112) =   "bmp(20):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                     _StyleDefs(113) =   "bmp(21):id=2,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                     _StyleDefs(114) =   "bmp(22):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
                     _StyleDefs(115) =   "bmp(23):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                     _StyleDefs(116) =   "bmp(24):id=2,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                     _StyleDefs(117) =   "bmp(25):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
                     _StyleDefs(118) =   "bmp(26):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                     _StyleDefs(119) =   "bmp(27):id=2,797v797v797v7wAAAA=="
                     _StyleDefs(120) =   "bmp(0):id=3,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
                     _StyleDefs(121) =   "bmp(1):id=3,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
                     _StyleDefs(122) =   "bmp(2):id=3,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
                     _StyleDefs(123) =   "bmp(3):id=3,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
                     _StyleDefs(124) =   "bmp(4):id=3,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
                     _StyleDefs(125) =   "bmp(5):id=3,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
                     _StyleDefs(126) =   "bmp(6):id=3,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
                     _StyleDefs(127) =   "bmp(7):id=3,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                     _StyleDefs(128) =   "bmp(8):id=3,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                     _StyleDefs(129) =   "bmp(9):id=3,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                     _StyleDefs(130) =   "bmp(10):id=3,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
                     _StyleDefs(131) =   "bmp(11):id=3,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                     _StyleDefs(132) =   "bmp(12):id=3,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                     _StyleDefs(133) =   "bmp(13):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
                     _StyleDefs(134) =   "bmp(14):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                     _StyleDefs(135) =   "bmp(15):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                     _StyleDefs(136) =   "bmp(16):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
                     _StyleDefs(137) =   "bmp(17):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                     _StyleDefs(138) =   "bmp(18):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                     _StyleDefs(139) =   "bmp(19):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
                     _StyleDefs(140) =   "bmp(20):id=3,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                     _StyleDefs(141) =   "bmp(21):id=3,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                     _StyleDefs(142) =   "bmp(22):id=3,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
                     _StyleDefs(143) =   "bmp(23):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                     _StyleDefs(144) =   "bmp(24):id=3,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                     _StyleDefs(145) =   "bmp(25):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
                     _StyleDefs(146) =   "bmp(26):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                     _StyleDefs(147) =   "bmp(27):id=3,797v797v797v7wAAAA=="
                  End
                  Begin XtremeCommandBars.BackstageButton ccButton1 
                     Height          =   495
                     Index           =   0
                     Left            =   2580
                     TabIndex        =   14
                     Top             =   180
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   873
                     _StockProps     =   79
                     Caption         =   "��ӿ���"
                  End
                  Begin XtremeCommandBars.BackstageButton ccButton1 
                     Height          =   495
                     Index           =   1
                     Left            =   2580
                     TabIndex        =   15
                     Top             =   810
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   873
                     _StockProps     =   79
                     Caption         =   "��ȡ����"
                  End
                  Begin XtremeCommandBars.BackstageButton ccButton1 
                     Height          =   495
                     Index           =   2
                     Left            =   2580
                     TabIndex        =   16
                     Top             =   2070
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   873
                     _StockProps     =   79
                     Caption         =   "���ɷ�����"
                  End
                  Begin XtremeCommandBars.BackstageButton ccButton1 
                     Height          =   495
                     Index           =   3
                     Left            =   2580
                     TabIndex        =   17
                     Top             =   1440
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   873
                     _StockProps     =   79
                     Caption         =   "��տ����б�"
                  End
                  Begin XtremeCommandBars.BackstageButton ccButton1 
                     Height          =   495
                     Index           =   4
                     Left            =   2580
                     TabIndex        =   18
                     Top             =   2700
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   873
                     _StockProps     =   79
                     Caption         =   "�˳�"
                  End
                  Begin VB.Label Label1 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "���ţ�"
                     Height          =   195
                     Left            =   180
                     TabIndex        =   19
                     Top             =   240
                     Width           =   540
                  End
               End
            End
         End
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   3480
            Left            =   10050
            TabIndex        =   20
            Top             =   5925
            Width           =   4095
            _cx             =   7223
            _cy             =   6138
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
            Appearance      =   2
            MousePointer    =   0
            Version         =   800
            BackColor       =   15465210
            ForeColor       =   -2147483630
            FrontTabColor   =   -2147483633
            BackTabColor    =   15465210
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "�Ƶ������ƺš��ջ���λ|���ٹ�ѡ|���Ŀ���"
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   3
            Position        =   0
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   0   'False
            ShowFocusRect   =   0   'False
            TabsPerPage     =   4
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
            Begin VB.PictureBox Picture4 
               BorderStyle     =   0  'None
               Height          =   3105
               Left            =   45
               ScaleHeight     =   3105
               ScaleWidth      =   4005
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   330
               Width           =   4005
               Begin VB.ComboBox Combo1 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   1500
                  Style           =   2  'Dropdown List
                  TabIndex        =   43
                  Top             =   120
                  Width           =   1995
               End
               Begin VB.TextBox Text2 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   1
                  Left            =   1500
                  TabIndex        =   42
                  Top             =   1980
                  Visible         =   0   'False
                  Width           =   1995
               End
               Begin VB.ComboBox Combo2 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   1500
                  Style           =   2  'Dropdown List
                  TabIndex        =   41
                  Top             =   600
                  Width           =   1995
               End
               Begin VB.ComboBox Combo3 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   1500
                  Style           =   2  'Dropdown List
                  TabIndex        =   40
                  Top             =   1080
                  Width           =   1995
               End
               Begin VB.TextBox Text2 
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   0
                  Left            =   1500
                  TabIndex        =   39
                  Top             =   1560
                  Width           =   1995
               End
               Begin VB.TextBox Text2 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   840
                  Index           =   2
                  Left            =   1500
                  MultiLine       =   -1  'True
                  TabIndex        =   38
                  Top             =   1980
                  Width           =   1995
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�����ˣ�"
                  Height          =   195
                  Index           =   0
                  Left            =   360
                  TabIndex        =   49
                  Top             =   203
                  Width           =   720
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "���ƺţ�"
                  Height          =   195
                  Index           =   1
                  Left            =   360
                  TabIndex        =   48
                  Top             =   683
                  Width           =   720
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�ջ���λ��"
                  Height          =   195
                  Index           =   2
                  Left            =   360
                  TabIndex        =   47
                  Top             =   2070
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "װж����"
                  Height          =   195
                  Index           =   3
                  Left            =   360
                  TabIndex        =   46
                  Top             =   1163
                  Width           =   720
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Ƶ��ˣ�"
                  Height          =   195
                  Index           =   4
                  Left            =   360
                  TabIndex        =   45
                  Top             =   1643
                  Width           =   720
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ע��"
                  Height          =   195
                  Index           =   5
                  Left            =   360
                  TabIndex        =   44
                  Top             =   2070
                  Width           =   540
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   3105
               Left            =   5040
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   330
               Width           =   4005
               _cx             =   7064
               _cy             =   5477
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
               GridRows        =   4
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmStorageRKCKNewFP.frx":10E7
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.PictureBox Picture3 
                  BorderStyle     =   0  'None
                  Height          =   3045
                  Left            =   30
                  ScaleHeight     =   3045
                  ScaleWidth      =   3945
                  TabIndex        =   22
                  TabStop         =   0   'False
                  Top             =   30
                  Width           =   3945
                  Begin VB.TextBox Text3 
                     Height          =   375
                     Left            =   1620
                     TabIndex        =   24
                     Top             =   360
                     Width           =   1935
                  End
                  Begin XtremeCommandBars.BackstageButton ccButton4 
                     Height          =   555
                     Left            =   1620
                     TabIndex        =   23
                     Top             =   960
                     Width           =   1875
                     _Version        =   1048578
                     _ExtentX        =   3307
                     _ExtentY        =   979
                     _StockProps     =   79
                     Caption         =   "����ѡ��������"
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ŀ�Ŀ��ţ�"
                     Height          =   195
                     Left            =   420
                     TabIndex        =   25
                     Top             =   420
                     Width           =   900
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   3105
               Left            =   4740
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   330
               Width           =   4005
               _cx             =   7064
               _cy             =   5477
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
               _GridInfo       =   $"frmStorageRKCKNewFP.frx":1161
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.Frame Frame3 
                  Caption         =   "���ٹ�ѡ"
                  Height          =   2925
                  Left            =   90
                  TabIndex        =   33
                  Top             =   90
                  Width           =   1875
                  Begin VB.CheckBox Check1 
                     Caption         =   "���ÿ��ٹ�ѡ"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   36
                     Top             =   300
                     Width           =   1575
                  End
                  Begin VB.CommandButton Command1 
                     Caption         =   "ѡ��"
                     Enabled         =   0   'False
                     Height          =   315
                     Left            =   120
                     TabIndex        =   35
                     Top             =   1140
                     Width           =   1215
                  End
                  Begin VB.TextBox Text1 
                     Appearance      =   0  'Flat
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   4
                     Left            =   120
                     TabIndex        =   34
                     Top             =   690
                     Width           =   1575
                  End
               End
               Begin VB.Frame Frame4 
                  Caption         =   "���ݲ���"
                  Height          =   2925
                  Left            =   2025
                  TabIndex        =   27
                  Top             =   90
                  Width           =   1890
                  Begin XtremeCommandBars.BackstageButton ccButton3 
                     Height          =   495
                     Index           =   0
                     Left            =   420
                     TabIndex        =   28
                     Top             =   240
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   873
                     _StockProps     =   79
                     Caption         =   "�������"
                  End
                  Begin XtremeCommandBars.BackstageButton ccButton3 
                     Height          =   495
                     Index           =   1
                     Left            =   420
                     TabIndex        =   29
                     Top             =   780
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   873
                     _StockProps     =   79
                     Caption         =   "ɾ����ǰ��"
                  End
                  Begin XtremeCommandBars.BackstageButton ccButton3 
                     Height          =   495
                     Index           =   2
                     Left            =   420
                     TabIndex        =   30
                     Top             =   1320
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   873
                     _StockProps     =   79
                     Caption         =   "ɾ����ǰ����������"
                  End
                  Begin XtremeCommandBars.BackstageButton ccButton3 
                     Height          =   495
                     Index           =   3
                     Left            =   420
                     TabIndex        =   31
                     Top             =   1860
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   873
                     _StockProps     =   79
                     Caption         =   "�޸�KG������"
                  End
                  Begin XtremeCommandBars.BackstageButton ccButton3 
                     Height          =   495
                     Index           =   4
                     Left            =   420
                     TabIndex        =   32
                     Top             =   2400
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   873
                     _StockProps     =   79
                     Caption         =   "ɾ��ѡ��"
                  End
               End
            End
         End
         Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
            Bindings        =   "frmStorageRKCKNewFP.frx":11DB
            Height          =   9375
            Left            =   30
            TabIndex        =   59
            Top             =   30
            Width           =   10005
            _ExtentX        =   17648
            _ExtentY        =   16536
            _LayoutType     =   0
            _RowHeight      =   17
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
            Splits(0).ShowCollapseExpandIcons=   0   'False
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
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
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=131345"
            Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=3281"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3175"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=131345"
            Splits(0)._ColumnProps(13)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=����"
            PrintInfos(0).PageFooterFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=����"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
            ColumnFooters   =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1.2
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
            DataView        =   2
            GroupByCaption  =   "����ʾ��ͼ"
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
            _StyleDefs(5)   =   ":id=0,.fontname=����"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.alignment=2,.bgcolor=&HF7FFFF&,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(8)   =   ":id=1,.fontname=����"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(12)  =   ":id=2,.fontname=����"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(15)  =   ":id=3,.fontname=����"
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
            _StyleDefs(26)  =   "Splits(0).Style:id=55,.parent=1,.valignment=2,.wraptext=-1"
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
Attribute VB_Name = "frmStorageRKCKNewFP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mvarObjectID As String
Private strSQL As String
Private A_rs As New RecordSet

Private A_rsKH As New RecordSet   '���ż�¼

Private A_PS_Checked As Long  '��ǰ�Ѿ�����ѡ��ƥ��
Private A_Qty_Checked As Double  '��ǰ�Ѿ�����ѡ�Ĺ�����
Private A_Meters_Checked As Double  '��ǰ�Ѿ�����ѡ������

Private iGroupColumn As Long
Private X As New XArrayDB
Private A_ItemID As Long

Private A_ID_RK As String   '���ɵ���ⵥ��G_BillCP�е�B_ID������Щ��ⵥ���⣩
Private A_RKCID As String   '���ĵ��ݱ���ַ���
Private A_ClientName As String

Private clsBingComboZD As New cls_Link_Data_Ctl
Private A_ActualInv As New RecordSet  '�ڻ�ȡ���ʱ���γɵ���ʵ����ϸ����

Private Const A_FPBObjID As String = "12B013"  '���ɵ�ɫ���������Ķ�����
Private BinderLoader As New cls_Link_Data_Ctl  'ɫ������װж��
Private Const A_BillType As String = "COL09"   'ɫ��������


Public Property Let ObjectID(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ObjectID = 5
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ObjectID
    ObjectID = mvarObjectID
End Property

'��ʼ�����ż�¼
Private Sub InitRs_KH()
    Set A_rsKH = New RecordSet
    A_rsKH.Fields.Append "B_KH", adVarChar, 100
    A_rsKH.Open
    
    TDBGrid2.DataSource = A_rsKH
End Sub

Private Sub InitFrm()
    '�Ƶ��ˣ�Ĭ�ϵ�¼ϵͳ���û�
    Text2(0).Text = Gm.SysID.SystemUser
    
    
    '��ʼ����ǰ����ѡ������
    A_PS_Checked = 0
    A_Qty_Checked = 0
    A_Meters_Checked = 0
    
    
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
    
    '��ʼ���Ƶ�
    Init_Combo_ZD
    
    
    '��ʾ��ǰ����ѡ������
    ShowCheckedQty
    
    '��ʼ�����ż�¼
    InitRs_KH
    
    
    '��ʼ�����ƺ�
    InitCPH
    
    '��ʼ���ͻ���
    InitSHR
    
    '��ʼ���ͻ�
    InitClients
    
    '���ÿؼ�����Ч��
    g_CJSuite.SetCodejockCtlTheme Me
End Sub

'��ȡ��Ʒ��棬���ؿ��ļ�¼��
Private Function GetFPInv() As RecordSet
    Dim rs As New RecordSet
    
    
    
    Set GetFPInv = rs.Clone
End Function

'���浱ǰ��Ʒ������ݣ��ȴ����������ʱʹ��
Private Sub SetActualInv()
    Set A_ActualInv = New RecordSet
    strSQL = "Select * From G_MidTableCPDRK"
    A_ActualInv.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    
End Sub

'���ݿ��Ż�ȡϸ�뵥�ϵ�����
Private Sub GetRS()
    On Error Resume Next
    Dim rs As New RecordSet
    '�����޸�ʱ��֪ͨ���������ˡ��Ժ��Ʒ��������ˡ�ֻ��ӡ��ϸ�뵥����
'    Dim strZPORCP As String   '����ʾ��Ʒ���Ǵ�Ʒ�����Ƕ���ʾ
'
'    '��ʼ�������ʱ����ʾ��Ʒ����Ʒ������ȫ����ʾ
'    strZPORCP = ""
'    If g_CPDRKZXSZP = 1 Then
'        strZPORCP = "001"
'    End If
'
'    If g_CPDRKZXSCP = 1 Then
'        strZPORCP = "002"
'    End If
    

    If A_rsKH.State <> adStateOpen Then
        MsgBox "û��¼�뿨�ţ�", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If A_rsKH.RecordCount <= 0 Then
        MsgBox "û��¼�뿨�ţ�", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If


    '����ͬʱ�ж���ͻ�
    If JudgeClientUnique = False Then
        MsgBox "�����д��ڶ���ͻ���", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If





    '��ȡָ�����ŵĿ������
    '1. ������м���ʱ��
    '2. д������ϸ����
    strSQL = "exec dbo.[P_InsertDRKCP_NFP] '" & GetKHString & "'"
    Debug.Print strSQL
    Gm.cnnTool.cnn.Execute strSQL

    '���浱ǰ��Ʒ������ݣ��ȴ����������ʱʹ��
    SetActualInv
    

    '���м���ʱ��ͳ�ƺ��γɿ�����ݵȴ���ʾ
    Set rs = New RecordSet
    strSQL = "exec dbo.P_GetDRKCP"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly


    FillUnConnectRecordSet rs, A_rs
    TDBGrid1.DataSource = A_rs


    Dim clsGridShow1 As New clsGridShow
    With clsGridShow1
        .ObjectID = "11S008"
        .InitClass TDBGrid1, 3
        .ShowGridFormat
    End With
    
    TDBGrid1.Columns("B_CheckID").ValueItems.Presentation = dbgCheckBox
    TDBGrid1.Columns("B_CheckID").ValueItems.Translate = True

    '��Ӻϼ�
    GetStat A_rs, TDBGrid1, "B_Qty"
    GetStat A_rs, TDBGrid1, "B_Number"



    '���÷���
    SetTDBGridColumnGroup "B_DateDJ"
    SetTDBGridColumnGroup "B_EDP"
    SetTDBGridColumnGroup "B_ClientName"
    SetTDBGridColumnGroup "B_PinM"
    SetTDBGridColumnGroup "B_MenFuMiChang"
    SetTDBGridColumnGroup "B_Color"
    SetTDBGridColumnGroup "B_SeHao"
    SetTDBGridColumnGroup "B_DTRK"
    


    '����ĳЩ�в�����
    SetTDBGridColumnDisabled "B_Qty"
    SetTDBGridColumnDisabled "B_KQty"
    SetTDBGridColumnDisabled "B_Number"
    SetTDBGridColumnDisabled "B_EDP"
    


    TDBGrid1.FetchRowStyle = True

    GetXArray A_rs

    A_rs.MoveFirst



    '��ձ�ѡ�е�ͳ��
    A_PS_Checked = 0
    A_Qty_Checked = 0
    A_Meters_Checked = 0
    
    
    'ͳ�Ƶ�ǰ����ѡ�ļ�¼
    GetQtyChecked
    
    TDBGrid1.Splits(1).ScrollBars = dbgBoth
End Sub




'��2017��4��14�� 16:36:31Ϊ�ִ����
'ͨ���ͻ��������Ż�ȡ��Ʒ���
Private Sub GetRsEx()
    On Error Resume Next
    Dim rs As New RecordSet


    strSQL = "Delete From G_MidTableCPDRK"
    Gm.cnnTool.cnn.Execute strSQL


    Dim szClientID As String
    Dim szDDH As String
    szClientID = UCListBox1.Text
    szDDH = UCListBox2.Text
    strSQL = "exec dbo.[P_InsertDRKCP_LD] '" & szClientID & "','" & szDDH & "'"
    Gm.cnnTool.cnn.Execute strSQL


    Set rs = New RecordSet
'    If g_CPFH_InvSeque = 0 Then
'        strSQL = "exec dbo.P_GetDRKCP"
'    Else
'        strSQL = "exec dbo.[P_GetDRKCP_Sequ] '" & GetKHString & "' "
'    End If
    strSQL = "exec dbo.P_GetDRKCP"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly


    FillUnConnectRecordSet rs, A_rs
    TDBGrid1.DataSource = A_rs


    Dim clsGridShow1 As New clsGridShow
    With clsGridShow1
        .InitClass TDBGrid1, 2
        .ShowGridFormat
        .ShowGridCtl
    End With


    '��Ӻϼ�
    GetStat A_rs, TDBGrid1, "B_Qty"
    GetStat A_rs, TDBGrid1, "B_Number"



    '���÷���
    SetTDBGridColumnGroup "B_DateDJ"
    SetTDBGridColumnGroup "B_EDP"
    SetTDBGridColumnGroup "B_ClientName"
    SetTDBGridColumnGroup "B_PinM"
    SetTDBGridColumnGroup "B_MenFuMiChang"
    SetTDBGridColumnGroup "B_Color"
    SetTDBGridColumnGroup "B_SeHao"
    SetTDBGridColumnGroup "B_DTRK"
    


    '����ĳЩ�в�����
    SetTDBGridColumnDisabled "B_Qty"
    SetTDBGridColumnDisabled "B_KQty"
    SetTDBGridColumnDisabled "B_Number"
    SetTDBGridColumnDisabled "B_EDP"
    


    TDBGrid1.FetchRowStyle = True

    GetXArray A_rs

    A_rs.MoveFirst



    '��ձ�ѡ�е�ͳ��
    A_PS_Checked = 0
    A_Qty_Checked = 0
    A_Meters_Checked = 0
    
    
    'ͳ�Ƶ�ǰ����ѡ�ļ�¼
    GetQtyChecked
    
    TDBGrid1.Splits(1).ScrollBars = dbgBoth
End Sub


Private Sub FillUnConnectRecordSet(ByRef sRs As RecordSet, ByRef tRs As RecordSet)
    On Error Resume Next
    Dim i As Long
    
    Set tRs = New RecordSet
    For i = 0 To sRs.Fields.Count - 1
        tRs.Fields.Append sRs.Fields(i).name, sRs.Fields(i).Type, sRs.Fields(i).DefinedSize, adFldIsNullable
    Next
    
    tRs.Open
    Do While Not sRs.EOF
        tRs.addnew
        For i = 0 To sRs.Fields.Count - 1
            tRs.Fields(i).Value = IIf(IsNull(sRs.Fields(i).Value), "", sRs.Fields(i).Value)
        Next
        tRs.Update
        sRs.movenext
    Loop
    
End Sub



'�������ͣ�
'�趨��¼���Ͱ������Լ���Ҫͳ�Ƶ��ֶκ󣬺����Զ�������ֶ��������е�λ�ý���ͳ��
Private Sub GetStat(ByRef rs As ADODB.RecordSet, ByRef TDBGrid1 As TDBGrid, ByVal m_Field As String)
    On Error Resume Next
    Dim i As Long
    Dim j As Long
    Dim Sum As Double
    
    
    Sum = 0
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
        Sum = Sum + rs.Fields(m_Field)
        rs.movenext
    Next
    
    TDBGrid1.Columns(j).FooterAlignment = dbgRight
    TDBGrid1.Columns(j).FooterText = Sum / 2
    TDBGrid1.Columns(j).FooterText = Format(TDBGrid1.Columns(j).FooterText, "###0.0#")
End Sub



Private Sub SetTDBGridColumnGroup(ByVal m_FieldName As String)
    Dim Col As TrueOleDBGrid80.Column
    Dim c As Variant
    
    Dim iFatherCol As Long
    
    'iGroupColumn = 1
    
    For Each Col In TDBGrid1.Columns
        If Col.DataField = m_FieldName Then
            
            If TDBGrid1.GroupColumns.Count = 0 Then
                iFatherCol = 0
                Set c = TDBGrid1.GroupColumns.add(TDBGrid1.Columns(m_FieldName).ColIndex, Col)
            Else
                iFatherCol = TDBGrid1.GroupColumns.Count - 1
                Set c = TDBGrid1.GroupColumns.add(TDBGrid1.Columns(m_FieldName).ColIndex, m_FieldName)
            End If
            
            iGroupColumn = iGroupColumn + 1
        End If
    Next
End Sub

Private Sub SetTDBGridColumnDisabled(ByVal m_FieldName As String)
    TDBGrid1.Columns(m_FieldName).Locked = True
    TDBGrid1.Columns(m_FieldName).AllowFocus = False
End Sub


Private Sub GetXArray(ByRef vRs As RecordSet)
    On Error GoTo IFERR
    Dim i As Long
    Dim j As Long
    Dim m_str As String
    Dim rs As New RecordSet
    
    Dim Row As Long, Col As Integer
    
    If vRs.State <> adStateOpen Then
        Exit Sub
    End If
    
    If vRs.RecordCount <= 0 Then
        Exit Sub
    End If
    
    Set rs = vRs.Clone
    '��������:
    '1.�е���ʼ���
    '2.�е���ֹ���
    '3.�е���ʼ���
    '4.�е���ֹ���
    '����������ǽ���¼�����������ݸ��Ƶ������ά������
    X.ReDim 0, rs.RecordCount - 1, 0, rs.Fields.Count - 1
    rs.MoveFirst
    For Row = X.LowerBound(1) To X.UpperBound(1)
        m_str = ""
        For Col = X.LowerBound(2) To X.UpperBound(2)
            X(Row, Col) = rs(Col).Value
            m_str = m_str & rs(Col).Value & "   "
        Next
        Debug.Print m_str
        rs.movenext
    Next
    
    '�������ά����󶨵�����ؼ�
    Set TDBGrid1.Array = X
    
    Exit Sub
    
IFERR:
    MsgBox Err.Description, vbOKOnly + vbInformation, "��ʾ"
    Exit Sub
End Sub


Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "���Ų���-ɾ����ǰѡ��"
            KH_Del
    End Select
End Sub



Private Sub ccButton2_Click(Index As Integer)
    Select Case Index
        Case 0
            SelectAll
        Case 1
            SelectAllNO
        Case 2
            SelectContrary
    End Select
    
    '��ʾ��ǰ����ѡ������
    GetQtyChecked
End Sub

Private Sub ccButton3_Click(Index As Integer)
    Select Case Index
    
        Case 0
            Detail_Add
        Case 1
            Detail_Del
            
        Case 2  'ɾ����ǰ�����µ���������
            DelKHAll
            
        Case 3
            EditeQty   '�޸ĵ�ǰ���ŵ�����
        Case 4   'ɾ��ѡ�е������е�����
            DelSelected
    End Select
End Sub

Private Sub DelSelected()
    On Error Resume Next
    Dim tdbgRow As Variant
    Dim lItemID As Long
    
    If IsProExists = False Then
        Exit Sub
    End If
    
    
    If MsgBox("��ȷ��Ҫɾ����ǰ��ѡ�е�������ô��", vbYesNo + vbDefaultButton2 + vbExclamation, "����") = vbNo Then
        Exit Sub
    End If
    
    
    For Each tdbgRow In TDBGrid1.SelBookmarks
        A_rs.bookmark = tdbgRow
        lItemID = A_rs!B_itemid
        
        A_rs.delete
        
        strSQL = "Delete from G_JRKBill where B_ItemID=" & lItemID
        Gm.cnnTool.cnn.Execute strSQL
    Next
End Sub

Private Sub ccButton4_Click()
    ChangeKH
End Sub

Private Sub ccButton5_Click(Index As Integer)
    Select Case Index
        Case 0  '�ִ�ģʽ�ĳ�Ʒ����ѯ
            GetRsEx
        Case 1   '���ɷ�����
            Call ccButton1_Click(2)
        Case 2
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    InitFrm
End Sub

'��UI����ʾ��ǰ����ѡ��ƥ���͹�����
Private Sub ShowCheckedQty()
    Label3(0).Caption = A_PS_Checked
    Label3(1).Caption = Format(A_Qty_Checked, "0.0")  '����
    Label3(2).Caption = Format(A_Meters_Checked, "0.0")  '����
End Sub

'ȫѡ
Private Sub SelectAll()
    Dim m_bmRow
    Dim m_bmCol
    Dim szTip As String
    
    szTip = "��ǰ����Ϊ�գ�" & vbNewLine
    szTip = szTip & "���ȵ��[��ȡ����]��ť��"
    
    
    If JudgeRsLawless(A_rs) = False Then
        MsgBox szTip, vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    
    m_bmRow = A_rs.bookmark
    m_bmCol = TDBGrid1.Col
    
    
    A_rs.MoveFirst
    Do While Not A_rs.EOF
        'If A_rs("B_CheckID") = 0 Then
            If A_rs("B_HJ") <> 1 Then
                A_rs("B_CheckID") = 1
            End If
        'End If
        A_rs.movenext
    Loop
    
    A_rs.bookmark = m_bmRow
    TDBGrid1.Col = m_bmCol
End Sub

'ȫ��ѡ
Private Sub SelectAllNO()
    Dim m_bmRow
    Dim m_bmCol
    Dim szTip As String
    
    szTip = "��ǰ����Ϊ�գ�" & vbNewLine
    szTip = szTip & "���ȵ��[��ȡ����]��ť��"
    
    
    If JudgeRsLawless(A_rs) = False Then
        MsgBox szTip, vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    m_bmRow = A_rs.bookmark
    m_bmCol = TDBGrid1.Col
    
    A_rs.MoveFirst
    Do While Not A_rs.EOF
        If A_rs("B_CheckID") = 1 Or A_rs("B_CheckID") = -1 Then
            A_rs("B_CheckID") = 0
        End If
        A_rs.movenext
    Loop
    
    
    A_rs.bookmark = m_bmRow
    TDBGrid1.Col = m_bmCol
End Sub

'��ѡ
Private Sub SelectContrary()
    Dim m_bmRow
    Dim m_bmCol
    Dim szTip As String
    
    szTip = "��ǰ����Ϊ�գ�" & vbNewLine
    szTip = szTip & "���ȵ��[��ȡ����]��ť��"
    
    
    If JudgeRsLawless(A_rs) = False Then
        MsgBox szTip, vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    m_bmRow = A_rs.bookmark
    m_bmCol = TDBGrid1.Col
    
    
    A_rs.MoveFirst
    Do While Not A_rs.EOF
        If A_rs("B_CheckID") = 1 Or A_rs("B_CheckID") = -1 Then
            A_rs("B_CheckID") = 0
        Else
            If A_rs("B_HJ") <> 1 Then
                A_rs("B_CheckID") = 1
            End If
        End If
        A_rs.movenext
    Loop
    
        A_rs.bookmark = m_bmRow
    TDBGrid1.Col = m_bmCol
End Sub

'���һ������
Private Sub KH_Add()
    If A_rsKH.State <> adStateOpen Then
        Exit Sub
    End If
    
    Dim szKH As String
    Dim cls1 As New clsFlowCard
    szKH = Trim(Text1(0).Text)
    
    If Val(szKH) <= 0 Then
        MsgBox "���ű���Ϊ�����֣�", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If cls1.CheckBCExists(szKH) = False Then
        MsgBox "��ǰ���Ų����ڣ�", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    Text1(0).Text = ""
    
    A_rsKH.addnew
    A_rsKH!B_KH = szKH
    A_rsKH.Update
End Sub

'ɾ��һ������
Private Sub KH_Del()
    If A_rsKH.State <> adStateOpen Then
        Exit Sub
    End If
    
    If A_rsKH.RecordCount < 1 Then
        Exit Sub
    End If
    
    A_rsKH.delete
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
'    TDBGrid1.Update
'    GetQtyChecked
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    GetQtyChecked
End Sub

Private Sub TDBGrid2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ActiveBar21.Bands("���Ų���").PopupMenu
    End If
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    '�ڿ��ŵ��ı����ϻس��������һ�����ŵ������������
    If Index = 0 Then
        Select Case KeyCode
            Case 13
                KH_Add
        End Select
    End If
    
    If Index = 4 Then
        Select Case KeyCode
            Case 13
                SelectFast
        End Select
        
    End If
End Sub


Private Function GetKHString() As String
    Dim cls1 As clsRecordset
    
    If A_rsKH.State <> adStateOpen Then
        Exit Function
    End If
    
    If A_rsKH.RecordCount <= 0 Then
        Exit Function
    End If
    
    Set cls1 = New clsRecordset
    GetKHString = cls1.RecordSetToString(A_rsKH, "B_KH", ",")
End Function


Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    If TDBGrid1.Columns("B_HJ").CellText(bookmark) = "1" Then
        RowStyle.BackColor = &H808000
    End If
End Sub

'��ȡ��ǰ��ѡ��ƥ���͹������ĺϼ�
Private Sub GetQtyChecked()
    On Error Resume Next
    A_PS_Checked = 0
    A_Qty_Checked = 0
    A_Meters_Checked = 0
    
    
    Dim m_CurCol As Long
    
    If A_rs.State <> adStateOpen Then
        Exit Sub
    End If
    
    Dim m_BookMark
    Dim rs As New RecordSet
    
    m_BookMark = A_rs.bookmark
    m_CurCol = TDBGrid1.Col
    
    
    Set rs = A_rs.Clone
    Debug.Print "����֮ǰ��¼����Ŀ����" & rs.RecordCount
    rs.Filter = " B_CheckID=1 or B_CheckID=-1"
    Debug.Print "����֮�󱻹�ѡ����Ŀ����" & rs.RecordCount
    If rs.RecordCount > 0 Then
        A_PS_Checked = A_PS_Checked + rs.RecordCount
        Do While Not rs.EOF
            A_Qty_Checked = A_Qty_Checked + rs("B_Qty") '����
            A_Meters_Checked = A_Meters_Checked + rs("B_KQty") '����
            rs.movenext
        Loop
    End If
    
    
    ShowCheckedQty
    
End Sub

'���ٹ�ѡ
'ͨ������ĺ���GetXArray01
'����¼������������(�����к�������)�����Ƶ���������
'FIND��ʱ���ǲ��ҷ��ϵ�ǰ�еĿͻ�Ʒ���ȵ���Ϣ��
'ͬʱ��λ��Ŀ�������ĵ�Ԫ����
Private Sub SelectFast()
    Dim m_Qty As Double
    Dim RowFound As Long

    Dim m_bmCol As Long  '��ǰ�е�INDEX
    
    If A_rs.State <> adStateOpen Then
        Exit Sub
    End If
    
    If A_rs.RecordCount <= 0 Then
        Exit Sub
    End If
    

    
    If Val(Trim(Text1(4).Text)) <= 0 Then
        MsgBox "��û��¼������ѯ������!", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    
    
    'ԭʼBOOKMARK
    m_bmCol = TDBGrid1.Col
    m_Qty = Val(Trim(Text1(4).Text))
    
    'X.FIND��������:
    'x.LowerBound(1):�ӵ�ǰ��¼�ĵڼ��п�ʼ����
    'TDBGrid1.Columns("B_Qty").ColIndex:׼�����ҵ��е�INDEX
    'm_Qty:��Ҫ�����ҵ���ֵ
    'XORDER_ASCEND:����ʽ,����,����
    'XCOMP_EQ:�ȽϷ�ʽ<��ǰΪ=>,��=,>,<,>=,<=
    'XTYPE_NUMBER:�����ҵ�Ŀ�����������<��ǰΪ��ֵ>
    RowFound = X.Find(X.LowerBound(1), TDBGrid1.Columns("B_Qty").ColIndex, m_Qty, XORDER_ASCEND, XCOMP_EQ, XTYPE_NUMBER)
    If RowFound >= 0 Then
        TDBGrid1.bookmark = RowFound + 1
    Else
        MsgBox "��ǰ¼�������δ�ҵ�!", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    
    '�����ǰ�����ļ�¼�Ѿ�����ѡ,��ôѰ����һ����������δ����ѡ�ļ�¼
    '���й�ѡ
    '������û���ҵ���Ҫ�˳��ú�����
    If Abs(Val(TDBGrid1.Columns("B_CheckID"))) = 1 Then
        If CheckNextRepeat(m_Qty) = True Then
            Text1(4).Text = ""
        End If
        Exit Sub
    End If
    
    
    '�ѵ�ǰ�еļ�¼���й�ѡ
    TDBGrid1.Columns("B_CheckID") = 1
    A_ItemID = A_rs("B_ItemID")
    TDBGrid1.Update
    
    
    '�ָ�����
    Text1(4).Text = ""
    
    TDBGrid1_AfterColEdit TDBGrid1.Columns("B_Qty").ColIndex
End Sub


Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Text1(4).Enabled = True
        Text1(4).Text = ""
        Command1.Enabled = True
        
        '����ǰ��¼��ɸѡ�ɵ�ǰ��һ���Ĳ�Ʒ
        FilterToCur
    Else
        Text1(4).Enabled = False
        Text1(4).Text = ""
        Command1.Enabled = False
        
        A_rs.Filter = ""
    End If
End Sub


'���������е�����Ϊ��ǰ�е���Ϣ
'�����ϵ�ǰ�еĿͻ�,Ʒ��,�ŷ��ȵȵ���Ϣ
Private Sub FilterToCur()
    Dim m_Client As String
    Dim m_PinMing As String
    Dim m_MenFu As String
    Dim m_Color As String
    Dim m_SH As String
    Dim m_Date As String
    Dim m_strFilter As String
    Dim m_RKDH As String  '��ⵥ�ݱ��
    
    
    If A_rs.State <> adStateOpen Then
        MsgBox "��¼��δ��,����ʹ�ñ�����!", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If A_rs.RecordCount <= 0 Then
        MsgBox "��¼������Ϊ��,����ʹ�ñ�����!", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    
    TDBGrid1.bookmark = 1
    m_Client = TDBGrid1.Columns("B_ClientName").Value
    m_PinMing = IIf(IsNull(TDBGrid1.Columns("B_PinM").Value), 0, TDBGrid1.Columns("B_PinM").Value)
    m_MenFu = IIf(IsNull(TDBGrid1.Columns("B_MenFuMiChang").Value), 0, TDBGrid1.Columns("B_MenFuMiChang").Value)
    m_Color = IIf(IsNull(TDBGrid1.Columns("B_Color").Value), 0, TDBGrid1.Columns("B_Color").Value)
    m_SH = IIf(IsNull(TDBGrid1.Columns("B_SeHao").Value), 0, TDBGrid1.Columns("B_SeHao").Value)
    'm_Date = TDBGrid1.Columns("B_Date190045").Value
    m_Date = TDBGrid1.Columns("B_DTRK").Value
    'm_RKDH = TDBGrid1.Columns("B_CodeID").Value
    

    
    
    m_strFilter = " B_ClientName='" & m_Client & "'"
    m_strFilter = m_strFilter & " And B_PinM='" & m_PinMing & "'"
    m_strFilter = m_strFilter & " And B_MenFuMiChang='" & m_MenFu & "'"
    m_strFilter = m_strFilter & " And B_Color='" & m_Color & "'"
    m_strFilter = m_strFilter & " And B_SeHao='" & m_SH & "'"
    'm_strFilter = m_strFilter & " And B_Date190045='" & m_Date & "'"
    m_strFilter = m_strFilter & " And B_DTRK='" & m_Date & "'"
    'm_strFilter = m_strFilter & " And B_CodeID='" & m_RKDH & "'"
    Debug.Print m_strFilter
    
    A_rs.Filter = m_strFilter
    If A_rs.RecordCount <= 0 Then
        MsgBox "���ϵ�ǰ��Ϣ�ļ�¼������!", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    Debug.Print m_strFilter
End Sub



'��ǰ��������ѡ�����������Ѿ�����ѡ�Ļ�
'��Ӻ���SelectFast01����������,����ѭ��ֱ���ҵ�
'���ϱ����ݵ�δ��ѡ�����ݽ��й�ѡ,���û�з�����
'�������Ѿ�����ѡ�ĸ�����,������ʾ������������ֻ��һ����¼
'����ֵ:FALSE��ʾû���ҵ�,TRUE��ʾ�ҵ���
Private Function CheckNextRepeat(ByVal vQty As Double) As Boolean
    On Error Resume Next
    CheckNextRepeat = False
    TDBGrid1.movenext
    Do While Not A_rs.EOF
        If Abs(Val(TDBGrid1.Columns("B_CheckID"))) <> 1 Then
        If A_rs!B_qty = vQty Then
            TDBGrid1.Columns("B_CheckID") = 1
            A_ItemID = A_rs("B_ItemID")
            TDBGrid1.Update
            
            CheckNextRepeat = True
            Exit Function
        End If
        End If
        A_rs.movenext
    Loop
    
    
    MsgBox "����ǰ��Ҫ��ѡ������:" & vQty & "�Ѿ�����ѡ!", vbOKOnly + vbInformation, "��ʾ"
End Function

'ɾ��һ���������
Private Sub Detail_Del()
    If IsProExists = False Then
        Exit Sub
    End If
    
    If MsgBox("��ȷ��Ҫɾ����ǰ��ѡ�е���ô��", vbYesNo + vbExclamation + vbDefaultButton2, "����") = vbNo Then
        Exit Sub
    End If
    
    On Error Resume Next
    Dim lItemID As Long
    lItemID = A_rs!B_itemid
    'strSQL = "Delete From G_JRKBill where B_ItemID=" & litemID
    strSQL = "exec dbo.usp_JRK_DelOne " & lItemID
    Debug.Print strSQL
    Gm.cnnTool.cnn.Execute strSQL
    
    Dim lBookmark As Long
    lBookmark = TDBGrid1.bookmark
    GetRS
    TDBGrid1.bookmark = lBookmark
    
End Sub

'���һ�����Ƶ���ϸ����
Private Sub Detail_Add()
    On Error Resume Next
    Dim m_str As String
    Dim m_Qty As Double '����
    Dim m_XS As Double  '�ӹ���ת����������ϵ��
    
    If A_rs.State <> adStateOpen Then
        MsgBox "�޿��ʱ�������ò�����", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If A_rs.RecordCount <= 0 Then
        MsgBox "�޿��ʱ�������ò�����", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    
    
    m_str = " ��ȷ��Ҫ���������Ϣ��һƥ��Ʒ��ô��" & vbNewLine
    'm_str = m_str & "���ڣ�" & A_rs!B_DateDJ & vbNewLine
    m_str = m_str & "����Ʒ��" & A_rs!B_CPType & vbNewLine
    m_str = m_str & "�ͻ���" & A_rs!B_ClientName & vbNewLine
    m_str = m_str & "Ʒ����" & A_rs!B_PinM & vbNewLine
    m_str = m_str & "�ŷ���" & A_rs!B_MenFuMiChang & vbNewLine
    m_str = m_str & "��ɫ��" & A_rs!B_color & vbNewLine
    m_str = m_str & "ɫ�ţ�" & A_rs!B_SeHao & vbNewLine
    m_str = m_str & "������ڣ�" & A_rs!B_DTRK & vbNewLine
    
    If MsgBox(m_str, vbYesNo + vbDefaultButton2 + vbInformation, "��ʾ") = vbNo Then
        Exit Sub
    End If
    
    
    m_Qty = Val(InputBox("��¼��������", "����һƥ��", 0))
    If m_Qty <= 0 Then
        MsgBox "��������Ϊ0�����ʧ�ܣ�", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    
    Dim rs As New RecordSet
    Dim rs1 As New RecordSet
    Dim cls1 As New clsFlowCard
    Dim m_BCJ As String  '������
    
    Dim PH1 As Long
    Dim PH2 As Long
    
    
    Set rs = New RecordSet
    strSQL = "Select * From G_JRKBill Where 1=0"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    Set rs1 = New RecordSet
    strSQL = "Select * From G_JRKBill Where B_ItemID=" & A_rs!B_itemid
    rs1.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
        
    If rs1!B_MS > 0 Then
        m_XS = rs1!B_MS / rs1!B_GJ
    Else
        m_XS = 0
    End If
    
    
    
    If rs1.RecordCount > 0 Then
        rs.addnew
        rs!B_id = rs1!B_id
        rs!B_ProcessName = rs1!B_ProcessName
        
'        If rs1!B_EDP = "001" Then
'            m_BCJ = cls1.CreateBarCodeJ(0)
'        End If
'
'        If rs1!B_EDP = "002" Then
'            m_BCJ = cls1.CreateBarCodeJ(1)
'        End If
        m_BCJ = ""
        rs!B_BC = m_BCJ
        
        rs!B_GJ = m_Qty
        rs!B_MS = m_XS * m_Qty
        rs!B_Date = rs1!B_Date
        rs!B_CUN = rs1!B_CUN
        rs!B_CN = rs1!B_CN
        rs!B_IP = rs1!B_IP
        rs!B_BCFC = rs1!B_BCFC
        rs!B_MF = rs1!B_MF
        
        '��ȡ��ǰƥ��
        PH1 = rs1!B_PH1
        GetPHEx rs1!B_id, PH1, PH2
        rs!B_PH1 = rs1!B_PH1
        rs!B_PH2 = PH2
        
        
        rs!B_EDP = rs1!B_EDP
        rs!B_Description = "������������ֶ����"
        rs!B_DateUP = Format(Now, "YYYY-MM-DD HH:MM:SS")
        rs!B_DTRK = rs1!B_DTRK
        
        rs!B_ZGZ = rs1!B_ZGZ
        rs!B_DZ = rs1!B_DZ
        rs!B_KJZ = rs1!B_KJZ
        rs!B_Class = rs1!B_Class
        rs!B_BZFS = rs1!B_BZFS
        rs!B_StaffNO = rs1!B_StaffNO
        rs!B_StaffName = rs1!B_StaffName
        rs!B_KJZNP = rs1!B_KJZNP
        
        rs.Update
    End If
    
    rs1.Close
    Set rs1 = Nothing
    
    rs.Close
    Set rs = Nothing
    
    
    Dim m_bmRow
    Dim m_bmCol
    
    m_bmRow = A_rs.bookmark
    m_bmCol = TDBGrid1.Col
    
    
    '�ֶ�ˢ��
    GetRS
    
    
    A_rs.bookmark = m_bmRow
    TDBGrid1.Col = m_bmCol
End Sub



'����ʽ����һϵ�е�������Ƿ��жϵ������
'����еĻ�����С����ſ�ʼ����
'û�еĻ������������һ�������ۼ�1��ȡ��ǰ�����
Private Sub GetPHEx(ByVal vBIDFC As Long, ByRef PH1 As Long, PH2 As Long)
    Dim rs As New RecordSet
    Dim m_ID As Long   '�ƻ�����G_CJBill�е�B_ID
    
    
    strSQL = "Select G_CJBill.*"
    strSQL = strSQL & " From G_CJBill, G_CJFlowBill"
    strSQL = strSQL & " Where G_CJBill.B_ID = G_CJFlowBill.B_BIDCJBill"
    strSQL = strSQL & " And G_CJFlowBill.B_ID=" & vBIDFC
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        PH1 = 0
        PH2 = 0
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    m_ID = rs!B_id
    PH1 = IIf(IsNull(rs!B_PIShu), 0, rs!B_PIShu)
    rs.Close
    Set rs = Nothing
    
    
    Set rs = New RecordSet
    strSQL = "Select B_PH2"
    strSQL = strSQL & " From"
    strSQL = strSQL & " ("
    strSQL = strSQL & "     Select B_PH2"
    strSQL = strSQL & "     From G_JRKBill"
    strSQL = strSQL & "     Where B_ID In"
    strSQL = strSQL & "     ("
    strSQL = strSQL & "         Select B_ID"
    strSQL = strSQL & "         From G_CJFlowBill"
    strSQL = strSQL & "         Where B_BIDCJBill=" & m_ID
    strSQL = strSQL & "     )"
        
        
    strSQL = strSQL & "     Union All"
        
    strSQL = strSQL & "     Select B_PH2"
    strSQL = strSQL & "     From G_JRKBillDraft"
    strSQL = strSQL & "     Where B_ID In"
    strSQL = strSQL & "     ("
    strSQL = strSQL & "         Select B_ID"
    strSQL = strSQL & "         From G_CJFlowBill"
    strSQL = strSQL & "         Where B_BIDCJBill=" & m_ID
    strSQL = strSQL & "     )"
    strSQL = strSQL & " ) as P"
    strSQL = strSQL & " Order By B_PH2"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    '1.��ǰ����Ϊ��һ��
    If rs.RecordCount <= 0 Then
        PH2 = 1
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    
    
    '2.��ǰ������Ҫ������С��һ���ϵ���ƥ��
    Dim i As Long
    i = 0
    Do While Not rs.EOF
        i = i + 1
        If Val(rs!B_PH2) <> i Then
            PH2 = i
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
        rs.movenext
    Loop
    
    
    
    '3.��ǰ����Ϊ�������+1
    rs.movelast
    PH2 = rs!B_PH2 + 1
    rs.Close
    Set rs = Nothing
End Sub

Private Sub ccButton1_Click(Index As Integer)
    Select Case Index
        Case 0
            KH_Add
        Case 1
            GetRS
        Case 2
            CreateCPRK
        Case 3  '��տ����б�
            ClearKHTable
        Case 4
            Unload Me
    End Select
End Sub

Private Sub ClearKHTable()
    If A_rsKH.State <> adStateOpen Then
        Exit Sub
    End If
    
    If A_rsKH.RecordCount <= 0 Then
        Exit Sub
    End If
    
    A_rsKH.MoveFirst
    Do While Not A_rsKH.EOF
        A_rsKH.delete
        A_rsKH.movenext
    Loop
End Sub

'��һ����¼��ת��Ϊһ���ַ���
'��strFilter���ȴ���0ʱ��ÿ����¼���ø��ַ������
'����Ļ��м䲻�����
'strFieldNameΪĿ���ֶε��ֶ���
Public Function RecordSetToString(ByRef rs As RecordSet, ByVal strFieldName As String, ByVal strFilter As String) As String
    Dim str As String
       
    str = ""
    rs.MoveFirst
    Do While Not rs.EOF
        If Len(Trim(strFilter)) > 0 Then
            str = str & rs(strFieldName) & strFilter
        Else
            str = str & rs(strFieldName)
        End If
        rs.movenext
    Loop
       
    If Len(Trim(strFilter)) > 0 Then
        str = Left(str, Len(str) - 1)
    End If
       
    '�γ��ַ����󣬼�¼���Զ��ƶ�����һ����¼��
    rs.MoveFirst
    RecordSetToString = str
End Function

'�ж��Ƿ���������
'���ڲ����������ڱ�����ȡ����
'�������ն����˲��ֻ���ȫ�����ݵķ�����
'���µ�ǰԤ�����������Ѿ��в��ֻ���ȫ���Ѿ�����
'�������ݵ��������FALSE�����򷵻�TRUE
Private Function HaveDirtyData() As Boolean
    Dim rs As New RecordSet
    strSQL = "exec dbo.[P_InsertDRKCP_NFPCur] '" & GetKHString & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    '��ȡ�����γɵĿ�棬�͵�ǰ���ʱ�Ŀ��Ƚ�
    '��Ŀ��������ͬ����Ϊ���������ݣ���Ʒ��ƥ����
    Dim szTip As String
    If rs.RecordCount <> A_ActualInv.RecordCount Then
        HaveDirtyData = False
        szTip = "���ڲ����������в��ֻ���ȫ�������Ѿ�������" & vbNewLine
        szTip = szTip & "�����»�ȡ������ݺ��ٷ�����"
        MsgBox szTip, vbOKOnly + vbInformation, "��ʾ"
    Else
        HaveDirtyData = True
    End If
    
    rs.Close
    Set rs = Nothing
End Function

'����FALSE�򲻿ɼ������ɷ�����
Private Function JudgeLawless() As Boolean
    If A_rs.State <> adStateOpen Then
        MsgBox "����¼�뿨�ź��ٻ�ȡ���ݣ�", vbOKOnly + vbInformation, "��ʾ"
        JudgeLawless = False
        Exit Function
    End If
    
    If A_rs.RecordCount <= 0 Then
        MsgBox "����¼�뿨�ź��ٻ�ȡ���ݣ�", vbOKOnly + vbInformation, "��ʾ"
        JudgeLawless = False
        Exit Function
    End If

    '1.�жϵ�ǰ��û�м�¼����ѡ
    Dim rs As New RecordSet
    Set rs = A_rs.Clone
    
    rs.Filter = " B_CheckID=1 or B_CheckID=-1"
    If rs.RecordCount <= 0 Then
        JudgeLawless = False
        MsgBox "��ǰû�м�¼����ѡ��", vbOKOnly + vbInformation, "��ʾ"
        Exit Function
    End If
    
    '�ڻ�ȡ����ʱ���Ѿ����˵������Ļ������ﲻ��Ҫ�ظ����
'    If JudgeFH = False Then
'        JudgeLawless = False
'        MsgBox "��ǰ����б䶯��������������������������⣬�����»�ȡ���ݺ��ٷ�����", vbOKOnly + vbInformation, "��ʾ"
'        Exit Function
'    End If
    
    
    '���������
    JudgeLawless = HaveDirtyData
End Function


'��2017��12��24�� 09:40:03
'������Ⱦ����Ǩ������ȫ�ֱ���A_ClientName����������λ���
Private Function JudgeClientUniqueKH() As Boolean
    Dim cls1 As New clsFlowCard
    Dim szKH As String
    Dim rs As RecordSet
    
    
    szKH = RecordSetToString(A_rsKH, "B_KH", ",")
    If cls1.CheckClientUnique(szKH, ",") = False Then
        JudgeClientUniqueKH = False
        A_ClientName = ""
    Else
        JudgeClientUniqueKH = True
        Set rs = cls1.GetClient4CNs(szKH, ",").Clone
        A_ClientName = rs!B_Clientid
    End If
    
End Function

'���ɳ�Ʒ���
Private Sub CreateCPRK()
    '���Ƿ���
    '1. �Ƿ������˿���
    '2. �Ƿ��й�ѡ������
    '3. ���������
    If JudgeLawless = False Then
        Exit Sub
    End If
    
    '���ɳ��ⵥ
    'CreateBillCK A_ID_RK, "120011", "G_DraftBillCP", "G_DraftBillDetailCP", "G_BillCP", "G_BillDetailCP"
    
    Dim szDraftMain As String
    Dim szDraftDetail As String
    Dim szFormalMain As String
    Dim szFormalDetail As String
    
    szDraftMain = "G_DraftBillColor"
    szDraftDetail = "G_DraftBillDetailColor"
    szFormalMain = "G_BillColor"
    szFormalDetail = "G_BillDetailColor"
    
    CreateBillCK A_ID_RK, A_FPBObjID, szDraftMain, _
        szDraftDetail, szFormalMain, szFormalDetail
        
        
    '��տ����б�
    ClearKHTable
    
    A_rs.Close
    Set A_rs = Nothing
End Sub

'�жϵ�ǰ��ʾ�ĳ�Ʒ��漯���Ƿ��Ѿ�����
'�����ǰ�Ŀ����б����Ѿ����ڷ�����¼���򷵻�FALSE�����ɼ�������
'���򷵻�TRUE����״̬�ɼ�������
Private Function JudgeFH() As Boolean
    Dim rs As RecordSet
    strSQL = "exec dbo.P_GetFH '" & GetKHString & "'"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        JudgeFH = False
    Else
        JudgeFH = True
    End If
    
    rs.Close
    Set rs = Nothing
End Function

'Ϊ��Ʒ������ϸ��׼��������
Private Function GetStatistic4DeliveryDetail(ByVal vID As Long) As RecordSet
    Dim rs As New RecordSet
    strSQL = "exec dbo.usp_GetStatistic4FPDelivery '" & vID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Set GetStatistic4DeliveryDetail = rs.Clone
    rs.Close
    Set rs = Nothing
    
End Function

'���ɳ��ⵥ
'vID:��ⵥ�ڱ�G_BillCP�е�B_ID����ⵥ��OBJECTID=120007
Private Sub CreateBillCK(ByVal vID As String, ByVal ObjectID As String, _
    ByVal m_DraftBill As String, ByVal m_DraftBillDetail As String, _
    ByVal m_Bill As String, ByVal m_BillDetail)
    
    Dim rs1 As New RecordSet
    Dim rs2 As New RecordSet
    Dim rs As New RecordSet
    Dim i As Long
    Dim arr
    Dim m_ID As Long   '����������G_BillCP�е�B_ID

    
    
    '���ⵥ�����ϵͳ����
    Dim m_CodeID As String
    Dim m_BID As String
    Dim m_Date As String
    Dim szTemp As String
    
    '����������ͻ���Ʒ����Ϣ
    Dim mvarClient As String

    '�����ɷ�������ʱ��������ɰ�ť
    ForbidBT_FH False
    
    Set rs = New RecordSet
    strSQL = "Select * From G_BL Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount <= 0 Then
        ForbidBT_FH True
        Exit Sub
    End If
    
    m_Date = Format(Now, "YYYY-MM-DD HH:MM:SS")
    m_BID = rs("B_BID")
    m_CodeID = GetFrameCodeDetail(ObjectID)
    rs.Close
    Set rs = Nothing
    
    
    '=================================
    Dim szZD As String   '�Ƶ�
    Dim szCPH As String  '���ƺ�
    Dim szSHDW As String  '�ջ���λ
    Dim szSHR As String   '�ͻ���
    
    szZD = clsBingComboZD.GetData  'ɫ������������
    szCPH = Combo2.Text   'ɫ���������ƺ�
    szSHDW = Trim$(Text2(1).Text)  '�ջ���λ
    'szSHR = Combo3.Text   'ɫ������װж��
    szSHR = BinderLoader.GetData  'ɫ������װж��
    '=================================
    
    
    '��ȡ�ͻ���
    If Len(A_ClientName) <= 0 Then
        A_ClientName = IIf(IsNull(A_rs!B_ClientName), "", A_rs!B_ClientName)
    End If
        
        
    'д��ݸ�����
    strSQL = "Insert Into " & m_DraftBill
    strSQL = strSQL & " (B_CodeID,B_Date,B_BID,B_ObjectID,B_CN,"
    '�ͻ��������ˡ����ƺš��ջ���λ
    strSQL = strSQL & " B_ClientID,B_Operator,B_PlateNumber,B_Revice,"
    'װж�����Ƶ�Ա����ע
    strSQL = strSQL & " B_Loader,B_Producer,B_Memo,B_BillType,B_UserName)"
    strSQL = strSQL & " Values"
    strSQL = strSQL & " ('" & m_CodeID & "','" & m_Date & "','" & m_BID & "','" & ObjectID & "','" & Gm.SysID.ComputerName & "','" & A_ClientName & "','" & szZD & "','" & szCPH & "','" & szSHDW & "','" & szSHR & "','" & Gm.SysID.SystemUser & "','" & Text2(2).Text & "','" & A_BillType & "','" & Gm.SysID.SystemUser & "')"
    Debug.Print strSQL
    Gm.cnnTool.cnn.Execute strSQL
    
    
    
    Set rs = New RecordSet
    strSQL = "Select * From " & m_DraftBill & " Where B_CodeID='" & m_CodeID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount <= 0 Then
        ForbidBT_FH True
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    
    m_ID = rs("B_ID")
    rs.Close
    Set rs = Nothing
    
    
    '=============================================================
    '�������ݵ���ʽ����
    '�Ƚ��ݸ�������ʵ���ʽ����Ϊ�����޸Ĵ����ʽ���е��ֶ�B_FPDID
    '������Ƕ�Ӧ��ʽ���
    Set rs = New RecordSet
    strSQL = "Select * from " & m_DraftBill & " Where B_ID='" & m_ID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    m_Date = Format(rs("B_Date"), "YYYY-MM-DD HH:MM:SS")
    strSQL = "Insert Into " & m_Bill
    strSQL = strSQL & " (B_Closed,B_ID,B_CodeID,B_Date,B_BID,B_ObjectID,"
    strSQL = strSQL & " B_ClientID,B_CN,B_Operator,"
    strSQL = strSQL & " B_PlateNumber,B_Revice,B_Loader,B_Producer,B_Memo,B_BillType,B_UserName)"
    strSQL = strSQL & " Values"
    
    szTemp = ""
    szTemp = " (1,'" & m_ID & "','" & rs("B_CodeID") & "','" & m_Date & "',"
    szTemp = szTemp & " '" & rs("B_BID") & "','" & rs("B_ObjectID") & "',"
    szTemp = szTemp & " '" & rs("B_ClientID") & "','" & rs!B_CN & "',"
    szTemp = szTemp & " '" & rs!B_Operator & "','" & rs!B_PlateNumber & "',"
    szTemp = szTemp & " '" & rs!B_Revice & "','" & rs!B_Loader & "',"
    szTemp = szTemp & " '" & rs!B_Producer & "','" & rs!B_memo & "','" & rs!B_BillType & "','" & rs!B_UserName & "')"
    strSQL = strSQL & szTemp
    Debug.Print strSQL
    Gm.cnnTool.cnn.Execute strSQL
    
    rs.Close
    Set rs = Nothing
    '=============================================================
    
    
    
    
    'д�����ʽ���е��ֶ�B_FPDID
    'ֻдUI�ϱ���ѡ�ļ�¼
    If A_rs.RecordCount > 0 Then
    A_rs.Filter = " B_CheckID=1 or B_CheckID=-1"
    If A_rs.RecordCount > 0 Then
    A_rs.MoveFirst
    Do While Not A_rs.EOF
        strSQL = "update G_JRKBill Set B_FPDID='" & m_ID & "' where B_itemID='" & A_rs!B_itemid & "'"
        Gm.cnnTool.cnn.Execute strSQL
        A_rs.movenext
    Loop
    End If
    End If
    
    
    'д��Ʒ��������ϸ������
    'ͳ�ƴ���G_JRKBill�е����ݣ��ٽ�ϼƻ�����G_CJBill
    '������ʾ�������Ĳ�Ʒ��Ϣ������
    Dim rs4Detail As New RecordSet
    Set rs4Detail = GetStatistic4DeliveryDetail(m_ID).Clone
    
    
    'ʹ�������ͳ������д���Ʒ������ϸ��
    'д����ϸ������ݽ����ǣ������ʹ��������ϵ���ֶ�
    '��Ʒ��Ϣ���Լƻ���
    Set rs = New RecordSet
    strSQL = "Select * from " & m_DraftBillDetail & " where 1=0"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs4Detail.RecordCount > 0 Then
    rs4Detail.MoveFirst
    Do While Not rs4Detail.EOF
        rs.addnew
        rs!B_id = m_ID
        rs!B_FCID = rs4Detail!B_id
        rs!B_ps = rs4Detail!B_ps 'ͳ��ƥ��
        rs!B_kg = rs4Detail!B_GJ 'ͳ�ƹ���
        rs!B_meter = rs4Detail!B_MS 'ͳ������
        rs!B_BoxQty = rs4Detail!B_GJNet '����
        rs.Update
        rs4Detail.movenext
    Loop
    End If
    rs.Close
    Set rs = Nothing
    

    '�������ݵ���ϸ����
    strSQL = "exec dbo.usp_Copy2FormalFPDeliveryDetail '" & m_ID & "'"
    Gm.cnnTool.cnn.Execute strSQL
    
    'ɾ���ݸ�����
    strSQL = "Delete From " & m_DraftBillDetail & " Where B_ID='" & m_ID & "'"
    Gm.cnnTool.cnn.Execute strSQL
    
    
    strSQL = "Delete From " & m_DraftBill & " Where B_ID='" & m_ID & "'"
    Gm.cnnTool.cnn.Execute strSQL
    
    MsgBox "���ɳɹ�!", vbOKOnly + vbInformation, "��ʾ"
    
    ForbidBT_FH True
    
    '�Զ��򿪷�����
    OpenBL Val(Trim(m_ID))
End Sub


'���ɵ��ݱ��B_CodeID
'���ݶ�������ñ�ΪG_BL�����ݴ�������������ţ����Զ���ȡ������ǰ׺����ʽ����
Public Function GetFrameCodeDetail(ByVal m_ObjectID As String) As String
    On Error Resume Next
    Dim strTmpBH As String
    Dim strTmpBHLast As String
    Dim strTmpMonth As String
    Dim strTmpDay As String
    Dim mstrSQL As String
    Dim rs As New RecordSet
    Dim gdateSystemDat As Date
    Dim rstemp As RecordSet
    Dim mvarm_BID As String
    Dim strSQL As String
    Dim m_DraftMainTable As String
    Dim m_MainTable As String
       
    Set rstemp = New RecordSet
    strSQL = "Select * From G_BL Where B_ObjectID='" & m_ObjectID & "'"
    rstemp.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    mvarm_BID = rstemp("B_BID")
    m_DraftMainTable = rstemp("B_DraftMainTable")
    m_MainTable = rstemp("B_MainTable")
    rstemp.Close
    Set rstemp = Nothing
      
    gdateSystemDat = Now
    Set rs = New RecordSet
    strTmpMonth = Trim(Month(gdateSystemDat))
    If Len(Trim(strTmpMonth)) = 1 Then
        strTmpMonth = "0" & strTmpMonth
    End If
    strTmpDay = Trim(Day(gdateSystemDat))
    If Len(Trim(strTmpDay)) = 1 Then
        strTmpDay = "0" & strTmpDay
    End If
    strTmpBH = Trim(mvarm_BID) & Trim(Year(gdateSystemDat)) & strTmpMonth & strTmpDay
    'Debug.Print strTmpBH
    mstrSQL = "Select CASE WHEN ISNULL(P1.B_CodeID,0)>ISNULL(P2.B_CodeID,0) THEN P1.B_CodeID"
    mstrSQL = mstrSQL & " Else P2.B_CodeID End as B_PCodeID"
    mstrSQL = mstrSQL & " From (Select Max(B_CodeID) as B_CodeID"
    mstrSQL = mstrSQL & " From " & m_DraftMainTable
    mstrSQL = mstrSQL & " Where B_CodeID Like '" & Trim(strTmpBH) & "%') as P1,"
    mstrSQL = mstrSQL & " (Select Max(B_CodeID) as B_CodeID"
    mstrSQL = mstrSQL & " From " & m_MainTable
    mstrSQL = mstrSQL & " Where B_CodeID Like '" & Trim(strTmpBH) & "%') as P2"
    'Debug.Print mstrSQL
    rs.Open mstrSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    If IsNull(rs(0)) Then
        '���û����ǰ�ļ�¼
        strTmpBH = strTmpBH & "0001"
        GetFrameCodeDetail = strTmpBH
    Else
        strTmpBHLast = Trim(str(Val(Mid(Trim(rs(0)), 12, 4)) + 1))
        If Len(Trim(strTmpBHLast)) < 4 Then
            strTmpBH = strTmpBH & String(4 - Len(Trim(strTmpBHLast)), "0") & strTmpBHLast
        Else
            strTmpBH = strTmpBH & strTmpBHLast
        End If
        GetFrameCodeDetail = strTmpBH
    End If
    rs.Close
    Set rs = Nothing
End Function



'�����ɺ�,�Զ��򿪳����뵥
Private Sub OpenBL(ByVal m_ID As String)
    Dim clsCommand1 As New clsCommand
    Dim cls1 As New clsIniFile
    
    clsCommand1.InitClass
    
    '��2017��11��19���޸�Ϊ�°汾�ķ�����
    clsCommand1.Execute "12B013", "ɫ��������", "EditObject", Nothing, m_ID
End Sub


Private Sub CreateBill(ByVal ObjectID As String, ByVal m_DraftBill As String, ByVal m_DraftBillDetail As String, ByVal m_Bill As String, ByVal m_BillDetail)
    Dim m_CodeID As String
    Dim m_BID As String
    Dim strSQL As String
    Dim rs As RecordSet
    Dim rs1 As RecordSet
    Dim rs2 As RecordSet
    Dim m_Date As String
    
    Dim m_ID As Long
    Dim rs3 As RecordSet
    
    
    
    If MsgBox("��ȷ��Ҫ����ǰ����ѡ����Ŀ������ⵥô?", vbYesNo + vbDefaultButton2 + vbInformation, "��ʾ") = vbNo Then
        Exit Sub
    End If
    
    
    
    Set rs = New RecordSet
    strSQL = "Select * From G_BL Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    '�����������ɵĵ��ݶ����Ų����ڵĻ����˳�
    If rs.RecordCount <= 0 Then
        Exit Sub
    End If
    
    
    
    '��ȡ��ѡ��ļ�¼�в�ͬ��B_ID(����ȡ��ͬ�ļƻ�����B_ID)
    '========================
    Dim m_ArrayID
    Dim i As Long
    m_ArrayID = Split(GetDistinctValueFromRSEx(A_rs, "B_ID", "B_CheckID"), ",")
    If UBound(m_ArrayID) < 0 Then
        Exit Sub
    End If
    '========================
    
    
    'CSBmk <1.�����ݸ��2.������ʽ��3.ɾ���ݸ��>
    
    m_BID = rs("B_BID")
    m_Date = Format(Now, "YYYY-MM-DD HH:MM:SS")
    
    rs.Close
    Set rs = Nothing
    
    
    Dim m_BarCode As String   '��ȡһ���ƻ����е�һ���׵��Ŀ���
    Dim m_FHProcessName As String  '����������ʾ�Ĺ�����
    
    

    
    
    Set rs2 = New RecordSet
    rs2.Fields.Append "B_Field0", adInteger
    rs2.Open
    
    
    '��ⵥ�ݱ���ַ������
    '��ͬ�ļƻ����ֱ����
    '(�����ǰ����ѡ�ļ�¼�д���3����ͬ�ļƻ���-����G_CJBill�е�B_ID,��ô����3�ų�Ʒ��ⵥ)
    A_RKCID = ""
    A_ID_RK = ""
    For i = 0 To UBound(m_ArrayID)
        A_rs.Filter = " B_ID=" & m_ArrayID(i)
        If A_rs.RecordCount > 0 Then
        
        
        'д��ݸ�����
        Set rs = New RecordSet
        strSQL = "Select * From " & m_DraftBill & " Where 1=0"
        rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        
        '��ȡ�ƻ����е�һ���׵���
        'm_BarCode = GetJHDGDHM(m_ArrayID(i))
        '��ȡ��ID���ƻ�����G_CJBill��B_ID���µ�һ����ʹ�õĸ׵���B_ID
        '�ڱ�G_MidTableCPDRK��
        m_BarCode = ""
        Set rs3 = New RecordSet
        strSQL = "exec dbo.P_Get13052501 " & m_ArrayID(i)
        rs3.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        If rs3.RecordCount > 0 Then
            m_BarCode = rs3!B_BarCode
        End If
        rs3.Close
        Set rs3 = Nothing
        
        
        
        
        '��ȡ��ǰ�ƻ�������������
        Set rs1 = New RecordSet
        strSQL = "Select * From G_CJBill Where B_ID=" & m_ArrayID(i)
        rs1.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        
        
        '��ȡ������������
        m_FHProcessName = GetFHProcessName(m_ArrayID(i))
        
        m_CodeID = GetFrameCodeDetail(ObjectID)
        rs.addnew
        rs!B_Codeid = m_CodeID
        rs!B_Date = m_Date
        rs!B_BID = m_BID
        rs!B_ObjectID = ObjectID
        rs!B_ComputerName = Gm.SysID.ComputerName
        
        rs!B_BarCode = m_BarCode
        rs!B_ChangDaiHao = rs1!B_ChangDaiHao
        rs!B_DDH = rs1!B_DingDanhao
        rs!B_ClientName = rs1!B_Client
        rs!B_MenFuMiChang = rs1!B_MenFu
        rs!B_PinM = IIf(IsNull(rs1!B_PMQZ), "", rs1!B_PMQZ) & IIf(IsNull(rs1!B_PinMingKH), "", rs1!B_PinMingKH)
        rs!B_SeHao = rs1!B_SeHao
        rs!B_SeXing = rs1!B_color
        rs!B_FHDProcess = m_FHProcessName
        rs!B_EDP = A_rs!B_EDP
        rs!B_BelongJHD = m_ArrayID(i)
        

        
        
        rs.Update
        m_ID = rs("B_ID")
        A_ID_RK = A_ID_RK & m_ID & ","
        'Debug.Print "���ɳ�Ʒ��ⵥ��B_IDΪ��" & m_ID
        rs.Close
        Set rs = Nothing
        
        rs1.Close
        Set rs1 = Nothing
        
        
        'д���ݵ��ݸ���ϸ��
        strSQL = " (B_ID=" & m_ArrayID(i) & " And B_CheckID=1)"
        strSQL = strSQL & " Or (B_ID=" & m_ArrayID(i) & " And B_CheckID=-1)"
        
        A_rs.Filter = strSQL
        A_rs.MoveFirst
        Do While Not A_rs.EOF
            
            'ֻ�в��ǺϼƵ���   �Ž��б���
            If A_rs("B_HJ") = 0 Then
                'д��ݸ����ϸ���ݵ�ʱ��ͬʱд��������G_JRKBill�е�B_ItemID
                strSQL = "Insert Into " & m_DraftBillDetail
                strSQL = strSQL & " (B_BelongItemID,B_ID,B_EDP,B_GoodsID,B_Specification,B_Color,B_SeHao,B_Qty,B_KQty)"
                strSQL = strSQL & " Values"
                strSQL = strSQL & " ('" & A_rs!B_itemid & "','" & m_ID & "','" & A_rs!B_EDP & "','" & A_rs("B_PinM") & "','" & A_rs("B_MenFuMiChang") & "','" & A_rs("B_Color") & "','" & A_rs("B_SeHao") & "','" & A_rs("B_Qty") & "','" & A_rs("B_KQty") & "')"
                Gm.cnnTool.cnn.Execute strSQL
                
                
                
                'ͬʱ��¼��ǰ��ϸ�ڱ�G_JRKBill�е�B_ItemID
                rs2.addnew
                rs2.Fields(0) = A_rs!B_itemid
            End If
            
            A_rs.movenext
        Loop
        
        
        
        
        '�������ݵ���ʽ����
        Set rs = New RecordSet
        strSQL = "Select * from " & m_DraftBill & " Where B_ID='" & m_ID & "'"
        rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        
        Set rs1 = New RecordSet
        strSQL = "Select * from " & m_Bill & " Where 1=0"
        rs1.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        rs1.addnew
        rs1!B_Closed = 1
        rs1!B_id = rs!B_id
        rs1!B_Codeid = rs!B_Codeid
        rs1!B_Date = rs!B_Date
        rs1!B_BID = rs!B_BID
        rs1!B_ObjectID = rs!B_ObjectID
        rs1!B_BarCode = rs!B_BarCode
        rs1!B_ChangDaiHao = rs!B_ChangDaiHao
        rs1!B_DDH = rs!B_DDH
        rs1!B_ClientName = rs!B_ClientName
        rs1!B_MenFuMiChang = rs!B_MenFuMiChang
        rs1!B_PinM = rs!B_PinM
        rs1!B_SeHao = rs!B_SeHao
        rs1!B_SeXing = rs!B_SeXing
        rs1!B_FHDProcess = rs!B_FHDProcess
        rs1!B_EDP = rs!B_EDP
        rs1!B_BelongJHD = rs!B_BelongJHD
        rs1.Update
        
        
        A_RKCID = A_RKCID & rs!B_Codeid & ","
        rs1.Close
        Set rs1 = Nothing
        
        rs.Close
        Set rs = Nothing
        
        
        
        '�������ݵ���ϸ����
        Set rs = New RecordSet
        strSQL = "Select * from " & m_DraftBillDetail & " Where B_ID='" & m_ID & "'"
        rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        Do While Not rs.EOF
            strSQL = "Insert Into " & m_BillDetail
            strSQL = strSQL & " (B_BelongItemID,B_ItemID,B_ID,B_EDP,B_GoodsID,B_Specification,B_Color,B_SeHao,B_Qty,B_KQty)"
            strSQL = strSQL & " Values"
            strSQL = strSQL & " ('" & rs!B_BelongItemID & "','" & rs("B_ItemID") & "','" & rs("B_ID") & "','" & rs!B_EDP & "','" & rs("B_GoodsID") & "','" & rs("B_Specification") & "','" & rs("B_Color") & "','" & rs("B_SeHao") & "','" & rs("B_Qty") & "','" & rs("B_KQty") & "')"
            Gm.cnnTool.cnn.Execute strSQL
            rs.movenext
        Loop
        rs.Close
        Set rs = Nothing
        
        
        'ɾ���ݸ�����
        strSQL = "Delete From " & m_DraftBillDetail & " Where B_ID='" & m_ID & "'"
        Gm.cnnTool.cnn.Execute strSQL
        
        
        strSQL = "Delete From " & m_DraftBill & " Where B_ID='" & m_ID & "'"
        Gm.cnnTool.cnn.Execute strSQL
                
        End If
    Next
    
    
    '��ʾ���ɵ���ⵥ�������Լ��ֱ�ĵ��ݱ��
    '=========================================
    A_RKCID = Left$(A_RKCID, Len(A_RKCID) - 1)
    A_ID_RK = Left$(A_ID_RK, Len(A_ID_RK) - 1)
    Dim m_StrTip As String
    Dim m_Array
    
    m_Array = Split(A_RKCID, ",")
    
    
    m_StrTip = "������" & UBound(m_ArrayID) + 1 & "����ⵥ!" & vbNewLine
    m_StrTip = m_StrTip & "���ݱ��Ϊ��" & vbNewLine
    
    For i = 0 To UBound(m_Array)
        m_StrTip = m_StrTip & m_Array(i) & vbNewLine
    Next
    
    'MsgBox m_StrTip, vbOKOnly + vbInformation, "��ʾ"
    '=========================================
End Sub

'���ֶ�vCheckFieldName�ϱ���ѡ�ļ�¼�Ż�ȡ��������ֶ��ϵ�ֵ�������ж��Ѿ������ظ��Ļ������л�ȡ
'vRs:��ַ�����ļ�¼��
'vFieldName:һ��Ϊ�����ֶΣ����ֶ��ϲ�Ϊ�ղŽ��в���
'vCheckFieldName:��ѡ�ֶε��ֶ�����һ��ΪB_CheckID
Private Function GetDistinctValueFromRSEx(ByRef vRs As RecordSet, ByVal vFieldName As String, ByVal vCheckFieldName As String) As String
    Dim rs As New RecordSet
    Dim rs1 As New RecordSet
    
    
    GetDistinctValueFromRSEx = ""
    
    
    Set rs = vRs.Clone
    rs1.Fields.Append "B_Field0", adVarChar, 100
    rs1.Open
    
    If rs.State <> adStateOpen Then
        Exit Function
    End If
    
    If rs.RecordCount <= 0 Then
        Exit Function
    End If
    
    
    rs.MoveFirst
    Do While Not rs.EOF
        If Len(IIf(IsNull(rs(vFieldName)), "", rs(vFieldName))) > 0 Then
            If Abs(IIf(IsNull(rs(vCheckFieldName)), 0, rs(vCheckFieldName))) = 1 Then
                rs1.Filter = " B_Field0='" & IIf(IsNull(rs(vFieldName)), "", rs(vFieldName)) & "'"
                If rs1.RecordCount <= 0 Then
                    rs1.addnew
                    rs1(0).Value = IIf(IsNull(rs(vFieldName)), "", rs(vFieldName))
                    rs1.Update
                End If
            End If
        End If
        rs.movenext
    Loop
    
    rs1.Filter = ""
    If rs1.RecordCount > 0 Then
        rs1.MoveFirst
        GetDistinctValueFromRSEx = ""
        Do While Not rs1.EOF
            GetDistinctValueFromRSEx = GetDistinctValueFromRSEx & rs1(0).Value & ","
            rs1.movenext
        Loop
        
        GetDistinctValueFromRSEx = Trim(Left(GetDistinctValueFromRSEx, Len(GetDistinctValueFromRSEx) - 1))
    End If
End Function


'�������ͣ�
'vCDH���ƻ�����G_CJBill��B_ID
'���ظüƻ����ڷ�������Ӧ����ʾ�Ĺ���
Private Function GetFHProcessName(ByVal vID As Long) As String
    Dim rs As RecordSet
    Dim m_FHDProcessNameString As String
    Dim m_CDH As String
    
    Set rs = New RecordSet
    strSQL = "Select * From G_CJBill Where B_ID=" & vID
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        m_CDH = rs!B_ChangDaiHao
    Else
        MsgBox "��ǰ�ƻ��������ڣ�", vbOKOnly + vbInformation, "��ʾ"
        Exit Function
    End If
    rs.Close
    Set rs = Nothing
    
    
    strSQL = "exec dbo.S_GetFHDProcessNameString '" & m_CDH & "'"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    m_FHDProcessNameString = ""
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            m_FHDProcessNameString = m_FHDProcessNameString & rs("B_ProcessName") & ","
            rs.movenext
        Loop
        
        m_FHDProcessNameString = Left(m_FHDProcessNameString, Len(m_FHDProcessNameString) - 1)
    End If
    rs.Close
    Set rs = Nothing
    
    GetFHProcessName = m_FHDProcessNameString
End Function

'��ʼ���Ƶ�
Private Sub Init_Combo_ZD()
    g_FunctTool.BindCombo Combo1, clsBingComboZD, "ɫ������������"
End Sub

Private Function JudgeRsLawless(ByRef vRs As RecordSet) As Boolean
    If vRs.State <> adStateOpen Then
        JudgeRsLawless = False
        Exit Function
    End If
    
    If vRs.RecordCount <= 0 Then
        JudgeRsLawless = False
        Exit Function
    End If
    
    JudgeRsLawless = True
End Function

'����ѡ�е������еĿ��ţ��������������ƻ�����
Private Sub ChangeKH()
    '��2017��12��28��ʵʩʱע�͵�
'    If TDBGrid1.SelBookmarks.Count <= 0 Then
'        MsgBox "û��ѡ���κ��У�", vbOKOnly + vbInformation, "��ʾ"
'        Exit Sub
'    End If
'
'
'
'    '�ж�Ŀ�Ŀ����Ƿ����
'    '==================================
'    Dim rs As RecordSet
'    Dim szKH As String
'    Dim lIDKH As Long   '�����ڱ�G_CJFlowBill�е�B_ID
'    szKH = Trim$(Text3.Text)
'    strSQL = "Select * From G_CJFlowBill where B_BarCode13='" & GetBarCode13(szKH) & "'"
'    Set rs = New RecordSet
'    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
'
'    If rs.RecordCount <= 0 Then
'        MsgBox "��ǰĿ�Ŀ��Ų����ڣ�", vbOKOnly + vbInformation, "��ʾ"
'        rs.Close
'        Set rs = Nothing
'        Exit Sub
'    End If
'
'    lIDKH = rs!B_ID
'    rs.Close
'    Set rs = Nothing
'    '==================================
'
'
'    '��ʾ
'    If MsgBox("��ȷ��Ҫ����ǰѡ�е������еĿ����޸�Ϊ��" & szKH & "ô��", vbYesNo + vbExclamation + vbDefaultButton2, "����") = vbNo Then
'        Exit Sub
'    End If
'
'
'    Dim tdbgRow As Variant
'    Dim lItemID As Long
'
'    For Each tdbgRow In TDBGrid1.SelBookmarks
'        A_rs.Bookmark = tdbgRow
'        lItemID = A_rs!B_itemid
'        strSQL = "Update G_JRKBill set B_BCFC='" & szKH & "',B_ID=" & lIDKH & " where B_ItemID=" & lItemID
'        Gm.cnnTool.cnn.Execute strSQL
'    Next
'
'    MsgBox "�޸���ϣ�", vbOKOnly + vbInformation, "��ʾ"
'
'    'ˢ������
'    GetRs
End Sub

Public Sub LoadObject()
    
End Sub

'ɾ����ǰ�����µ���������
Private Sub DelKHAll()
    If IsProExists = False Then
        Exit Sub
    End If
    
    If MsgBox("��ȷ��Ҫɾ����ǰѡ���п����µ���������ô��", vbYesNo + vbExclamation + vbDefaultButton2, "����") = vbNo Then
        Exit Sub
    End If
    
    
    On Error Resume Next
    Dim szKH As String
    szKH = A_rs!B_KH
    'strSQL = "Delete From G_JRKBill where B_BCFC='" & szKH & "'"
    strSQL = "exec dbo.usp_JRK_DelBCFC '" & szKH & "'"
    Gm.cnnTool.cnn.Execute strSQL
    
    Dim lBookmark As Long
    lBookmark = TDBGrid1.bookmark
    GetRS
    TDBGrid1.bookmark = lBookmark
End Sub

Private Sub EditeQty()
    '��2017��12��28��ʵʩʱע�͵�
'    Dim dKG As Double
'    Dim dMeters As Double
'
'    dKG = IIf(IsNull(A_rs!B_Qty), 0, A_rs!B_Qty)
'    dMeters = IIf(IsNull(A_rs!B_KQty), 0, A_rs!B_KQty)
'
'    Dim frm1 As New frmStorageRKCKEdit01
'    Dim lItemID As Long
'    lItemID = A_rs!B_itemid
'    With frm1
'        .KG = dKG
'        .Meters = dMeters
'        .Show vbModal
'    End With
'
'    If frm1.Saved = False Then
'        Unload frm1
'        Exit Sub
'    End If
'
'    dKG = frm1.KG
'    dMeters = frm1.Meters
'
'    Unload frm1
'    strSQL = "Update G_JRKBill Set B_GJ=" & dKG & " ,B_MS='" & dMeters & "' Where B_ItemID=" & lItemID
'    Gm.cnnTool.cnn.Execute strSQL
'
'    Dim lBookmark As Long
'    lBookmark = TDBGrid1.Bookmark
'    GetRs
'    TDBGrid1.Bookmark = lBookmark
End Sub

'��2017��10��25�� 18:53:54�滻Ϊ����ķ���
'Private Sub EditeQty()
'    If MsgBox("��ȷ��Ҫ�޸ĵ�ǰѡ���е�����ô��", vbYesNo + vbExclamation + vbDefaultButton2, "����") = vbNo Then
'        Exit Sub
'    End If
'
'    Dim dQty As Double
'    Dim litemID As Long
'    Dim QtyOld As Double
'
'    litemID = A_rs!B_ItemID
'    QtyOld = IIf(IsNull(A_rs!B_Qty), 0, A_rs!B_Qty)
'
'    dQty = Val(InputBox("¼��������", "�޸�����", 0))
'
'
'    strSQL = "Update G_JRKBill Set B_GJ=" & dQty & " Where B_ItemID=" & litemID
'    gm.cnnTool.cnn.Execute strSQL
'
'
'    Dim lBookmark As Long
'    lBookmark = TDBGrid1.Bookmark
'    GetRs
'    TDBGrid1.Bookmark = lBookmark
'End Sub



Private Sub InitCPH()
    On Error Resume Next
    Combo2.Clear
    
    strSQL = g_FunctTool.GetSQL("ɫ���������ƺ�")
    
    Dim rs As New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    Do While Not rs.EOF
        Combo2.AddItem IIf(IsNull(rs!B_PlateNumber), "", rs!B_PlateNumber)
        rs.movenext
    Loop

    rs.Close
    Set rs = Nothing
End Sub


Private Sub InitSHR()
    On Error Resume Next
    Combo3.Clear
    
'    Dim strSQL As String
'    Dim rs As RecordSet
'    Set rs = New RecordSet
'    strSQL = "Select B_SHR From G_SHR Order by B_SHR"
'    rs.Open strSQL, gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
'
'    Do While Not rs.EOF
'        Combo3.AddItem rs!B_SHR
'        rs.movenext
'    Loop
'
'    rs.Close
'    Set rs = Nothing

    Set BinderLoader = New cls_Link_Data_Ctl
    g_FunctTool.BindCombo Combo3, BinderLoader, "ɫ������װж��"
End Sub



'���ݶѺŻ�ȡ��Ʒ����ģʽ���жϿͻ������Ƿ�Ψһ
Private Function JudgeClientUniqueBlock() As Boolean
    JudgeClientUniqueBlock = True
    
    If A_rsKH.State <> adStateOpen Then
        Exit Function
    End If
    
    If A_rsKH.RecordCount <= 0 Then
        Exit Function
    End If
    
    Dim szKH As String
    Dim rs As RecordSet
    szKH = RecordSetToString(A_rsKH, "B_KH", ",")
    
    
    Set rs = New RecordSet
    strSQL = "exec dbo.[P_GetClientNameByBlock] '" & szKH & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount > 1 Then
        JudgeClientUniqueBlock = False
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    If rs.RecordCount > 0 Then
        A_ClientName = rs!B_Client
    End If
    
    rs.Close
    Set rs = Nothing
End Function


Private Function JudgeClientUnique() As Boolean
    JudgeClientUnique = JudgeClientUniqueKH
End Function


Private Sub InitClients()
    With UCListBox1
        .ConnectionString = Gm.cnnTool.cnnStr
        .sql = "SELECT B_ClientID, B_ClientName FROM G_ContactCompany WHERE 1=1 AND B_ContactType='�ͻ�'"
        .Refresh
    End With
End Sub

Private Sub InitDDH()
    Dim strSQL As String
    Dim szClientID As String
    szClientID = UCListBox1.Text
    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT B_DingDanHao"
    strSQL = strSQL & " From G_CJBill"
    strSQL = strSQL & " WHERE B_Client ="
    strSQL = strSQL & " ("
    strSQL = strSQL & "     SELECT G_ContactCompany.B_ClientName"
    strSQL = strSQL & "     FROM G_ContactCompany WHERE B_ClientID='" & szClientID & "'"
    strSQL = strSQL & " )"

    Debug.Print strSQL
    With UCListBox2
        .ConnectionString = Gm.cnnTool.cnnStr
        .sql = strSQL
        .Refresh
    End With
End Sub

Private Sub UCListBox1_Change()
    InitDDH
End Sub

'����FALSE��ʾ�����ڣ�����Ϊ����
Private Function IsProExists() As Boolean
    IsProExists = True
    '��2017��12��28��ע�͵�
'    Dim cls1 As New clsDataBase
'    IsProExists = cls1.JudgeDBObjExists("usp_JRK_DelOne")
'
'    Dim szErr As String
'    szErr = "�ͻ��˰汾ȱ�ݣ�����ϵ������ṩ���¸��°���" & vbNewLine
'    szErr = szErr & "2017��7��15�� - ɾ����Ʒ�����Լ���"
'
'
'    If IsProExists = False Then
'        MsgBox szErr, vbOKOnly + vbInformation, "��ʾ"
'        Exit Function
'    End If
'
'    IsProExists = cls1.JudgeDBObjExists("usp_JRK_DelBCFC")
'
'    If IsProExists = False Then
'        MsgBox szErr, vbOKOnly + vbInformation, "��ʾ"
'    End If
End Function

'���÷�����ť����ֹ����
Private Sub ForbidBT_FH(ByVal vEnabled As Boolean)
    ccButton1(2).Enabled = vEnabled
End Sub

