VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmOrderProduct 
   Caption         =   "成品合同"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17940
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
   ScaleHeight     =   10950
   ScaleWidth      =   17940
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   10950
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17940
      _LayoutVersion  =   1
      _ExtentX        =   31644
      _ExtentY        =   19315
      _DataPath       =   ""
      Bands           =   "frmOrderProduct.frx":0000
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   10815
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   16095
         _cx             =   28390
         _cy             =   19076
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
         _GridInfo       =   $"frmOrderProduct.frx":01C8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   3060
            Left            =   30
            ScaleHeight     =   3060
            ScaleWidth      =   16035
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   465
            Width           =   16035
            Begin XtremeSuiteControls.PushButton PushButton6 
               Height          =   555
               Left            =   9360
               TabIndex        =   21
               Top             =   1950
               Width           =   1635
               _Version        =   1048578
               _ExtentX        =   2884
               _ExtentY        =   979
               _StockProps     =   79
               Caption         =   "设置备注"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox CheckBox3 
               Height          =   255
               Left            =   14940
               TabIndex        =   22
               Top             =   2760
               Width           =   255
               _Version        =   1048578
               _ExtentX        =   450
               _ExtentY        =   450
               _StockProps     =   79
               Enabled         =   0   'False
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox CheckBox2 
               Height          =   315
               Left            =   13860
               TabIndex        =   23
               Top             =   2730
               Width           =   375
               _Version        =   1048578
               _ExtentX        =   661
               _ExtentY        =   556
               _StockProps     =   79
               Enabled         =   0   'False
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox CheckBox1 
               Height          =   195
               Left            =   12720
               TabIndex        =   24
               Top             =   2790
               Width           =   315
               _Version        =   1048578
               _ExtentX        =   556
               _ExtentY        =   344
               _StockProps     =   79
               Enabled         =   0   'False
               UseVisualStyle  =   -1  'True
            End
            Begin MSAdodcLib.Adodc Adodc1 
               Height          =   330
               Left            =   3720
               Top             =   480
               Visible         =   0   'False
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   582
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   8
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   "Adodc1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   14760
               TabIndex        =   25
               Top             =   180
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
               Left            =   2100
               TabIndex        =   26
               Top             =   180
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               Enabled         =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit2 
               Height          =   375
               Left            =   13200
               TabIndex        =   27
               Top             =   180
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               Enabled         =   0   'False
               BackColor       =   14737632
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit3 
               Height          =   375
               Left            =   2100
               TabIndex        =   28
               Top             =   800
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               Enabled         =   0   'False
               BackColor       =   14737632
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit4 
               Height          =   375
               Left            =   9420
               TabIndex        =   29
               Top             =   800
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit5 
               Height          =   375
               Left            =   13200
               TabIndex        =   30
               Top             =   800
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               Enabled         =   0   'False
               BackColor       =   14737632
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit6 
               Height          =   375
               Left            =   2100
               TabIndex        =   31
               Top             =   1420
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               Enabled         =   0   'False
               BackColor       =   14737632
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit7 
               Height          =   375
               Left            =   2100
               TabIndex        =   32
               Top             =   2040
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit8 
               Height          =   375
               Left            =   9420
               TabIndex        =   33
               Top             =   1420
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               Enabled         =   0   'False
            End
            Begin XtremeSuiteControls.ComboBox ComboBox1 
               Height          =   315
               Left            =   5820
               TabIndex        =   34
               Top             =   830
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
            End
            Begin XtremeSuiteControls.ComboBox ComboBox2 
               Height          =   315
               Left            =   5820
               TabIndex        =   35
               Top             =   1450
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
            End
            Begin XtremeSuiteControls.ComboBox ComboBox3 
               Height          =   315
               Left            =   13380
               TabIndex        =   36
               Top             =   2520
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
            End
            Begin XtremeSuiteControls.ComboBox ComboBox4 
               Height          =   315
               Left            =   13200
               TabIndex        =   37
               Top             =   1450
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
            End
            Begin XtremeSuiteControls.ComboBox ComboBox5 
               Height          =   315
               Left            =   5820
               TabIndex        =   38
               Top             =   2070
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
            End
            Begin XtremeSuiteControls.DateTimePicker DateTimePicker1 
               Height          =   375
               Left            =   5820
               TabIndex        =   39
               Top             =   180
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   68
               Format          =   1
            End
            Begin XtremeSuiteControls.DateTimePicker DateTimePicker2 
               Height          =   375
               Left            =   9420
               TabIndex        =   40
               Top             =   180
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   68
               Format          =   1
            End
            Begin XtremeSuiteControls.ComboBox ComboBox6 
               Height          =   315
               Left            =   13200
               TabIndex        =   41
               Top             =   2070
               Visible         =   0   'False
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit9 
               Height          =   375
               Left            =   2100
               TabIndex        =   42
               Top             =   2640
               Visible         =   0   'False
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.Label Label15 
               Height          =   375
               Left            =   8280
               TabIndex        =   63
               Top             =   1425
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "制单人："
            End
            Begin XtremeSuiteControls.Label Label14 
               Height          =   255
               Left            =   4800
               TabIndex        =   62
               Top             =   2100
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "业务跟单："
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label13 
               Height          =   255
               Left            =   1020
               TabIndex        =   61
               Top             =   2100
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "合同签订地："
            End
            Begin XtremeSuiteControls.Label Label12 
               Height          =   255
               Left            =   12120
               TabIndex        =   60
               Top             =   1480
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "结算方式："
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label11 
               Height          =   270
               Left            =   12420
               TabIndex        =   59
               Top             =   2520
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   476
               _StockProps     =   79
               Caption         =   "计价单位："
            End
            Begin XtremeSuiteControls.Label Label10 
               Height          =   255
               Left            =   4800
               TabIndex        =   58
               Top             =   1480
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "合同开票："
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label9 
               Height          =   255
               Left            =   1020
               TabIndex        =   57
               Top             =   1480
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "合同金额："
            End
            Begin XtremeSuiteControls.Label Label8 
               Height          =   255
               Left            =   12120
               TabIndex        =   56
               Top             =   860
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "合同订单数："
            End
            Begin XtremeSuiteControls.Label Label7 
               Height          =   255
               Left            =   8280
               TabIndex        =   55
               Top             =   855
               Width           =   1095
               _Version        =   1048578
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "合同号："
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label6 
               Height          =   255
               Left            =   4800
               TabIndex        =   54
               Top             =   860
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "交货方式："
               ForeColor       =   255
            End
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Left            =   1020
               TabIndex        =   53
               Top             =   860
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "客户简称："
            End
            Begin XtremeSuiteControls.Label Label4 
               Height          =   255
               Left            =   12120
               TabIndex        =   52
               Top             =   240
               Width           =   495
               _Version        =   1048578
               _ExtentX        =   873
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "客户："
               ForeColor       =   255
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "交货期："
               Height          =   195
               Left            =   8280
               TabIndex        =   51
               Top             =   270
               Width           =   720
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   255
               Left            =   4800
               TabIndex        =   50
               Top             =   240
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "单据日期："
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   255
               Left            =   1020
               TabIndex        =   49
               Top             =   240
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "单据编号："
               ForeColor       =   0
            End
            Begin XtremeSuiteControls.Label Label16 
               Height          =   315
               Left            =   12420
               TabIndex        =   48
               Top             =   2490
               Width           =   735
               _Version        =   1048578
               _ExtentX        =   1296
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "货清:"
            End
            Begin XtremeSuiteControls.Label Label17 
               Height          =   195
               Left            =   13560
               TabIndex        =   47
               Top             =   2550
               Width           =   375
               _Version        =   1048578
               _ExtentX        =   661
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "票清:"
            End
            Begin XtremeSuiteControls.Label Label18 
               Height          =   255
               Left            =   14640
               TabIndex        =   46
               Top             =   2520
               Width           =   435
               _Version        =   1048578
               _ExtentX        =   767
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "款清:"
            End
            Begin XtremeSuiteControls.Label Label21 
               Height          =   255
               Left            =   8280
               TabIndex        =   45
               Top             =   2100
               Width           =   495
               _Version        =   1048578
               _ExtentX        =   873
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "备注:"
            End
            Begin XtremeSuiteControls.Label Label22 
               Height          =   255
               Left            =   12120
               TabIndex        =   44
               Top             =   2100
               Visible         =   0   'False
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "包装方式:"
            End
            Begin XtremeSuiteControls.Label Label23 
               Height          =   375
               Left            =   1020
               TabIndex        =   43
               Top             =   2640
               Visible         =   0   'False
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "款            号："
            End
         End
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   7230
            Left            =   30
            TabIndex        =   3
            Top             =   3555
            Width           =   16035
            _cx             =   28284
            _cy             =   12753
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
            Caption         =   "合同及色布计划|色布计划|白坯计划|订单图样|白坯构成|合同扫描件|色布采购|打箱计划|辅料计划|裁剪计划|缝制计划|包装计划"
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
            Picture(0)      =   "frmOrderProduct.frx":024D
            Picture(1)      =   "frmOrderProduct.frx":05E7
            Picture(2)      =   "frmOrderProduct.frx":0981
            Picture(3)      =   "frmOrderProduct.frx":0F1B
            Picture(4)      =   "frmOrderProduct.frx":12B5
            Picture(5)      =   "frmOrderProduct.frx":164F
            Picture(6)      =   "frmOrderProduct.frx":19E9
            Picture(7)      =   "frmOrderProduct.frx":1F83
            Picture(8)      =   "frmOrderProduct.frx":231D
            Picture(9)      =   "frmOrderProduct.frx":26B7
            Picture(10)     =   "frmOrderProduct.frx":2811
            Picture(11)     =   "frmOrderProduct.frx":2BAB
            Begin C1SizerLibCtl.C1Elastic C1Elastic14 
               Height          =   7200
               Left            =   20655
               TabIndex        =   89
               TabStop         =   0   'False
               Top             =   15
               Width           =   15000
               _cx             =   26458
               _cy             =   12700
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
               GridRows        =   5
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmOrderProduct.frx":3145
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.TextBox Text1 
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   825
                  Left            =   1710
                  TabIndex        =   97
                  Top             =   6285
                  Width           =   13200
               End
               Begin VB.PictureBox Picture7 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   645
                  Left            =   90
                  ScaleHeight     =   615
                  ScaleWidth      =   14790
                  TabIndex        =   91
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   14820
                  Begin XtremeSuiteControls.PushButton PushButton8 
                     Height          =   435
                     Left            =   5760
                     TabIndex        =   92
                     Top             =   60
                     Width           =   1755
                     _Version        =   1048578
                     _ExtentX        =   3096
                     _ExtentY        =   767
                     _StockProps     =   79
                     Caption         =   "预览图片"
                     UseVisualStyle  =   -1  'True
                  End
                  Begin MSComDlg.CommonDialog CommonDialog3 
                     Left            =   0
                     Top             =   0
                     _ExtentX        =   847
                     _ExtentY        =   847
                     _Version        =   393216
                  End
                  Begin XtremeSuiteControls.PushButton PushButton9 
                     Height          =   435
                     Left            =   3420
                     TabIndex        =   93
                     Top             =   60
                     Width           =   1755
                     _Version        =   1048578
                     _ExtentX        =   3096
                     _ExtentY        =   767
                     _StockProps     =   79
                     Caption         =   "上传图片"
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.PushButton PushButton10 
                     Height          =   435
                     Left            =   1080
                     TabIndex        =   94
                     Top             =   60
                     Width           =   1755
                     _Version        =   1048578
                     _ExtentX        =   3096
                     _ExtentY        =   767
                     _StockProps     =   79
                     Caption         =   "删除图片"
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.Label Label24 
                     Height          =   315
                     Left            =   7620
                     TabIndex        =   95
                     Top             =   60
                     Width           =   7035
                     _Version        =   1048578
                     _ExtentX        =   12409
                     _ExtentY        =   556
                     _StockProps     =   79
                     Caption         =   "请提供JPG或者BMP格式图片，在切换页面之前,请先上传图片,会失去图片"
                     ForeColor       =   8421631
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
               End
               Begin VB.PictureBox Picture6 
                  Height          =   5430
                  Left            =   90
                  ScaleHeight     =   5370
                  ScaleWidth      =   14760
                  TabIndex        =   90
                  TabStop         =   0   'False
                  Top             =   795
                  Width           =   14820
               End
               Begin VB.Label Label25 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "备注："
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   27.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   825
                  Left            =   90
                  TabIndex        =   96
                  Top             =   6285
                  Width           =   1560
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic11 
               Height          =   7200
               Left            =   19755
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   15
               Width           =   15000
               _cx             =   26458
               _cy             =   12700
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
               _GridInfo       =   $"frmOrderProduct.frx":31D2
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin TrueOleDBGrid80.TDBGrid TDBGrid8 
                  Height          =   6555
                  Left            =   30
                  TabIndex        =   79
                  Top             =   615
                  Width           =   14940
                  _ExtentX        =   26353
                  _ExtentY        =   11562
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
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=4339"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
                  Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
                  Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(6)=   "Column(1).Width=4339"
                  Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
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
               Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar27 
                  Height          =   555
                  Left            =   30
                  TabIndex        =   80
                  Top             =   30
                  Width           =   14940
                  _LayoutVersion  =   1
                  _ExtentX        =   26353
                  _ExtentY        =   979
                  _DataPath       =   ""
                  Bands           =   "frmOrderProduct.frx":3256
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic8 
               Height          =   7200
               Left            =   19455
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   15
               Width           =   15000
               _cx             =   26458
               _cy             =   12700
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
               _GridInfo       =   $"frmOrderProduct.frx":49F6
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar26 
                  Height          =   555
                  Left            =   30
                  TabIndex        =   76
                  Top             =   30
                  Width           =   14940
                  _LayoutVersion  =   1
                  _ExtentX        =   26353
                  _ExtentY        =   979
                  _DataPath       =   ""
                  Bands           =   "frmOrderProduct.frx":4A7A
               End
               Begin TrueOleDBGrid80.TDBGrid TDBGrid7 
                  Height          =   6555
                  Left            =   30
                  TabIndex        =   77
                  Top             =   615
                  Width           =   14940
                  _ExtentX        =   26353
                  _ExtentY        =   11562
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
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=4339"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
                  Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
                  Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(6)=   "Column(1).Width=4339"
                  Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   7200
               Left            =   1020
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   15
               Width           =   15000
               _cx             =   26458
               _cy             =   12700
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
               _GridInfo       =   $"frmOrderProduct.frx":6758
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
                  Height          =   6660
                  Left            =   30
                  TabIndex        =   87
                  Top             =   510
                  Width           =   14940
                  _ExtentX        =   26353
                  _ExtentY        =   11748
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
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=4339"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
                  Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
                  Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(6)=   "Column(1).Width=4339"
                  Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
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
                  MultiSelect     =   2
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
               Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar210 
                  Height          =   450
                  Left            =   30
                  TabIndex        =   88
                  Top             =   30
                  Width           =   14940
                  _LayoutVersion  =   1
                  _ExtentX        =   26353
                  _ExtentY        =   794
                  _DataPath       =   ""
                  Bands           =   "frmOrderProduct.frx":67DC
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic9 
               Height          =   7200
               Left            =   17655
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   15
               Width           =   15000
               _cx             =   26458
               _cy             =   12700
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
               _GridInfo       =   $"frmOrderProduct.frx":746C
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin TrueOleDBGrid80.TDBGrid TDBGrid3 
                  Height          =   6675
                  Left            =   30
                  TabIndex        =   73
                  Top             =   495
                  Width           =   14940
                  _ExtentX        =   26353
                  _ExtentY        =   11774
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
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=4339"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
                  Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
                  Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(6)=   "Column(1).Width=4339"
                  Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
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
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
                  _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(12)  =   ":id=2,.fontname=宋体"
                  _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
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
                  _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.alignment=2,.valignment=2"
                  _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                  _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgpicMode=2,.bgbmp=2"
                  _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3,.bgpicMode=2,.bgbmp=1"
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
               Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar24 
                  Height          =   465
                  Left            =   30
                  TabIndex        =   74
                  Top             =   30
                  Width           =   14940
                  _LayoutVersion  =   1
                  _ExtentX        =   26353
                  _ExtentY        =   820
                  _DataPath       =   ""
                  Bands           =   "frmOrderProduct.frx":7505
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   7200
               Left            =   17955
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   15
               Width           =   15000
               _cx             =   26458
               _cy             =   12700
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
               _GridInfo       =   $"frmOrderProduct.frx":B8B1
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic C1Elastic4 
                  Height          =   7140
                  Left            =   30
                  TabIndex        =   7
                  TabStop         =   0   'False
                  Top             =   30
                  Width           =   14940
                  _cx             =   26353
                  _cy             =   12594
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
                  _GridInfo       =   $"frmOrderProduct.frx":B94A
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar23 
                     Height          =   435
                     Left            =   30
                     TabIndex        =   65
                     Top             =   30
                     Width           =   1110
                     _LayoutVersion  =   1
                     _ExtentX        =   1958
                     _ExtentY        =   767
                     _DataPath       =   ""
                     Bands           =   "frmOrderProduct.frx":B9E2
                  End
                  Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
                     Height          =   6675
                     Left            =   30
                     TabIndex        =   66
                     Top             =   465
                     Width           =   1110
                     _ExtentX        =   1958
                     _ExtentY        =   11774
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
                     Splits(0)._ColumnProps(1)=   "Column(0).Width=4339"
                     Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                     Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
                     Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
                     Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                     Splits(0)._ColumnProps(6)=   "Column(1).Width=4339"
                     Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
                     Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
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
                     _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
                     _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=134"
                     _StyleDefs(12)  =   ":id=2,.fontname=宋体"
                     _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
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
                     _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.alignment=2,.valignment=2"
                     _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                     _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgpicMode=2,.bgbmp=1"
                     _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                     _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                     _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
                     _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
                     _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgpicMode=2"
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
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   7200
               Left            =   18255
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   15
               Width           =   15000
               _cx             =   26458
               _cy             =   12700
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
               _GridInfo       =   $"frmOrderProduct.frx":E0D2
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic C1Elastic10 
                  Height          =   7140
                  Left            =   30
                  TabIndex        =   9
                  TabStop         =   0   'False
                  Top             =   30
                  Width           =   14940
                  _cx             =   26353
                  _cy             =   12594
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
                  _GridInfo       =   $"frmOrderProduct.frx":E158
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin VB.PictureBox Picture2 
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     ForeColor       =   &H80000008&
                     Height          =   690
                     Left            =   30
                     ScaleHeight     =   660
                     ScaleWidth      =   1080
                     TabIndex        =   68
                     TabStop         =   0   'False
                     Top             =   30
                     Width           =   1110
                     Begin XtremeSuiteControls.PushButton PushButton2 
                        Height          =   435
                        Left            =   5760
                        TabIndex        =   69
                        Top             =   60
                        Width           =   1755
                        _Version        =   1048578
                        _ExtentX        =   3096
                        _ExtentY        =   767
                        _StockProps     =   79
                        Caption         =   "预览图片"
                        UseVisualStyle  =   -1  'True
                     End
                     Begin MSComDlg.CommonDialog CommonDialog1 
                        Left            =   0
                        Top             =   0
                        _ExtentX        =   847
                        _ExtentY        =   847
                        _Version        =   393216
                     End
                     Begin XtremeSuiteControls.PushButton PushButton4 
                        Height          =   435
                        Left            =   3420
                        TabIndex        =   70
                        Top             =   60
                        Width           =   1755
                        _Version        =   1048578
                        _ExtentX        =   3096
                        _ExtentY        =   767
                        _StockProps     =   79
                        Caption         =   "上传图片"
                        UseVisualStyle  =   -1  'True
                     End
                     Begin XtremeSuiteControls.PushButton PushButton7 
                        Height          =   435
                        Left            =   1080
                        TabIndex        =   71
                        Top             =   60
                        Width           =   1755
                        _Version        =   1048578
                        _ExtentX        =   3096
                        _ExtentY        =   767
                        _StockProps     =   79
                        Caption         =   "删除图片"
                        UseVisualStyle  =   -1  'True
                     End
                     Begin XtremeSuiteControls.Label Label20 
                        Height          =   195
                        Left            =   7620
                        TabIndex        =   72
                        Top             =   180
                        Width           =   3435
                        _Version        =   1048578
                        _ExtentX        =   6059
                        _ExtentY        =   344
                        _StockProps     =   79
                        Caption         =   "在切换页面之前,请先上传图片,会失去图片"
                        ForeColor       =   8421631
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
                  End
                  Begin VB.PictureBox Picture3 
                     Height          =   6420
                     Left            =   30
                     ScaleHeight     =   6360
                     ScaleWidth      =   1050
                     TabIndex        =   67
                     TabStop         =   0   'False
                     Top             =   720
                     Width           =   1110
                  End
                  Begin TrueOleDBGrid80.TDBGrid TDBGrid4 
                     Height          =   7110
                     Left            =   30
                     TabIndex        =   0
                     Top             =   30
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   12541
                     _LayoutType     =   0
                     _RowHeight      =   18
                     _WasPersistedAsPixels=   0
                     Columns(0)._VlistStyle=   0
                     Columns(0)._MaxComboItems=   5
                     Columns(0).DataField=   ""
                     Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                     Columns.Count   =   1
                     Splits(0)._UserFlags=   0
                     Splits(0).RecordSelectors=   0   'False
                     Splits(0).RecordSelectorWidth=   953
                     Splits(0)._SavedRecordSelectors=   0   'False
                     Splits(0).ScrollBars=   2
                     Splits(0).AllowColSelect=   0   'False
                     Splits(0).DividerColor=   15790320
                     Splits(0).SpringMode=   0   'False
                     Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                     Splits(0)._ColumnProps(0)=   "Columns.Count=1"
                     Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
                     Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                     Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                     Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
                     Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                     Splits.Count    =   1
                     PrintInfos(0)._StateFlags=   3
                     PrintInfos(0).Name=   "piInternal 0"
                     PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
                     PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
                     PrintInfos(0).PageHeaderHeight=   0
                     PrintInfos(0).PageFooterHeight=   0
                     PrintInfos.Count=   1
                     AllowUpdate     =   0   'False
                     DefColWidth     =   0
                     HeadLines       =   1.5
                     FootLines       =   2
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
                     _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=114,.bold=0,.fontsize=825,.italic=0"
                     _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
                     _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
                     _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=134"
                     _StyleDefs(8)   =   ":id=1,.fontname=宋体"
                     _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                     _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
                     _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=134"
                     _StyleDefs(12)  =   ":id=2,.fontname=宋体"
                     _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
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
                     _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.alignment=2,.valignment=2"
                     _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                     _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgpicMode=2,.bgbmp=1"
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
                     _StyleDefs(40)  =   "Named:id=33:Normal"
                     _StyleDefs(41)  =   ":id=33,.parent=0"
                     _StyleDefs(42)  =   "Named:id=34:Heading"
                     _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                     _StyleDefs(44)  =   ":id=34,.wraptext=-1"
                     _StyleDefs(45)  =   "Named:id=35:Footing"
                     _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                     _StyleDefs(47)  =   "Named:id=36:Selected"
                     _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                     _StyleDefs(49)  =   "Named:id=37:Caption"
                     _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
                     _StyleDefs(51)  =   "Named:id=38:HighlightRow"
                     _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                     _StyleDefs(53)  =   "Named:id=39:EvenRow"
                     _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                     _StyleDefs(55)  =   "Named:id=40:OddRow"
                     _StyleDefs(56)  =   ":id=40,.parent=33"
                     _StyleDefs(57)  =   "Named:id=41:RecordSelector"
                     _StyleDefs(58)  =   ":id=41,.parent=34"
                     _StyleDefs(59)  =   "Named:id=42:FilterBar"
                     _StyleDefs(60)  =   ":id=42,.parent=33"
                     _StyleDefs(61)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIyanIyanIyanIyanIyanIya"
                     _StyleDefs(62)  =   "bmp(1):id=1,nIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIya"
                     _StyleDefs(63)  =   "bmp(2):id=1,nIyanIyanAAAAJSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSm"
                     _StyleDefs(64)  =   "bmp(3):id=1,pZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpQAAAJyurZyurZyurZyurZyurZyurZyu"
                     _StyleDefs(65)  =   "bmp(4):id=1,rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyu"
                     _StyleDefs(66)  =   "bmp(5):id=1,rZyurQAAAKW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2"
                     _StyleDefs(67)  =   "bmp(6):id=1,taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2tQAAAK2+va2+va2+va2+va2+va2+va2+va2+"
                     _StyleDefs(68)  =   "bmp(7):id=1,va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
                     _StyleDefs(69)  =   "bmp(8):id=1,vQAAAK2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
                     _StyleDefs(70)  =   "bmp(9):id=1,va2+va2+va2+va2+va2+va2+va2+va2+va2+vQAAALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
                     _StyleDefs(71)  =   "bmp(10):id=1,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAA"
                     _StyleDefs(72)  =   "bmp(11):id=1,ALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
                     _StyleDefs(73)  =   "bmp(12):id=1,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
                     _StyleDefs(74)  =   "bmp(13):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3P"
                     _StyleDefs(75)  =   "bmp(14):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
                     _StyleDefs(76)  =   "bmp(15):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
                     _StyleDefs(77)  =   "bmp(16):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAM7X1s7X"
                     _StyleDefs(78)  =   "bmp(17):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
                     _StyleDefs(79)  =   "bmp(18):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1gAAAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
                     _StyleDefs(80)  =   "bmp(19):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1gAAANbj59bj59bj"
                     _StyleDefs(81)  =   "bmp(20):id=1,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
                     _StyleDefs(82)  =   "bmp(21):id=1,59bj59bj59bj59bj59bj5wAAANbj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
                     _StyleDefs(83)  =   "bmp(22):id=1,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj5wAAAN7r797r797r797r"
                     _StyleDefs(84)  =   "bmp(23):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
                     _StyleDefs(85)  =   "bmp(24):id=1,797r797r797r797r7wAAAN7r797r797r797r797r797r797r797r797r797r797r797r797r797r"
                     _StyleDefs(86)  =   "bmp(25):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r7wAAAN7r797r797r797r797r"
                     _StyleDefs(87)  =   "bmp(26):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
                     _StyleDefs(88)  =   "bmp(27):id=1,797r797r797r7wAAAA=="
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   7200
               Left            =   18555
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   15
               Width           =   15000
               _cx             =   26458
               _cy             =   12700
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
               _GridInfo       =   $"frmOrderProduct.frx":E1F0
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar25 
                  Height          =   420
                  Left            =   30
                  TabIndex        =   11
                  Top             =   30
                  Width           =   14940
                  _LayoutVersion  =   1
                  _ExtentX        =   26353
                  _ExtentY        =   741
                  _DataPath       =   ""
                  Bands           =   "frmOrderProduct.frx":E288
               End
               Begin TrueOleDBGrid80.TDBGrid TDBGrid5 
                  Height          =   6720
                  Left            =   30
                  TabIndex        =   12
                  Top             =   450
                  Width           =   14940
                  _ExtentX        =   26353
                  _ExtentY        =   11853
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
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=4339"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
                  Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
                  Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(6)=   "Column(1).Width=4339"
                  Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
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
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
                  _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(12)  =   ":id=2,.fontname=宋体"
                  _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
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
                  _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.alignment=2,.valignment=2"
                  _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                  _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgpicMode=2,.bgbmp=1"
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
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   7200
               Left            =   18855
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   15
               Width           =   15000
               _cx             =   26458
               _cy             =   12700
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
               _GridInfo       =   $"frmOrderProduct.frx":1079E
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.PictureBox Picture5 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   585
                  Left            =   30
                  ScaleHeight     =   555
                  ScaleWidth      =   14910
                  TabIndex        =   15
                  TabStop         =   0   'False
                  Top             =   30
                  Width           =   14940
                  Begin XtremeSuiteControls.PushButton PushButton3 
                     Height          =   435
                     Left            =   6300
                     TabIndex        =   16
                     Top             =   60
                     Width           =   1755
                     _Version        =   1048578
                     _ExtentX        =   3096
                     _ExtentY        =   767
                     _StockProps     =   79
                     Caption         =   "选择图片"
                     UseVisualStyle  =   -1  'True
                  End
                  Begin MSComDlg.CommonDialog CommonDialog2 
                     Left            =   0
                     Top             =   0
                     _ExtentX        =   847
                     _ExtentY        =   847
                     _Version        =   393216
                  End
                  Begin XtremeSuiteControls.PushButton PushButton5 
                     Height          =   435
                     Left            =   1800
                     TabIndex        =   17
                     Top             =   60
                     Width           =   1755
                     _Version        =   1048578
                     _ExtentX        =   3096
                     _ExtentY        =   767
                     _StockProps     =   79
                     Caption         =   "上传图片"
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.Label Label19 
                     Height          =   195
                     Left            =   9300
                     TabIndex        =   18
                     Top             =   180
                     Width           =   3435
                     _Version        =   1048578
                     _ExtentX        =   6059
                     _ExtentY        =   344
                     _StockProps     =   79
                     Caption         =   "在切换页面之前,请先上传图片,会失去图片"
                     ForeColor       =   8421631
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
               End
               Begin VB.PictureBox Picture4 
                  Height          =   6525
                  Left            =   30
                  ScaleHeight     =   6465
                  ScaleWidth      =   14880
                  TabIndex        =   14
                  TabStop         =   0   'False
                  Top             =   645
                  Width           =   14940
               End
            End
            Begin TrueOleDBGrid80.TDBGrid TDBGrid6 
               Height          =   7200
               Left            =   19155
               TabIndex        =   19
               Top             =   15
               Width           =   15000
               _ExtentX        =   26458
               _ExtentY        =   12700
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
               Splits(0)._ColumnProps(1)=   "Column(0).Width=4339"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=4339"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic12 
               Height          =   7200
               Left            =   20055
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   15
               Width           =   15000
               _cx             =   26458
               _cy             =   12700
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
               _GridInfo       =   $"frmOrderProduct.frx":10822
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin TrueOleDBGrid80.TDBGrid TDBGrid9 
                  Height          =   6555
                  Left            =   30
                  TabIndex        =   82
                  Top             =   615
                  Width           =   14940
                  _ExtentX        =   26353
                  _ExtentY        =   11562
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
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=4339"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
                  Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
                  Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(6)=   "Column(1).Width=4339"
                  Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
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
               Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar28 
                  Height          =   555
                  Left            =   30
                  TabIndex        =   83
                  Top             =   30
                  Width           =   14940
                  _LayoutVersion  =   1
                  _ExtentX        =   26353
                  _ExtentY        =   979
                  _DataPath       =   ""
                  Bands           =   "frmOrderProduct.frx":108A6
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic13 
               Height          =   7200
               Left            =   20355
               TabIndex        =   84
               TabStop         =   0   'False
               Top             =   15
               Width           =   15000
               _cx             =   26458
               _cy             =   12700
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
               _GridInfo       =   $"frmOrderProduct.frx":1282E
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin TrueOleDBGrid80.TDBGrid TDBGrid10 
                  Height          =   6555
                  Left            =   30
                  TabIndex        =   85
                  Top             =   615
                  Width           =   14940
                  _ExtentX        =   26353
                  _ExtentY        =   11562
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
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=4339"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
                  Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
                  Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(6)=   "Column(1).Width=4339"
                  Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
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
               Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar29 
                  Height          =   555
                  Left            =   30
                  TabIndex        =   86
                  Top             =   30
                  Width           =   14940
                  _LayoutVersion  =   1
                  _ExtentX        =   26353
                  _ExtentY        =   979
                  _DataPath       =   ""
                  Bands           =   "frmOrderProduct.frx":128B2
               End
            End
         End
         Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar22 
            Height          =   405
            Left            =   30
            TabIndex        =   64
            Top             =   30
            Width           =   16035
            _LayoutVersion  =   1
            _ExtentX        =   28284
            _ExtentY        =   714
            _DataPath       =   ""
            Bands           =   "frmOrderProduct.frx":154F6
         End
      End
   End
End
Attribute VB_Name = "frmOrderProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private theClientID As String
Private theBLTool As New clsAutoCreateBL
Private Const theObjectID As String = "12B001"  '订单单据对象编号
Private oComboBinder As clsCJComboLinker

Private theVsFlexGridTool As New clsVsFlexGrid
Private rsdetail As New ADODB.RecordSet
 
Private theRsBill As RecordSet   '主表记录集
Private theRsDetail As RecordSet  '明细表记录集
Private rsdetailwhite As RecordSet  '白坯明细表记录集
Private rsdetailColor As RecordSet

Public theID As Long   '全局主表自增字段

Private theidwhite As Long
Private rss As RecordSet
Private clsBL As clsBL
Private theidColor As String
Private cls1 As New clsPicture
Private cls2 As New clsPicture
Private szFile As String
Private szFile2 As String
Private szFile3 As String
Private rsgrid4 As RecordSet
Private rsgrid5 As RecordSet
Private rsgrid6 As RecordSet
Private rsgrid7 As RecordSet '箱数计划记录集
Private rsgrid8 As RecordSet '辅料计划记录集
Private rsgrid9 As RecordSet '裁剪计划记录集
Private rsgrid10 As RecordSet '缝制计划记录集
Private Type RGB
        Red   As Byte
        Green   As Byte
        Blue   As Byte
End Type

Private strSQL As String

Public mvarObjectID As String

Private note As String

Public report1 As String
'验证身份和时间
Private clspI As New clspI

Private Logger As New clsFile
Private theLogFile As String
Private rssport As RecordSet
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Property Set Rsreport(ByRef vData As RecordSet)
    Set rssport = vData

End Property


Public Property Let ObjectID(ByVal vData As String)
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
    ObjectID = mvarObjectID
End Property


'登账为正式数据
Private Sub save()

    theBLTool.Update
End Sub


Private Sub GetCurDetail()
'    Set theRsDetail = New RecordSet
'    strSQL = "Select * from G_DraftBillDetailOrder where B_ID='" & theID & "'"
'    theRsDetail.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
'    theRsDetail
      Detail
End Sub


Private Sub AddNewDetail()
'    VSFlexGrid1.AddItem ""

    '弹出页面用于编辑数据，保存数据到草稿明细表中
    '刷新网格数据 - 从草稿明细表中获取数据
    GetCurDetail
    'datediff()
End Sub

Private Function GetOperator() As String
    GetOperator = Gm.SysID.SystemUserName
End Function

'新增一个主表
Private Sub AddNewBill()
    ClearAll
    note = ""
    FlatEdit1.Text = GetCodeID
    delivery '--交货方式
    baozhuang
    Whether '--是否开票
'    Valuation '--计价单位
    Settlement '--结算方式
    Business '--业务跟单
    DateTimePicker1.Value = Now
    DateTimePicker2.Value = Now
    FlatEdit8.Text = GetOperator
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim a As String
    Dim sql As String
    sql1 = "select *from G_DraftBillOrder"
    rs.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.AddNew
   
    a = Format(Now, "YYYY-MM-DD")
    rs!B_datecreate = a
    rs.Update
    theID = rs!B_id
    '白坯计划新增
    sql = "delete from G_DraftBillOrder where B_ID='" & theID & "'"
    Gm.cnnTool.cnn.Execute sql
  
    DraftDetail
    whitedetail
    pattern
    pictureorder
    WhiteComposition
    Colordetail
'    whitedetail
     SetBillState False
     SetAuditState 0
     ActiveBar22.Bands("Band1").Tools("审核").Enabled = False
     ActiveBar22.Bands("Band1").Tools("作废").Enabled = False
     ActiveBar22.Bands("Band1").Tools("取消作废").Enabled = False
     ActiveBar22.Bands("Band1").Tools("作废图片").Visible = False
     sumall
     setRs
     setRs1
     auxiliary
End Sub

Private Function GetCodeID() As String
    GetCodeID = theBLTool.GetFrameCodeDetail(theObjectID)
End Function

'新增进行主表和明细表清空
Private Sub ClearAll()
     On Error Resume Next
    Dim o As Object
    
    For Each o In Me.Controls
        Select Case TypeName(o)
        
            Case "FlatEdit"
                o.Text = ""
            Case "ComboBox"
                o.Text = ""
        End Select
    Next
End Sub

'检查是否存在数据 1.存在B_ID 进行数据填写，2.没有进行新增
Private Sub checkMain()
     Dim rs As New RecordSet
     Dim sql As String
     sql = "select * from G_BillOrder order by B_ID desc"
     rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
     'rss.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

     If rs.RecordCount > 0 Then
        theID = rs!B_id
        SetAuditState IIf(IsNull(rs!B_Audit), 0, rs!B_Audit)
        
        openbill

        SetBillState True
'        savewhitebill

     Else
        Set rss = New RecordSet
        Dim sql1 As String
        sql1 = "select *from G_DraftBillOrder where 1=0"
        rss.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        rss.AddNew
        Dim b As String
        b = Format(Now, "YYYY-MM-DD")
        rss!B_datecreate = b
        rss.Update
        theID = rss!B_id
'         Set rss = Nothing
'         rss.Close
        DraftDetail
        Dim sql2 As String
        sql2 = "delete from G_DraftBillOrder where B_ID='" & theID & "'"
        Gm.cnnTool.cnn.Execute sql2
        SetAuditState 0
     End If
'        Set rss = Nothing
'
End Sub
'打箱计划
Private Sub setRs()
    Set rsgrid7 = New RecordSet
    rsgrid7.Fields.Append "B_id", adVarChar, 100
    rsgrid7.Fields.Append "B_OrderCode", adVarChar, 100
    rsgrid7.Fields.Append "B_goodsid", adVarChar, 100
    rsgrid7.Fields.Append "B_goodsName", adVarChar, 100
    rsgrid7.Fields.Append "B_size", adVarChar, 100
    rsgrid7.Fields.Append "B_Width", adVarChar, 100
    rsgrid7.Fields.Append "B_Weight", adVarChar, 100
    rsgrid7.Fields.Append "B_patterncode", adVarChar, 100
    rsgrid7.Fields.Append "B_colorid", adVarChar, 100
    rsgrid7.Fields.Append "B_color", adVarChar, 100
    rsgrid7.Fields.Append "B_Hex", adVarChar, 100
    rsgrid7.Fields.Append "B_CasePack", adVarChar, 100
    rsgrid7.Fields.Append "B_boxname", adVarChar, 100
    rsgrid7.Fields.Append "B_boxgg", adVarChar, 100
    rsgrid7.Fields.Append "B_memo", adVarChar, 100
    rsgrid7.Fields.Append "B_orderitemid", adVarChar, 100
    rsgrid7.Fields.Append "B_orderid", adVarChar, 100
    rsgrid7.Open
    
    TDBGrid7.DataSource = rsgrid7
    setrsDetail
End Sub
Private Sub setrsDetail()
    TDBGrid7.Columns("B_OrderCode").Caption = "订单号"
    TDBGrid7.Columns("B_goodsName").Caption = "品名"
'     TDBGrid7.Columns("B_goodsID").Caption = "品名"
    TDBGrid7.Columns("B_size").Caption = "尺寸"
    TDBGrid7.Columns("B_Width").Caption = "门幅"
    TDBGrid7.Columns("B_Weight").Caption = "克重"
    TDBGrid7.Columns("B_patterncode").Caption = "色号"
    TDBGrid7.Columns("B_color").Caption = "颜色"
    TDBGrid7.Columns("B_Hex").Caption = "色块"
    TDBGrid7.Columns("B_CasePack").Caption = "每箱数量"
    TDBGrid7.Columns("B_boxname").Caption = "箱名称"
    TDBGrid7.Columns("B_boxgg").Caption = "箱规格"
    TDBGrid7.Columns("B_memo").Caption = "备注"
    
    TDBGrid7.Columns("B_OrderCode").width = 1000
    TDBGrid7.Columns("B_goodsName").width = 1200
    TDBGrid7.Columns("B_Width").width = 1000
    TDBGrid7.Columns("B_Weight").width = 1000
    TDBGrid7.Columns("B_patterncode").width = 1000
    TDBGrid7.Columns("B_color").width = 1000
    TDBGrid7.Columns("B_Hex").width = 1000
    TDBGrid7.Columns("B_CasePack").width = 1500
    TDBGrid7.Columns("B_boxname").width = 1500
    TDBGrid7.Columns("B_boxgg").width = 1500
    
    TDBGrid7.Columns("B_size").Locked = True
    TDBGrid7.Columns("B_OrderCode").Locked = True
    TDBGrid7.Columns("B_goodsName").Locked = True
    TDBGrid7.Columns("B_Width").Locked = True
    TDBGrid7.Columns("B_Weight").Locked = True
    TDBGrid7.Columns("B_patterncode").Locked = True
    TDBGrid7.Columns("B_color").Locked = True
    TDBGrid7.Columns("B_Hex").Locked = True
   
    TDBGrid7.Columns("B_id").Visible = False
    TDBGrid7.Columns("B_id").Locked = True
    TDBGrid7.Columns("B_id").AllowSizing = False
    TDBGrid7.Columns("B_goodsid").Visible = False
    TDBGrid7.Columns("B_goodsid").Locked = True
    TDBGrid7.Columns("B_goodsid").AllowSizing = False
    TDBGrid7.Columns("B_colorid").Visible = False
    TDBGrid7.Columns("B_colorid").Locked = True
    TDBGrid7.Columns("B_colorid").AllowSizing = False
    TDBGrid7.Columns("B_orderitemid").Visible = False
    TDBGrid7.Columns("B_orderitemid").Locked = True
    TDBGrid7.Columns("B_orderitemid").AllowSizing = False
    TDBGrid7.Columns("B_orderid").Visible = False
    TDBGrid7.Columns("B_orderid").Locked = True
    TDBGrid7.Columns("B_orderid").AllowSizing = False
    
    TDBGrid7.Columns("B_Hex").FetchStyle = True
    TDBGrid7.HoldFields
    TDBGrid7.MarqueeStyle = dbgHighlightRow
End Sub
'裁剪计划
Private Sub setRs1()
    Set rsgrid9 = New RecordSet
    rsgrid9.Fields.Append "B_id", adVarChar, 100
    rsgrid9.Fields.Append "B_itemCode", adVarChar, 100
    rsgrid9.Fields.Append "B_label", adVarChar, 100
    rsgrid9.Fields.Append "B_size", adVarChar, 100
    rsgrid9.Fields.Append "B_colorid", adVarChar, 100
    rsgrid9.Fields.Append "B_color", adVarChar, 100
    rsgrid9.Fields.Append "B_Hex", adVarChar, 100
    rsgrid9.Fields.Append "B_BarCode", adVarChar, 100
    rsgrid9.Fields.Append "B_ChiCun", adVarChar, 100
    rsgrid9.Fields.Append "B_qtyall", adVarChar, 100
    rsgrid9.Fields.Append "B_quantity", adVarChar, 100
    rsgrid9.Fields.Append "B_boxqty", adVarChar, 100
    rsgrid9.Fields.Append "B_memo", adVarChar, 100
    rsgrid9.Fields.Append "B_orderitemid", adVarChar, 100
    rsgrid9.Fields.Append "B_orderid", adVarChar, 100
    rsgrid9.Open
    
    TDBGrid9.DataSource = rsgrid9
    setrsDetail1
End Sub
Private Sub setrsDetail1()
    TDBGrid9.Columns("B_itemCode").Caption = "编号"
    TDBGrid9.Columns("B_label").Caption = "主唛"
    TDBGrid9.Columns("B_size").Caption = "Size"
    TDBGrid9.Columns("B_color").Caption = "颜色"
    TDBGrid9.Columns("B_Hex").Caption = "色块"
    TDBGrid9.Columns("B_BarCode").Caption = "条形码"
    TDBGrid9.Columns("B_ChiCun").Caption = "尺寸"
    TDBGrid9.Columns("B_qtyall").Caption = "总条数"
    TDBGrid9.Columns("B_quantity").Caption = "每箱数量"
    TDBGrid9.Columns("B_boxqty").Caption = "箱数"
    TDBGrid9.Columns("B_memo").Caption = "备注"
    

    TDBGrid9.Columns("B_color").width = 1000
    TDBGrid9.Columns("B_Hex").width = 1000
     TDBGrid9.Columns("B_itemCode").width = 1000
    TDBGrid9.Columns("B_label").width = 1000
     TDBGrid9.Columns("B_size").width = 1000
    TDBGrid9.Columns("B_Hex").width = 1000
     TDBGrid9.Columns("B_ChiCun").width = 1000
    TDBGrid9.Columns("B_qtyall").width = 1000
     TDBGrid9.Columns("B_quantity").width = 1000
    TDBGrid9.Columns("B_boxqty").width = 1000
    
    
    TDBGrid9.Columns("B_color").Locked = True
    TDBGrid9.Columns("B_Hex").Locked = True
   TDBGrid9.Columns("B_boxqty").Locked = True
   
    TDBGrid9.Columns("B_id").Visible = False
    TDBGrid9.Columns("B_id").Locked = True
    TDBGrid9.Columns("B_id").AllowSizing = False
        TDBGrid9.Columns("B_boxqty").Visible = False
    TDBGrid9.Columns("B_boxqty").Locked = True
    TDBGrid9.Columns("B_boxqty").AllowSizing = False
        TDBGrid9.Columns("B_quantity").Visible = False
    TDBGrid9.Columns("B_quantity").Locked = True
    TDBGrid9.Columns("B_quantity").AllowSizing = False
   
    TDBGrid9.Columns("B_colorid").Visible = False
    TDBGrid9.Columns("B_colorid").Locked = True
    TDBGrid9.Columns("B_colorid").AllowSizing = False
    TDBGrid9.Columns("B_orderitemid").Visible = False
    TDBGrid9.Columns("B_orderitemid").Locked = True
    TDBGrid9.Columns("B_orderitemid").AllowSizing = False
    TDBGrid9.Columns("B_orderid").Visible = False
    TDBGrid9.Columns("B_orderid").Locked = True
    TDBGrid9.Columns("B_orderid").AllowSizing = False
    
    TDBGrid9.Columns("B_Hex").FetchStyle = True
    TDBGrid9.HoldFields
    TDBGrid9.MarqueeStyle = dbgHighlightRow
End Sub
'缝制计划
Private Sub setRs2()
    Set rsgrid10 = New RecordSet
    rsgrid10.Fields.Append "B_id", adVarChar, 100
    rsgrid10.Fields.Append "B_itemCode", adVarChar, 100
    rsgrid10.Fields.Append "B_KuanHao", adVarChar, 100  '款号
    rsgrid10.Fields.Append "B_label", adVarChar, 100
    rsgrid10.Fields.Append "B_size", adVarChar, 100
    rsgrid10.Fields.Append "B_colorid", adVarChar, 100
    rsgrid10.Fields.Append "B_color", adVarChar, 100
    rsgrid10.Fields.Append "B_Hex", adVarChar, 100
    rsgrid10.Fields.Append "B_BarCode", adVarChar, 100
    rsgrid10.Fields.Append "B_ChiCun", adVarChar, 100
    rsgrid10.Fields.Append "B_qtyall", adVarChar, 100
    rsgrid10.Fields.Append "B_quantity", adVarChar, 100
    rsgrid10.Fields.Append "B_boxqty", adVarChar, 100
    rsgrid10.Fields.Append "B_process", adVarChar, 100
    rsgrid10.Fields.Append "B_memo", adVarChar, 100
    rsgrid10.Fields.Append "B_orderitemid", adVarChar, 100
    rsgrid10.Fields.Append "B_orderid", adVarChar, 100
    
    rsgrid10.Open
    
    TDBGrid10.DataSource = rsgrid10
    setrsDetail2
End Sub
Private Sub setrsDetail2()
    TDBGrid10.Columns("B_itemCode").Caption = "编号"
    TDBGrid10.Columns("B_KuanHao").Caption = "款号"
    TDBGrid10.Columns("B_label").Caption = "主唛"
    TDBGrid10.Columns("B_size").Caption = "Size"
    TDBGrid10.Columns("B_color").Caption = "颜色"
    TDBGrid10.Columns("B_Hex").Caption = "色块"
    TDBGrid10.Columns("B_BarCode").Caption = "条形码"
    TDBGrid10.Columns("B_ChiCun").Caption = "尺寸"
    TDBGrid10.Columns("B_qtyall").Caption = "总条数"
    TDBGrid10.Columns("B_quantity").Caption = "每箱数量"
    TDBGrid10.Columns("B_boxqty").Caption = "箱数"
    TDBGrid10.Columns("B_process").Caption = "缝制工艺"
    TDBGrid10.Columns("B_memo").Caption = "备注"
    
    TDBGrid10.Columns("B_process").ValueItems.Presentation = dbgComboBox
    
    TDBGrid10.Columns("B_color").width = 1000
    TDBGrid10.Columns("B_Hex").width = 1000
     TDBGrid10.Columns("B_itemCode").width = 1000
     TDBGrid10.Columns("B_KuanHao").width = 1000
    TDBGrid10.Columns("B_label").width = 1000
     TDBGrid10.Columns("B_size").width = 1000
    TDBGrid10.Columns("B_Hex").width = 1000
     TDBGrid10.Columns("B_ChiCun").width = 1000
    TDBGrid10.Columns("B_qtyall").width = 1000
     TDBGrid10.Columns("B_quantity").width = 1000
    TDBGrid10.Columns("B_boxqty").width = 1000
    
    TDBGrid10.Columns("B_process").Locked = True
    TDBGrid10.Columns("B_color").Locked = True
    TDBGrid10.Columns("B_Hex").Locked = True
   TDBGrid10.Columns("B_boxqty").Locked = True
   
    TDBGrid10.Columns("B_id").Visible = False
    TDBGrid10.Columns("B_id").Locked = True
    TDBGrid10.Columns("B_id").AllowSizing = False
       TDBGrid10.Columns("B_boxqty").Visible = False
    TDBGrid10.Columns("B_boxqty").Locked = True
    TDBGrid10.Columns("B_boxqty").AllowSizing = False
       TDBGrid10.Columns("B_quantity").Visible = False
    TDBGrid10.Columns("B_quantity").Locked = True
    TDBGrid10.Columns("B_quantity").AllowSizing = False
   
    TDBGrid10.Columns("B_colorid").Visible = False
    TDBGrid10.Columns("B_colorid").Locked = True
    TDBGrid10.Columns("B_colorid").AllowSizing = False
    TDBGrid10.Columns("B_orderitemid").Visible = False
    TDBGrid10.Columns("B_orderitemid").Locked = True
    TDBGrid10.Columns("B_orderitemid").AllowSizing = False
    TDBGrid10.Columns("B_orderid").Visible = False
    TDBGrid10.Columns("B_orderid").Locked = True
    TDBGrid10.Columns("B_orderid").AllowSizing = False
    
    TDBGrid10.Columns("B_Hex").FetchStyle = True
    TDBGrid10.HoldFields
    TDBGrid10.MarqueeStyle = dbgHighlightRow
End Sub


Private Sub InitLayout()
    'C1Tab1.TabVisible(1) = False
    
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
    DateTimePicker1.Value = Now
    DateTimePicker2.Value = Now
    
   
''    FlatEdit8.Text=
'    Detail
'    setVSFlexGrid
'    '初始化单据对象的工具类
'    Set theBLTool = New clsAutoCreateBL
'
'    '传入单据编号，获取如下数据：
'    '1. 可获取最新的单据编号
'    '2. 获取4个表组成的套表
'    theBLTool.InitCls theObjectID
End Sub

Private Sub ActiveBar210_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
                Case "打印当前订单"
                    prit1
    End Select
End Sub

Private Sub ActiveBar22_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)

    Select Case Tool.name
        
        Case "保存"
            saveALL
        Case "主表 - 新增"
            AddNewBill
        Case "新增行"
            VSFlexGrid1_null
        Case "第一单"
            movefrist
        Case "前一单"
            MovePreview
        Case "后一单"
            movenext
        Case "最后单"
            movelast
        Case "退出"
             Unload Me
             theID = 0
        Case "删除行"
            Deleterow
        Case "删除"
            delete
        Case "复制行"
            Copyrow
        Case "复制多行"
            CopyrowAll
        Case "审核"
            Audit 1
        Case "取消审核"
            Audit 0
        Case "生成色布计划"
            CopyToColor
        Case "全部生成色布计划"
           CopyToColorAll
        Case "保存样式"
            setGridStyle
        Case "作废"
            invalid 1
        Case "取消作废"
            invalid 0
        Case "色布采购"
            colorcast
        Case "生成箱数计划"
            boxadd
        Case "此行生成裁剪计划"
            cjaddone
        Case "全部生成裁剪计划"
            cjaddall
        Case "此行生成缝制计划"
            fzaddone
        Case "全部生成缝制计划"
            fzaddall
        Case "全部生成箱数计划"
            boxall
    End Select
End Sub
'合同中色布采购
Private Sub colorcast()
    On Error Resume Next
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If ActiveBar22.Bands("Band1").Tools("审核图片").Visible = True Then
        Dim frm1 As New frmOrderColor
        frm1.id = rsdetail!B_ItemID
        frm1.FlatEdit1.Text = rsdetail!B_ordercode
        frm1.FlatEdit3.Text = rsdetail!B_Width
        frm1.FlatEdit4.Text = rsdetail!B_weight
        frm1.FlatEdit2.Text = rsdetail!B_GoodsID
        frm1.Show vbModal
        rsgrid6.requery
    Else
        MsgBox "此单没有审核", vbInformation, "提示"
        Exit Sub
    End If
    
End Sub
'生成裁剪计划
Private Sub cjaddone()
    On Error Resume Next
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If ActiveBar22.Bands("Band1").Tools("审核图片").Visible = True Then
        rsgrid9.AddNew
        rsgrid9!B_hex = rsdetail!B_hex
        rsgrid9!B_colorid = rsdetail!B_sid
        rsgrid9!B_color = rsdetail!B_color
        rsgrid9!B_orderitemid = rsdetail!B_ItemID
        rsgrid9!B_OrderID = theID
        rsgrid9.Update
        C1Tab1.CurrTab = 9
    Else
        MsgBox "此单没有审核", vbInformation, "提示"
        Exit Sub
    End If
    
End Sub
Private Sub cjaddall()
    Dim rs As New RecordSet
    Set rs = rsdetail.Clone
      On Error Resume Next
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If ActiveBar22.Bands("Band1").Tools("审核图片").Visible = True Then
        rs.MoveFirst
        Do While Not rs.EOF
            rsgrid9.AddNew
            rsgrid9!B_hex = rs!B_hex
            rsgrid9!B_colorid = rs!B_sid
            rsgrid9!B_color = rs!B_color
            rsgrid9!B_orderitemid = rs!B_ItemID
            rsgrid9!B_OrderID = theID
            rsgrid9.Update
             rs.movenext
        Loop
        C1Tab1.CurrTab = 9
    Else
        MsgBox "此单没有审核", vbInformation, "提示"
        Exit Sub
    End If
End Sub
'生成缝制计划
Private Sub fzaddone()
    On Error Resume Next
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If ActiveBar22.Bands("Band1").Tools("审核图片").Visible = True Then
        rsgrid10.AddNew
        rsgrid10!B_hex = rsdetail!B_hex
        rsgrid10!B_colorid = rsdetail!B_sid
        rsgrid10!B_color = rsdetail!B_color
        rsgrid10!B_orderitemid = rsdetail!B_ItemID
         rsgrid10!B_KuanHao = rsdetail!B_KuanHao
        rsgrid10!B_OrderID = theID
       
        rsgrid10.Update
        C1Tab1.CurrTab = 10
    Else
        MsgBox "此单没有审核", vbInformation, "提示"
        Exit Sub
    End If
    
End Sub
Private Sub fzaddall()
    Dim rs As New RecordSet
    Set rs = rsdetail.Clone
      On Error Resume Next
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If ActiveBar22.Bands("Band1").Tools("审核图片").Visible = True Then
        rs.MoveFirst
        Do While Not rs.EOF
            rsgrid10.AddNew
            rsgrid10!B_hex = rs!B_hex
            rsgrid10!B_colorid = rs!B_sid
            rsgrid10!B_color = rs!B_color
            rsgrid10!B_orderitemid = rs!B_ItemID
            rsgrid10!B_KuanHao = rsdetail!B_KuanHao
            rsgrid10!B_OrderID = theID
            rsgrid10.Update
             rs.movenext
        Loop
        C1Tab1.CurrTab = 10
    Else
        MsgBox "此单没有审核", vbInformation, "提示"
        Exit Sub
    End If
End Sub
Private Sub invalid(ByVal a As Long)
   
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim sql2 As String
    
    sql = "select * from G_Billorder where B_ID='" & theID & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
     If rs.RecordCount > 0 Then
        Dim strSQL1 As String
        Dim rs1 As New RecordSet
        strSQL1 = "select * from G_SystemUser where B_UserName='" & Gm.SysID.SystemUser & "'"
        rs1.Open strSQL1, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
        If rs1!B_SuperAdmin <> 1 Then
            MsgBox "您不是超级管理员，不能进行修改", vbInformation, "提示"
            Exit Sub
        End If
        If a = 1 Then
            SetInvalidState True
'            If report1 <> "" Then
'                rssport.requery
'                report1 = ""
'            End If
            sql1 = "update G_BillOrder set B_invalid=1 where B_ID='" & theID & "' "
            ActiveBar22.Bands("Band1").Tools("作废").Enabled = False
            ActiveBar22.Bands("Band1").Tools("取消作废").Enabled = True
            
            Debug.Print sql1
            Gm.cnnTool.cnn.Execute sql1
        End If
        If a = 0 Then
            SetInvalidState False
'            If report1 <> "" Then
'                rssport.requery
'                report1 = ""
'            End If
            ActiveBar22.Bands("Band1").Tools("作废").Enabled = True
            ActiveBar22.Bands("Band1").Tools("取消作废").Enabled = False
            sql1 = "update G_BillOrder set B_invalid=0 where B_ID='" & theID & "' "
            Gm.cnnTool.cnn.Execute sql1
        End If
    End If
    
End Sub


Private Sub Audit(ByVal a As Long)
    Dim sql3 As String
    Dim s As String
    s = Format(Now, "YYYY-MM-DD")
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    
    

    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_BIllorder where B_ID='" & theID & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        Dim strSQL1 As String
        Dim rs1 As New RecordSet
        strSQL1 = "select * from G_SystemUser where B_UserName='" & Gm.SysID.SystemUser & "'"
        rs1.Open strSQL1, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
        If rs1!B_SuperAdmin <> 1 Then
            MsgBox "您不是超级管理员，不能进行修改", vbInformation, "提示"
            Exit Sub
        End If
        If a = 1 Then
            SetAuditState 1
            
           sql3 = "update G_BillOrder set B_DateAudit='" & s & "' where B_ID='" & theID & "'"
           Gm.cnnTool.cnn.Execute sql3
            
            
            If report1 <> "" Then
                rssport.requery
                report1 = ""
            End If
        End If
        If a = 0 Then
            SetAuditState 0
            If report1 <> "" Then
                rssport.requery
                report1 = ""
            End If
        End If
    End If
End Sub


'Private Sub Command1_Click()
'    Picture6.Picture = LoadPicture("C:\Users\Administrator\Desktop\公司文件\白玉兰工艺卡软件\工艺卡基础资料图片\花型\AAA.jpg")
'End Sub

Private Sub Form_Load()
'        theid = 0
        theLogFile = App.Path & "\log.txt"
        
        
'        Logger.WriteFileContent theLogFile, "进入FORMLOAD"
        
        SetBillState False
        
'        Logger.WriteFileContent theLogFile, "SetBillState执行完毕"
'        InitRsBill
'        InitRsDetail
        FlatEdit1.Text = GetCodeID
        
        InitLayout
        delivery '--交货方式
        baozhuang
        Whether '--是否开票
'        Valuation '--计价单位
        Settlement '--结算方式
        Business '--业务跟单
        FlatEdit8.Text = GetOperator
      
 
        
        
        If theID <= 0 Then
            AddNewBill
        End If
         
        
'       checkMain
       Label16.Visible = False
       Label17.Visible = False
       Label18.Visible = False
       CheckBox1.Visible = False
       CheckBox2.Visible = False
       CheckBox3.Visible = False
       Label11.Visible = False
       ComboBox3.Visible = False
       note = ""
       C1Tab1.TabIndex = 0
        setRs
        setRs1
        setRs2
        HT_Name '合同命名
 
        
         TDBGrid1.HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid1.Splits(0).HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid1.HighlightRowStyle.ForeColor = &H80000008
    TDBGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
    TDBGrid1.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell
    
    TDBGrid2.HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid2.Splits(0).HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid2.HighlightRowStyle.ForeColor = &H80000008
    TDBGrid2.MarqueeStyle = dbgHighlightRowRaiseCell
    TDBGrid2.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell
    
        TDBGrid3.HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid3.Splits(0).HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid3.HighlightRowStyle.ForeColor = &H80000008
    TDBGrid3.MarqueeStyle = dbgHighlightRowRaiseCell
    TDBGrid3.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell
    
            TDBGrid6.HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid6.Splits(0).HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid6.HighlightRowStyle.ForeColor = &H80000008
    TDBGrid6.MarqueeStyle = dbgHighlightRowRaiseCell
    TDBGrid6.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell
    
             TDBGrid7.HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid7.Splits(0).HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid7.HighlightRowStyle.ForeColor = &H80000008
    TDBGrid7.MarqueeStyle = dbgHighlightRowRaiseCell
    TDBGrid7.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell
    
             TDBGrid8.HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid8.Splits(0).HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid8.HighlightRowStyle.ForeColor = &H80000008
    TDBGrid8.MarqueeStyle = dbgHighlightRowRaiseCell
    TDBGrid8.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell
    
             TDBGrid9.HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid9.Splits(0).HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid9.HighlightRowStyle.ForeColor = &H80000008
    TDBGrid9.MarqueeStyle = dbgHighlightRowRaiseCell
    TDBGrid9.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell
    
    TDBGrid10.HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid10.Splits(0).HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid10.HighlightRowStyle.ForeColor = &H80000008
    TDBGrid10.MarqueeStyle = dbgHighlightRowRaiseCell
    TDBGrid10.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell
    
       '草稿表中有数据先显示
         reduction
End Sub

Private Sub reduction()
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    sql = "SELECT DISTINCT B_id FROM G_draftBillDetailorder where B_username='" & Gm.SysID.SystemUser & "' and isnull(B_ContractLogodetail,0)=1 ORDER BY B_id DESC "
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If MsgBox("有未保存的草稿数据,是否应用", vbYesNo + vbDefaultButton1 + vbInformation, "提示") = vbYes Then
            rs.MoveFirst
            theID = rs!B_id
            DraftDetail
        Else
            rs.MoveFirst
            sql1 = "delete from G_draftBillDetailorder where B_id='" & rs!B_id & "'"
            Gm.cnnTool.cnn.Execute sql1
        End If
    Else
        Exit Sub
    End If
End Sub

'将合同号名字替换（根据不同工厂不同命名）
Private Sub HT_Name()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "select * from G_Config_FormCtlShow where B_sid='合同号'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        Label7.Caption = rs!B_Caption
        
    End If
End Sub
Private Sub saveALL()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
    
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    
    
    If clspI.authenticate(theID) = False Then
        Exit Sub
    End If
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
    Dim sql3 As String
    Dim rs3 As New RecordSet
    Dim rs4 As New RecordSet
    Dim sql4 As String
    Dim Rs5 As New RecordSet
    Dim sql5 As String
    
    
    sql = "select * from G_BillOrder where B_PactCode='" & Trim(FlatEdit4.Text) & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    sql2 = "select *from G_BillOrder where B_id='" & theID & "'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs2.RecordCount > 0 Then
        If rs2!B_PactCode <> FlatEdit4.Text Then
            If rs.RecordCount > 0 Then
                MsgBox "合同号已经存在，不能重复", vbInformation, "提示"
                Exit Sub
            End If
        End If
    Else
        If rs.RecordCount > 0 Then
            MsgBox "合同号已经存在，不能重复", vbInformation, "提示"
            Exit Sub
        End If
    End If
    sql1 = "select * from G_BilldetailOrder where B_orderCode in (select B_orderCode from G_DraftBilldetailOrder where B_ID='" & theID & "')"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs1.RecordCount > 0 Then
        MsgBox "订单号" & rs1!B_ordercode & "，不能重复", vbInformation, "提示"
        
        Exit Sub
    End If
     
     
    If Trim(FlatEdit2.Text) = "" Then
            MsgBox "客户不能为空", vbInformation, "提示"
            Exit Sub
        End If
        If Trim(ComboBox1.Text) = "" Then
            MsgBox "交货方式不能为空", vbInformation, "提示"
            Exit Sub
        End If
        If Trim(FlatEdit4.Text) = "" Then
            MsgBox "合同号不能为空", vbInformation, "提示"
            Exit Sub
        End If
         If Trim(ComboBox2.Text) = "" Then
            MsgBox "合同开票不能为空", vbInformation, "提示"
            Exit Sub
        End If
'        If Trim(ComboBox3.Text) = "" Then
'            MsgBox "计价单位不能为空", vbInformation, "提示"
'            Exit Sub
'        End If
        If Trim(ComboBox4.Text) = "" Then
            MsgBox "结算方式不能为空", vbInformation, "提示"
            Exit Sub
        End If
         If Trim(ComboBox5.Text) = "" Then
            MsgBox "业务跟单不能为空", vbInformation, "提示"
            Exit Sub
        End If
        If rsdetail.RecordCount <= 0 Then
            MsgBox "明细表没有数据不能保存", vbInformation, "提示"
            Exit Sub
        End If
        Debug.Print Picture4.Picture
        
     
        If Picture1.Picture.Handle <> 0 Then
            MsgBox "合同扫描件为空", vbInformation, "提示"
            Exit Sub
        Else
            sql3 = "select * from G_Imageorder where B_ID='" & theID & "'"
            rs3.Open sql3, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
            If rs3.RecordCount <= 0 Then
                MsgBox "先上传合同扫描件", vbInformation, "提示"
                Exit Sub
            End If
        End If
           
  
        savemain
        savedetail
'        savewhitebill
        MsgBox "保存成功", vbInformation, "提示"
        
        '删除草稿表
        deletedraft
        SetBillState True
        sql4 = "select *From G_Billorder where B_id='" & theID & "'"
        rs4.Open sql4, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If IIf(IsNull(rs4!B_Audit), 0, rs4!B_Audit) = 0 Then
            ActiveBar22.Bands("Band1").Tools("审核").Enabled = True
            ActiveBar22.Bands("Band1").Tools("取消审核").Enabled = False
        Else
            ActiveBar22.Bands("Band1").Tools("审核").Enabled = False
            ActiveBar22.Bands("Band1").Tools("取消审核").Enabled = True
            ActiveBar22.Bands("Band1").Tools("作废").Enabled = True
            ActiveBar22.Bands("Band1").Tools("取消作废").Enabled = False
            ActiveBar22.Bands("Band1").Tools("作废图片").Visible = False
        End If
        If IIf(IsNull(rs4!B_invalid), 0, rs4!B_invalid) = 0 Then
            ActiveBar22.Bands("Band1").Tools("作废").Enabled = True
            ActiveBar22.Bands("Band1").Tools("取消作废").Enabled = False
            ActiveBar22.Bands("Band1").Tools("作废图片").Visible = False
        Else
            ActiveBar22.Bands("Band1").Tools("作废").Enabled = False
            ActiveBar22.Bands("Band1").Tools("取消作废").Enabled = True
            ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True
        End If
        pattern
End Sub

Private Sub savemain()
        
        Dim rs4 As New RecordSet
        Dim sql4 As String
     
        sql4 = "Select B_SID From G_Employee  where B_Name='" & Trim(ComboBox5.Text) & "'"
        rs4.Open sql4, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        Dim rs1 As New RecordSet
        Dim sql1 As String
        sql1 = "select * from G_BillOrder where B_ID='" & theID & "'"
        rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs1.RecordCount > 0 Then
            Dim d As String
            Dim T As String
            Dim h As String
              If "否 " = ComboBox2.Text Then
                    h = 0
                Else
                    h = 1
                End If
            
            Dim rs2 As New RecordSet
            Dim sql2 As String
             d = Format(DateTimePicker1.Value, "YYYY-MM-DD")
             T = Format(DateTimePicker2.Value, "YYYY-MM-DD")
             sql2 = "exec usp_saveBillupdate '" & theID & "','" & Trim(FlatEdit1.Text) & "','" & d & "','" & T & "','" & theClientID & "','" & Trim(ComboBox1.Text) & "','" & Trim(FlatEdit4.Text) & "','" & Trim(FlatEdit5.Text) & "','" & Trim(FlatEdit6.Text) & "','" & h & "','" & Trim(ComboBox3.Text) & "','" & Trim(ComboBox4.Text) & "','" & Trim(FlatEdit7.Text) & "','" & rs4!B_sid & "','" & Trim(CheckBox1.Value) & "','" & Trim(CheckBox2.Value) & "','" & Trim(CheckBox3.Value) & "','" & note & "','" & ComboBox6.Text & "','" & Trim(FlatEdit9.Text) & "',''"
             rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        Else
        Dim rs As New RecordSet
        Dim sql As String
        Dim a As String
        Dim b As String
        Dim c  As Integer
        If "否 " = ComboBox2.Text Then
            c = 0
        Else
            c = 1
        End If
        a = Format(DateTimePicker1.Value, "YYYY-MM-DD")
        b = Format(DateTimePicker2.Value, "YYYY-MM-DD")
        
        sql = "exec usp_saveBill '" & theID & "','" & Trim(FlatEdit1.Text) & "','" & a & "','" & b & "','" & theClientID & "','" & Trim(ComboBox1.Text) & "','" & Trim(FlatEdit4.Text) & "','" & Trim(FlatEdit5.Text) & "','" & Trim(FlatEdit6.Text) & "','" & c & "','" & Trim(ComboBox3.Text) & "','" & Trim(ComboBox4.Text) & "','" & Trim(FlatEdit7.Text) & "','" & rs4!B_sid & "','" & Gm.SysID.SystemUser & "','" & Trim(CheckBox1.Value) & "','" & Trim(CheckBox2.Value) & "','" & Trim(CheckBox3.Value) & "','" & note & "','" & ComboBox6.Text & "','" & Trim(FlatEdit9.Text) & "','1',''"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        
        
        End If
End Sub
Private Sub savedetail()
        Dim rs1 As New RecordSet
        Dim sql1 As String
        sql1 = "select *from G_BillDetailOrder where B_ID='" & theID & "'"
        rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs1.RecordCount > 0 Then
''            Dim sql2 As String
''            sql2 = "delete from G_BillDetailOrder where B_ID='" & theID & "'"
''            Gm.cnnTool.cnn.Execute sql2
                Exit Sub
        End If
        Dim rs As New RecordSet
        Dim sql As String
        sql = "select * from G_BillDetailOrder where 1=0 order by B_OrderCode,B_itemid"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rsdetail.RecordCount > 0 Then
            rsdetail.MoveFirst
        
            Do While Not rsdetail.EOF
                rs.AddNew
                rs!B_id = theID
                rs!B_ItemID = rsdetail!B_ItemID
                rs!B_ordercode = rsdetail!B_ordercode
                rs!B_GoodsID = rsdetail!B_GoodsID
                rs!B_Width = rsdetail!B_Width
                rs!B_weight = rsdetail!B_weight
                rs!B_colorid = rsdetail!B_sid
                rs!B_Positivefabric = rsdetail!B_Positivefabric
                rs!B_Middlefabric = rsdetail!B_Middlefabric
                rs!B_Backfabric = rsdetail!B_Backfabric
                rs!B_qty = rsdetail!B_qty
                rs!B_BoxQty = rsdetail!B_BoxQty
                rs!B_QtyPerbox = rsdetail!B_QtyPerbox
                rs!B_price = rsdetail!B_price
                rs!B_Sum = rsdetail!B_Sum
                rs!B_MemoDetail = rsdetail!B_MemoDetail
                rs!B_color = rsdetail!B_color
                rs!B_GoodManual = rsdetail!B_GoodManual
                rs!B_process = rsdetail!B_process
                rs!B_Packaging = rsdetail!B_Packaging
                rs!B_PositiveFactory = rsdetail!B_PositiveFactory
                rs!B_MiddleFactory = rsdetail!B_MiddleFactory
                rs!B_BackFactory = rsdetail!B_BackFactory
                
                rs!B_Width2 = rsdetail!B_Width2
                rs!B_Weight2 = rsdetail!B_Weight2
                rs!B_Width3 = rsdetail!B_Width3
                rs!B_Weight3 = rsdetail!B_Weight3
                rs!B_seam = rsdetail!B_seam
                rs!B_Beatbox = rsdetail!B_Beatbox
                rs!B_size = rsdetail!B_size
                
                rs.Update
                rsdetail.movenext
            Loop
        End If
        
        
        
End Sub

'删除草稿表
Private Sub deletedraft()
    Dim sql As String
    sql = "delete from G_DraftBillOrder where B_ID='" & theID & "'"
    Gm.cnnTool.cnn.Execute sql
    Dim sql1 As String
    sql1 = "delete from G_DraftBillDetailOrder where B_ID='" & theID & "'"
    Gm.cnnTool.cnn.Execute sql1
     
End Sub
'将主表和明细表全部删除
Private Sub delete()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If

    If clspI.authenticate(theID) = False Then
        Exit Sub
    End If
    Dim rs  As New RecordSet
    Dim sql2 As String
    sql2 = "select *from G_BillOrder where B_ID='" & theID & "'"
    rs.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        Dim sql As String
        sql = "delete from G_BillOrder where B_ID='" & theID & "'"
        Gm.cnnTool.cnn.Execute sql
        Dim sql1 As String
        sql1 = "delete from G_BillDetailOrder where B_ID='" & theID & "'"
        Gm.cnnTool.cnn.Execute sql1
        
        DeletewhiteAll
        DeleteColorAll
        DeletePicture
        DeletePicture_1
        deleteWhiteComposition
        MsgBox "删除成功", vbInformation, "提示"
        AddNewBill
    '白坯计划删除
'    Deletewhite
    End If
End Sub
Private Sub movefrist()
    Dim sql2 As String
    Dim rs2 As New RecordSet
    
    If TDBGrid1.ApproxCount > 0 Then
        sql2 = "select * from G_Billorder where B_ID='" & theID & "'"
        rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs2.RecordCount <= 0 Then
            If MsgBox("表中有数据,是否需要保存", vbYesNo + vbDefaultButton2 + vbInformation, "") = vbNo Then
                Dim sql3 As String
                sql3 = "delete from G_draftBilldetailorder where B_ID='" & theID & "'"
                Gm.cnnTool.cnn.Execute sql3
                rsdetail.requery
            Else
                saveALL
            End If
        End If
    End If
    Dim rs As New RecordSet
    Dim sql As String
    Dim sql1 As String
    Dim rs1 As New RecordSet
    sql1 = "select * from G_systemUser where B_username='" & Gm.SysID.SystemUser & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs1!B_SuperAdmin = 1 Then
        sql = "select top 1 * from G_BillOrder where B_ContractLogo=1"
    Else
        sql = "select top 1 * from G_BillOrder where B_Username='" & Gm.SysID.SystemUser & "' and B_ContractLogo=1"
        
    End If
    
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     If rs.RecordCount <= 0 Then
            MsgBox "当前没有数据", vbInformation, "提示"
     Else
            theID = rs!B_id
            SetAuditState IIf(IsNull(rs!B_Audit), 0, rs!B_Audit)
            openbill
            SetBillState True
     End If
    
End Sub
Private Sub MovePreview()
     Dim sql2 As String
    Dim rs2 As New RecordSet
    If TDBGrid1.ApproxCount > 0 Then
        sql2 = "select * from G_Billorder where B_ID='" & theID & "'"
        rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs2.RecordCount <= 0 Then
            If MsgBox("表中有数据,是否需要保存", vbYesNo + vbDefaultButton2 + vbInformation, "") = vbNo Then
                Dim sql3 As String
                sql3 = "delete from G_draftBilldetailorder where B_ID='" & theID & "'"
                Gm.cnnTool.cnn.Execute sql3
                rsdetail.requery
            Else
                saveALL
            End If
        End If
    End If
     Dim rs As New RecordSet
     Dim sql As String
     Dim sql1 As String
    Dim rs1 As New RecordSet
    sql1 = "select * from G_systemUser where B_username='" & Gm.SysID.SystemUser & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs1!B_SuperAdmin = 1 Then
        sql = "select top 1 * from G_BillOrder where B_ID<'" & theID & "' and B_ContractLogo=1  Order by B_ID desc"
    Else
        sql = "select top 1 * from G_BillOrder where B_ID<'" & theID & "' and B_ContractLogo=1 and B_username='" & Gm.SysID.SystemUser & "'  Order by B_ID desc"
    End If
     
      rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
      
      If rs.RecordCount <= 0 Then
        MsgBox "已经是第一单了", vbOKOnly + vbInformation, "提示"
     Else
        theID = rs!B_id
        SetAuditState IIf(IsNull(rs!B_Audit), 0, rs!B_Audit)
        openbill
        SetBillState True
     End If
End Sub
Private Sub movenext()
     Dim sql2 As String
    Dim rs2 As New RecordSet
    If TDBGrid1.ApproxCount > 0 Then
        sql2 = "select * from G_Billorder where B_ID='" & theID & "'"
        rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs2.RecordCount <= 0 Then
            If MsgBox("表中有数据,是否需要保存", vbYesNo + vbDefaultButton2 + vbInformation, "") = vbNo Then
                Dim sql3 As String
                sql3 = "delete from G_draftBilldetailorder where B_ID='" & theID & "'"
                Gm.cnnTool.cnn.Execute sql3
                rsdetail.requery
            Else
                saveALL
            End If
        End If
    End If
     Dim rs As New RecordSet
     Dim sql As String
    Dim sql1 As String
    Dim rs1 As New RecordSet
    sql1 = "select * from G_systemUser where B_username='" & Gm.SysID.SystemUser & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs1!B_SuperAdmin = 1 Then
        sql = "select top 1 * from G_BillOrder where B_ID>'" & theID & "' and B_ContractLogo=1  Order by B_ID asc"
    Else
        sql = "select top 1 * from G_BillOrder where B_ID>'" & theID & "' and B_ContractLogo=1 and B_username='" & Gm.SysID.SystemUser & "'  Order by B_ID asc"
    End If
     
      rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
      If rs.RecordCount <= 0 Then
        MsgBox "已经是最后一单了", vbOKOnly + vbInformation, "提示"
    Else
        theID = rs!B_id
        SetAuditState IIf(IsNull(rs!B_Audit), 0, rs!B_Audit)
        openbill
        SetBillState True
    End If
   
End Sub
Private Sub movelast()
    Dim sql2 As String
    Dim rs2 As New RecordSet
    If TDBGrid1.ApproxCount > 0 Then
        sql2 = "select * from G_Billorder where B_ID='" & theID & "'"
        rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs2.RecordCount <= 0 Then
            If MsgBox("表中有数据,是否需要保存", vbYesNo + vbDefaultButton2 + vbInformation, "") = vbNo Then
                Dim sql3 As String
                sql3 = "delete from G_draftBilldetailorder where B_ID='" & theID & "'"
                Gm.cnnTool.cnn.Execute sql3
                rsdetail.requery
            Else
                saveALL
            End If
        End If
    End If
    Debug.Print Gm.SysID.SystemUser
    Dim rs As New RecordSet
    Dim sql As String
    Dim sql1 As String
    Dim rs1 As New RecordSet
    sql1 = "select * from G_systemUser where B_username='" & Gm.SysID.SystemUser & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs1!B_SuperAdmin = 1 Then
        sql = "Select top 1 * from G_BillOrder where B_ContractLogo=1 order by B_ID desc"
    Else
        sql = "Select top 1 * from G_BillOrder where B_Username='" & Gm.SysID.SystemUser & "'and B_ContractLogo=1 order  by B_ID desc"
    End If
    
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     If rs.RecordCount <= 0 Then
        MsgBox "当前没有任何数据！", vbOKOnly + vbInformation, "提示"
    Else
        theID = rs!B_id
        SetAuditState IIf(IsNull(rs!B_Audit), 0, rs!B_Audit)
        openbill   '根据全局变量Sid打开单据，主表明细表显示到UI的对应位置
        SetBillState True
    End If
End Sub

Public Sub openbill()
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim rs3 As New RecordSet
    Dim sql3 As String
    If theID <= 0 Then
        Exit Sub
    End If
    
    
    sql = "select *from G_BillOrder where B_ID='" & theID & "' "
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    sql1 = "select *from G_SystemUser where B_UserName='" & IIf(IsNull(rs!B_username), "", rs!B_username) & "' "
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic

    FlatEdit1.Text = IIf(IsNull(rs!B_Codeid), "", rs!B_Codeid)
    DateTimePicker1.Value = IIf(IsNull(rs!B_Date), "", rs!B_Date)
    DateTimePicker2.Value = IIf(IsNull(rs!B_Deliverydate), "", rs!B_Deliverydate)
    FlatEdit2.Text = getClientName(rs!B_Clientid)
    theClientID = rs!B_Clientid
    FlatEdit3.Text = getAlias(rs!B_Clientid)
    ComboBox1.Text = IIf(IsNull(rs!B_DeliveryType), "", rs!B_DeliveryType)
    FlatEdit4.Text = IIf(IsNull(rs!B_PactCode), "", rs!B_PactCode)
    FlatEdit5.Text = IIf(IsNull(rs!B_PactQty), "", rs!B_PactQty)
    FlatEdit6.Text = IIf(IsNull(rs!B_PactBoxQty), "", rs!B_PactBoxQty)
    If 1 = rs!B_invoice Then
        ComboBox2.Text = "是"
    Else
        ComboBox2.Text = "否"
    End If
'    ComboBox2.Text = IIf(IsNull(rs!B_invoice), "", rs!B_invoice)
    ComboBox3.Text = IIf(IsNull(rs!B_Denominated), "", rs!B_Denominated)
    ComboBox4.Text = IIf(IsNull(rs!B_BalanceType), "", rs!B_BalanceType)
    FlatEdit7.Text = IIf(IsNull(rs!B_signed), "", rs!B_signed)
    FlatEdit8.Text = GetUserName(IIf(IsNull(rs!B_username), "", rs!B_username))
    FlatEdit9.Text = IIf(IsNull(rs!B_StyleHao), "", rs!B_StyleHao)
    ComboBox5.Text = getBusinessName(rs!B_BusinessGD)
    ComboBox6.Text = IIf(IsNull(rs!B_package), "", rs!B_package)
    
'    ComboBox5.Text = IIf(IsNull(rs!B_BusinessGD), "", rs!B_BusinessGD)
    
    'FlatEdit8.Text = IIf(IsNull(rs1!B_UserDes), "", rs1!B_UserDes)
    CheckBox1.Value = IIf(IsNull(rs!B_cargoClear), 0, rs!B_cargoClear)
    CheckBox2.Value = IIf(IsNull(rs!B_ticketClear), 0, rs!B_ticketClear)
    CheckBox3.Value = IIf(IsNull(rs!B_paragraphClear), 0, rs!B_paragraphClear)
    
    ActiveBar22.Bands("Band1").Tools("审核").Enabled = False
    ActiveBar22.Bands("Band1").Tools("取消审核").Enabled = True
    
    If IIf(IsNull(rs!B_invalid), 0, rs!B_invalid) = 0 Then
        ActiveBar22.Bands("Band1").Tools("作废").Enabled = True
        ActiveBar22.Bands("Band1").Tools("取消作废").Enabled = False
        ActiveBar22.Bands("Band1").Tools("作废图片").Visible = False
    Else
        ActiveBar22.Bands("Band1").Tools("取消作废").Enabled = True
        ActiveBar22.Bands("Band1").Tools("作废").Enabled = False
        ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True
        
    End If
    Dim sql2 As String
    Dim rs2 As New RecordSet
    sql2 = "select * from G_BillOrder  where B_ID='" & theID & "'"
    Debug.Print sql2
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs2!B_Audit = 1 Then
        SetAuditState 1
    Else
        SetAuditState 0
    End If
    note = IIf(IsNull(rs!B_memo), "", rs!B_memo)
    
    sql3 = "select *from G_Billorder where B_ID='" & theID & "'"
    rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs3.RecordCount > 0 Then
         SetBillState True
    End If
    
    
    Detail
End Sub

Private Sub PushButton1_Click()
    Dim frm1 As New frmPopupClient
    frm1.Show vbModal
    theClientID = frm1.clientid
    FlatEdit3.Text = frm1.Alias
    FlatEdit2.Text = frm1.ClientName
    Unload frm1
End Sub
'包装方式
Private Sub baozhuang()
    Dim rs As New RecordSet
    Dim sql As String
    ComboBox6.Clear
    sql = "Select B_SID From G_PackWay Where 1=1"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

    Do While Not rs.EOF
        ComboBox6.AddItem rs!B_sid
        rs.movenext
    Loop
End Sub
'--交货方式
Private Sub delivery()
    ComboBox1.Clear
    Dim rs As RecordSet
    Set rs = New RecordSet
    Dim sql As String
    sql = "Select B_SID From G_Delivery Where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     Do While Not rs.EOF
        ComboBox1.AddItem rs!B_sid
        rs.movenext
    Loop
End Sub
'--是否开票
Private Sub Whether()
    ComboBox2.Clear
    ComboBox2.AddItem "是"
    ComboBox2.AddItem "否"
    ComboBox2.Text = "是"
End Sub
'--计价单位
'Private Sub Valuation()
'    ComboBox3.Clear
'    ComboBox3.AddItem "米数"
'    ComboBox3.AddItem "公斤数"
'    ComboBox3.Text = "公斤数"
'End Sub
'--结算方式
Private Sub Settlement()
    ComboBox4.Clear
    Dim rs As RecordSet
    Set rs = New RecordSet
    Dim sql As String
    sql = "Select B_SID From G_BalanceContract Where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    Do While Not rs.EOF
        ComboBox4.AddItem rs!B_sid
        rs.movenext
    Loop
'    ComboBox4.ListIndex = 2
End Sub
'--业务跟单
Private Sub Business()
    ComboBox5.Clear
    Dim rs As RecordSet
    Set rs = New RecordSet
    Dim sql As String
    sql = "Select B_SID,B_Name From G_Employee Where B_Department='业务员'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
'    Do While Not rs.EOF
'        ComboBox5.AddItem rs!B_Name, , rs!B_SID
'        rs.MoveNext
'    Loop
    Set oComboBinder = New clsCJComboLinker
    oComboBinder.InitCls ComboBox5, "B_Name", "B_SID", rs

End Sub

'明细表初始化
Private Sub Detail()
    On Error Resume Next
    Dim sql As String
    Set rsdetail = New RecordSet
    sql = "select B_itemid,B_OrderCode,B_SourceOrderCode,B_KuanHao,B_GoodsID,b.B_Name,B_GoodManual,B_Size,B_Width,B_Weight,B_patterncode,B_color,a.B_Hex,a.B_SID,"
    sql = sql & " B_PatternCode2,B_ColorID2,B_Color2,B_MianLiaoQty2,B_ComputUnit2,B_MianLiaoPrice2,B_PatternCode3,B_ColorID3,B_Color3,B_MianLiaoQty3,B_ComputUnit3,B_MianLiaoPrice3,B_Positivefabric,B_Middlefabric,B_Backfabric,B_Qty,B_BoxQty,B_QtyPerbox,B_Price,B_Sum,B_Process,B_Packaging,B_MemoDetail,B_PositiveFactory,B_MiddleFactory,B_BackFactory,B_Width2,B_Weight2,B_Width3,B_Weight3,B_seam,B_Beatbox"
    sql = sql & ",B_MianLiaoQty,B_ComputUnit,B_MianLiaoPrice,B_HomeMill,B_ProcessQty,B_ProcessPrice,B_ProcessMoney"
    sql = sql & " from G_BillDetailOrder left outer join G_Color a on  G_BillDetailOrder.B_ColorID=a.B_SID "
    sql = sql & " left outer join G_Product b on  G_BillDetailOrder.B_GoodsID=b.B_SID "
     sql = sql & " where B_ID='" & theID & "' Order by B_OrderCode,B_itemid asc"
    Debug.Print sql
    rsdetail.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
 
    TDBGrid1.DataSource = rsdetail
   
   
    TDBGrid1.Columns("B_Hex").FetchStyle = True
    
    
    whitedetail
  
    Colordetail

    pattern

    pictureorder

    WhiteComposition
       
    ColorProcure
    
    boxPlan  '打箱计划
    
    tailorPlan '裁剪计划
    sewPlan    '缝制计划
    auxiliary
    
    OpenImageBZ
    
    SetGrid
    HT_Name2
End Sub
'草稿明细表
Private Sub DraftDetail()
    Dim sql As String
    Set rsdetail = New RecordSet
    sql = "select B_itemid,B_OrderCode,B_SourceOrderCode,B_KuanHao,B_GoodsID,b.B_Name,B_GoodManual,B_Size,B_Width,B_Weight,B_patterncode,B_color,a.B_Hex,a.B_SID,"
    sql = sql & " B_Positivefabric,B_Middlefabric,B_Backfabric,B_Qty,B_BoxQty,B_QtyPerbox,B_Price,B_Sum,B_Process,B_Packaging,B_MemoDetail,B_PositiveFactory,B_MiddleFactory,B_BackFactory,B_Width2,B_Weight2,B_Width3,B_Weight3,B_seam,B_Beatbox,B_Size"
    sql = sql & " from G_draftBillDetailOrder left outer join G_Color a on  G_draftBillDetailOrder.B_ColorID=a.B_SID "
    sql = sql & " left outer join G_Product b on  G_draftBillDetailOrder.B_GoodsID=b.B_SID "
     sql = sql & " where B_ID='" & theID & "' Order by B_OrderCode,B_itemid asc"
Debug.Print sql
    rsdetail.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    TDBGrid1.DataSource = rsdetail
    
    TDBGrid1.Columns("B_Hex").FetchStyle = True
    SetGrid
    
    Dim sql2 As String
    Dim rs2 As New RecordSet
    sql2 = "select * from G_Config_FormCtlShow where B_sid='订单号' "
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs2.RecordCount > 0 Then
        TDBGrid1.Columns("B_OrderCode").Caption = rs2!B_Caption
    End If
End Sub

Private Sub setGridShow()
    Dim cls1 As New clsGridShow
    With cls1
        .ObjectID = "11S066"
        .InitClass TDBGrid1, 3
        .ShowGridFormat
    End With
End Sub

'保存网格样式
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
                strSQL = "update G_BLSField set B_GridWidth='" & dWidth & "' where B_ObjectID='11S066' and B_FieldName='" & szFieldName & "'"
                Gm.cnnTool.cnn.Execute strSQL
            End If
        End If
    Next
    
End Sub

Private Sub SetGrid()
    setGridShow
'
'
'      If rsDetail.RecordCount > 0 Then
'         If Len(rsDetail!B_Colorid) > 0 Then
'              TDBGrid1.Columns("B_Hex").FetchStyle = True
'        End If
'    End If
'    TDBGrid1.Columns("B_itemid").Caption = ""
'    TDBGrid1.Columns("B_OrderCode").Caption = "订单号"
'    TDBGrid1.Columns("B_GoodsID").Caption = "品名"
'    TDBGrid1.Columns("B_Width").Caption = "门幅"
'    TDBGrid1.Columns("B_Weight").Caption = "克重"
'    TDBGrid1.Columns("B_ColorID").Caption = "颜色"
'    TDBGrid1.Columns("B_HX").Caption = "花型"
'    TDBGrid1.Columns("B_PatternCode").Caption = "色号/花号"
'    TDBGrid1.Columns("B_Meter").Caption = "米数"
'    TDBGrid1.Columns("B_KG").Caption = "公斤数"
'    TDBGrid1.Columns("B_Qty").Caption = "码数"
'    TDBGrid1.Columns("B_Price").Caption = "单价"
'    TDBGrid1.Columns("B_Sum").Caption = "金额"
'    TDBGrid1.Columns("B_BackMaterial").Caption = "计价单位"
'    TDBGrid1.Columns("B_MemoDetail").Caption = "后道工序"
'    TDBGrid1.Columns("B_Hex").Caption = "颜色标识"
'
    TDBGrid1.Columns("B_itemid").width = 0
    TDBGrid1.Columns("B_itemid").Visible = False
    TDBGrid1.Columns("B_itemid").AllowSizing = False
        TDBGrid1.Columns("B_PositiveFactory").width = 0
    TDBGrid1.Columns("B_PositiveFactory").Visible = False
    TDBGrid1.Columns("B_PositiveFactory").AllowSizing = False
        TDBGrid1.Columns("B_MiddleFactory").width = 0
    TDBGrid1.Columns("B_MiddleFactory").Visible = False
    TDBGrid1.Columns("B_MiddleFactory").AllowSizing = False
        TDBGrid1.Columns("B_BackFactory").width = 0
    TDBGrid1.Columns("B_BackFactory").Visible = False
    TDBGrid1.Columns("B_BackFactory").AllowSizing = False
    
    
    TDBGrid1.Columns("B_Width2").width = 0
    TDBGrid1.Columns("B_Width2").Visible = False
    TDBGrid1.Columns("B_Width2").AllowSizing = False
    TDBGrid1.Columns("B_Weight2").width = 0
    TDBGrid1.Columns("B_Weight2").Visible = False
    TDBGrid1.Columns("B_Weight2").AllowSizing = False
    TDBGrid1.Columns("B_Width3").width = 0
    TDBGrid1.Columns("B_Width3").Visible = False
    TDBGrid1.Columns("B_Width3").AllowSizing = False
    TDBGrid1.Columns("B_Weight3").width = 0
    TDBGrid1.Columns("B_Weight3").Visible = False
    TDBGrid1.Columns("B_Weight3").AllowSizing = False
      TDBGrid1.Columns("B_seam").width = 0
    TDBGrid1.Columns("B_seam").Visible = False
    TDBGrid1.Columns("B_seam").AllowSizing = False
    TDBGrid1.Columns("B_Beatbox").width = 0
    TDBGrid1.Columns("B_Beatbox").Visible = False
    TDBGrid1.Columns("B_Beatbox").AllowSizing = False
'        TDBGrid1.Columns("B_size").width = 0
'    TDBGrid1.Columns("B_size").Visible = False
'    TDBGrid1.Columns("B_size").AllowSizing = False



    TDBGrid1.Columns("B_Price").NumberFormat = "0.00"
    TDBGrid1.Columns("B_Sum").NumberFormat = "0.00"
       TDBGrid1.Columns("B_GoodsID").width = 0
    TDBGrid1.Columns("B_GoodsID").Visible = False
    TDBGrid1.Columns("B_GoodsID").AllowSizing = False
'    TDBGrid1.Columns("B_Hex").width = 0
'    TDBGrid1.Columns("B_Hex").Visible = False
'    TDBGrid1.Columns("B_Hex").AllowSizing = False
'    '设置宽度
'    TDBGrid1.Columns("B_Qty").width = 1200
'
'    TDBGrid1.Columns("B_MemoDetail").width = 1300
'    TDBGrid1.Columns("B_OrderCode").width = 800
'    TDBGrid1.Columns("B_GoodsID").width = 2000
'    TDBGrid1.Columns("B_Width").width = 800
'    TDBGrid1.Columns("B_Weight").width = 800
'    TDBGrid1.Columns("B_ColorID").width = 1500
'    TDBGrid1.Columns("B_HX").width = 1000
'    TDBGrid1.Columns("B_PatternCode").width = 1500
'    TDBGrid1.Columns("B_Meter").width = 800
'    TDBGrid1.Columns("B_KG").width = 1000
'    TDBGrid1.Columns("B_Price").width = 800
'    TDBGrid1.Columns("B_Sum").width = 1200
'    TDBGrid1.Columns("B_Hex").width = 1000
'    TDBGrid1.Columns("B_Price").NumberFormat = "0.0"
'    TDBGrid1.Columns("B_Sum").NumberFormat = "0.00"
'    TDBGrid1.Columns("B_Qty").NumberFormat = "0.00"
    
    TDBGrid1.MarqueeStyle = dbgHighlightRow
    '小计
    sumall
  
    
    TDBGrid1.HoldFields
End Sub
'将合同号名字替换（根据不同工厂不同命名）
Private Sub HT_Name2()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "select * from G_Config_FormCtlShow where B_sid='订单号' "
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        TDBGrid1.Columns("B_OrderCode").Caption = rs!B_Caption

        TDBGrid3.Columns("B_ItemIDB").Caption = rs!B_Caption
        TDBGrid2.Columns("B_ItemIDB").Caption = rs!B_Caption
        TDBGrid4.Columns("B_OrderCode").Caption = rs!B_Caption
        TDBGrid5.Columns("B_ItemIDB").Caption = rs!B_Caption
        TDBGrid6.Columns("B_ItemIDB").Caption = rs!B_Caption
        TDBGrid7.Columns("B_OrderCode").Caption = rs!B_Caption
    End If
End Sub



Private Sub PushButton6_Click()
    Dim frm1 As New frmnote
    frm1.FlatEdit1.Text = note
    frm1.Show vbModal
        If frm1.bsave = True Then
            note = frm1.FlatEdit1.Text
        End If
    Unload frm1
End Sub

Private Sub PushButton7_Click()
    Dim sql As String
    If theID > 0 Then
        sql = "delete from G_image where B_ID='" & theID & "'"
        Gm.cnnToolImage.cnn.Execute sql
    End If
    Picture3.Picture = Nothing
End Sub
'包装计划预览图片
Private Sub PushButton8_Click()
 If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    On Error GoTo IFERR
    With CommonDialog3
        .ShowOpen
        szFile3 = .FileName
    End With
    
    If Len(szFile3) <= 0 Then
        Exit Sub
    End If
    cls1.InitCls szFile3, Picture6
    
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "图片格式不支持，请提供JPG或者BMP格式图片！" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub
'包装计划上传图片
Private Sub PushButton9_Click()
 Dim sql As String
    Dim rs As New RecordSet
    If TDBGrid1.ApproxCount <= 0 Then
        MsgBox "当前没有订单号不能上传", vbInformation, "提示"
        Exit Sub
    End If
    If szFile3 <> "" Then
'                 需要引用：Microsoft Scripting Runtime
        Dim fso As New FileSystemObject
    
        Dim lLength As Long
        Dim oFile As File
    
        Set oFile = fso.GetFile(szFile3)
        
        sql = "select * from G_ImageSize"
        rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
        
        
        If oFile.Size / 1000000 > rs!B_size Then
            MsgBox "图片太大不能上传", vbInformation, "提示"
            Exit Sub
        End If
    
        '获取的长度的单位是：字节
        
        saveImageBZ
       
            MsgBox "图片上传成功", vbInformation, "提示"
       
    End If
End Sub

'包装计划删除图片
Private Sub PushButton10_Click()
Dim sql As String
    If theID > 0 Then
        sql = "delete from WVAccountImage.dbo.G_image_NEW_BZ where B_OrderID='" & theID & "'"
        Gm.cnnToolImage.cnn.Execute sql
    
    End If
    Picture6.Picture = Nothing
End Sub
'进行弹出
Private Sub TDBGrid1_DblClick()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If rsdetail.RecordCount > 0 Then
       VSFlexGrid1_UPdate
    Else
        VSFlexGrid1_null
    End If
End Sub

Private Sub VSFlexGrid1_null()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    
    
      If clspI.authenticate(theID) = False Then
            Exit Sub
      End If
     
    Dim frm1 As New frmOrderProduct_Edit
'    frm1.Valuation = ComboBox3.Text
    frm1.client = theClientID
    frm1.id = theID
    frm1.FlatEdit1.TabIndex = 0
    frm1.Show vbModal

    Unload frm1
    rsdetail.requery
    
'    rsDetail.requery
'    setgrid

    sumprice
End Sub

Private Sub VSFlexGrid1_UPdate()
     On Error Resume Next
    
    Dim rs  As RecordSet
    Dim rs1 As New RecordSet
    Dim sql1 As String
    sql1 = "select * from G_BillOrder where B_ID='" & theID & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs1.RecordCount > 0 Then
        Set rs = New RecordSet
        Dim sql2 As String
        sql2 = "select * from G_BillDetailOrder where B_itemid='" & rsdetail!B_ItemID & "' "
        rs.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    Else
    Set rs = New RecordSet
    Dim sql As String
       sql = "select * from G_DraftBillDetailOrder where B_itemid='" & rsdetail!B_ItemID & "' "
       rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    End If
    Dim frm1 As New frmOrderProduct_Edit
    frm1.itemid = rsdetail!B_ItemID
    frm1.FlatEdit1.Text = IIf(IsNull(rs!B_ordercode), "", rs!B_ordercode)
    frm1.FlatEdit2.Text = GetProductName(rs!B_GoodsID)
    frm1.Productid = IIf(IsNull(rs!B_GoodsID), "", rs!B_GoodsID)
    frm1.FlatEdit3.Text = IIf(IsNull(rs!B_Width), "", rs!B_Width)
    frm1.FlatEdit4.Text = IIf(IsNull(rs!B_weight), "", rs!B_weight)
    frm1.colorid = IIf(IsNull(rs!B_colorid), "", rs!B_colorid)
    frm1.FlatEdit15.Text = IIf(IsNull(rs!B_patterncode), "", rs!B_patterncode)
    frm1.FlatEdit11.Text = IIf(IsNull(rs!B_color), "", rs!B_color)
    frm1.FlatEdit5.Text = IIf(IsNull(rs!B_Positivefabric), "", rs!B_Positivefabric)
    frm1.FlatEdit6.Text = IIf(IsNull(rs!B_Middlefabric), "", rs!B_Middlefabric)
    frm1.FlatEdit14.Text = IIf(IsNull(rs!B_Backfabric), "", rs!B_Backfabric)
    frm1.FlatEdit7.Text = IIf(IsNull(rs!B_qty), "", rs!B_qty)
    frm1.FlatEdit12.Text = IIf(IsNull(rs!B_BoxQty), "", rs!B_BoxQty)
    frm1.FlatEdit8.Text = IIf(IsNull(rs!B_QtyPerbox), "", rs!B_QtyPerbox)
    frm1.FlatEdit13.Text = IIf(IsNull(rs!B_MemoDetail), "", rs!B_MemoDetail)
    frm1.FlatEdit9.Text = rs!B_price
    frm1.FlatEdit10.Text = rs!B_Sum
    frm1.ComboBox2.Text = rs!B_GoodManual
    frm1.FlatEdit20.Text = rs!B_process
    frm1.FlatEdit16.Text = rs!B_Packaging
    frm1.FlatEdit17.Text = getClientName(rs!B_PositiveFactory)
    frm1.Positiveid = rs!B_PositiveFactory
    frm1.FlatEdit18.Text = getClientName(rs!B_MiddleFactory)
    frm1.Middleid = rs!B_MiddleFactory
    frm1.FlatEdit19.Text = getClientName(rs!B_BackFactory)
    frm1.backid = rs!B_BackFactory
    frm1.FlatEdit21.Text = rsdetail!B_Width2
    frm1.FlatEdit22.Text = rsdetail!B_Weight2
    frm1.FlatEdit23.Text = rsdetail!B_Width3
    frm1.FlatEdit24.Text = rsdetail!B_Weight3
    frm1.FlatEdit25.Text = rsdetail!B_seam
    frm1.FlatEdit26.Text = rsdetail!B_Beatbox
    frm1.FlatEdit27.Text = rsdetail!B_size
    frm1.FlatEdit28.Text = rsdetail!B_SourceOrderCode
    frm1.FlatEdit29.Text = rs!B_KuanHao
    
    frm1.FlatEdit30.Text = IIf(IsNull(rs!B_MianLiaoQty), "", rs!B_MianLiaoQty)
    frm1.ComboBox1.Text = IIf(IsNull(rs!B_ComputUnit), "", rs!B_ComputUnit)
    frm1.FlatEdit32.Text = IIf(IsNull(rs!B_MianLiaoPrice), "", rs!B_MianLiaoPrice)
    frm1.FlatEdit33.Text = IIf(IsNull(rs!B_HomeMill), "", rs!B_HomeMill)
    frm1.FlatEdit31.Text = IIf(IsNull(rs!B_ProcessQty), "", rs!B_ProcessQty)
    frm1.FlatEdit34.Text = IIf(IsNull(rs!B_ProcessPrice), "", rs!B_ProcessPrice)
    frm1.FlatEdit35.Text = IIf(IsNull(rs!B_ProcessMoney), "", rs!B_ProcessMoney)
    
    frm1.FlatEdit36.Text = IIf(IsNull(rs!B_PatternCode2), "", rs!B_PatternCode2)
    frm1.colorid2 = IIf(IsNull(rs!B_ColorID2), "", rs!B_ColorID2)
    frm1.FlatEdit38.Text = IIf(IsNull(rs!B_Color2), "", rs!B_Color2)
    frm1.FlatEdit39.Text = IIf(IsNull(rs!B_MianLiaoQty2), "", rs!B_MianLiaoQty2)
    frm1.ComboBox3.Text = IIf(IsNull(rs!B_ComputUnit2), "", rs!B_ComputUnit2)
    frm1.FlatEdit40.Text = IIf(IsNull(rs!B_MianLiaoPrice2), "", rs!B_MianLiaoPrice2)
    
    frm1.FlatEdit41.Text = IIf(IsNull(rs!B_PatternCode3), "", rs!B_PatternCode3)
    frm1.colorid3 = IIf(IsNull(rs!B_ColorID3), "", rs!B_ColorID3)
    frm1.FlatEdit42.Text = IIf(IsNull(rs!B_Color3), "", rs!B_Color3)
    frm1.FlatEdit43.Text = IIf(IsNull(rs!B_MianLiaoQty3), "", rs!B_MianLiaoQty3)
    frm1.ComboBox4.Text = IIf(IsNull(rs!B_ComputUnit3), "", rs!B_ComputUnit3)
    frm1.FlatEdit44.Text = IIf(IsNull(rs!B_MianLiaoPrice3), "", rs!B_MianLiaoPrice3)
    
    
    frm1.client = theClientID
    frm1.itemid = rs!B_ItemID
    frm1.id = theID
    frm1.OpenImage
    
    frm1.Show vbModal
    
    Unload frm1
    
    rsdetail.requery
'    requery
    
    rsdetailwhite.requery
    rsdetailColor.requery
    rsdetail.MoveFirst
'    setgrid
    sumprice
    Exit Sub
'IFERR:
'
'    MsgBox "请点击有数据的地方", vbOKOnly + vbInformation, "提示"

End Sub

'进行明细表的数据刷新
Private Sub requery()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_BillDetailorder where B_ID='" & theID & "'"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        Detail
    Else
        DraftDetail
    End If
    
    
End Sub
Private Sub Sum()
    Dim dMS As Double
    Dim dgj As Double
    Dim dSum As Double
    Do While Not rsdetail.EOF
        dMS = dMS + rsdetail!B_meter
        dgj = dgj + rsdetail!B_kg
        dSum = dSum + rsdetail!B_Sum
        rsdetail.movenext
    Loop
  
'    TDBGrid1.Columns("B_PS").FooterText = dPS
'    TDBGrid1.Columns("B_jz").FooterText = dgj
'    TDBGrid1.Columns("RowIndex").FooterText = "合计"
End Sub

'合同订单数和合同金额自动生成
Private Sub sumprice()
    Dim dingdanshu As Long
    Dim dprice As Double
    Dim sql1 As String
    Dim rs1 As New RecordSet
    
    dprice = 0
    Do While Not rsdetail.EOF
        dprice = dprice + IIf(IsNull(rsdetail!B_Sum), 0, rsdetail!B_Sum)
        rsdetail.movenext
    Loop
    sql1 = "select * from G_Billorder where B_ID='" & theID & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs1.RecordCount > 0 Then
        Dim sql3 As String
        Dim rs3 As New RecordSet
        sql3 = "select distinct B_OrderCode from G_Billdetailorder where B_ID='" & theID & "'"
        rs3.Open sql3, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        
        dingdanshu = rs3.RecordCount
        
        Dim rs As New RecordSet
        Dim sql As String
        sql = "update G_BillOrder set B_PactQty='" & dingdanshu & "',B_PactBoxQty='" & dprice & "' where B_ID='" & theID & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    Else
        Dim sql4 As String
        Dim rs4 As New RecordSet
        sql4 = "select distinct B_OrderCode from G_draftBilldetailorder where B_ID='" & theID & "'"
        rs4.Open sql4, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        dingdanshu = rs4.RecordCount
    End If
        FlatEdit5.Text = dingdanshu
        FlatEdit6.Text = dprice
     If rsdetail.RecordCount > 0 Then
        rsdetail.movelast
    End If
End Sub
'主表中制单人名称
Private Function GetUserName(ByVal UserName As String) As String
     Dim rs As New RecordSet
     Dim sql As String
     sql = "select * from G_SystemUser where B_UserName='" & UserName & "'"
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     If rs.RecordCount > 0 Then
        GetUserName = rs!B_UserDes
     Else
        GetUserName = ""
     End If
End Function
'主表中客户ID获取客户名称
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
'获取客户的别称
Private Function getAlias(ByVal client As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "Select B_Alias From G_ContactCompany Where B_ClientID='" & client & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     If rs.RecordCount > 0 Then
        getAlias = rs!B_Alias
     Else
        getAlias = ""
     End If
End Function
'获取颜色名称
Private Function GetColorName(ByVal color As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "Select B_Name From G_Color Where B_SID='" & color & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     If rs.RecordCount > 0 Then
        GetColorName = rs!B_name
     Else
        GetColorName = ""
     End If
End Function
'获取成品名称
Private Function GetProductName(ByVal color As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "Select B_Name From G_Product Where B_SID='" & color & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     If rs.RecordCount > 0 Then
        GetProductName = rs!B_name
     Else
        GetProductName = ""
     End If
End Function
'获取业务跟单的名称
Private Function getBusinessName(ByVal name As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "Select B_Name From G_Employee  where B_SID='" & name & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        getBusinessName = rs!B_name
    Else
        getBusinessName = ""
    End If
    
End Function

'删除行
Private Sub Deleterow()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If


    On Error GoTo IFERR
      If clspI.authenticate(theID) = False Then
            Exit Sub
      End If
    If rsdetail.RecordCount <= 0 Then
        Exit Sub
    End If
    
    If MsgBox("确定是否删除", vbYesNo + vbDefaultButton2 + vbInformation, "") = vbNo Then
        Exit Sub
    End If
    
    Dim sql As String
    sql = "delete from G_BillDetailOrder where B_itemid='" & rsdetail!B_ItemID & "'"
    Gm.cnnTool.cnn.Execute sql
    Dim sql1 As String
    sql1 = "delete from G_DraftBillDetailOrder where B_itemid='" & rsdetail!B_ItemID & "'"
    Gm.cnnTool.cnn.Execute sql1
    rsdetail.requery
    sumprice
    sumall
    Exit Sub
IFERR:
    
    MsgBox "请点击有数据的地方", vbOKOnly + vbInformation, "提示"
End Sub
'小计
Private Sub sumall()
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    a = 0
    b = 0
    c = 0
    d = 0
    If rsdetail.RecordCount <= 0 Then
      
        a = 0
        b = 0
        c = 0
        d = 0
    Else
        rsdetail.MoveFirst
        Do While Not rsdetail.EOF
            a = a + IIf(IsNull(rsdetail!B_BoxQty), 0, rsdetail!B_BoxQty)
            b = b + IIf(IsNull(rsdetail!B_Sum), 0, rsdetail!B_Sum)
            c = c + IIf(IsNull(rsdetail!B_QtyPerbox), 0, rsdetail!B_QtyPerbox)
            d = d + IIf(IsNull(rsdetail!B_qty), 0, rsdetail!B_qty)
            rsdetail.movenext
        Loop
        rsdetail.MoveFirst
    End If
    TDBGrid1.Columns("B_OrderCode").FooterText = "合计"
    TDBGrid1.Columns("B_BoxQty").FooterText = "" & a & ""
    TDBGrid1.Columns("B_sum").FooterText = "" & b & ""
    TDBGrid1.Columns("B_QtyPerbox").FooterText = "" & c & ""
    TDBGrid1.Columns("B_Qty").FooterText = "" & d & ""
End Sub
'网格右键
Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ActiveBar22.Bands("网格右键").PopupMenu
    End If
End Sub
Private Sub SetBillState(ByVal vState As Boolean)
    ActiveBar22.Bands("Band1").Tools("已保存").Visible = vState

    ActiveBar22.RecalcLayout
End Sub
Private Sub SetInvalidState(ByVal vState As Boolean)
    ActiveBar22.Bands("Band1").Tools("作废图片").Visible = vState

    ActiveBar22.RecalcLayout
End Sub
Private Sub SetAuditState(ByVal vState As Long)
        Dim sql3 As String
       
        If vState = 1 Then
            
            Dim sql As String
            sql = "update G_BillOrder set B_Audit=1 where B_ID='" & theID & "'"
            Gm.cnnTool.cnn.Execute sql
            vState = True
            ActiveBar22.Bands("Band1").Tools("审核").Enabled = False
            ActiveBar22.Bands("Band1").Tools("取消审核").Enabled = True
            Dim s As String
'            s = Format(Now, "YYYY-MM-DD")
'            sql3 = "update G_BillOrder B_DateAudit='" & s & "' B_ID='" & theid & "'"
'            Gm.cnnTool.cnn.Execute sql3
            
            
            CheckBox1.Enabled = True
            CheckBox2.Enabled = True
            CheckBox3.Enabled = True
        End If
        If vState = 0 Then
           Dim sql1 As String
            sql1 = "update G_BillOrder set B_Audit=0 where B_ID='" & theID & "'"
            Gm.cnnTool.cnn.Execute sql1
            C1Tab1.CurrTab = 0
            vState = False
            ActiveBar22.Bands("Band1").Tools("审核").Enabled = True
            ActiveBar22.Bands("Band1").Tools("取消审核").Enabled = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
        End If
        
        ActiveBar22.Bands("Band1").Tools("审核图片").Visible = vState
        If vState = True Then
            C1Tab1.TabVisible(1) = True
            C1Tab1.TabVisible(2) = True
'            C1Tab1.TabVisible(3) = True
            C1Tab1.TabVisible(4) = True
'            C1Tab1.TabVisible(5) = True
            C1Tab1.TabVisible(6) = True
            C1Tab1.TabVisible(7) = True
            C1Tab1.TabVisible(8) = True
            C1Tab1.TabVisible(9) = True
            C1Tab1.TabVisible(10) = True
            ActiveBar22.Bands("网格右键").Tools("新增行").Enabled = False
            ActiveBar22.Bands("网格右键").Tools("删除行").Enabled = False
            ActiveBar22.Bands("网格右键").Tools("复制行").Enabled = False
            ActiveBar22.Bands("网格右键").Tools("生成色布计划").Enabled = True
        End If
        If vState = False Then
            C1Tab1.TabVisible(1) = False
            C1Tab1.TabVisible(2) = False
'            C1Tab1.TabVisible(3) = False
            C1Tab1.TabVisible(4) = False
'            C1Tab1.TabVisible(5) = False
            C1Tab1.TabVisible(6) = False
            C1Tab1.TabVisible(7) = False
            C1Tab1.TabVisible(8) = False
            C1Tab1.TabVisible(9) = False
            C1Tab1.TabVisible(10) = False
            ActiveBar22.Bands("网格右键").Tools("新增行").Enabled = True
            ActiveBar22.Bands("网格右键").Tools("删除行").Enabled = True
            ActiveBar22.Bands("网格右键").Tools("复制行").Enabled = True
            ActiveBar22.Bands("网格右键").Tools("生成色布计划").Enabled = False
        End If
        ActiveBar22.RecalcLayout

End Sub
'复制行
Private Sub Copyrow()
    On Error GoTo IFERR
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    
    
      If clspI.authenticate(theID) = False Then
            Exit Sub
      End If
    
    If rsdetail.RecordCount <= 0 Then
        Exit Sub
    End If
     Dim frm1 As New frmOrderProduct_Edit
    frm1.FlatEdit1.Text = IIf(IsNull(rsdetail!B_ordercode), "", rsdetail!B_ordercode)
    frm1.FlatEdit2.Text = GetProductName(rsdetail!B_GoodsID)
    frm1.Productid = IIf(IsNull(rsdetail!B_GoodsID), "", rsdetail!B_GoodsID)
    frm1.FlatEdit3.Text = IIf(IsNull(rsdetail!B_Width), "", rsdetail!B_Width)
    frm1.FlatEdit4.Text = IIf(IsNull(rsdetail!B_weight), "", rsdetail!B_weight)
    frm1.colorid = IIf(IsNull(rsdetail!B_sid), "", rsdetail!B_sid)
    frm1.FlatEdit15.Text = IIf(IsNull(rsdetail!B_patterncode), "", rsdetail!B_patterncode)
    frm1.FlatEdit11.Text = IIf(IsNull(rsdetail!B_color), "", rsdetail!B_color)
    If IIf(IsNull(rsdetail!B_hex), "", rsdetail!B_hex) <> "" Then
        frm1.Picture2.BackColor = IIf(IsNull(rsdetail!B_hex), "", rsdetail!B_hex)
    End If
    frm1.FlatEdit5.Text = IIf(IsNull(rsdetail!B_Positivefabric), "", rsdetail!B_Positivefabric)
    frm1.FlatEdit6.Text = IIf(IsNull(rsdetail!B_Middlefabric), "", rsdetail!B_Middlefabric)
    frm1.FlatEdit14.Text = IIf(IsNull(rsdetail!B_Backfabric), "", rsdetail!B_Backfabric)
    frm1.FlatEdit7.Text = IIf(IsNull(rsdetail!B_qty), "", rsdetail!B_qty)
    frm1.FlatEdit8.Text = IIf(IsNull(rsdetail!B_BoxQty), "", rsdetail!B_BoxQty)
    frm1.FlatEdit12.Text = IIf(IsNull(rsdetail!B_QtyPerbox), "", rsdetail!B_QtyPerbox)
    frm1.FlatEdit13.Text = IIf(IsNull(rsdetail!B_MemoDetail), "", rsdetail!B_MemoDetail)
    frm1.FlatEdit9.Text = rsdetail!B_price
    frm1.FlatEdit10.Text = rsdetail!B_Sum
    frm1.ComboBox2.Text = rsdetail!B_GoodManual
    frm1.FlatEdit20.Text = rsdetail!B_process
    frm1.FlatEdit16.Text = rsdetail!B_Packaging
    frm1.FlatEdit17.Text = getClientName(rsdetail!B_PositiveFactory)
    frm1.Positiveid = rsdetail!B_PositiveFactory
    frm1.FlatEdit18.Text = getClientName(rsdetail!B_MiddleFactory)
    frm1.Middleid = rsdetail!B_MiddleFactory
    frm1.FlatEdit19.Text = getClientName(rsdetail!B_BackFactory)
    frm1.backid = rsdetail!B_BackFactory
    
    frm1.FlatEdit21.Text = rsdetail!B_Width2
    frm1.FlatEdit22.Text = rsdetail!B_Weight2
    frm1.FlatEdit23.Text = rsdetail!B_Width3
    frm1.FlatEdit24.Text = rsdetail!B_Weight3
    frm1.FlatEdit25.Text = rsdetail!B_seam
    frm1.FlatEdit26.Text = rsdetail!B_Beatbox
    frm1.FlatEdit27.Text = rsdetail!B_size
    frm1.FlatEdit28.Text = rsdetail!B_SourceOrderCode
    
    frm1.FlatEdit29.Text = rsdetail!B_KuanHao
    
    frm1.FlatEdit30.Text = IIf(IsNull(rsdetail!B_MianLiaoQty), "", rsdetail!B_MianLiaoQty)
    frm1.ComboBox1.Text = IIf(IsNull(rsdetail!B_ComputUnit), "", rsdetail!B_ComputUnit)
    frm1.FlatEdit32.Text = IIf(IsNull(rsdetail!B_MianLiaoPrice), "", rsdetail!B_MianLiaoPrice)
    frm1.FlatEdit33.Text = IIf(IsNull(rsdetail!B_HomeMill), "", rsdetail!B_HomeMill)
    frm1.FlatEdit31.Text = IIf(IsNull(rsdetail!B_ProcessQty), "", rsdetail!B_ProcessQty)
    frm1.FlatEdit34.Text = IIf(IsNull(rsdetail!B_ProcessPrice), "", rsdetail!B_ProcessPrice)
    frm1.FlatEdit35.Text = IIf(IsNull(rsdetail!B_ProcessMoney), "", rsdetail!B_ProcessMoney)
    

    
    frm1.client = theClientID
    frm1.id = theID
    frm1.Show vbModal
    rsdetail.requery
    SetGrid
    rsdetail.MoveFirst
    sumprice
    Exit Sub
IFERR:
     MsgBox "请点击有数据的地方", vbOKOnly + vbInformation, "提示"
End Sub

'复制多行
Private Sub CopyrowAll()
    Dim rs3 As RecordSet
    Dim sql3 As String
    Dim rs1 As New RecordSet
    Dim sql1 As String
    sql1 = "select *from G_BillOrder where B_id='" & theID & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs1.RecordCount > 0 Then
        Detail_save
    Else
        Dim tdbgRow As Variant
        For Each tdbgRow In TDBGrid1.SelBookmarks
            rsdetail.bookmark = tdbgRow
                Set rs3 = New RecordSet
                sql3 = "exec usp_savedetailProduct '" & theID & "','" & rsdetail!B_ordercode & "','" & rsdetail!B_GoodsID & "','" & rsdetail!B_Width & "','" & rsdetail!B_weight & "','" & rsdetail!B_patterncode & "','" & rsdetail!B_Positivefabric & "','" & rsdetail!B_Middlefabric & "','" & rsdetail!B_Backfabric & "','" & rsdetail!B_qty & "','" & rsdetail!B_BoxQty & "','" & rsdetail!B_QtyPerbox & "','" & rsdetail!B_price & "','" & rsdetail!B_Sum & "','" & rsdetail!B_MemoDetail & "','" & rsdetail!B_color & "','" & rsdetail!B_sid & "'"
                sql3 = sql3 & ",'" & rsdetail!B_GoodManual & "','" & rsdetail!B_PositiveFactory & "','" & rsdetail!B_MiddleFactory & "','" & rsdetail!B_BackFactory & "','" & rsdetail!B_process & "','" & rsdetail!B_Packaging & "'"
                sql3 = sql3 & ",'" & rsdetail!B_Width2 & "','" & rsdetail!B_Weight2 & "','" & rsdetail!B_Width3 & "','" & rsdetail!B_Weight3 & "','" & rsdetail!B_seam & "','" & rsdetail!B_Beatbox & "','" & rsdetail!B_size & "','" & Gm.SysID.SystemUser & "','1','" & rsdetail!B_SourceOrderCode & "'"
                sql3 = sql3 & ",'" & rsdetail!B_KuanHao & "','" & rsdetail!B_MianLiaoQty & "','" & rsdetail!B_ComputUnit & "','" & rsdetail!B_MianLiaoPrice & "','" & rsdetail!B_HomeMill & "','" & rsdetail!B_ProcessQty & "','" & rsdetail!B_ProcessPrice & "','" & rsdetail!B_ProcessMoney & "'"
                sql3 = sql3 & ",'" & rsdetail!B_PatternCode2 & "','" & rsdetail!B_ColorID2 & "','" & rsdetail!B_Color2 & "','" & rsdetail!B_MianLiaoQty2 & "','" & rsdetail!B_ComputUnit2 & "','" & rsdetail!B_MianLiaoPrice2 & "'"
                sql3 = sql3 & ",'" & rsdetail!B_PatternCode3 & "','" & rsdetail!B_ColorID3 & "','" & rsdetail!B_Color3 & "','" & rsdetail!B_MianLiaoQty3 & "','" & rsdetail!B_ComputUnit3 & "','" & rsdetail!B_MianLiaoPrice3 & "'"
                rs3.Open sql3, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        Next
    End If
    rsdetail.requery
End Sub
'明细表中有数据，进行删除
Private Sub Detail_save()
    Dim sql As String
    sql = "delete from G_DraftBillDetailOrder where B_ID='" & theID & "'"
    Gm.cnnTool.cnn.Execute sql
    
    Dim rs1 As New RecordSet
    Dim sql1 As String
    
    
    Dim tdbgRow As Variant
    For Each tdbgRow In TDBGrid1.SelBookmarks
        rsdetail.bookmark = tdbgRow
            Set rs1 = New RecordSet
            sql1 = "exec usp_savedetailProduct '" & theID & "','" & rsdetail!B_ordercode & "','" & rsdetail!B_GoodsID & "','" & rsdetail!B_Width & "','" & rsdetail!B_weight & "','" & rsdetail!B_patterncode & "','" & rsdetail!B_Positivefabric & "','" & rsdetail!B_Middlefabric & "','" & rsdetail!B_Backfabric & "','" & rsdetail!B_qty & "','" & rsdetail!B_BoxQty & "','" & rsdetail!B_QtyPerbox & "','" & rsdetail!B_price & "','" & rsdetail!B_Sum & "','" & rsdetail!B_MemoDetail & "','" & rsdetail!B_color & "','" & rsdetail!B_sid & "'"
            sql1 = sql1 & ",'" & rsdetail!B_GoodManual & "','" & rsdetail!B_PositiveFactory & "','" & rsdetail!B_MiddleFactory & "','" & rsdetail!B_BackFactory & "','" & rsdetail!B_process & "','" & rsdetail!B_Packaging & "'"
            sql1 = sql1 & ",'" & rsdetail!B_Width2 & "','" & rsdetail!B_Weight2 & "','" & rsdetail!B_Width3 & "','" & rsdetail!B_Weight3 & "','" & rsdetail!B_seam & "','" & rsdetail!B_Beatbox & "','" & rsdetail!B_size & "','" & Gm.SysID.SystemUser & "','1',''"
            sql1 = sql1 & ",'" & rsdetail!B_KuanHao & "','" & rsdetail!B_MianLiaoQty & "','" & rsdetail!B_ComputUnit & "','" & rsdetail!B_MianLiaoPrice & "','" & rsdetail!B_HomeMill & "','" & rsdetail!B_ProcessQty & "','" & rsdetail!B_ProcessPrice & "','" & rsdetail!B_ProcessMoney & "'"
            sql1 = sql1 & ",'" & rsdetail!B_PatternCode2 & "','" & rsdetail!B_ColorID2 & "','" & rsdetail!B_Color2 & "','" & rsdetail!B_MianLiaoQty2 & "','" & rsdetail!B_ComputUnit2 & "','" & rsdetail!B_MianLiaoPrice2 & "'"
            sql1 = sql1 & ",'" & rsdetail!B_PatternCode3 & "','" & rsdetail!B_ColorID3 & "','" & rsdetail!B_Color3 & "','" & rsdetail!B_MianLiaoQty3 & "','" & rsdetail!B_ComputUnit3 & "','" & rsdetail!B_MianLiaoPrice3 & "'"
            rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    Next

    
    Dim sql2 As String
    sql2 = "insert into G_BillDetailOrder (B_itemid,B_ID,B_OrderCode,B_GoodsID,B_Width,B_Weight,B_patterncode,B_Positivefabric,B_Middlefabric,B_Backfabric,B_Qty,B_BoxQty,B_QtyPerbox,B_Price,B_Sum,B_MemoDetail,B_color,B_colorid,B_GoodManual,B_PositiveFactory,B_MiddleFactory,B_BackFactory,B_Process,B_Packaging,B_Width2,B_Weight2,B_Width3,B_Weight3,B_seam,B_Beatbox,B_size,B_username,B_ContractLogodetail,B_SourceOrderCode,B_KuanHao,B_MianLiaoQty,B_ComputUnit,B_MianLiaoPrice,B_HomeMill,B_ProcessQty,B_ProcessPrice,B_ProcessMoney)"
    sql2 = sql2 & "   select B_itemid,B_ID,B_OrderCode,B_GoodsID,B_Width,B_Weight,B_patterncode,B_Positivefabric,B_Middlefabric,B_Backfabric,B_Qty,B_BoxQty,B_QtyPerbox,B_Price,B_Sum,B_MemoDetail,B_color,B_colorid,B_GoodManual,B_PositiveFactory,B_MiddleFactory,B_BackFactory,B_Process,B_Packaging,B_Width2,B_Weight2,B_Width3,B_Weight3,B_seam,B_Beatbox,B_size,B_username,B_ContractLogodetail,B_SourceOrderCode,B_KuanHao,B_MianLiaoQty,B_ComputUnit,B_MianLiaoPrice,B_HomeMill,B_ProcessQty,B_ProcessPrice,B_ProcessMoney from G_DraftBillDetailOrder  where B_ID='" & theID & "'"
    Gm.cnnTool.cnn.Execute sql2
    Dim sql3 As String
    sql3 = "delete from G_DraftBillDetailOrder where B_ID='" & theID & "'"
    Gm.cnnTool.cnn.Execute sql3
End Sub

Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, _
    ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, _
    ByVal CellStyle As TrueOleDBGrid80.StyleDisp)

    Dim ys As RGB
    On Error Resume Next
    Debug.Print TDBGrid1.Columns("B_Hex").CellValue(bookmark)
    CellStyle.BackColor = TDBGrid1.Columns("B_Hex").CellValue(bookmark)
     CellStyle.ForeColor = TDBGrid1.Columns("B_Hex").CellValue(bookmark)
End Sub

'---------------------------------------------------------------色布计划---------------------------------------------------------

Private Sub ActiveBar24_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
                Case "新增"
                
                    draftColordetail_null
                Case "删除"
                    DeleteColor
                Case "复制行"
                    copyColor
                Case "全部删除"
                    DeleteColorAll
                Case "打印流程卡"
                    card
                Case "打印全部流程卡"
                    cardAll
                Case "打印当前行染厂"
                    departprint
                Case "染厂派工单"
'                    depart

                    departprint
                    
                Case "打印当前行深加工"
                    processprint
                Case "深加工派工单"
                    process
                Case "保存样式"
                    setGridStyle2
                Case "打印单行"
                    PrintHang
                Case "色布采购"
                    colorplancast
                Case "生成白坯计划"
                     GrBPJHD
                
    End Select
End Sub

'生成白坯计划单
Private Sub GrBPJHD()

Dim sql3 As String
Dim rs3 As New RecordSet
   
   sql3 = "SELECT SUM(isnull(B_KG,0))AS B_BoxQty FROM G_BillDetailColor "
   sql3 = sql3 & " WHERE B_ItemIDB='" & rsdetailColor!B_ItemIDB & "' "
   sql3 = sql3 & " AND B_width='" & rsdetailColor!B_Width & "' AND B_weight='" & rsdetailColor!B_weight & "'"
   rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
   
'   Dim frm1 As New frmWhite_Edit
'       frm1.id = theid
'       frm1.FlatEdit6.Text = rs3!B_BoxQty
'    Debug.Print theid
'    frm1.Show vbModal
'   Unload frm1
   If yanzhenWhite(theID) = False Then
        Exit Sub
    End If
    Dim frm1 As New frmWhite_Edit
'    frm1.theidwhite = theidwhite
    frm1.id = theID
     frm1.FlatEdit6.Text = rs3!B_BoxQty
    Debug.Print theID
    frm1.Show vbModal
    Unload frm1
    
    
    rsdetailwhite.requery
    Dim sql As String
    Dim rs As New RecordSet
    sql = "select b.B_UserDes from G_BillWhite a left outer join G_systemUser b on a.B_UserName=b.B_UserName  where B_belongorderid='" & theID & "'"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs.RecordCount > 0 Then
        ActiveBar23.Bands("Band1").Tools("制单").Caption = "" & rs!B_UserDes & ""
    End If

End Sub

'色布计划中色布采购
Private Sub colorplancast()
    On Error Resume Next
    If TDBGrid3.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If ActiveBar22.Bands("Band1").Tools("审核图片").Visible = True Then
        Dim frm1 As New frmOrderProductColor_Edit
        frm1.colororderid = rsdetailColor!B_ItemID
        frm1.FlatEdit39.Text = rsdetailColor!B_ItemIDB
'        frm1.C1Tab1 = 1
        
        frm1.Show vbModal
        rsgrid6.requery
    Else
        MsgBox "此单没有审核", vbInformation, "提示"
        Exit Sub
    End If
    
End Sub
Private Sub TDBGrid3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ActiveBar24.Bands("网格右键").PopupMenu
    End If
End Sub
Private Sub setGridShow2()
    Dim cls1 As New clsGridShow
    With cls1
        .ObjectID = "11S005"
        .InitClass TDBGrid3, 3
        .ShowGridFormat
    End With
End Sub

Private Sub PrintHang()
    If TDBGrid3.ApproxCount <= 0 Then
        Exit Sub
    End If
    If rsdetailColor!B_ItemID = "" Then
        Exit Sub
    End If
    
    Dim sql As String
    Dim rs As New RecordSet
 
'    Dim a As String
'    Dim b As String
'    a = Format(DTPicker1.Value, "YYYY-MM-DD")
'    b = Format(DTPicker2.Value, "YYYY-MM-DD")
    sql = "exec usp_dingdandetailreport '','','" & rsdetailColor!B_ItemID & "','" & Gm.SysID.SystemUserName & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockReadOnly
    
    Dim frm1 As New frmModBLRPreviewOri
    Set frm1.RecordSet = rs.Clone
    frm1.ObjectID = "22B025"
    frm1.Show
    Unload frm1
End Sub
Private Sub copyColor()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    On Error GoTo IFERR
    If rsdetailColor.RecordCount <= 0 Then
        Exit Sub
    End If
    Dim rs  As RecordSet
    Set rs = New RecordSet
    Dim sql2 As String
    sql2 = "select * from G_BillDetailColor where B_itemid='" & rsdetailColor!B_ItemID & "' "
    rs.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    Dim frm1 As New frmOrderProductColor_Edit
'    frm1.itemid = rs!B_itemid
  
    frm1.ordertocolorid = IIf(IsNull(rs!B_orderitemid), 0, rs!B_orderitemid)
    If IIf(IsNull(rs!B_GroupID), "", rs!B_GroupID) <> "" Then
     frm1.lGroupID = rs!B_GroupID
    End If
    frm1.id = theID
    frm1.Label26.Caption = Val(rsdetailColor!rowIndex)
    frm1.FlatEdit3.Text = IIf(IsNull(rs!B_ItemIDB), "", rs!B_ItemIDB)
    frm1.FlatEdit2.Text = IIf(IsNull(rs!B_GoodsNameAlias), "", rs!B_GoodsNameAlias)
    frm1.FlatEdit4.Text = IIf(IsNull(rs!B_Width), "", rs!B_Width)
    frm1.FlatEdit10.Text = IIf(IsNull(rs!B_weight), "", rs!B_weight)
    frm1.colorid = IIf(IsNull(rsdetailColor!B_colorid), "", rsdetailColor!B_colorid)
    frm1.FlatEdit11.Text = IIf(IsNull(rsdetailColor!B_orderColor), "", rsdetailColor!B_orderColor)
    frm1.FlatEdit1.Text = IIf(IsNull(rsdetailColor!B_Producer), "", rsdetailColor!B_Producer)
    frm1.FlatEdit5.Text = IIf(IsNull(rs!B_SeHao), "", rs!B_SeHao)
'    If Trim(rs!B_Meter) > 0 Then
        frm1.FlatEdit6.Text = rsdetailColor!B_meter
'        frm1.FlatEdit6.Enabled = True
'         frm1.FlatEdit7.Enabled = False
'         frm1.FlatEdit7.BackColor = &HC0C0C0
'    End If
'    If Trim(rs!B_KG) > 0 Then
        frm1.FlatEdit7.Text = rsdetailColor!B_kg
'        frm1.FlatEdit6.Enabled = False
'         frm1.FlatEdit7.Enabled = True
'         frm1.FlatEdit6.BackColor = &HC0C0C0
'    End If
'    frm1.FlatEdit6.Text = IIf(IsNull(rs!B_meter), "", rs!B_meter)
'    frm1.FlatEdit7.Text = IIf(IsNull(rs!B_KG), "", rs!B_KG)
    frm1.FlatEdit8.Text = IIf(IsNull(rsdetailColor!B_depart), "", rsdetailColor!B_depart)
    frm1.departid = IIf(IsNull(rsdetailColor!B_departid), "", rsdetailColor!B_departid)
    frm1.FlatEdit9.Text = IIf(IsNull(rsdetailColor!B_department), "", rsdetailColor!B_department)
    frm1.departmentid = IIf(IsNull(rsdetailColor!B_departmentid), "", rsdetailColor!B_departmentid)
    frm1.FlatEdit18.Text = IIf(IsNull(rs!B_phone4), "", rs!B_phone4)
    If IIf(IsNull(rsdetailColor!B_departdate), "", rsdetailColor!B_departdate) = "" Then
        frm1.DTPicker1.Value = Now
    Else
        frm1.DTPicker1.Value = rsdetailColor!B_departdate
    End If
'    frm1.DTPicker1.Value = IIf(IsNull(rsdetailColor!B_departdate), "", rsdetailColor!B_departdate)
'    frm1.FlatEdit12.Text = IIf(IsNull(rs!B_flowCard), "", rs!B_flowCard)
    frm1.FlatEdit13.Text = IIf(IsNull(rs!B_departdannote), "", rs!B_departdannote)
    
    frm1.FlatEdit14.Text = IIf(IsNull(rsdetailColor!B_processunit), "", rsdetailColor!B_processunit)
    frm1.processid = IIf(IsNull(rsdetailColor!B_processunitid), "", rsdetailColor!B_processunitid)
    frm1.FlatEdit15.Text = IIf(IsNull(rsdetailColor!B_processdocumentary), "", rsdetailColor!B_processdocumentary)
    frm1.processmentid = IIf(IsNull(rsdetailColor!B_processdocumentaryid), "", rsdetailColor!B_processdocumentaryid)
    
    frm1.FlatEdit22.Text = IIf(IsNull(rs!B_phone1), "", rs!B_phone1)
   If IIf(IsNull(rsdetailColor!B_processdate), "", rsdetailColor!B_processdate) = "" Then
        frm1.DTPicker2.Value = Now
    Else
        frm1.DTPicker2.Value = rsdetailColor!B_processdate
    End If
    frm1.FlatEdit16.Text = IIf(IsNull(rs!B_processcost), "", rs!B_processcost)
'    frm1.DTPicker2.Value = IIf(IsNull(rs!B_processdate), "", rs!B_processdate)
    frm1.FlatEdit17.Text = IIf(IsNull(rsdetailColor!B_processnote), "", rsdetailColor!B_processnote)
    
    frm1.FlatEdit25.Text = IIf(IsNull(rsdetailColor!B_processunit2), "", rsdetailColor!B_processunit2)
    frm1.processid2 = IIf(IsNull(rsdetailColor!B_processunitid2), "", rsdetailColor!B_processunitid2)
    frm1.FlatEdit24.Text = IIf(IsNull(rsdetailColor!B_processdocumentary2), "", rsdetailColor!B_processdocumentary2)
    frm1.processmentid2 = IIf(IsNull(rsdetailColor!B_processdocumentaryid2), "", rsdetailColor!B_processdocumentaryid2)
    frm1.FlatEdit23.Text = IIf(IsNull(rsdetailColor!B_phone2), "", rsdetailColor!B_phone2)
   If IIf(IsNull(rs!B_processdate2), "", rs!B_processdate2) = "" Then
        frm1.DTPicker3.Value = Now
    Else
        frm1.DTPicker3.Value = rs!B_processdate2
    End If
    frm1.FlatEdit27.Text = IIf(IsNull(rs!B_processCost2), "", rs!B_processCost2)
'    frm1.DTPicker2.Value = IIf(IsNull(rs!B_processdate), "", rs!B_processdate)
    frm1.FlatEdit26.Text = IIf(IsNull(rs!B_processnote2), "", rs!B_processnote2)
    
    frm1.FlatEdit30.Text = IIf(IsNull(rsdetailColor!B_processunit3), "", rsdetailColor!B_processunit3)
    frm1.processid3 = IIf(IsNull(rsdetailColor!B_processunitid3), "", rsdetailColor!B_processunitid3)
    frm1.FlatEdit29.Text = IIf(IsNull(rsdetailColor!B_processdocumentary3), "", rsdetailColor!B_processdocumentary3)
    frm1.processmentid3 = IIf(IsNull(rsdetailColor!B_processdocumentaryid3), "", rsdetailColor!B_processdocumentaryid3)
    frm1.FlatEdit28.Text = IIf(IsNull(rs!B_phone3), "", rs!B_phone3)
   If IIf(IsNull(rs!B_processdate3), "", rs!B_processdate3) = "" Then
        frm1.DTPicker4.Value = Now
    Else
        frm1.DTPicker4.Value = rs!B_processdate3
    End If
    frm1.FlatEdit31.Text = IIf(IsNull(rs!B_processCost3), "", rs!B_processCost3)
'    frm1.DTPicker2.Value = IIf(IsNull(rs!B_processdate), "", rs!B_processdate)
    frm1.FlatEdit32.Text = IIf(IsNull(rs!B_processnote3), "", rs!B_processnote3)
    
    
     If IIf(IsNull(rsdetailColor!B_Progressprocess), "", rsdetailColor!B_Progressprocess) = "" Then
            frm1.ComboBox5.Text = ""
     Else
        frm1.ComboBox5.Text = GetProgressCraftCT(rs!B_Progressprocess)
     End If
'    frm1.ComboBox5.Text = GetProgressCraftCT(rs!B_Progressprocess)
    frm1.FlatEdit19.Text = IIf(IsNull(rs!B_Paper), "", rs!B_Paper)
    frm1.FlatEdit20.Text = IIf(IsNull(rs!B_pocket), "", rs!B_pocket)
    frm1.FlatEdit21.Text = IIf(IsNull(rs!B_Empty), "", rs!B_Empty)
    If IIf(IsNull(rsdetailColor!B_packagstyle), "", rsdetailColor!B_packagstyle) = "" Then
            frm1.ComboBox4.Text = ""
     Else
        frm1.ComboBox4.Text = GetB_packagstyle(rsdetailColor!B_packagstyle)
     End If
'    frm1.ComboBox4.Text = GetB_packagstyle(rs!B_packagstyle)
    frm1.FlatEdit12.Text = IIf(IsNull(rs!B_departCost), "", rs!B_departCost)
    frm1.FlatEdit33.Text = IIf(IsNull(rs!B_fold), "", rs!B_fold)
    frm1.FlatEdit34.Text = IIf(IsNull(rs!B_Cast), "", rs!B_Cast)
    frm1.FlatEdit35.Text = IIf(IsNull(rs!B_PracticeCast), "", rs!B_PracticeCast)
    frm1.ComboBox1.Text = IIf(IsNull(rs!B_LabelTemplate), "", rs!B_LabelTemplate)
    frm1.ComboBox2.Text = IIf(IsNull(rs!B_DetailTemplate), "", rs!B_DetailTemplate)
    
    frm1.FlatEdit36.Text = IIf(IsNull(rs!B_DepartColor), "", rs!B_DepartColor)
    frm1.Show vbModal
    Unload frm1
    rsdetailColor.requery
    Exit Sub
IFERR:
    MsgBox " 点有数据的地方", vbOKOnly + vbInformation, "提示"
End Sub

'保存网格样式
Private Sub setGridStyle2()
    Dim i As Long
    Dim strSQL As String
    Dim dWidth As Integer
    Dim szFieldName As String
    
    For i = 0 To TDBGrid3.Columns.Count - 1
        If TDBGrid3.Columns(i).width > 0 Then
            If TDBGrid3.Columns(i).Visible = True Then
                szFieldName = TDBGrid3.Columns(i).DataField
                dWidth = TDBGrid3.Columns(i).width
                strSQL = "update G_BLSField set B_GridWidth='" & dWidth & "' where B_ObjectID='11S005' and B_FieldName='" & szFieldName & "'"
                Gm.cnnTool.cnn.Execute strSQL
            End If
        End If
    Next
    
End Sub

Private Sub CopyToColor()
        If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
            MsgBox "此单已被作废", vbInformation, "提示"
            Exit Sub
        End If


        Set clspI = New clspI
'        If clspI.authenticate(theid) = False Then
'            Exit Sub
'        End If
    
        If rsdetail.RecordCount <= 0 Then
            Exit Sub
        End If
        Dim rs As New RecordSet
        Dim sql As String
        sql = "select * from G_BillColor where B_Belongorderid='" & theID & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
'        If rs.RecordCount > 0 Then
'            MsgBox "色布计划已经存在", vbInformation, "提示"
'            Exit Sub
'        End If

        If TDBGrid3.ApproxCount <= 0 Then
            savecolormain
        Else
            theidColor = rs!B_id
        End If
        SaveColorDetail
        
        MsgBox "复制成功", vbInformation, "提示"
        ActiveBar24.Bands("Band1").Tools("制单人").Caption = Gm.SysID.SystemUserName
'        Debug.Print ActiveBar24.Bands("Band1").Tools("制单人").Caption

        ActiveBar24.RecalcLayout
        C1Tab1.CurrTab = 1
       pattern
End Sub
Private Sub CopyToColorAll()
        If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
            MsgBox "此单已被作废", vbInformation, "提示"
            Exit Sub
        End If
        Set clspI = New clspI
        If rsdetail.RecordCount <= 0 Then
            Exit Sub
        End If
'        Dim rs As New RecordSet
'        Dim sql As String
'        sql = "select * from G_BillColor where B_Belongorderid='" & theid & "'"
'        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
  If rsdetail.RecordCount > 0 Then
            rsdetail.MoveFirst
        
    Do While Not rsdetail.EOF
    
        Dim rs As New RecordSet
        Dim sql As String
        sql = "select * from G_BillColor where B_Belongorderid='" & theID & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        If TDBGrid3.ApproxCount <= 0 Then
            savecolormain
        Else
            theidColor = rs!B_id
        End If
        
        SaveColorDetail
        rs.Clone
        Set rs = Nothing
        rsdetail.movenext
     Loop
     
End If
        
        MsgBox "复制成功", vbInformation, "提示"
        ActiveBar24.Bands("Band1").Tools("制单人").Caption = Gm.SysID.SystemUserName
'        Debug.Print ActiveBar24.Bands("Band1").Tools("制单人").Caption

        ActiveBar24.RecalcLayout
        C1Tab1.CurrTab = 1
       pattern
End Sub
Private Sub savecolormain()
     Set clsBL = New clsBL
    Dim sql As String
            Dim rs As New RecordSet
            sql = "select * from G_DraftBillColor where 1=1 "
            rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            rs.AddNew
            Dim a As String
            a = Format(Now, "YYYY-MM-DD")
            rs!B_datecreate = a
            rs.Update
            theidColor = rs!B_id
            
            Dim rs1 As New RecordSet
            Dim sql1 As String
            sql1 = "select * from G_BillColor where 1=1 "
            rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            rs1.AddNew
               rs1!B_id = theidColor
               Dim b As String
               b = Format(Now, "YYYY-MM-DD")
               rs1!B_datecreate = b
               rs1!B_username = Gm.SysID.SystemUser
               rs1!B_BID = B_BID_CC
               rs1!B_ObjectID = B_ObjectID_CC
               rs1!B_BillType = B_BillType_CC
               rs1!B_Codeid = clsBL.GetFrameCodeDetail_01(B_ObjectID_CC)
               rs1!B_BelongOrderID = theID
               rs1.Update
               Dim rs2 As New RecordSet
               Dim sql2 As String
               sql2 = "delete from G_DraftBillColor where B_ID='" & theidColor & "'"
               Gm.cnnTool.cnn.Execute sql2
End Sub

'从表G_BillDetailColor获取当前最新一个条码的自增数字
Private Function GetNewBCIncr() As Long
    Dim rs As New RecordSet
    strSQL = "select top 1 * from G_BillDetailColor order by B_BCIncr desc"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Dim lRturn As Long
    If rs.RecordCount <= 0 Then
        lRturn = 1
    Else
        lRturn = IIf(IsNull(rs!B_BCIncr), 0, rs!B_BCIncr) + 1
    End If
    
    rs.Close
    Set rs = Nothing
    
    GetNewBCIncr = lRturn
End Function

'传入参数：任意长度的自增数字的字符串类型
'返回值：返回BC13条码的前面12位字符
Private Function FillGetBC12(ByVal vIncr As String) As String
    Dim cls1 As New clsString
    Dim szReturn As String
    
    szReturn = cls1.FillRepeat(vIncr, 11, "0", True)
    szReturn = COLORBC13FIRST & szReturn
    
    FillGetBC12 = szReturn
End Function

Private Function GetBC13(ByVal vBC12 As String) As String
    Dim szRturn As String
    szRturn = GetEAN13CheckOut(vBC12)
    
    GetBC13 = vBC12 & szRturn
End Function

'获取最新的一个13位条码
Private Function GetBC13Ex() As String
    Dim szIncr As String
    szIncr = GetNewBCIncr
    
    Dim szBC12 As String
    szBC12 = FillGetBC12(GetNewBCIncr)
    
    GetBC13Ex = GetBC13(szBC12)
End Function
'生成色布计划
Private Sub SaveColorDetail()
    Dim itemid As String
    Dim rs As RecordSet
    Dim rs1 As RecordSet
    Dim sql As String
    Dim sql1 As String
    
    Dim lIncr As Long
    Dim szBC13 As String
    Dim i As Long, lGroupID As Long
    Dim szOrderCode As String
    i = 1
    lGroupID = 1
    
'    rsdetail.MoveFirst
'    Do While Not rsdetail.EOF
        
        '============================
        '更换订单号则初始化行计数器和分组号
        If szOrderCode <> rsdetail!B_ordercode Then
            If Len(szOrderCode) > 0 Then
                i = 1
                lGroupID = 1
            End If
        End If
        szOrderCode = rsdetail!B_ordercode
        '============================
        
        
        Set rs = New RecordSet
        sql = "select * from G_DraftBillDetailColor where 1=0"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        rs.AddNew
        rs!B_datecreate = Now
        rs.Update
        itemid = rs!B_ItemID
        
        Set rs1 = New RecordSet
        sql1 = "select * from G_BillDetailColor where 1=1"
        rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        '          rs1.AddNew
        Dim sql2 As String
         
         
        '获取最新的一个条码的自增数字
        lIncr = GetNewBCIncr
        szBC13 = GetBC13(FillGetBC12(lIncr))
        
        sql2 = "exec usp_savetoColor '" & theidColor & "','" & itemid & "','" & rsdetail!B_ordercode & "','" & rsdetail!B_GoodsID & "'"
        sql2 = sql2 & ",'" & rsdetail!B_Width & "','" & rsdetail!B_weight & "','" & rsdetail!B_sid & "','','" & rsdetail!B_patterncode & "','','','','" & lIncr & "','" & szBC13 & "','" & lGroupID & "','" & rsdetail!B_color & "','" & rsdetail!B_ItemID & "'"
        Debug.Print sql2
        Gm.cnnTool.cnn.Execute sql2
        
        '          rs1.Update
        Dim sql3 As String
        sql3 = "delete from G_DraftBillDetailColor where B_itemid='" & itemid & "'"
        Gm.cnnTool.cnn.Execute sql3
        
        
        '============================
        '达到指定行数则初始化计数器，分组号+1
        If i = 4 Then
            i = 1
            lGroupID = lGroupID + 1
        End If
        i = i + 1
        '============================
        
        
'        rsdetail.movenext
'    Loop
    rsdetailColor.requery
    If rsdetailColor.RecordCount > 0 Then
        If Len(rsdetailColor!B_colorid) > 0 Then
        TDBGrid3.Columns("B_Hex").FetchStyle = True
        End If
    End If
    'rsDetail.requery
'    rsdetail.MoveFirst
    ActiveBar24.RecalcLayout
End Sub

'获取本合同下的色布计划
Private Sub Colordetail()
    Set rsdetailColor = New RecordSet
    Dim sql As String
    sql = "exec usp_SelectColor '" & theID & "'"
    Debug.Print sql
    rsdetailColor.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    TDBGrid3.DataSource = rsdetailColor
    
    Dim sql1 As String
    Dim rs As New RecordSet
    sql1 = "select b.B_UserDes from G_BillColor  a left outer join  G_SystemUser  b on a.B_username=b.B_username where B_belongorderid='" & theID & "'"
    Debug.Print sql1
    rs.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Dim szUserName As String
    If rs.RecordCount > 0 Then
        szUserName = IIf(IsNull(rs!B_UserDes), "", rs!B_UserDes)
    End If
    
    If C1Tab1.TabVisible(1) = True Then
        ActiveBar24.Bands("Band1").Tools("制单人").Caption = "" & szUserName & ""
        ActiveBar24.RecalcLayout
    End If
    setColor
    TDBGrid3.Columns("B_Hex").FetchStyle = True
End Sub
'色布计划
Private Sub setColor()
    setGridShow2
'    TDBGrid3.Columns("B_ItemIDB").Caption = "订单号"
'    TDBGrid3.Columns("B_GoodsNameAlias").Caption = "品名"
'    TDBGrid3.Columns("B_width").Caption = "门幅"
'    TDBGrid3.Columns("B_weight").Caption = "克重"
'    TDBGrid3.Columns("B_Color").Caption = "颜色"
'    TDBGrid3.Columns("B_Hex").Caption = "颜色标识"
'    TDBGrid3.Columns("B_Producer").Caption = "花型"
'    TDBGrid3.Columns("B_SeHao").Caption = "色号"
'    TDBGrid3.Columns("B_meter").Caption = "米数"
'    TDBGrid3.Columns("B_KG").Caption = "公斤数"
'    TDBGrid3.Columns("B_depart").Caption = "染厂"
'    TDBGrid3.Columns("B_Department").Caption = "染厂跟单"
'    TDBGrid3.Columns("B_departdate").Caption = "染厂交期"
'    TDBGrid3.Columns("B_departCost").Caption = "染厂加工费"
'    TDBGrid3.Columns("B_departdannote").Caption = "染厂备注"
'    TDBGrid3.Columns("B_flowCardprint").Caption = "流程卡打印次数"
'    TDBGrid3.Columns("B_departdanprint").Caption = "染厂打印次数"
'    TDBGrid3.Columns("B_processunit").Caption = "深加工单位"
'    TDBGrid3.Columns("B_processdocumentary").Caption = "深加工跟单"
'    TDBGrid3.Columns("B_processdate").Caption = "深加工交期"
'    TDBGrid3.Columns("B_processCost").Caption = "深加工加工费"
'    TDBGrid3.Columns("B_processnote").Caption = "深加工备注"
'    TDBGrid3.Columns("B_processprint").Caption = "深加工派工单打印次数"
'    TDBGrid3.Columns("B_Progressprocess").Caption = "进度工序"
'    TDBGrid3.Columns("B_Paper").Caption = "纸管"
'    TDBGrid3.Columns("B_pocket").Caption = "袋重"
'    TDBGrid3.Columns("B_Empty").Caption = "空加"
'    TDBGrid3.Columns("B_packagstyle").Caption = "包装方式"
'    TDBGrid3.Columns("B_Paper").NumberFormat = "0.00"
'    TDBGrid3.Columns("B_Empty").NumberFormat = "0.00"
'    TDBGrid3.Columns("B_pocket").NumberFormat = "0.00"
'
    TDBGrid3.Columns("B_Department").width = 0
    TDBGrid3.Columns("B_Department").Visible = False
    TDBGrid3.Columns("B_Department").AllowSizing = False
    TDBGrid3.Columns("B_departdate").width = 0
    TDBGrid3.Columns("B_departdate").Visible = False
    TDBGrid3.Columns("B_departdate").AllowSizing = False
    TDBGrid3.Columns("B_departCost").width = 0
    TDBGrid3.Columns("B_departCost").Visible = False
    TDBGrid3.Columns("B_departCost").AllowSizing = False
    TDBGrid3.Columns("B_departdannote").width = 0
    TDBGrid3.Columns("B_departdannote").Visible = False
    TDBGrid3.Columns("B_departdannote").AllowSizing = False
    TDBGrid3.Columns("B_processunit").width = 0
    TDBGrid3.Columns("B_processunit").Visible = False
    TDBGrid3.Columns("B_processunit").AllowSizing = False

    TDBGrid3.Columns("B_processdocumentary").width = 0
    TDBGrid3.Columns("B_processdocumentary").Visible = False
    TDBGrid3.Columns("B_processdocumentary").AllowSizing = False
    TDBGrid3.Columns("B_processdate").width = 0
    TDBGrid3.Columns("B_processdate").Visible = False
    TDBGrid3.Columns("B_processdate").AllowSizing = False
    TDBGrid3.Columns("B_processCost").width = 0
    TDBGrid3.Columns("B_processCost").Visible = False
    TDBGrid3.Columns("B_processCost").AllowSizing = False
        TDBGrid3.Columns("B_processnote").width = 0
    TDBGrid3.Columns("B_processnote").Visible = False
    TDBGrid3.Columns("B_processnote").AllowSizing = False
        TDBGrid3.Columns("B_Progressprocess").width = 0
    TDBGrid3.Columns("B_Progressprocess").Visible = False
    TDBGrid3.Columns("B_Progressprocess").AllowSizing = False

    TDBGrid3.Columns("RowIndex").width = 0
    TDBGrid3.Columns("RowIndex").Visible = False
     TDBGrid3.Columns("RowIndex").AllowSizing = False
    TDBGrid3.Columns("B_ItemID").width = 0
    TDBGrid3.Columns("B_ItemID").Visible = False
     TDBGrid3.Columns("B_ItemID").AllowSizing = False
    TDBGrid3.Columns("B_Colorid").width = 0
    TDBGrid3.Columns("B_Colorid").Visible = False
    TDBGrid3.Columns("B_Colorid").AllowSizing = False
'    TDBGrid3.Columns("B_departid").width = 0
'    TDBGrid3.Columns("B_departid").Visible = False
'    TDBGrid3.Columns("B_departid").AllowSizing = False
    TDBGrid3.Columns("B_Departmentid").width = 0
    TDBGrid3.Columns("B_Departmentid").Visible = False
    TDBGrid3.Columns("B_Departmentid").AllowSizing = False
    TDBGrid3.Columns("B_processunitid").width = 0
    TDBGrid3.Columns("B_processunitid").Visible = False
    TDBGrid3.Columns("B_processunitid").AllowSizing = False
    TDBGrid3.Columns("B_processdocumentaryid").width = 0
    TDBGrid3.Columns("B_processdocumentaryid").Visible = False
    TDBGrid3.Columns("B_processdocumentaryid").AllowSizing = False
    TDBGrid3.Columns("B_GoodsNameAlias").width = 1600
    TDBGrid3.Columns("B_width").width = 1000
    TDBGrid3.Columns("B_weight").width = 1000
    TDBGrid3.Columns("B_orderColor").width = 1500
     TDBGrid3.Columns("B_Hex").width = 1000
    TDBGrid3.Columns("B_Producer").width = 1000
    TDBGrid3.Columns("B_SeHao").width = 1000
    TDBGrid3.Columns("B_meter").width = 1000
    TDBGrid3.Columns("B_KG").width = 1000
    TDBGrid3.Columns("B_depart").width = 1300
    TDBGrid3.Columns("B_department").width = 1500
    TDBGrid3.Columns("B_departdate").width = 1500
    TDBGrid3.Columns("B_flowCardprint").width = 1200
    TDBGrid3.Columns("B_departdanprint").width = 1200
     TDBGrid3.Columns("B_processprint").width = 1300
    TDBGrid3.Columns("B_Paper").width = 800
    TDBGrid3.Columns("B_pocket").width = 800
    TDBGrid3.Columns("B_Empty").width = 800
    TDBGrid3.Columns("B_packagstyle").width = 1200
    setgridView
    TDBGrid3.MarqueeStyle = dbgHighlightRow
'    If rsdetailColor.RecordCount > 0 Then
'        If Len(rsdetailColor!B_Colorid) > 0 Then
'
'        TDBGrid3.Columns("B_Hex").FetchStyle = True
'        End If
'    End If
'    TDBGrid3.HoldFields
End Sub

Private Sub setgridView()
        TDBGrid3.Columns("B_LabelTemplate").width = 0
    TDBGrid3.Columns("B_LabelTemplate").Visible = False
    TDBGrid3.Columns("B_LabelTemplate").AllowSizing = False
        TDBGrid3.Columns("B_DetailTemplate").width = 0
    TDBGrid3.Columns("B_DetailTemplate").Visible = False
    TDBGrid3.Columns("B_DetailTemplate").AllowSizing = False

    TDBGrid3.Columns("B_fold").width = 0
    TDBGrid3.Columns("B_fold").Visible = False
    TDBGrid3.Columns("B_fold").AllowSizing = False
    TDBGrid3.Columns("B_Cast").width = 0
    TDBGrid3.Columns("B_Cast").Visible = False
    TDBGrid3.Columns("B_Cast").AllowSizing = False
    TDBGrid3.Columns("B_PracticeCast").width = 0
    TDBGrid3.Columns("B_PracticeCast").Visible = False
    TDBGrid3.Columns("B_PracticeCast").AllowSizing = False

    TDBGrid3.Columns("B_processunitid2").width = 0
    TDBGrid3.Columns("B_processunitid2").Visible = False
    TDBGrid3.Columns("B_processunitid2").AllowSizing = False
    TDBGrid3.Columns("B_processdocumentaryid2").width = 0
    TDBGrid3.Columns("B_processdocumentaryid2").Visible = False
    TDBGrid3.Columns("B_processdocumentaryid2").AllowSizing = False
       TDBGrid3.Columns("B_processunitid3").width = 0
    TDBGrid3.Columns("B_processunitid3").Visible = False
    TDBGrid3.Columns("B_processunitid3").AllowSizing = False
    TDBGrid3.Columns("B_processdocumentaryid3").width = 0
    TDBGrid3.Columns("B_processdocumentaryid3").Visible = False
    TDBGrid3.Columns("B_processdocumentaryid3").AllowSizing = False

        TDBGrid3.Columns("B_processunit2").width = 0
    TDBGrid3.Columns("B_processunit2").Visible = False
    TDBGrid3.Columns("B_processunit2").AllowSizing = False

    TDBGrid3.Columns("B_processdocumentary2").width = 0
    TDBGrid3.Columns("B_processdocumentary2").Visible = False
    TDBGrid3.Columns("B_processdocumentary2").AllowSizing = False
    TDBGrid3.Columns("B_processdate2").width = 0
    TDBGrid3.Columns("B_processdate2").Visible = False
    TDBGrid3.Columns("B_processdate2").AllowSizing = False
    TDBGrid3.Columns("B_processCost2").width = 0
    TDBGrid3.Columns("B_processCost2").Visible = False
    TDBGrid3.Columns("B_processCost2").AllowSizing = False
        TDBGrid3.Columns("B_processnote2").width = 0
    TDBGrid3.Columns("B_processnote2").Visible = False
    TDBGrid3.Columns("B_processnote2").AllowSizing = False

        TDBGrid3.Columns("B_processunit3").width = 0
    TDBGrid3.Columns("B_processunit3").Visible = False
    TDBGrid3.Columns("B_processunit3").AllowSizing = False

    TDBGrid3.Columns("B_processdocumentary3").width = 0
    TDBGrid3.Columns("B_processdocumentary3").Visible = False
    TDBGrid3.Columns("B_processdocumentary3").AllowSizing = False
    TDBGrid3.Columns("B_processdate").width = 0
    TDBGrid3.Columns("B_processdate3").Visible = False
    TDBGrid3.Columns("B_processdate3").AllowSizing = False
    TDBGrid3.Columns("B_processCost3").width = 0
    TDBGrid3.Columns("B_processCost3").Visible = False
    TDBGrid3.Columns("B_processCost3").AllowSizing = False
        TDBGrid3.Columns("B_processnote3").width = 0
    TDBGrid3.Columns("B_processnote3").Visible = False
    TDBGrid3.Columns("B_processnote3").AllowSizing = False
            TDBGrid3.Columns("B_phone1").width = 0
    TDBGrid3.Columns("B_phone1").Visible = False
    TDBGrid3.Columns("B_phone1").AllowSizing = False
                TDBGrid3.Columns("B_phone2").width = 0
    TDBGrid3.Columns("B_phone2").Visible = False
    TDBGrid3.Columns("B_phone2").AllowSizing = False
                TDBGrid3.Columns("B_phone3").width = 0
    TDBGrid3.Columns("B_phone3").Visible = False
    TDBGrid3.Columns("B_phone3").AllowSizing = False
                TDBGrid3.Columns("B_phone4").width = 0
    TDBGrid3.Columns("B_phone4").Visible = False
    TDBGrid3.Columns("B_phone4").AllowSizing = False
End Sub
Private Sub TDBGrid3_FetchCellStyle(ByVal Condition As Integer, _
    ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, _
    ByVal CellStyle As TrueOleDBGrid80.StyleDisp)
    On Error Resume Next
    Dim ys As RGB
    CellStyle.BackColor = TDBGrid3.Columns("B_Hex").CellValue(bookmark)
    CellStyle.ForeColor = TDBGrid3.Columns("B_Hex").CellValue(bookmark)
End Sub
Private Sub TDBGrid3_DblClick()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If rsdetailColor.RecordCount > 0 Then
       Colordetail_UPdate
    Else
        draftColordetail_null
    End If
End Sub

Private Sub Colordetail_UPdate()
     On Error GoTo IFERR
    If rsdetailColor.RecordCount <= 0 Then
        Exit Sub
    End If
    Dim rs  As RecordSet
    Set rs = New RecordSet
    Dim sql2 As String
    sql2 = "select * from G_BillDetailColor where B_itemid='" & rsdetailColor!B_ItemID & "' "
    rs.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Dim frm1 As New frmOrderProductColor_Edit
    
    frm1.colororderid = rsdetailColor!B_ItemID
    frm1.FlatEdit39.Text = rsdetailColor!B_ItemIDB
    frm1.itemid = rs!B_ItemID
    frm1.id = theID
    frm1.Label26.Caption = Val(rsdetailColor!rowIndex)
    frm1.FlatEdit3.Text = IIf(IsNull(rs!B_ItemIDB), "", rs!B_ItemIDB)
    frm1.FlatEdit2.Text = IIf(IsNull(rs!B_GoodsNameAlias), "", rs!B_GoodsNameAlias)
    frm1.FlatEdit4.Text = IIf(IsNull(rs!B_Width), "", rs!B_Width)
    frm1.FlatEdit10.Text = IIf(IsNull(rs!B_weight), "", rs!B_weight)
    frm1.colorid = IIf(IsNull(rsdetailColor!B_colorid), "", rsdetailColor!B_colorid)
    frm1.FlatEdit11.Text = IIf(IsNull(rsdetailColor!B_orderColor), "", rsdetailColor!B_orderColor)
    frm1.FlatEdit1.Text = IIf(IsNull(rsdetailColor!B_Producer), "", rsdetailColor!B_Producer)
    frm1.FlatEdit5.Text = IIf(IsNull(rs!B_SeHao), "", rs!B_SeHao)
    frm1.FlatEdit37.Text = IIf(IsNull(rs!B_BoxQty), "", rs!B_BoxQty)
'    If Trim(rs!B_Meter) > 0 Then
        frm1.FlatEdit6.Text = rsdetailColor!B_meter
'        frm1.FlatEdit6.Enabled = True
'         frm1.FlatEdit7.Enabled = False
'         frm1.FlatEdit7.BackColor = &HC0C0C0
'    End If
'    If Trim(rs!B_KG) > 0 Then
        frm1.FlatEdit7.Text = rsdetailColor!B_kg
'        frm1.FlatEdit6.Enabled = False
'         frm1.FlatEdit7.Enabled = True
'         frm1.FlatEdit6.BackColor = &HC0C0C0
'    End If
'    frm1.FlatEdit6.Text = IIf(IsNull(rs!B_meter), "", rs!B_meter)
'    frm1.FlatEdit7.Text = IIf(IsNull(rs!B_KG), "", rs!B_KG)
    frm1.FlatEdit8.Text = IIf(IsNull(rsdetailColor!B_depart), "", rsdetailColor!B_depart)
    frm1.departid = IIf(IsNull(rsdetailColor!B_departid), "", rsdetailColor!B_departid)
    frm1.FlatEdit9.Text = IIf(IsNull(rsdetailColor!B_department), "", rsdetailColor!B_department)
    frm1.departmentid = IIf(IsNull(rsdetailColor!B_departmentid), "", rsdetailColor!B_departmentid)
    frm1.FlatEdit18.Text = IIf(IsNull(rs!B_phone4), "", rs!B_phone4)
    If IIf(IsNull(rsdetailColor!B_departdate), "", rsdetailColor!B_departdate) = "" Then
        frm1.DTPicker1.Value = Now
    Else
        frm1.DTPicker1.Value = rsdetailColor!B_departdate
    End If
'    frm1.DTPicker1.Value = IIf(IsNull(rsdetailColor!B_departdate), "", rsdetailColor!B_departdate)
'    frm1.FlatEdit12.Text = IIf(IsNull(rs!B_flowCard), "", rs!B_flowCard)
    frm1.FlatEdit13.Text = IIf(IsNull(rs!B_departdannote), "", rs!B_departdannote)
    
    frm1.FlatEdit14.Text = IIf(IsNull(rsdetailColor!B_processunit), "", rsdetailColor!B_processunit)
    frm1.processid = IIf(IsNull(rsdetailColor!B_processunitid), "", rsdetailColor!B_processunitid)
    frm1.FlatEdit15.Text = IIf(IsNull(rsdetailColor!B_processdocumentary), "", rsdetailColor!B_processdocumentary)
    frm1.processmentid = IIf(IsNull(rsdetailColor!B_processdocumentaryid), "", rsdetailColor!B_processdocumentaryid)
    frm1.FlatEdit22.Text = IIf(IsNull(rs!B_phone1), "", rs!B_phone1)
   If IIf(IsNull(rsdetailColor!B_processdate), "", rsdetailColor!B_processdate) = "" Then
        frm1.DTPicker2.Value = Now
    Else
        frm1.DTPicker2.Value = rsdetailColor!B_processdate
    End If
    frm1.FlatEdit16.Text = IIf(IsNull(rs!B_processcost), "", rs!B_processcost)
'    frm1.DTPicker2.Value = IIf(IsNull(rs!B_processdate), "", rs!B_processdate)
    frm1.FlatEdit17.Text = IIf(IsNull(rsdetailColor!B_processnote), "", rsdetailColor!B_processnote)
    
    frm1.FlatEdit25.Text = IIf(IsNull(rsdetailColor!B_processunit2), "", rsdetailColor!B_processunit2)
    frm1.processid2 = IIf(IsNull(rsdetailColor!B_processunitid2), "", rsdetailColor!B_processunitid2)
    frm1.FlatEdit24.Text = IIf(IsNull(rsdetailColor!B_processdocumentary2), "", rsdetailColor!B_processdocumentary2)
    frm1.processmentid2 = IIf(IsNull(rsdetailColor!B_processdocumentaryid2), "", rsdetailColor!B_processdocumentaryid2)
    frm1.FlatEdit23.Text = IIf(IsNull(rsdetailColor!B_phone2), "", rsdetailColor!B_phone2)
   If IIf(IsNull(rs!B_processdate2), "", rs!B_processdate2) = "" Then
        frm1.DTPicker3.Value = Now
    Else
        frm1.DTPicker3.Value = rs!B_processdate2
    End If
    frm1.FlatEdit27.Text = IIf(IsNull(rs!B_processCost2), "", rs!B_processCost2)
'    frm1.DTPicker2.Value = IIf(IsNull(rs!B_processdate), "", rs!B_processdate)
    frm1.FlatEdit26.Text = IIf(IsNull(rs!B_processnote2), "", rs!B_processnote2)
    
    frm1.FlatEdit30.Text = IIf(IsNull(rsdetailColor!B_processunit3), "", rsdetailColor!B_processunit3)
    frm1.processid3 = IIf(IsNull(rsdetailColor!B_processunitid3), "", rsdetailColor!B_processunitid3)
    frm1.FlatEdit29.Text = IIf(IsNull(rsdetailColor!B_processdocumentary3), "", rsdetailColor!B_processdocumentary3)
    frm1.processmentid3 = IIf(IsNull(rsdetailColor!B_processdocumentaryid3), "", rsdetailColor!B_processdocumentaryid3)
    frm1.FlatEdit28.Text = IIf(IsNull(rs!B_phone3), "", rs!B_phone3)
   If IIf(IsNull(rs!B_processdate3), "", rs!B_processdate3) = "" Then
        frm1.DTPicker4.Value = Now
    Else
        frm1.DTPicker4.Value = rs!B_processdate3
    End If
    frm1.FlatEdit31.Text = IIf(IsNull(rs!B_processCost3), "", rs!B_processCost3)
'    frm1.DTPicker2.Value = IIf(IsNull(rs!B_processdate), "", rs!B_processdate)
    frm1.FlatEdit32.Text = IIf(IsNull(rs!B_processnote3), "", rs!B_processnote3)
    
    
     If IIf(IsNull(rsdetailColor!B_Progressprocess), "", rsdetailColor!B_Progressprocess) = "" Then
            frm1.ComboBox5.Text = ""
     Else
        frm1.ComboBox5.Text = GetProgressCraftCT(rs!B_Progressprocess)
     End If
'    frm1.ComboBox5.Text = GetProgressCraftCT(rs!B_Progressprocess)
    frm1.FlatEdit19.Text = IIf(IsNull(rs!B_Paper), "", rs!B_Paper)
    frm1.FlatEdit20.Text = IIf(IsNull(rs!B_pocket), "", rs!B_pocket)
    frm1.FlatEdit21.Text = IIf(IsNull(rs!B_Empty), "", rs!B_Empty)
    If IIf(IsNull(rsdetailColor!B_packagstyle), "", rsdetailColor!B_packagstyle) = "" Then
            frm1.ComboBox4.Text = ""
     Else
        frm1.ComboBox4.Text = GetB_packagstyle(rsdetailColor!B_packagstyle)
     End If
'    frm1.ComboBox4.Text = GetB_packagstyle(rs!B_packagstyle)
    frm1.FlatEdit12.Text = IIf(IsNull(rs!B_departCost), "", rs!B_departCost)
    frm1.FlatEdit33.Text = Format(IIf(IsNull(rs!B_fold), "", rs!B_fold), "0.000")
    frm1.FlatEdit34.Text = IIf(IsNull(rs!B_Cast), "", rs!B_Cast)
    frm1.FlatEdit35.Text = IIf(IsNull(rs!B_PracticeCast), "", rs!B_PracticeCast)
    frm1.ComboBox1.Text = IIf(IsNull(rs!B_LabelTemplate), "", rs!B_LabelTemplate)
    frm1.ComboBox2.Text = IIf(IsNull(rs!B_DetailTemplate), "", rs!B_DetailTemplate)
    
    frm1.FlatEdit36.Text = IIf(IsNull(rs!B_DepartColor), "", rs!B_DepartColor)
    
    frm1.FlatEdit51.Text = IIf(IsNull(rs!B_PBGuige), "", rs!B_PBGuige)
    frm1.FlatEdit52.Text = IIf(IsNull(rs!B_PBPhone), "", rs!B_PBPhone)
    frm1.FlatEdit53.Text = IIf(IsNull(rs!B_PBDiZhi), "", rs!B_PBDiZhi)
    frm1.FlatEdit54.Text = IIf(IsNull(rs!B_SBDiZhi), "", rs!B_SBDiZhi)
    
    frm1.FlatEdit55.Text = IIf(IsNull(rs!B_TiaoShu), "", rs!B_TiaoShu)
    frm1.FlatEdit56.Text = IIf(IsNull(rs!B_TiaoZhong), "", rs!B_TiaoZhong)
    frm1.FlatEdit57.Text = IIf(IsNull(rs!B_ColorQty), "", rs!B_ColorQty)
    frm1.FlatEdit58.Text = Format(IIf(IsNull(rs!B_ColorZL), "", rs!B_ColorZL), "0.000")
    
        If rs!B_ItemID <> "" Then
                Dim rs1 As New RecordSet
                Dim sql As String
                Dim clsFile01 As New clsFile
                Dim szPic As String
                
                sql = "select * from G_image_NEW where B_BDCItemID='" & rs!B_ItemID & "'"
                rs1.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
                
                If rs1.RecordCount > 0 Then
                    szPic = App.Path & "\temp\" & rs1!B_id & rs1!B_BDCItemID & ".JPG"
                    Debug.Print szPic
                    
'                    clsFile01.DownloadPic rs1!B_picture, szPic
'                    cls1.InitCls szPic, frm1.Picture5
                    
                    PicShow2Ctl rs1!B_picture, frm1.Picture5
                    'PicShow2Ctl rs!B_picture, Picture3
                Else
                    frm1.Picture5.Picture = Nothing
                End If
    
       End If
    
    frm1.Show vbModal
    Unload frm1
    rsdetailColor.requery
    Exit Sub
IFERR:
    MsgBox " 点有数据的地方", vbOKOnly + vbInformation, "提示"
End Sub

Private Sub draftColordetail_null()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If

    Dim frm1 As New frmOrderProductColor_Edit
'    frm1.Valuation = ComboBox3.Text
     frm1.id = theID
    frm1.Show vbModal
    Unload frm1
    rsdetailColor.requery
End Sub
Private Sub DeleteColor()
      If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If TDBGrid3.ApproxCount <= 0 Then
        Exit Sub
    End If
    If rsdetailColor.RecordCount = 1 Then
            Dim sql1 As String
           sql1 = "delete from G_BillColor where B_belongorderid='" & theID & "'"
           Gm.cnnTool.cnn.Execute sql1
    End If
    Dim rs As New RecordSet
    Dim sql As String
    sql = "delete from G_BillDetailColor where B_itemid='" & rsdetailColor!B_ItemID & "'"
    Gm.cnnTool.cnn.Execute sql
    rsdetailColor.requery
End Sub
'合同删除，一起删除
Private Sub DeleteColorAll()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If yanzhenColor(theID) = False Then
        Exit Sub
    End If
    
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_BillColor where B_belongorderid='" & theID & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
    Dim sql1 As String
    sql1 = "delete from G_BilldetailColor where B_ID='" & rs!B_id & "'"
    Gm.cnnTool.cnn.Execute sql1
    Dim sql2 As String
    sql2 = "delete from G_BillColor where B_ID='" & rs!B_id & "'"
    Gm.cnnTool.cnn.Execute sql2
    End If
    
    '色布计划
    Colordetail
End Sub

Private Function GetProgressCraftCT(ByVal da As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select B_SID from G_ProgressCraftCT where B_SID='" & da & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        GetProgressCraftCT = rs!B_sid
    Else
        GetProgressCraftCT = ""
    End If
End Function
Private Function GetB_packagstyle(ByVal da As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "Select B_SID From G_PackWay where B_SID='" & da & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        GetB_packagstyle = rs!B_sid
    Else
        GetB_packagstyle = ""
    End If
End Function

Private Sub card()
    If rsdetailColor.RecordCount <= 0 Then
        Exit Sub
    End If
    Dim rs As New RecordSet
    Dim sql As String
    sql = "exec usp_printCard '" & theID & "','" & rsdetailColor!B_ItemID & "','" & Gm.SysID.SystemUserName & "'"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     Dim frm1 As New ActiveReport2
    frm1.itmeid = rsdetailColor!B_ItemID
    frm1.flowCardprint = IIf(IsNull(rsdetailColor!B_flowCardprint), 0, rsdetailColor!B_flowCardprint)
    Set frm1.rs = rs.Clone
    frm1.Show vbModal
    rsdetailColor.requery
End Sub

Private Sub cardAll()
    If rsdetailColor.RecordCount <= 0 Then
        Exit Sub
    End If
    Dim rs As RecordSet
    Dim frm1 As ActiveReport2
    rsdetailColor.MoveFirst
    Do While Not rsdetailColor.EOF
        Set frm1 = New ActiveReport2
        Set rs = New RecordSet
        frm1.itmeid = rsdetailColor!B_ItemID
         frm1.flowCardprint = IIf(IsNull(rsdetailColor!B_flowCardprint), 0, rsdetailColor!B_flowCardprint)
        Dim sql As String
        sql = "exec usp_printCard '" & theID & "','" & rsdetailColor!B_ItemID & "','" & Gm.SysID.SystemUserName & "'"
        Debug.Print sql
         rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        Set frm1.rs = rs.Clone
        frm1.Show vbModal
        rsdetailColor.movenext
    Loop
     rsdetailColor.requery
End Sub
Private Sub departprint()
     If rsdetailColor.RecordCount <= 0 Then
        Exit Sub
     End If
        
    If IIf(IsNull(rsdetailColor!B_depart), "", rsdetailColor!B_depart) = "" Then
        MsgBox "请先填写染厂", vbInformation, "提示"
        Exit Sub
    End If
    Dim sql As String
    Dim sql1 As String
    Dim rs As New RecordSet
    Dim rs1 As New RecordSet
    sql1 = "select * from G_BillColor where B_belongorderID='" & theID & "'"
    rs.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
'    sql = "exec usp_depart '" & theid & "','" & rsdetailColor!B_itemid & "'"
    sql = "exec usp_depart '" & theID & "','" & rsdetailColor!B_orderColor & "','" & rsdetailColor!B_departid & "'"
    Debug.Print sql
'    sql = "exec usp_depart '" & theID & "'"
'
'    Dim frm1 As New ActiveReport7
'    frm1.id = rs!B_id
''    frm1.itemid = rsdetailColor!B_itemid
''    frm1.color = rsdetailColor!B_orderColor
'    frm1.depart = rsdetailColor!B_departid
'    frm1.departdanprint = IIf(IsNull(rsdetailColor!B_departdanprint), 0, rsdetailColor!B_departdanprint)
'    frm1.sql = sql
'
'    frm1.Show vbModal
'    Unload frm1

    sql = "exec usp_depart '" & theID & "','" & rsdetailColor!B_orderColor & "','" & rsdetailColor!B_departid & "'"
    Debug.Print sql
    rs1.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

    Dim frm1 As New frmModBLRPreviewOri
    Set frm1.RecordSet = rs1.Clone
'    frm1.obj = "11S067"
    frm1.ObjectID = "22B138"
    frm1.Show
    
    rsdetailColor.requery
End Sub
'打印染厂派工单
Private Sub depart()
    If rsdetailColor.RecordCount <= 0 Then
        Exit Sub
    End If
    If IIf(IsNull(rsdetailColor!B_depart), "", rsdetailColor!B_depart) = "" Then
        MsgBox "请先填写染厂", vbInformation, "提示"
        Exit Sub
    End If
    Dim sql As String
    Dim a As String
    Dim sql1 As String
    Dim rs As New RecordSet
    a = ""
'    sql = "exec usp_depart '" & theID & "','" & rsdetailColor!B_ItemID & "'"
    sql = "exec usp_departall '" & theID & "','" & a & "'"
    Debug.Print sql
    sql1 = "select * from G_BillColor where B_belongorderID='" & theID & "'"
    rs.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
'    Dim frm1 As New ActiveReport1
'    frm1.id = rs!B_id
'    frm1.itemid = 0
'    frm1.departdanprint = IIf(IsNull(rsdetailColor!B_departdanprint), 0, rsdetailColor!B_departdanprint)
'    frm1.sql = sql
'    Debug.Print sql
'    frm1.Show vbModal
'    Unload frm1
    
    Dim frm1 As New frmModBLRPreviewOri
    Set frm1.RecordSet = rs.Clone
'    frm1.obj = "11S067"
    frm1.ObjectID = "22B138"
    frm1.Show
    
    
    rsdetailColor.requery

End Sub
Private Sub processprint()
     If rsdetailColor.RecordCount <= 0 Then
        Exit Sub
    End If
    If IIf(IsNull(rsdetailColor!B_processunit), "", rsdetailColor!B_processunit) = "" Then
        MsgBox "请先填写深加工", vbInformation, "提示"
        Exit Sub
    End If
    Dim sql As String
    Dim sql1 As String
    Dim rs As New RecordSet
    sql1 = "select * from G_BillColor where B_belongorderID='" & theID & "'"
    rs.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
'    sql = "exec usp_depart '" & theID & "','" & rsdetailColor!B_ItemID & "'"
    sql = "exec usp_process '" & theID & "','" & rsdetailColor!B_ItemID & "'"
'    sql = "exec usp_depart '" & theID & "'"
'
 Dim frm1 As New ActiveReport4
    frm1.id = rs!B_id
     frm1.itemid = rsdetailColor!B_ItemID
    frm1.processprint = IIf(IsNull(rsdetailColor!B_processprint), 0, rsdetailColor!B_processprint)
    frm1.sql = sql

    frm1.Show vbModal
      Unload frm1
    rsdetailColor.requery
End Sub
'深加工派工单
Private Sub process()
    If rsdetailColor.RecordCount <= 0 Then
        Exit Sub
    End If
    
      If IIf(IsNull(rsdetailColor!B_processunit), "", rsdetailColor!B_processunit) = "" Then
        MsgBox "请先填写深加工", vbInformation, "提示"
        Exit Sub
    End If
    Dim sql As String
    Dim a As String
    Dim sql1 As String
    Dim rs As New RecordSet
    Dim rs1 As New RecordSet
    a = ""
    sql = "exec usp_process '" & theID & "','" & a & "'"
'    sql = "exec usp_process '" & theID & "'"
    sql1 = "select * from G_BillColor where B_belongorderID='" & theID & "'"
    rs.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
'    Dim frm1 As New ActiveReport4
'    frm1.id = rs!B_id
'     frm1.itemid = 0
'    frm1.processprint = IIf(IsNull(rsdetailColor!B_processprint), 0, rsdetailColor!B_processprint)
'    frm1.sql = sql
''    Set frm1.rs = rs.Clone
'    frm1.Show vbModal
'      Unload frm1

    sql = "exec usp_process '" & theID & "','" & a & "'"
   rs1.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

    Dim frm1 As New frmModBLRPreviewOri
    Set frm1.RecordSet = rs1.Clone
'    frm1.obj = "11S067"
    frm1.ObjectID = "22B139"
    frm1.Show


    rsdetailColor.requery
End Sub
Public Function yanzhenColor(ByVal theID As String) As Boolean
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
    sql = "Select * From G_SystemUser where B_UserName='" & Gm.SysID.SystemUser & "'"
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs!B_SuperAdmin = 1 Then
        yanzhenColor = True
        Exit Function
    End If
    
    sql1 = "select distinct B_date from G_BilldetailColor where B_ID=(select B_ID from G_BillColor where B_belongorderid='" & theID & "')"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
'    sql2 = "select * from G_BillColor where B_Belongorderid='" & theid & "'"
    sql2 = "SELECT * FROM G_UserPro WHERE B_username='" & Gm.SysID.SystemUser & "' AND B_objectid='11S005'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs2.RecordCount > 0 Then
        If IIf(IsNull(rs2!B_new), 0, rs2!B_new) = 1 Then
            yanzhenColor = True
        Else
            yanzhenColor = False
            MsgBox "请设置权限", vbInformation, "提示"
            Exit Function
        End If
        If IIf(IsNull(rs1!B_Date), "", rs1!B_Date) <> "" Then
            If DateDiff("s", rs1!B_Date, Now) > 84600 Then
                yanzhenColor = False
                MsgBox "已经超过制作本单据的时间不能进行修改", vbInformation, "提示"
            Else
                yanzhenColor = True
            End If
        End If
    Else
        yanzhenColor = False
        MsgBox "你没有此权限", vbInformation, "提示"
    End If
End Function



'---------------------------------------------------------------白坯计划---------------------------------------------------------

Private Sub ActiveBar23_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
                Case "新增"
                    draftwhitedetail_null
                Case "删除"
                    Deletewhite
                Case "复制行"
                    copywhite
                Case "打印当前订单"
                    printwhite
                Case "打印全部订单"
                    printwhiteAll
                Case "保存样式"
                    setGridStyle3
    End Select
End Sub
Private Sub setGridShow3()
    Dim cls1 As New clsGridShow
    With cls1
        .ObjectID = "11S006"
        .InitClass TDBGrid2, 3
        .ShowGridFormat
    End With
End Sub

'保存网格样式
Private Sub setGridStyle3()
    Dim i As Long
    Dim strSQL As String
    Dim dWidth As Integer
    Dim szFieldName As String
    
    For i = 0 To TDBGrid2.Columns.Count - 1
        If TDBGrid2.Columns(i).width > 0 Then
            If TDBGrid2.Columns(i).Visible = True Then
                szFieldName = TDBGrid2.Columns(i).DataField
                dWidth = TDBGrid2.Columns(i).width
                strSQL = "update G_BLSField set B_GridWidth='" & dWidth & "' where B_ObjectID='11S006' and B_FieldName='" & szFieldName & "'"
                Gm.cnnTool.cnn.Execute strSQL
            End If
        End If
    Next
    
End Sub
Private Sub copywhite()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If yanzhenWhite(theID) = False Then
        Exit Sub
    End If
     If rsdetailwhite.RecordCount <= 0 Then
        Exit Sub
    End If
     Dim frm1 As New frmWhite_Edit
     frm1.id = theID
     frm1.OrderID = rsdetailwhite!B_ItemIDB
     frm1.FlatEdit2.Text = rsdetailwhite!B_GoodsNameAlias
     frm1.FlatEdit5.Text = rsdetailwhite!B_GoodsID
     frm1.ComboBox2.Text = rsdetailwhite!B_Width
     frm1.ComboBox3.Text = rsdetailwhite!B_UnitWeight
     frm1.FlatEdit6.Text = rsdetailwhite!B_BoxQty
     frm1.FlatEdit7.Text = rsdetailwhite!B_MemoDetail
     frm1.FlatEdit1.Text = rsdetailwhite!B_Maohight
     frm1.DTPicker1.Value = rsdetailwhite!B_Deliverydate
'    frm1.Check1.Value = rsdetailwhite!B_Deliverydate
     frm1.Whiteid = IIf(IsNull(rsdetailwhite!B_sid), "", rsdetailwhite!B_sid)
     frm1.Check1.Value = IIf(IsNull(rsdetailwhite!B_intype), "", rsdetailwhite!B_intype)
     frm1.client = IIf(IsNull(rsdetailwhite!B_Clientid), "", rsdetailwhite!B_Clientid)
     frm1.FlatEdit11.Text = IIf(IsNull(rsdetailwhite!B_ClientName), "", rsdetailwhite!B_ClientName)
    frm1.Show vbModal
    Unload frm1
    rsdetailwhite.requery
End Sub
Private Sub Deletewhite()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If yanzhenWhite(theID) = False Then
        Exit Sub
    End If
    If rsdetailwhite.RecordCount = 1 Then
            Dim sql1 As String
           sql1 = "delete from G_BillWhite where B_belongorderid='" & theID & "'"
           Gm.cnnTool.cnn.Execute sql1
    End If
    If rsdetailwhite.RecordCount > 0 Then
        Dim rs As New RecordSet
        Dim sql As String
        sql = "delete from G_BillDetailWhite where B_itemid='" & rsdetailwhite!B_ItemID & "'"
        Gm.cnnTool.cnn.Execute sql
    End If
    rsdetailwhite.requery
End Sub
'合同删除，一起删除
Private Sub DeletewhiteAll()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_BillWhite where B_belongorderid='" & theID & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
    Dim sql1 As String
    sql1 = "delete from G_BilldetailWhite where B_ID='" & rs!B_id & "'"
    Gm.cnnTool.cnn.Execute sql1
    Dim sql2 As String
    sql2 = "delete from G_BillWhite where B_ID='" & rs!B_id & "'"
    Gm.cnnTool.cnn.Execute sql2
    End If
End Sub

'绑定tbgrid2的数据
Private Sub whitedetail()
    Set rsdetailwhite = New RecordSet
    Dim sql As String
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
    sql = "exec usp_SelectWhite '" & theID & "'"
    rsdetailwhite.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    TDBGrid2.DataSource = rsdetailwhite

    sql1 = "select * from G_BillWhite where B_belongorderid='" & theID & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs1.RecordCount > 0 Then
            sql2 = "select * from G_Systemuser where B_Username='" & rs1!B_username & "'"
            rs2.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
              If rs2.RecordCount > 0 Then
                    Dim szUserName As String
                    szUserName = IIf(IsNull(rs2!B_UserDes), "", rs2!B_UserDes)
                    ActiveBar23.Bands("Band1").Tools("制单").Caption = szUserName
                    ActiveBar23.RecalcLayout
               End If
    End If
  
    setTGB_two
End Sub

Private Sub setTGB_two()
    setGridShow3
'    TDBGrid2.Columns("B_ItemIDB").Caption = "订单号"
'    TDBGrid2.Columns("B_GoodsNameAlias").Caption = "品名"
'    TDBGrid2.Columns("B_Goodsid").Caption = "白坯名称"
'    TDBGrid2.Columns("B_Width").Caption = "门幅"
'    TDBGrid2.Columns("B_UnitWeight").Caption = "克重"
'    TDBGrid2.Columns("B_UnitWeight").width = 1500
'    TDBGrid2.Columns("B_BoxQty").Caption = "数量KG"
'    TDBGrid2.Columns("B_BoxQty").width = 1500
'    TDBGrid2.Columns("B_print").Caption = "打印白坯流转卡次数"
'    TDBGrid2.Columns("B_print").width = 1500
'    TDBGrid2.Columns("B_MemoDetail").Caption = "备注"
    TDBGrid2.Columns("B_SID").width = 0
    TDBGrid2.Columns("B_SID").Visible = False
    TDBGrid2.Columns("B_SID").AllowSizing = False
    TDBGrid2.Columns("B_Itemid").width = 0
    TDBGrid2.Columns("B_Itemid").Visible = False
    TDBGrid2.Columns("B_Itemid").AllowSizing = False
'     TDBGrid2.Columns("B_MaoHight").width = 0
'    TDBGrid2.Columns("B_MaoHight").Visible = False
'    TDBGrid2.Columns("B_MaoHight").AllowSizing = False
       TDBGrid2.Columns("B_goodMaohight").width = 0
    TDBGrid2.Columns("B_goodMaohight").Visible = False
    TDBGrid2.Columns("B_goodMaohight").AllowSizing = False
           TDBGrid2.Columns("B_ClientID").width = 0
    TDBGrid2.Columns("B_ClientID").Visible = False
    TDBGrid2.Columns("B_ClientID").AllowSizing = False
           TDBGrid2.Columns("B_ClientName").width = 0
    TDBGrid2.Columns("B_ClientName").Visible = False
    TDBGrid2.Columns("B_ClientName").AllowSizing = False
    TDBGrid2.Columns("B_intype").ValueItems.Presentation = dbgCheckBox
'
'
     TDBGrid2.MarqueeStyle = dbgHighlightRow
'    TDBGrid2.HoldFields
sumall3
End Sub
'小计
Private Sub sumall3()
    Dim a As Long
   
    Dim rs As New RecordSet
    a = 0
    Set rs = rsdetailwhite.Clone
    If rs.RecordCount <= 0 Then
      
        a = 0
      
    Else
        rs.MoveFirst
        Do While Not rs.EOF
            a = a + IIf(IsNull(rs!B_BoxQty), 0, rs!B_BoxQty)
           
            rs.movenext
        Loop
        rs.MoveFirst
    End If
    TDBGrid2.Columns("B_ItemIDB").FooterText = "合计"
    TDBGrid2.Columns("B_BoxQty").FooterText = "" & a & ""

End Sub

Private Sub TDBGrid2_DblClick()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If rsdetailwhite.RecordCount > 0 Then
       whitedetail_UPdate
    Else
        draftwhitedetail_null
    End If
End Sub
Public Function yanzhenWhite(ByVal theID As String) As Boolean
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
    sql = "Select * From G_SystemUser where B_UserName='" & Gm.SysID.SystemUser & "'"
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs!B_SuperAdmin = 1 Then
        yanzhenWhite = True
        Exit Function
    End If
    
    sql1 = "select * from G_BilldetailWhite where B_ID=(select B_ID from G_BillWhite where B_belongorderid='" & theID & "')"
    Debug.Print sql1
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
'    sql2 = "select * from G_BillWhite where B_belongorderid='" & theid & "'"
    sql2 = "SELECT * FROM G_UserPro WHERE B_username='" & Gm.SysID.SystemUser & "' AND B_objectid='11S006'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs2.RecordCount > 0 Then
        If IIf(IsNull(rs2!B_new), 0, rs2!B_new) = 1 Then
            yanzhenWhite = True
        Else
            yanzhenWhite = False
            MsgBox "请设置权限", vbInformation, "提示"
            Exit Function
        End If
        If rs1.RecordCount > 0 Then
            If IIf(IsNull(rs1!B_Date), "", rs1!B_Date) <> "" Then
                If DateDiff("s", rs1!B_Date, Now) > 84600 Then
                    yanzhenWhite = False
                    MsgBox "已经超过制作本单据的时间不能进行修改", vbInformation, "提示"
                Else
                    yanzhenWhite = True
                End If
            End If
        End If
    Else
        yanzhenWhite = False
        MsgBox "你没有此权限", vbInformation, "提示"
    End If
End Function
Private Sub draftwhitedetail_null()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If

    If yanzhenWhite(theID) = False Then
        Exit Sub
    End If
    Dim frm1 As New frmWhite_Edit
'    frm1.theidwhite = theidwhite
    frm1.id = theID
    Debug.Print theID
    frm1.Show vbModal
    Unload frm1
    rsdetailwhite.requery
    Dim sql As String
    Dim rs As New RecordSet
    sql = "select b.B_UserDes from G_BillWhite a left outer join G_systemUser b on a.B_UserName=b.B_UserName  where B_belongorderid='" & theID & "'"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs.RecordCount > 0 Then
        ActiveBar23.Bands("Band1").Tools("制单").Caption = "" & rs!B_UserDes & ""
    End If
End Sub
Private Sub whitedetail_UPdate()
    On Error GoTo IFERR
    If rsdetailwhite.RecordCount <= 0 Then
        Exit Sub
    End If
    Dim rs  As RecordSet
'    Dim rs1 As New RecordSet
'    Dim sql1 As String
'    sql1 = "select * from G_Billwhite where B_BelongOrderID='" & theID & "'"
'    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        Set rs = New RecordSet
        Dim sql2 As String
        sql2 = "select * from G_BillDetailwhite where B_itemid='" & rsdetailwhite!B_ItemID & "' "
        Debug.Print sql2
        rs.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        
        Dim sql3 As String
        Dim rs3 As New RecordSet
        sql3 = "select * from G_ContactCompany where B_Clientid='" & rs!B_supplier & "'"
        Debug.Print sql3
        rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
    Dim frm1 As New frmWhite_Edit
    frm1.itemid = rs!B_ItemID

    frm1.id = theID
    frm1.OrderID = IIf(IsNull(rs!B_ItemIDB), "", rs!B_ItemIDB)
   'frm1.ComboBox1.Text = IIf(IsNull(rs!B_ItemIDB), "", rs!B_ItemIDB)
    frm1.FlatEdit2.Text = IIf(IsNull(rs!B_GoodsNameAlias), "", rs!B_GoodsNameAlias)
    frm1.ComboBox2.Text = IIf(IsNull(rs!B_Width), "", rs!B_Width)
    frm1.ComboBox3.Text = IIf(IsNull(rs!B_UnitWeight), "", rs!B_UnitWeight)
    frm1.FlatEdit5.Text = IIf(IsNull(rsdetailwhite!B_GoodsID), "", rsdetailwhite!B_GoodsID)
    frm1.FlatEdit6.Text = IIf(IsNull(rs!B_BoxQty), "", rs!B_BoxQty)
    frm1.FlatEdit7.Text = IIf(IsNull(rs!B_MemoDetail), "", rs!B_MemoDetail)
    frm1.Whiteid = IIf(IsNull(rsdetailwhite!B_sid), "", rsdetailwhite!B_sid)
    frm1.FlatEdit1.Text = IIf(IsNull(rs!B_Maohight), "", rs!B_Maohight)
    frm1.DTPicker1.Value = IIf(IsNull(rs!B_Deliverydate), "", rs!B_Deliverydate)
    frm1.Check1.Value = IIf(IsNull(rs!B_intype), 0, rs!B_intype)
    If rs3.RecordCount > 0 Then
        frm1.client = IIf(IsNull(rs!B_supplier), "", rs!B_supplier)
        frm1.FlatEdit11.Text = IIf(IsNull(rs3!B_ClientName), "", rs3!B_ClientName)
    End If
'    frm1.theidwhite = theidwhite
    frm1.Show vbModal

    
    Unload frm1
    rsdetailwhite.requery
 
    Exit Sub
IFERR:
    
    MsgBox "请点击有数据的地方", vbOKOnly + vbInformation, "提示"

End Sub
'白坯计划打印
Private Sub printwhite()
    If rsdetailwhite.RecordCount <= 0 Then
        Exit Sub
    End If
     Dim rs As New RecordSet
    Dim sql As String
    sql = "exec usp_printWhite '" & theID & "','" & rsdetailwhite!B_ItemID & "'"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
'    Dim frm1 As New ActiveReport3
'    frm1.itmeid = rs!B_ItemID
'    frm1.flowCardprint = IIf(IsNull(rs!B_print), 0, rs!B_print)
'    Set frm1.rs = rs.Clone
'    frm1.Show vbModal
'    rsdetailwhite.requery
'    Unload frm1

    Dim frm1 As New frmModBLRPreviewOri
    Set frm1.RecordSet = rs.Clone
'    frm1.obj = "11S067"
    frm1.ObjectID = "22B140"
    frm1.Show


End Sub
Private Sub printwhiteAll()
     If rsdetailwhite.RecordCount <= 0 Then
        Exit Sub
    End If
     Dim rs As RecordSet
    Dim frm1 As New ActiveReport3
    Dim sql As String
    Do While Not rsdetailwhite.EOF
    Set rs = New RecordSet
    sql = "exec usp_printWhite '" & theID & "','" & rsdetailwhite!B_ItemID & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    Set frm1 = New ActiveReport3
    frm1.itmeid = rs!B_ItemID
    frm1.flowCardprint = IIf(IsNull(rs!B_print), 0, rs!B_print)
    Set frm1.rs = rs.Clone
    frm1.Show vbModal
    rsdetailwhite.movenext
    Loop
    rsdetailwhite.requery
End Sub


'-----------------------------------------------------------------------订单图样---------------------------------------------
Private Sub pattern()
    Set rsgrid4 = New RecordSet
    Dim sql As String
    sql = "select distinct B_OrderCode from G_Billdetailorder where B_ID='" & theID & "'"
    rsgrid4.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    TDBGrid4.DataSource = rsgrid4
    TDBGrid4.Columns("B_OrderCode").Caption = "订单号"
    TDBGrid4.MarqueeStyle = dbgHighlightRow
     If rsgrid4.RecordCount > 0 Then
          rsgrid4.MoveFirst
     End If
     
    Dim sql2 As String
    Dim rs2 As New RecordSet
    sql2 = "select * from G_Config_FormCtlShow where B_sid='订单号' "
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs2.RecordCount > 0 Then
        
        TDBGrid4.Columns("B_OrderCode").Caption = rs2!B_Caption
      
    End If
End Sub

Private Sub C1Tab1_Click()
     If C1Tab1.CurrTab = 3 Then
        If rsgrid4.RecordCount > 0 Then
                If rsgrid4!B_ordercode = "" Then
                        Exit Sub
                End If
                Dim rs As New RecordSet
                Dim sql As String
                Dim clsFile01 As New clsFile
                Dim szPic As String
                
                sql = "select * from G_Image where B_ID='" & theID & "' and B_itemid='" & rsgrid4!B_ordercode & "'"
                rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
                
                If rs.RecordCount > 0 Then
                    szPic = App.Path & "\temp\" & rs!B_id & rs!B_ItemID & ".JPG"
                    Debug.Print szPic
                    
                    clsFile01.DownloadPic rs!B_picture, szPic
                    cls1.InitCls szPic, Picture3
                    
                    'PicShow2Ctl rs!B_picture, Picture3
                Else
                    Picture3.Picture = Nothing
                End If
        End If
     End If
     
     
     If C1Tab1.CurrTab = 5 Then
        pictureorder
     End If
     
End Sub

'预览图片
Private Sub PushButton2_Click()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    On Error GoTo IFERR
    
    With CommonDialog1
        .ShowOpen
   
        szFile = .FileName
    End With
    
    If Len(szFile) <= 0 Then
        Exit Sub
    End If
    cls1.InitCls szFile, Picture3
    
'    Picture3.Picture = LoadPicture(szFile)
    
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "图片格式不支持，请提供JPG或者BMP格式图片！" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub
Private Sub PushButton4_Click()
    Dim sql As String
    Dim rs As New RecordSet
    If TDBGrid4.ApproxCount <= 0 Then
        MsgBox "当前没有订单号不能上传", vbInformation, "提示"
        Exit Sub
    End If
    If szFile <> "" Then
'                 需要引用：Microsoft Scripting Runtime
        Dim fso As New FileSystemObject
    
        Dim lLength As Long
        Dim oFile As File
    
        Set oFile = fso.GetFile(szFile)
        
        sql = "select * from G_ImageSize"
        rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
        
        
        If oFile.Size / 1000000 > rs!B_size Then
            MsgBox "图片太大不能上传", vbInformation, "提示"
            Exit Sub
        End If
    
        '获取的长度的单位是：字节
        saveImage
       
            MsgBox "图片上传成功", vbInformation, "提示"
       
    End If
End Sub
'Private Sub FileLen()
'    '需要引用：Microsoft Scripting Runtime
'    Dim fso As New FileSystemObject
'
'    Dim lLength As Long
'    Dim oFile As File
'
'    Set oFile = fso.GetFile("D:\4Trans\花型\云朵.jpg")
'
'    MsgBox oFile.Size '获取的长度的单位是：字节
'End Sub

Private Sub saveImage()
  
    If rsgrid4!B_ordercode = "" Then
'        saveImage = False
        Exit Sub
    End If
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_Image where 1=0"
    rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
    
     If szFile = "" Then
    Else
        Dim rs1 As New RecordSet
        Dim sql1 As String
        sql1 = "select * from G_Image where B_ID='" & theID & "' and B_itemid='" & rsgrid4!B_ordercode & "'"
        rs1.Open sql1, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
        If rs1.RecordCount > 0 Then
'            Dim sql2 As String
'            sql2 = "update G_Image set B_Picture='" & szFile & "' where  B_ID='" & theID & "' and B_itemid='" & rsgrid4!B_OrderCode & "'"
'            Gm.cnnToolImage.cnn.Execute sql2
            
            PicSaveToDB rs1!B_picture, szFile
            rs1.Update
            rs1.Close
            Set rs1 = Nothing
        Else
            rs.AddNew
            rs!B_id = theID
            PicSaveToDB rs!B_picture, szFile
            rs!B_ItemID = rsgrid4!B_ordercode
            rs.Update
            rs.Close
            Set rs = Nothing
        End If
        
    End If
End Sub
Private Sub saveImageBZ()
  
    If rsgrid4!B_ordercode = "" Then
'        saveImage = False
        Exit Sub
    End If
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from WVAccountImage.dbo.G_image_NEW_BZ where 1=0"
    rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
    
     If szFile3 = "" Then
    Else
        Dim rs1 As New RecordSet
        Dim sql1 As String
        sql1 = "select * from WVAccountImage.dbo.G_image_NEW_BZ where B_OrderID='" & theID & "'"
        rs1.Open sql1, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
        If rs1.RecordCount > 0 Then
        
            PicSaveToDB rs1!B_picture, szFile3
            rs1!B_memo = IIf(IsNull(Text1.Text), "", Text1.Text)
            rs1.Update
            rs1.Close
            Set rs1 = Nothing
        Else
            rs.AddNew
            rs!B_OrderID = theID
            PicSaveToDB rs!B_picture, szFile3
            rs!B_memo = IIf(IsNull(Text1.Text), "", Text1.Text)
            rs.Update
            rs.Close
            Set rs = Nothing
        End If
        
    End If
End Sub
'上传图片到服务器
'fld：记录集中的字段
'vFilePath：图片文件的绝对路径，包含图片文件名和扩展名
Private Sub PicSaveToDB(ByRef fld As ADODB.Field, ByVal vFilePath As String)
    Const blocksize = 4096
    Dim bytedata() As Byte
    Dim numblocks As Long
    Dim filelength As Long
    Dim leftover As Long
    Dim sourcefile As Long
    Dim i As Long
    sourcefile = FreeFile
    
    Open Trim(vFilePath) For Binary Access Read As sourcefile
    filelength = LOF(sourcefile)
    
    If filelength = 0 Then
        Close sourcefile
        'MsgBox Trim(vFilePath) & "无内容或不存在！"
    Else
        numblocks = filelength \ blocksize
        leftover = filelength Mod blocksize
        fld.Value = Null
        
        ReDim bytedata(blocksize)
        
        For i = 1 To numblocks
            Get sourcefile, , bytedata()
            fld.AppendChunk bytedata()
        Next
        

        ReDim bytedata(leftover)
        Get sourcefile, , bytedata()
        fld.AppendChunk bytedata()
        Close sourcefile
    End If
End Sub
'从DB中下载图片并且显示到UI的图片控件上
'vRs：包含有图片资源的数据源
'vPicField：保存图片文件的字段名
'oCtl：用于显示的控件。PictureBox、Image
Private Sub PicShow2Ctl(ByRef vFld As ADODB.Field, ByRef oCtl As Object)
    'On Error GoTo IFERR
    Dim Stream As ADODB.Stream
    Set Stream = New ADODB.Stream
    

    oCtl.Picture = LoadPicture("")
    If Not IsNull(vFld) Then
        Stream.Type = adTypeBinary
        Stream.Open
        Stream.Write vFld
        Stream.SaveToFile "filename", adSaveCreateOverWrite
        'Stream.SaveToFile "c:\aaa.jpg", adSaveCreateOverWrite
        
        szFile = LoadPicture("filename")
'        Debug.Print FileName
        oCtl.Picture = LoadPicture("filename")
        Stream.Close
    End If
    
    Set Stream = Nothing
'    Exit Sub
'IFERR:
'    Dim szErr As String
'    szErr = "错误发生于下载图片中，" & Err.Description
'    MsgBox szErr
End Sub

Private Sub TDBGrid4_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If TDBGrid4.ApproxCount > 0 Then
        If rsgrid4!B_ordercode = "" Then
            Exit Sub
        End If
        Dim rs As New RecordSet
        Dim sql As String
        Dim clsFile01 As New clsFile
        Dim szPic As String
        
        sql = "select * from G_Image where B_ID='" & theID & "' and B_itemid='" & rsgrid4!B_ordercode & "'"
        Debug.Print sql
        rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            szPic = App.Path & "\temp\" & rs!B_id & rs!B_ItemID & ".JPG"
            Debug.Print szPic
            clsFile01.DownloadPic rs!B_picture, szPic
            cls1.InitCls szPic, Picture3
            
            'PicShow2Ctl rs!B_picture, Picture3
        Else
            Picture3.Picture = Nothing
        End If
        
        Set rs = Nothing
        Set clsFile01 = Nothing
    End If
End Sub

Private Sub DeletePicture()
    Dim sql As String
    sql = "delete from G_image where B_ID='" & theID & "'"
    Gm.cnnToolImage.cnn.Execute sql
End Sub

'-----------------------------------------------------------------------白坯构成---------------------------------------------
Private Sub ActiveBar25_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
                Case "新增"
                    rsgrid5_Null
                Case "删除"
                    rsgrid5_delete
                Case "复制行"
                    rsgrid5_copy
                Case "保存样式"
                    setGridStyle4
    End Select
End Sub

Public Sub LoadObject()
    
End Sub
Private Sub setGridShow4()
    Dim cls1 As New clsGridShow
    With cls1
        .ObjectID = "11S007"
        .InitClass TDBGrid5, 3
        .ShowGridFormat
    End With
End Sub

'保存网格样式
Private Sub setGridStyle4()
    Dim i As Long
    Dim strSQL As String
    Dim dWidth As Integer
    Dim szFieldName As String
    
    For i = 0 To TDBGrid5.Columns.Count - 1
        If TDBGrid5.Columns(i).width > 0 Then
            If TDBGrid5.Columns(i).Visible = True Then
                szFieldName = TDBGrid5.Columns(i).DataField
                dWidth = TDBGrid5.Columns(i).width
                strSQL = "update G_BLSField set B_GridWidth='" & dWidth & "' where B_ObjectID='11S007' and B_FieldName='" & szFieldName & "'"
                Gm.cnnTool.cnn.Execute strSQL
            End If
        End If
    Next
    
End Sub

'初始化白坯构成
Private Sub WhiteComposition()
    Dim sql As String
    Set rsgrid5 = New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
    sql = "exec usp_Composition '" & theID & "'"
    rsgrid5.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid5.DataSource = rsgrid5
  
    sql1 = "select distinct B_UserName from G_WhiteComposition where B_id='" & theID & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
  
    If rs1.RecordCount > 0 Then
            sql2 = "select * from G_Systemuser where B_Username='" & rs1!B_username & "'"
            rs2.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
              If rs2.RecordCount > 0 Then
                    Dim szUserName As String
                    szUserName = IIf(IsNull(rs2!B_UserDes), "", rs2!B_UserDes)
                    ActiveBar25.Bands("Band1").Tools("制单").Caption = szUserName
                    ActiveBar25.RecalcLayout
               End If
    Else
        ActiveBar25.Bands("Band1").Tools("制单").Caption = ""
        ActiveBar25.RecalcLayout
    End If
    setWhiteComposition
End Sub
'设置构成样式
Private Sub setWhiteComposition()
    setGridShow4
'    TDBGrid5.Columns("B_Breed").Caption = "投料品种"
'    TDBGrid5.Columns("B_ItemIDB").Caption = "订单号"
'    TDBGrid5.Columns("B_GoodsNameAlias").Caption = "品名"
'    TDBGrid5.Columns("B_GoodsName").Caption = "品名名称"
'    TDBGrid5.Columns("B_Suppliers").Caption = "供应商"
'    TDBGrid5.Columns("B_Width").Caption = "规格"
''    TDBGrid5.Columns("B_UnitWeight").Caption = "克重"
'    TDBGrid5.Columns("B_StorageWay").Caption = "入库方式"
'    TDBGrid5.Columns("B_TransfersItemIDB").Caption = "调拨单号"
'    TDBGrid5.Columns("B_TransfersSuppliers").Caption = "调拨供应商"
'    TDBGrid5.Columns("B_BreedNum").Caption = "品种数量"
'    TDBGrid5.Columns("B_Breed").width = 1000
'    TDBGrid5.Columns("B_Width").width = 1000
''    TDBGrid5.Columns("B_UnitWeight").width = 1000
'
    TDBGrid5.Columns("B_supplement").ValueItems.Presentation = dbgCheckBox
    TDBGrid5.Columns("B_itemid").width = 0
    TDBGrid5.Columns("B_itemid").Visible = False
    TDBGrid5.Columns("B_itemid").AllowSizing = False
    TDBGrid5.Columns("B_Suppliersid").width = 0
    TDBGrid5.Columns("B_Suppliersid").Visible = False
    TDBGrid5.Columns("B_Suppliersid").AllowSizing = False
    TDBGrid5.Columns("B_TransfersSuppliersid").width = 0
    TDBGrid5.Columns("B_TransfersSuppliersid").Visible = False
    TDBGrid5.Columns("B_TransfersSuppliersid").AllowSizing = False
    TDBGrid5.MarqueeStyle = dbgHighlightRow
'
''    setWhiteComposition_Edit
'    TDBGrid5.HoldFields
End Sub

Private Sub TDBGrid5_DblClick()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If rsgrid5.RecordCount <= 0 Then
        rsgrid5_Null
        ActiveBar25.RecalcLayout
    Else
        rsgrid5_update
    End If
    
End Sub
Public Function yanzhenWhiteComposition(ByVal theID As String) As Boolean
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
    sql = "Select * From G_SystemUser where B_UserName='" & Gm.SysID.SystemUser & "'"
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs!B_SuperAdmin = 1 Then
        yanzhenWhiteComposition = True
        Exit Function
    End If
    
    sql1 = "select * from G_WhiteComposition where B_ID='" & theID & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
   
    sql2 = "select distinct B_UserName from G_WhiteComposition  where B_ID='" & theID & "'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs2.RecordCount > 0 Then
        If rs2!B_username = Gm.SysID.SystemUser Then
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
            If DateDiff("s", rs1!B_Date, Now) > 84600 Then
                yanzhenWhiteComposition = False
                MsgBox "已经超过制作本单据的时间不能进行修改", vbInformation, "提示"
            Else
                yanzhenWhiteComposition = True
            End If
        End If
    End If
End Function
Private Sub rsgrid5_Null()
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If yanzhenWhiteComposition(theID) = False Then
        Exit Sub
    End If
    Dim frm1 As New frmWhite_composition
    frm1.id = theID
    Set frm1.Rs5 = rsgrid5
    frm1.Show vbModal
        Dim sql As String
    Dim rs As New RecordSet
    sql = "select distinct b.B_UserDes from G_WhiteComposition a left outer join G_systemUser b on a.B_UserName=b.B_UserName  where B_id='" & theID & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs.RecordCount > 0 Then
        ActiveBar25.Bands("Band1").Tools("制单").Caption = "" & rs!B_UserDes & ""
    End If
End Sub

Private Sub rsgrid5_update()
'    Dim frm1 As New frmWhite_composition
'    frm1.id = theID
'    frm1.itemid = rsgrid5!b_itemid
    rsgrid5Update
'    frm1.Show vbModal
'    Unload frm1
    rsgrid5.requery
End Sub
Private Sub rsgrid5Update()
    If rsgrid5!B_Breed = "原料" Then
        If rsgrid5!B_StorageWay = "采购" Then
            CompositionUPdate_1
        Else
            CompositionUPdate_2
        End If
    Else
         If rsgrid5!B_StorageWay = "采购" Then
            CompositionUPdate_3
        Else
            CompositionUPdate_4
        End If
    End If
End Sub
Private Sub CompositionUPdate_1()
       Dim frm1 As New frmWhite_composition
        frm1.ComboBox1.Text = "采购"
        frm1.id = theID
        Set frm1.Rs5 = rsgrid5
        frm1.itemid = rsgrid5!B_ItemID
        frm1.FlatEdit1.Text = rsgrid5!B_ItemIDB
        frm1.FlatEdit6.Text = rsgrid5!B_BreedNum
        frm1.FlatEdit2.Text = rsgrid5!B_GoodsName
        frm1.OriginalProduct = rsgrid5!B_GoodsNameAlias
        frm1.FlatEdit4.Text = rsgrid5!B_Width
        frm1.FlatEdit3.Text = rsgrid5!B_TransfersItemIDB
        frm1.FlatEdit7.Text = IIf(IsNull(rsgrid5!B_suppliers), "", rsgrid5!B_suppliers)
        frm1.Originalsuppliers = IIf(IsNull(rsgrid5!B_suppliersid), "", rsgrid5!B_suppliersid)
        frm1.FlatEdit14.Text = IIf(IsNull(rsgrid5!B_memo), "", rsgrid5!B_memo)
        frm1.FlatEdit17.Text = IIf(IsNull(rsgrid5!B_price), "", rsgrid5!B_price)
        frm1.CheckBox1.Value = IIf(IsNull(rsgrid5!B_supplement), 0, rsgrid5!B_supplement)
        frm1.CheckBox3.Value = IIf(IsNull(rsgrid5!B_checkbox3), 0, rsgrid5!B_checkbox3)
        frm1.Show vbModal
        rsgrid5.requery
End Sub
Private Sub CompositionUPdate_2()
       Dim frm1 As New frmWhite_composition
        frm1.ComboBox1.Text = "调拨"
        frm1.id = theID
        Set frm1.Rs5 = rsgrid5
        frm1.itemid = rsgrid5!B_ItemID
        frm1.FlatEdit1.Text = rsgrid5!B_ItemIDB
        frm1.FlatEdit6.Text = rsgrid5!B_BreedNum
        frm1.FlatEdit2.Text = rsgrid5!B_GoodsName
        frm1.OriginalProduct = rsgrid5!B_GoodsNameAlias
        frm1.FlatEdit4.Text = rsgrid5!B_Width
        frm1.FlatEdit3.Text = rsgrid5!B_TransfersItemIDB
        frm1.FlatEdit7.Text = IIf(IsNull(rsgrid5!B_TransfersSuppliers), "", rsgrid5!B_TransfersSuppliers)
        frm1.Originalsuppliers = IIf(IsNull(rsgrid5!B_TransfersSuppliersid), "", rsgrid5!B_TransfersSuppliersid)
        frm1.FlatEdit14.Text = IIf(IsNull(rsgrid5!B_memo), "", rsgrid5!B_memo)
        frm1.FlatEdit17.Text = IIf(IsNull(rsgrid5!B_price), "", rsgrid5!B_price)
        frm1.CheckBox1.Value = IIf(IsNull(rsgrid5!B_supplement), 0, rsgrid5!B_supplement)
        frm1.CheckBox3.Value = IIf(IsNull(rsgrid5!B_checkbox3), 0, rsgrid5!B_checkbox3)
        frm1.Show vbModal
         rsgrid5.requery
End Sub
Private Sub CompositionUPdate_3()
        Dim rs As New RecordSet
        Dim sql As String
        sql = "select B_Width,B_UnitWeight from G_WhiteComposition where B_itemid='" & rsgrid5!B_ItemID & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
       Dim frm1 As New frmWhite_composition
        frm1.ComboBox2.Text = "采购"
        frm1.id = theID
        Set frm1.Rs5 = rsgrid5
        frm1.Whiteitemid = rsgrid5!B_ItemID
        frm1.FlatEdit8.Text = rsgrid5!B_ItemIDB
        frm1.FlatEdit12.Text = rsgrid5!B_BreedNum
        frm1.FlatEdit9.Text = rsgrid5!B_GoodsName
        frm1.whiteProduct = rsgrid5!B_GoodsNameAlias
        frm1.FlatEdit11.Text = rs!B_Width
        frm1.FlatEdit13.Text = rs!B_UnitWeight
        frm1.FlatEdit5.Text = rsgrid5!B_TransfersItemIDB
        frm1.FlatEdit10.Text = rsgrid5!B_suppliers
        frm1.producerid = rsgrid5!B_suppliersid
         frm1.FlatEdit15.Text = IIf(IsNull(rsgrid5!B_memo), "", rsgrid5!B_memo)
         frm1.FlatEdit16.Text = IIf(IsNull(rsgrid5!B_price), "", rsgrid5!B_price)
         frm1.CheckBox2.Value = IIf(IsNull(rsgrid5!B_supplement), 0, rsgrid5!B_supplement)
        frm1.Show vbModal
         rsgrid5.requery
End Sub
Private Sub CompositionUPdate_4()
       Dim rs As New RecordSet
        Dim sql As String
        sql = "select B_Width,B_UnitWeight from G_WhiteComposition where B_itemid='" & rsgrid5!B_ItemID & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
       Dim frm1 As New frmWhite_composition
        frm1.ComboBox2.Text = "调拨"
        frm1.id = theID
        Set frm1.Rs5 = rsgrid5
        frm1.Whiteitemid = rsgrid5!B_ItemID
        frm1.FlatEdit8.Text = rsgrid5!B_ItemIDB
        frm1.FlatEdit12.Text = rsgrid5!B_BreedNum
        frm1.FlatEdit9.Text = rsgrid5!B_GoodsName
        frm1.whiteProduct = rsgrid5!B_GoodsNameAlias
        frm1.FlatEdit11.Text = rs!B_Width
        frm1.FlatEdit13.Text = rs!B_UnitWeight
        frm1.FlatEdit5.Text = rsgrid5!B_TransfersItemIDB
        frm1.FlatEdit10.Text = rsgrid5!B_TransfersSuppliers
        frm1.producerid = rsgrid5!B_TransfersSuppliersid
        frm1.FlatEdit15.Text = IIf(IsNull(rsgrid5!B_memo), "", rsgrid5!B_memo)
        frm1.FlatEdit16.Text = IIf(IsNull(rsgrid5!B_price), "", rsgrid5!B_price)
        frm1.CheckBox2.Value = IIf(IsNull(rsgrid5!B_supplement), 0, rsgrid5!B_supplement)
        frm1.Show vbModal
         rsgrid5.requery
End Sub
Private Sub rsgrid5_delete()
        If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
            MsgBox "此单已被作废", vbInformation, "提示"
            Exit Sub
        End If
        If yanzhenWhiteComposition(theID) = False Then
            Exit Sub
        End If
        If rsgrid5.RecordCount > 0 Then
        Dim sql As String
        sql = "delete from G_WhiteComposition where B_itemid='" & rsgrid5!B_ItemID & "'"
        Gm.cnnTool.cnn.Execute sql
        End If
         rsgrid5.requery
End Sub
'合同删除进行白坯构成全部删除
Private Sub deleteWhiteComposition()
        Dim sql As String
        sql = "delete from G_WhiteComposition where B_ID='" & theID & "'"
        Gm.cnnTool.cnn.Execute sql
End Sub
'白坯构成的复制行
Private Sub rsgrid5_copy()
        If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
            MsgBox "此单已被作废", vbInformation, "提示"
            Exit Sub
        End If
        If yanzhenWhiteComposition(theID) = False Then
            Exit Sub
        End If
        If rsgrid5.RecordCount <= 0 Then
            Exit Sub
        End If
        
         If rsgrid5!B_Breed = "原料" Then
            If rsgrid5!B_StorageWay = "采购" Then
                    CompositionCopy_1
            Else
                    CompositionCopy_2
            End If
         Else
            If rsgrid5!B_StorageWay = "采购" Then
                    CompositionCopy_3
            Else
                    CompositionCopy_4
            End If
         End If
End Sub
Private Sub CompositionCopy_1()
       Dim frm1 As New frmWhite_composition
        frm1.ComboBox1.Text = "采购"
        frm1.id = theID
        Set frm1.Rs5 = rsgrid5
        frm1.FlatEdit1.Text = rsgrid5!B_ItemIDB
        frm1.FlatEdit6.Text = rsgrid5!B_BreedNum
        frm1.FlatEdit2.Text = rsgrid5!B_GoodsName
        frm1.OriginalProduct = rsgrid5!B_GoodsNameAlias
        frm1.FlatEdit4.Text = rsgrid5!B_Width
        frm1.FlatEdit3.Text = rsgrid5!B_TransfersItemIDB
        frm1.FlatEdit7.Text = rsgrid5!B_suppliers
        frm1.Originalsuppliers = rsgrid5!B_suppliersid
        frm1.FlatEdit14.Text = IIf(IsNull(rsgrid5!B_memo), "", rsgrid5!B_memo)
        frm1.FlatEdit17.Text = IIf(IsNull(rsgrid5!B_price), "", rsgrid5!B_price)
        frm1.CheckBox1.Value = IIf(IsNull(rsgrid5!B_supplement), 0, rsgrid5!B_supplement)
        frm1.Show vbModal
        rsgrid5.requery
End Sub
Private Sub CompositionCopy_2()
       Dim frm1 As New frmWhite_composition
        frm1.ComboBox1.Text = "调拨"
        frm1.id = theID
        Set frm1.Rs5 = rsgrid5
        frm1.FlatEdit1.Text = rsgrid5!B_ItemIDB
        frm1.FlatEdit6.Text = rsgrid5!B_BreedNum
        frm1.FlatEdit2.Text = rsgrid5!B_GoodsName
        frm1.OriginalProduct = rsgrid5!B_GoodsNameAlias
        frm1.FlatEdit4.Text = rsgrid5!B_Width
        frm1.FlatEdit3.Text = rsgrid5!B_TransfersItemIDB
        frm1.FlatEdit7.Text = rsgrid5!B_TransfersSuppliers
        frm1.Originalsuppliers = rsgrid5!B_TransfersSuppliersid
        frm1.FlatEdit14.Text = IIf(IsNull(rsgrid5!B_memo), "", rsgrid5!B_memo)
        frm1.FlatEdit17.Text = IIf(IsNull(rsgrid5!B_price), "", rsgrid5!B_price)
        frm1.CheckBox1.Value = IIf(IsNull(rsgrid5!B_supplement), 0, rsgrid5!B_supplement)
        frm1.Show vbModal
         rsgrid5.requery
End Sub
Private Sub CompositionCopy_3()
        Dim rs As New RecordSet
        Dim sql As String
        sql = "select B_Width,B_UnitWeight from G_WhiteComposition where B_itemid='" & rsgrid5!B_ItemID & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
       Dim frm1 As New frmWhite_composition
        frm1.ComboBox2.Text = "采购"
        frm1.id = theID
        Set frm1.Rs5 = rsgrid5
        frm1.FlatEdit8.Text = rsgrid5!B_ItemIDB
        frm1.FlatEdit12.Text = rsgrid5!B_BreedNum
        frm1.FlatEdit9.Text = rsgrid5!B_GoodsName
        frm1.whiteProduct = rsgrid5!B_GoodsNameAlias
        frm1.FlatEdit11.Text = rs!B_Width
        frm1.FlatEdit13.Text = rs!B_UnitWeight
        frm1.FlatEdit5.Text = rsgrid5!B_TransfersItemIDB
        frm1.FlatEdit10.Text = rsgrid5!B_suppliers
        frm1.producerid = rsgrid5!B_suppliersid
        frm1.FlatEdit15.Text = IIf(IsNull(rsgrid5!B_memo), "", rsgrid5!B_memo)
        frm1.FlatEdit16.Text = IIf(IsNull(rsgrid5!B_price), "", rsgrid5!B_price)
        frm1.CheckBox2.Value = IIf(IsNull(rsgrid5!B_supplement), 0, rsgrid5!B_supplement)
        frm1.Show vbModal
         rsgrid5.requery
End Sub


Private Sub CompositionCopy_4()
       Dim rs As New RecordSet
        Dim sql As String
        sql = "select B_Width,B_UnitWeight from G_WhiteComposition where B_itemid='" & rsgrid5!B_ItemID & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
       Dim frm1 As New frmWhite_composition
        frm1.ComboBox2.Text = "调拨"
        frm1.id = theID
        Set frm1.Rs5 = rsgrid5
        frm1.FlatEdit8.Text = rsgrid5!B_ItemIDB
        frm1.FlatEdit12.Text = rsgrid5!B_BreedNum
        frm1.FlatEdit9.Text = rsgrid5!B_GoodsName
        frm1.whiteProduct = rsgrid5!B_GoodsNameAlias
        frm1.FlatEdit11.Text = rs!B_Width
        frm1.FlatEdit13.Text = rs!B_UnitWeight
        frm1.FlatEdit5.Text = rsgrid5!B_TransfersItemIDB
        frm1.FlatEdit10.Text = rsgrid5!B_TransfersSuppliers
        frm1.producerid = rsgrid5!B_TransfersSuppliersid
        frm1.FlatEdit15.Text = IIf(IsNull(rsgrid5!B_memo), "", rsgrid5!B_memo)
        frm1.FlatEdit16.Text = IIf(IsNull(rsgrid5!B_price), "", rsgrid5!B_price)
        frm1.CheckBox2.Value = IIf(IsNull(rsgrid5!B_supplement), 0, rsgrid5!B_supplement)
        frm1.Show vbModal
         rsgrid5.requery
End Sub

'-----------------------------------------------------------------------合同扫描---------------------------------------------

Private Sub pictureorder()
                On Error GoTo IFERR
                Dim rs As New RecordSet
                Dim sql As String
                Dim clsFile01 As New clsFile
                Dim szPic As String
                
                Dim clsP1 As New clsPicture
                
                sql = "select * from G_Imageorder where B_ID='" & theID & "'"
                rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
                
                If rs.RecordCount > 0 Then
                    szPic = App.Path & "\temp\合同扫描件 " & rs!B_id & ".JPG"
                    clsFile01.DownloadPic rs!B_picture, szPic
                    'Set cls2 = New clsPicture
                    cls2.InitCls szPic, Picture4

                    'Picture4.Picture = LoadPicture(szPic)
                     
                    'PicShow2Ctl rs!B_picture, Picture3
                Else
                    'Picture4.Picture = Nothing
                End If
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "查询是否有图片文件夹" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub

Private Sub OpenPactPic()
    Dim filep1
    filep1 = ShellExecute(0, "open", szFile2, 0, 0, 1)
End Sub
Private Sub OpenPactPic1()
    Dim filep
    filep = ShellExecute(0, "open", szFile, 0, 0, 1)
End Sub

Private Sub Picture4_DblClick()
    OpenPactPic
End Sub
Private Sub Picture3_DblClick()
    OpenPactPic1
End Sub


Private Sub PushButton3_Click()
     If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    On Error GoTo IFERR
    
    With CommonDialog2
        .ShowOpen
   
        szFile2 = .FileName
    End With
    
    If Len(szFile2) <= 0 Then
        Exit Sub
    End If
    
    'Dim cls1 As New clsPicture
    cls2.InitCls szFile2, Picture4
    'Set cls1 = Nothing
    
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "图片格式不支持，请提供JPG或者BMP格式图片！" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub
Private Sub PushButton5_Click()
    Dim sql As String
    Dim rs As New RecordSet
    If szFile2 <> "" Then
        Dim fso As New FileSystemObject
    
        Dim lLength As Long
        Dim oFile As File
    
        Set oFile = fso.GetFile(szFile2)
        
        sql = "select * from G_ImageSize"
        rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
        
        
        If oFile.Size / 1000000 > rs!B_size Then
            MsgBox "图片太大不能上传", vbInformation, "提示"
            Exit Sub
        End If
    
        saveImageorder
        MsgBox "图片上传成功", vbInformation, "提示"
    End If
End Sub

Private Sub saveImageorder()
 
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_Imageorder where 1=0"
    rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
    
     If szFile2 = "" Then
    Else
        Dim rs1 As New RecordSet
        Dim sql1 As String
        sql1 = "select * from G_Imageorder where B_ID='" & theID & "'"
        rs1.Open sql1, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
        If rs1.RecordCount > 0 Then
'            Dim sql2 As String
'            sql2 = "update G_Image set B_Picture='" & szFile & "' where  B_ID='" & theID & "' and B_itemid='" & rsgrid4!B_OrderCode & "'"
'            Gm.cnnToolImage.cnn.Execute sql2
            
            PicSaveToDB rs1!B_picture, szFile2
            rs1.Update
            rs1.Close
            Set rs1 = Nothing
        Else
            rs.AddNew
            rs!B_id = theID
            PicSaveToDB rs!B_picture, szFile2
        
            rs.Update
            rs.Close
            Set rs = Nothing
        End If
        
    End If
End Sub

Private Sub DeletePicture_1()
    Dim sql As String
    sql = "delete from G_imageorder where B_ID='" & theID & "'"
    Gm.cnnToolImage.cnn.Execute sql
End Sub



'---------------------------------------------------------------------------------------------
'色布采购的列表 '
'初始化白坯构成
Private Sub ColorProcure()
    Dim sql As String
    Set rsgrid6 = New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
    sql = "exec usp_SelectColororder_billorder '" & theID & "'"
    rsgrid6.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid6.DataSource = rsgrid6
    setgrid6
End Sub
Private Sub setgrid6()
    TDBGrid6.Columns("B_ItemIDB").Caption = "订单号"
    TDBGrid6.Columns("B_ClientName").Caption = "色布供应商"
    TDBGrid6.Columns("B_Width").Caption = "门幅"
    TDBGrid6.Columns("B_weight").Caption = "克重"
    TDBGrid6.Columns("B_GoodsNameAlias").Caption = "品名"
    TDBGrid6.Columns("B_Name").Caption = "颜色"
    TDBGrid6.Columns("B_SeHao").Caption = "色号"
    TDBGrid6.Columns("B_ps").Caption = "匹数"
    TDBGrid6.Columns("B_kg").Caption = "公斤"
    TDBGrid6.Columns("B_meter").Caption = "米数"
    TDBGrid6.Columns("B_qty").Caption = "码数"
    TDBGrid6.Columns("B_departdate").Caption = "交期"
    TDBGrid6.Columns("B_MemoDetail").Caption = "备注"
    
    TDBGrid6.Columns("B_ItemIDB").width = 900
    TDBGrid6.Columns("B_Width").width = 800
    TDBGrid6.Columns("B_weight").width = 800
    TDBGrid6.Columns("B_SeHao").width = 900
    TDBGrid6.Columns("B_ps").width = 800
    TDBGrid6.Columns("B_kg").width = 800
    TDBGrid6.Columns("B_meter").width = 800
    TDBGrid6.Columns("B_qty").width = 800
    
    TDBGrid6.Columns("B_Clientid").Visible = False
    TDBGrid6.Columns("B_Clientid").AllowSizing = False
    TDBGrid6.Columns("B_Clientid").Locked = True
        TDBGrid6.Columns("B_color").Visible = False
    TDBGrid6.Columns("B_color").AllowSizing = False
    TDBGrid6.Columns("B_color").Locked = True
      TDBGrid6.Columns("B_orderitemid").Visible = False
    TDBGrid6.Columns("B_orderitemid").AllowSizing = False
    TDBGrid6.Columns("B_orderitemid").Locked = True
    
    TDBGrid6.MarqueeStyle = dbgHighlightRow
    TDBGrid6.HoldFields
End Sub
Private Sub TDBGrid6_DblClick()
    If TDBGrid6.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim frm1 As New frmOrderColor
    frm1.id = rsgrid6!B_orderitemid
    frm1.Show vbModal
    rsgrid6.requery
End Sub



'----------------------------------------------箱数计划------------------------------------------------
Private Sub boxPlan()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "exec usp_BoxPlan '" & theID & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    setRs
    Do While Not rs.EOF
        rsgrid7.AddNew
        rsgrid7!B_id = rs!B_id
        rsgrid7!B_ordercode = rs!B_ordercode
        rsgrid7!B_GoodsID = rs!B_GoodsID
        rsgrid7!B_GoodsName = rs!B_GoodsName
        rsgrid7!B_size = IIf(IsNull(rs!B_size), "", rs!B_size)
        rsgrid7!B_Width = rs!B_Width
        rsgrid7!B_weight = rs!B_weight
        rsgrid7!B_patterncode = rs!B_patterncode
        rsgrid7!B_colorid = rs!B_colorid
        rsgrid7!B_color = rs!B_color
        rsgrid7!B_hex = rs!B_hex
        rsgrid7!B_CasePack = rs!B_CasePack
        rsgrid7!B_boxname = rs!B_boxname
        rsgrid7!B_boxgg = rs!B_boxgg
        rsgrid7!B_orderitemid = rs!B_orderitemid
        rsgrid7!B_OrderID = rs!B_OrderID
        rsgrid7.Update
        rs.movenext
    Loop
End Sub
Private Sub ActiveBar26_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    
        Case "新增"
            boxadd
        Case "修改"
            boxupd
        Case "删除"
            boxdel
        Case "保存"
            boxsave
        Case "打印"
            Printbox
    End Select
End Sub
Private Sub TDBGrid7_FetchCellStyle(ByVal Condition As Integer, _
    ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, _
    ByVal CellStyle As TrueOleDBGrid80.StyleDisp)

'    Dim ys As RGB
    On Error Resume Next
    Debug.Print TDBGrid7.Columns("B_Hex").CellValue(bookmark)
    CellStyle.BackColor = TDBGrid7.Columns("B_Hex").CellValue(bookmark)
     CellStyle.ForeColor = TDBGrid7.Columns("B_Hex").CellValue(bookmark)
End Sub
'箱数计划新增
Private Sub boxadd()
    On Error Resume Next
    
        If rsdetail.RecordCount <= 0 Then
            Exit Sub
        End If
       rsgrid7.AddNew
        rsgrid7!B_id = ""
        rsgrid7!B_ordercode = rsdetail!B_ordercode
        rsgrid7!B_GoodsID = rsdetail!B_GoodsID
        rsgrid7!B_GoodsName = rsdetail!B_name
        rsgrid7!B_size = rsdetail!B_size
        rsgrid7!B_Width = rsdetail!B_Width
        rsgrid7!B_weight = rsdetail!B_weight
        rsgrid7!B_patterncode = rsdetail!B_patterncode
        rsgrid7!B_colorid = rsdetail!B_sid
        rsgrid7!B_color = rsdetail!B_color
        rsgrid7!B_hex = rsdetail!B_hex
        rsgrid7!B_CasePack = rsdetail!B_CasePack
        rsgrid7!B_boxname = rsdetail!B_boxname
        rsgrid7!B_boxgg = rsdetail!B_boxgg
        rsgrid7!B_orderitemid = rsdetail!B_ItemID
        rsgrid7!B_OrderID = theID
        rsgrid7.Update
    C1Tab1.CurrTab = 7
End Sub
'箱数计划新增
Private Sub boxall()
    On Error Resume Next
    
        If rsdetail.RecordCount <= 0 Then
            Exit Sub
        End If
        
        
        rsdetail.MoveFirst
        Do While Not rsdetail.EOF
            rsgrid7.AddNew
             rsgrid7!B_id = ""
             rsgrid7!B_ordercode = rsdetail!B_ordercode
             rsgrid7!B_GoodsID = rsdetail!B_GoodsID
             rsgrid7!B_GoodsName = rsdetail!B_name
             rsgrid7!B_size = rsdetail!B_size
             rsgrid7!B_Width = rsdetail!B_Width
             rsgrid7!B_weight = rsdetail!B_weight
             rsgrid7!B_patterncode = rsdetail!B_patterncode
             rsgrid7!B_colorid = rsdetail!B_sid
             rsgrid7!B_color = rsdetail!B_color
             rsgrid7!B_hex = rsdetail!B_hex
             rsgrid7!B_CasePack = rsdetail!B_CasePack
             rsgrid7!B_boxname = rsdetail!B_boxname
             rsgrid7!B_boxgg = rsdetail!B_boxgg
             rsgrid7!B_orderitemid = rsdetail!B_ItemID
             rsgrid7!B_OrderID = theID
             rsgrid7.Update
             
             rsdetail.movenext
        Loop
    C1Tab1.CurrTab = 7
    If rsgrid7.RecordCount > 0 Then
        rsgrid7.MoveFirst
    End If
End Sub


'箱数计划修改
Private Sub boxupd()
    On Error Resume Next
    If rsgrid7.RecordCount <= 0 Then
        Exit Sub
    End If
    Dim rs As New RecordSet
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim sql1 As String
    If Val(rsgrid7!B_CasePack) <= 0 Or rsgrid7!B_boxname = "" Then
            Exit Sub
    End If
    
    
    sql = "select * from  G_billbox where B_id='" & rsgrid7!B_id & "' "
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount <= 0 Then
        MsgBox "此行数据还没有保存", vbInformation, "提示"
        Exit Sub
    Else
        sql1 = "update G_billbox set B_CasePack='" & rsgrid7!B_CasePack & "',B_boxname='" & rsgrid7!B_boxname & "',B_boxgg='" & rsgrid7!B_boxgg & "',B_memo='" & rsgrid7!B_memo & "' where B_id='" & rsgrid7!B_id & "'"
        rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    End If
    Debug.Print sql1
End Sub
'箱数计划删除
Private Sub boxdel()
    On Error Resume Next
    If rsgrid7.RecordCount <= 0 Then
        Exit Sub
    End If
    
    Dim sql As String
    Dim rs As New RecordSet
    If MsgBox("是否删除", vbInformation + vbYesNo + vbDefaultButton1, "提示") = vbYes Then
        sql = "delete from G_billbox where B_id='" & rsgrid7!B_id & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        rsgrid7.delete
        If rsgrid7.RecordCount > 0 Then
            rsgrid7.MoveFirst
        End If
    End If
End Sub

Private Sub TDBGrid7_BeforeColUpdate(ByVal colIndex As Integer, OldValue As Variant, Cancel As Integer)
    If colIndex = TDBGrid7.Columns("B_CasePack").colIndex Then
        TDBGrid7.Columns("B_CasePack").Value = Abs(Val(TDBGrid7.Columns("B_CasePack").Value))
    End If
End Sub
'全部保存的时候保存箱数计划
Private Sub boxsave()
    Dim rs As New RecordSet
    Set rs = rsgrid7.Clone
       Dim rs1 As RecordSet
    Dim sql1 As String
    Dim rs2 As RecordSet
    Dim sql2 As String
    Dim rs3 As RecordSet
    Dim sql3 As String
    
    TDBGrid7.Update
    
    If rsgrid7.RecordCount <= 0 Then
        Exit Sub
    End If
    If boxbool = False Then
        
        Exit Sub
    End If
    
    rs.MoveFirst
    Do While Not rs.EOF
        Set rs1 = New RecordSet
        sql1 = "select * from G_billbox where B_id='" & rs!B_id & "'"
        rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        If rs1.RecordCount > 0 Then
            Set rs2 = New RecordSet
            sql2 = "update G_billbox set B_CasePack='" & rs!B_CasePack & "',B_boxname='" & rs!B_boxname & "',B_boxgg='" & rs!B_boxgg & "',B_memo='" & rs!B_memo & "',B_size='" & rs!B_size & "' where B_id='" & rs!B_id & "'"
            rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        Else
             Set rs3 = New RecordSet
             sql3 = "select * from G_billbox where 1=1"
             rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
             rs3.AddNew
            rs3!B_ordercode = rs!B_ordercode
            rs3!B_size = rs!B_size
            rs3!B_GoodsID = rs!B_GoodsID
            rs3!B_Width = rs!B_Width
            rs3!B_weight = rs!B_weight
            rs3!B_patterncode = rs!B_patterncode
            rs3!B_colorid = rs!B_colorid
            rs3!B_color = rs!B_color
            rs3!B_CasePack = rs!B_CasePack
            rs3!B_boxname = rs!B_boxname
            rs3!B_boxgg = rs!B_boxgg
            rs3!B_memo = rs!B_memo
            rs3!B_orderitemid = rs!B_orderitemid
            rs3!B_OrderID = rs!B_OrderID
            rs3.Update
        End If
        rs.movenext
    Loop
    boxPlan
End Sub
Private Function boxbool() As Boolean
    boxbool = False
    Dim a As String
    Dim b As String
    Dim rs As New RecordSet
    Dim rs1 As New RecordSet
    Set rs = rsgrid7.Clone
    Set rs1 = rsgrid7.Clone
    
    rs.MoveFirst
    Do While Not rs.EOF
        If Val(rs!B_CasePack) <= 0 Or rs!B_boxname = "" Then
            MsgBox "有数据为空,不能保存", vbInformation, "提示"
            boxbool = False
            Exit Function
        End If
        rs.movenext
    Loop
    
    rs.MoveFirst
    Do While Not rs.EOF
        a = rs!B_boxname
        b = rs!B_boxgg
        rs1.MoveFirst
        Do While Not rs1.EOF
            If a = rs1!B_boxname Then
                If b <> rs1!B_boxgg Then
                    MsgBox "同箱名称下箱规格不一样,不能保存", vbInformation, "提示"
                    boxbool = False
                    Exit Function
                End If
            End If
            rs1.movenext
        Loop
        rs.movenext
    Loop
    
    
    
    
    boxbool = True
    
End Function

'打箱计划的打印
Private Sub Printbox()
    Dim sql As String
    Dim rs As New RecordSet
'    sql = "SELECT * FROM G_billbox WHERE B_orderid='" & theid & "'"""
'    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    sql = "exec usp_billboxPrint'" & theID & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
         Dim frm1 As New frmModBLRPreviewOriColor
        Set frm1.RecordSet = rs.Clone
            
        frm1.ObjectID = "22B137"
        frm1.Show vbModal
    End If
End Sub




'-------------------------------------------辅料计划---------------------------------------------

Private Sub auxiliary()
    Dim sql As String
    Set rsgrid8 = New RecordSet
    sql = "exec usp_SelectAuxiliary '" & theID & "'"
    rsgrid8.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid8.DataSource = rsgrid8
    setgrid8
    If rsgrid8.RecordCount > 0 Then
        rsgrid8.MoveFirst
    End If
End Sub
Private Sub setgrid8()
   TDBGrid8.Columns("B_Name").Caption = "辅料名称"
    TDBGrid8.Columns("B_Alias").Caption = "别名"
    TDBGrid8.Columns("B_specifications").Caption = "规格"
    TDBGrid8.Columns("B_ClientName").Caption = "供应商"
    TDBGrid8.Columns("B_Qty").Caption = "数量"
    TDBGrid8.Columns("B_BoxQty").Caption = "箱数"
    TDBGrid8.Columns("B_Price").Caption = "单价"
    TDBGrid8.Columns("B_Sum").Caption = "金额"
    TDBGrid8.Columns("B_Memo").Caption = "备注"
    
    TDBGrid8.Columns("B_Price").NumberFormat = "0.00"
    TDBGrid8.Columns("B_Sum").NumberFormat = "0.00"
    
    TDBGrid8.Columns("B_id").Visible = False
    TDBGrid8.Columns("B_id").AllowSizing = False
    TDBGrid8.Columns("B_id").Locked = True
    
    TDBGrid8.Columns("B_autoid").Visible = False
    TDBGrid8.Columns("B_autoid").AllowSizing = False
    TDBGrid8.Columns("B_autoid").Locked = True
        TDBGrid8.Columns("B_auxiliary").Visible = False
    TDBGrid8.Columns("B_auxiliary").AllowSizing = False
    TDBGrid8.Columns("B_auxiliary").Locked = True
        TDBGrid8.Columns("B_ClientID").Visible = False
    TDBGrid8.Columns("B_ClientID").AllowSizing = False
    TDBGrid8.Columns("B_ClientID").Locked = True
    
    TDBGrid8.MarqueeStyle = dbgHighlightRow
    TDBGrid8.HoldFields
    sumall8
End Sub
Private Sub ActiveBar27_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    
    
        Case "新增"
            auxiliary_add
        Case "复制"
            auxiliary_copy
        Case "删除"
           auxiliary_del
    End Select
End Sub
Private Sub auxiliary_del()
    If TDBGrid8.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim sql As String
    Dim rs As New RecordSet
    If MsgBox("是否删除", vbInformation + vbYesNo + vbDefaultButton1, "提示") = vbYes Then
        sql = "delete from G_auxiliary where B_id='" & rsgrid8!B_id & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        rsgrid8.requery
        If rsgrid8.RecordCount > 0 Then
            rsgrid8.MoveFirst
        End If
    End If
End Sub
Private Sub auxiliary_add()
    Dim frm1 As New frmOrderProduct_auxiliary
    frm1.autoid = theID
    frm1.Show vbModal
    rsgrid8.requery
    Unload frm1
End Sub
Private Sub auxiliary_copy()
    On Error Resume Next
    
    If TDBGrid8.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    Dim frm1 As New frmOrderProduct_auxiliary
    frm1.autoid = theID
    frm1.auxiliaryid = rsgrid8!B_auxiliary
    frm1.FlatEdit2.Text = rsgrid8!B_name
    frm1.FlatEdit3.Text = rsgrid8!B_Alias
    frm1.FlatEdit4.Text = rsgrid8!B_specifications
    frm1.Positiveid = rsgrid8!B_Clientid
    frm1.FlatEdit17.Text = rsgrid8!B_ClientName
    frm1.FlatEdit7.Text = rsgrid8!B_qty
    frm1.FlatEdit8.Text = rsgrid8!B_BoxQty
    frm1.FlatEdit9.Text = rsgrid8!B_price
    frm1.FlatEdit10.Text = rsgrid8!B_Sum
    frm1.FlatEdit13.Text = rsgrid8!B_memo
    
    frm1.Show vbModal
    rsgrid8.requery
    Unload frm1
End Sub

Private Sub sumall8()
    Dim rs As New RecordSet
    Set rs = rsgrid8.Clone
    
    Dim a As Long
    Dim b As Long
    Dim c As Double
    Dim d As String
    a = 0
    b = 0
    c = 0
   
  
'    rs.MoveFirst
    Do While Not rs.EOF
        a = a + IIf(IsNull(rs!B_qty), 0, rs!B_qty)
        b = b + IIf(IsNull(rs!B_BoxQty), 0, rs!B_BoxQty)
        c = c + IIf(IsNull(rs!B_Sum), 0, rs!B_Sum)
        
        rs.movenext
    Loop
  
    d = Format(c, "0.00")
    TDBGrid8.Columns("B_Name").FooterText = "合计"
    TDBGrid8.Columns("B_qty").FooterText = "" & a & ""
    TDBGrid8.Columns("B_BoxQty").FooterText = "" & b & ""
   
    TDBGrid8.Columns("B_Sum").FooterText = "" & d & ""
End Sub

Private Sub TDBGrid8_DblClick()
    If TDBGrid8.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    Dim frm1 As New frmOrderProduct_auxiliary
    frm1.autoid = rsgrid8!B_autoid
    frm1.id = rsgrid8!B_id
    frm1.auxiliaryid = rsgrid8!B_auxiliary
    frm1.FlatEdit2.Text = rsgrid8!B_name
    frm1.FlatEdit3.Text = rsgrid8!B_Alias
    frm1.FlatEdit4.Text = rsgrid8!B_specifications
    frm1.Positiveid = rsgrid8!B_Clientid
    frm1.FlatEdit17.Text = rsgrid8!B_ClientName
    frm1.FlatEdit7.Text = rsgrid8!B_qty
    frm1.FlatEdit8.Text = rsgrid8!B_BoxQty
    frm1.FlatEdit9.Text = rsgrid8!B_price
    frm1.FlatEdit10.Text = rsgrid8!B_Sum
    frm1.FlatEdit13.Text = rsgrid8!B_memo
    
    frm1.Show vbModal
    rsgrid8.requery
    Unload frm1
End Sub


'----------------------------------------------裁剪计划------------------------------------------------
Private Sub tailorPlan()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "exec usp_tailorPlan '" & theID & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    setRs1
    
    Do While Not rs.EOF
        rsgrid9.AddNew
        rsgrid9!B_id = rs!B_id
        rsgrid9!B_itemCode = rs!B_itemCode
        rsgrid9!B_label = rs!B_label
        rsgrid9!B_size = rs!B_size
        rsgrid9!B_colorid = rs!B_colorid
        rsgrid9!B_color = rs!B_color
        rsgrid9!B_hex = rs!B_hex
        rsgrid9!B_BarCode = rs!B_BarCode
        rsgrid9!B_ChiCun = rs!B_ChiCun
        rsgrid9!B_qtyall = rs!B_qtyall
        rsgrid9!B_quantity = rs!B_quantity
        rsgrid9!B_BoxQty = rs!B_BoxQty
        rsgrid9!B_memo = rs!B_memo
        rsgrid9!B_orderitemid = rs!B_orderitemid
        rsgrid9!B_OrderID = rs!B_OrderID
        rsgrid9.Update
        rs.movenext
    Loop
End Sub
Private Sub ActiveBar28_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "修改"
            tailorupd
        Case "删除"
            tailordel
        Case "保存"
            tailorsave
        Case "此行生成缝制计划"
            cjtofzaddone
        Case "全部生成缝制计划"
            cjtofzaddall
        Case "打印"
            Printtailor
    End Select
End Sub
Private Sub TDBGrid9_FetchCellStyle(ByVal Condition As Integer, _
    ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, _
    ByVal CellStyle As TrueOleDBGrid80.StyleDisp)

'    Dim ys As RGB
    On Error Resume Next
    Debug.Print TDBGrid9.Columns("B_Hex").CellValue(bookmark)
    CellStyle.BackColor = TDBGrid9.Columns("B_Hex").CellValue(bookmark)
     CellStyle.ForeColor = TDBGrid9.Columns("B_Hex").CellValue(bookmark)
End Sub

'裁剪计划修改
Private Sub tailorupd()
    On Error Resume Next
    If rsgrid9.RecordCount <= 0 Then
        Exit Sub
    End If
    Dim rs As New RecordSet
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim sql1 As String
   
    sql = "select * from  G_Billtailor where B_id='" & rsgrid9!B_id & "' "
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount <= 0 Then
        MsgBox "此行数据还没有保存", vbInformation, "提示"
        Exit Sub
    Else
        sql1 = "update G_Billtailor set B_itemCode='" & rsgrid9!B_itemCode & "',B_label='" & rsgrid9!B_label & "',B_size='" & rsgrid9!B_size & "',B_BarCode='" & rsgrid9!B_BarCode & "' "
        sql1 = sql1 & ",B_ChiCun='" & rsgrid9!B_ChiCun & "',B_qtyall='" & rsgrid9!B_qtyall & "',B_quantity='" & rsgrid9!B_quantity & "'"
        sql1 = sql1 & ",B_boxqty='" & rsgrid9!B_BoxQty & "',B_memo='" & rsgrid9!B_memo & "' where B_id='" & rsgrid9!B_id & "'"
        rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    End If
    
End Sub
'裁剪计划删除
Private Sub tailordel()
    On Error Resume Next
    If rsgrid9.RecordCount <= 0 Then
        Exit Sub
    End If
    
    Dim sql As String
    Dim rs As New RecordSet
    If MsgBox("是否删除", vbInformation + vbYesNo + vbDefaultButton1, "提示") = vbYes Then
        sql = "delete from G_Billtailor where B_id='" & rsgrid9!B_id & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        rsgrid9.delete
        If rsgrid9.RecordCount > 0 Then
            rsgrid9.MoveFirst
        End If
    End If
End Sub

Private Sub TDBGrid9_BeforeColUpdate(ByVal colIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim a As Long
    Dim b As Double
    If colIndex = TDBGrid9.Columns("B_qtyall").colIndex Then
        If Abs(Val(TDBGrid9.Columns("B_quantity").Value)) > 0 Then
            TDBGrid9.Columns("B_qtyall").Value = Abs(Val(TDBGrid9.Columns("B_qtyall").Value))
               a = Abs(Val(TDBGrid9.Columns("B_qtyall").Value)) / Abs(Val(TDBGrid9.Columns("B_quantity").Value))
                b = Abs(Val(TDBGrid9.Columns("B_qtyall").Value)) / Abs(Val(TDBGrid9.Columns("B_quantity").Value))
                If b > a Then
                    TDBGrid9.Columns("B_boxqty").Value = a + 1
                Else
                    TDBGrid9.Columns("B_boxqty").Value = a
                End If
        End If
    End If
    If colIndex = TDBGrid9.Columns("B_quantity").colIndex Then
        TDBGrid9.Columns("B_quantity").Value = Abs(Val(TDBGrid9.Columns("B_quantity").Value))
        
        a = Abs(Val(TDBGrid9.Columns("B_qtyall").Value)) / Abs(Val(TDBGrid9.Columns("B_quantity").Value))
        b = Abs(Val(TDBGrid9.Columns("B_qtyall").Value)) / Abs(Val(TDBGrid9.Columns("B_quantity").Value))
        If b > a Then
            TDBGrid9.Columns("B_boxqty").Value = a + 1
         Else
                    TDBGrid9.Columns("B_boxqty").Value = a
        End If
    End If
End Sub
'全部保存的时候保存裁剪计划
Private Sub tailorsave()
    Dim rs As New RecordSet
    Set rs = rsgrid9.Clone
       Dim rs1 As RecordSet
    Dim sql1 As String
    Dim rs2 As RecordSet
    Dim sql2 As String
    Dim rs3 As RecordSet
    Dim sql3 As String
    
    
    If rsgrid9.RecordCount <= 0 Then
        Exit Sub
    End If
    If tailorbool = False Then
        
        Exit Sub
    End If
    
 TDBGrid9.Update
 
    rs.MoveFirst
    Do While Not rs.EOF
        Set rs1 = New RecordSet
        sql1 = "select * from G_Billtailor where B_id='" & rs!B_id & "'"
        rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        Debug.Print rs.RecordCount
        If rs1.RecordCount > 0 Then
            Set rs2 = New RecordSet
            sql2 = "update G_Billtailor set B_itemCode='" & rs!B_itemCode & "',B_label='" & rs!B_label & "',B_size='" & rs!B_size & "',B_BarCode='" & rs!B_BarCode & "' "
            sql2 = sql2 & ",B_ChiCun='" & rs!B_ChiCun & "',B_qtyall='" & rs!B_qtyall & "',B_quantity='" & rs!B_quantity & "'"
            sql2 = sql2 & ",B_boxqty='" & rs!B_BoxQty & "',B_memo='" & rs!B_memo & "' where B_id='" & rs!B_id & "'"
            rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        Else
             Set rs3 = New RecordSet
             sql3 = "select * from G_Billtailor where 1=1"
             rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            rs3.AddNew
'            rs3!B_ID = rsgrid9!B_ID
            rs3!B_itemCode = rs!B_itemCode
            rs3!B_label = rs!B_label
            rs3!B_size = rs!B_size
            rs3!B_colorid = rs!B_colorid
            rs3!B_color = rs!B_color
'            rs3!B_hex = rsgrid9!B_hex
            rs3!B_BarCode = rs!B_BarCode
            rs3!B_ChiCun = rs!B_ChiCun
            rs3!B_qtyall = rs!B_qtyall
            rs3!B_quantity = rs!B_quantity
            rs3!B_BoxQty = rs!B_BoxQty
            rs3!B_memo = rs!B_memo
            rs3!B_orderitemid = rs!B_orderitemid
            rs3!B_OrderID = rs!B_OrderID
            rs3.Update
        End If
        rs.movenext
    Loop
    tailorPlan
End Sub
Private Function tailorbool() As Boolean
    tailorbool = False

    Dim rs As New RecordSet
    Set rs = rsgrid9.Clone


    rs.MoveFirst
    Do While Not rs.EOF
        If Val(rs!B_qtyall) <= 0 Then
            MsgBox "数量小于0,不能保存", vbInformation, "提示"
            tailorbool = False
            Exit Function
        End If
        rs.movenext
    Loop

    tailorbool = True
    
End Function
'网格右键
Private Sub TDBGrid9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ActiveBar28.Bands("网格右键").PopupMenu
    End If
End Sub
'生成缝制计划
Private Sub cjtofzaddone()
    On Error Resume Next
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If ActiveBar22.Bands("Band1").Tools("审核图片").Visible = True Then
        rsgrid10.AddNew
        rsgrid10!B_hex = rsgrid9!B_hex
        rsgrid10!B_colorid = rsgrid9!B_colorid
        rsgrid10!B_color = rsgrid9!B_color
        rsgrid10!B_orderitemid = rsgrid9!B_orderitemid
        
        rsgrid10!B_itemCode = rsgrid9!B_itemCode
        rsgrid10!B_label = rsgrid9!B_label
        rsgrid10!B_size = rsgrid9!B_size
        rsgrid10!B_BarCode = rsgrid9!B_BarCode
        rsgrid10!B_ChiCun = rsgrid9!B_ChiCun
        rsgrid10!B_qtyall = rsgrid9!B_qtyall
        rsgrid10!B_quantity = rsgrid9!B_quantity
        rsgrid10!B_BoxQty = rsgrid9!B_BoxQty
        rsgrid10!B_memo = rsgrid9!B_memo
    
        rsgrid10!B_OrderID = theID
        rsgrid10.Update
        C1Tab1.CurrTab = 10
    Else
        MsgBox "此单没有审核", vbInformation, "提示"
        Exit Sub
    End If
    
End Sub
Private Sub cjtofzaddall()
    Dim rs As New RecordSet
    Set rs = rsgrid9.Clone
      On Error Resume Next
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    If ActiveBar22.Bands("Band1").Tools("作废图片").Visible = True Then
        MsgBox "此单已被作废", vbInformation, "提示"
        Exit Sub
    End If
    If ActiveBar22.Bands("Band1").Tools("审核图片").Visible = True Then
        rs.MoveFirst
        Do While Not rs.EOF
            rsgrid10.AddNew
            rsgrid10!B_hex = rs!B_hex
            rsgrid10!B_colorid = rs!B_colorid
            rsgrid10!B_color = rs!B_color
            
            rsgrid10!B_itemCode = rs!B_itemCode
            rsgrid10!B_label = rs!B_label
            rsgrid10!B_size = rs!B_size
            rsgrid10!B_BarCode = rs!B_BarCode
            rsgrid10!B_ChiCun = rs!B_ChiCun
            rsgrid10!B_qtyall = rs!B_qtyall
            rsgrid10!B_quantity = rs!B_quantity
            rsgrid10!B_BoxQty = rs!B_BoxQty
            rsgrid10!B_memo = rs!B_memo
            
            rsgrid10!B_orderitemid = rs!B_orderitemid
            rsgrid10!B_OrderID = theID
            rsgrid10.Update
             rs.movenext
        Loop
        C1Tab1.CurrTab = 10
    Else
        MsgBox "此单没有审核", vbInformation, "提示"
        Exit Sub
    End If
End Sub
'裁剪计划打印
Private Sub Printtailor()
    
    Dim sql As String
    Dim rs As New RecordSet
'    sql = "SELECT * FROM G_Billtailor WHERE B_orderid='" & theid & "'"""
'    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
sql = "exec usp_BilltailorPrint'" & theID & "'"
rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
         Dim frm1 As New frmModBLRPreviewOriColor
        Set frm1.RecordSet = rs.Clone
            
        frm1.ObjectID = "22B135"
        frm1.Show vbModal
    End If
End Sub



'----------------------------------------------缝制计划------------------------------------------------
Private Sub sewPlan()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "exec usp_sewPlan '" & theID & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    setRs2
    
    Do While Not rs.EOF
        rsgrid10.AddNew
        rsgrid10!B_id = rs!B_id
        rsgrid10!B_itemCode = rs!B_itemCode
        rsgrid10!B_label = rs!B_label
        rsgrid10!B_size = rs!B_size
        rsgrid10!B_colorid = rs!B_colorid
        rsgrid10!B_color = rs!B_color
        
        rsgrid10!B_hex = IIf(IsNull(rs!B_hex), "", rs!B_hex)
        rsgrid10!B_BarCode = rs!B_BarCode
        rsgrid10!B_ChiCun = rs!B_ChiCun
        rsgrid10!B_qtyall = rs!B_qtyall
        rsgrid10!B_quantity = rs!B_quantity
        rsgrid10!B_BoxQty = rs!B_BoxQty
        rsgrid10!B_memo = rs!B_memo
        rsgrid10!B_orderitemid = rs!B_orderitemid
        rsgrid10!B_OrderID = rs!B_OrderID
        rsgrid10!B_process = rs!B_process
        rsgrid10!B_KuanHao = IIf(IsNull(rs!B_KuanHao), "", rs!B_KuanHao)
        rsgrid10.Update
        rs.movenext
    Loop
End Sub
Private Sub ActiveBar29_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    
      
        Case "修改"
            sewupd
        Case "删除"
            sewdel
        Case "保存"
            sewsave
        Case "打印"
            Printsew
        Case "洗标和织标位置"
              WAW
        Case "水洗标内容"
             washing
        Case "洗涤标位置"
            Woven
        
    End Select
End Sub
Private Sub TDBGrid10_FetchCellStyle(ByVal Condition As Integer, _
    ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, _
    ByVal CellStyle As TrueOleDBGrid80.StyleDisp)

'    Dim ys As RGB
    On Error Resume Next
    Debug.Print TDBGrid10.Columns("B_Hex").CellValue(bookmark)
    CellStyle.BackColor = TDBGrid10.Columns("B_Hex").CellValue(bookmark)
     CellStyle.ForeColor = TDBGrid10.Columns("B_Hex").CellValue(bookmark)
End Sub

'裁剪计划修改
Private Sub sewupd()
    On Error Resume Next
    If rsgrid10.RecordCount <= 0 Then
        Exit Sub
    End If
    Dim rs As New RecordSet
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim sql1 As String
   
    sql = "select * from  G_BillSew where B_id='" & rsgrid10!B_id & "' "
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount <= 0 Then
        MsgBox "此行数据还没有保存", vbInformation, "提示"
        Exit Sub
    Else
        sql1 = "update G_BillSew set B_itemCode='" & rsgrid10!B_itemCode & "',B_label='" & rsgrid10!B_label & "',B_size='" & rsgrid10!B_size & "',B_BarCode='" & rsgrid10!B_BarCode & "' "
        sql1 = sql1 & ",B_ChiCun='" & rsgrid10!B_ChiCun & "',B_qtyall='" & rsgrid10!B_qtyall & "',B_quantity='" & rsgrid10!B_quantity & "'"
        sql1 = sql1 & ",B_boxqty='" & rsgrid10!B_BoxQty & "',B_memo='" & rsgrid10!B_memo & "',B_process='" & rsgrid10!B_process & "' where B_id='" & rsgrid10!B_id & "'"
        rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    End If
    
End Sub
'裁剪计划删除
Private Sub sewdel()
    On Error Resume Next
    If rsgrid10.RecordCount <= 0 Then
        Exit Sub
    End If
    
    Dim sql As String
    Dim rs As New RecordSet
    If MsgBox("是否删除", vbInformation + vbYesNo + vbDefaultButton1, "提示") = vbYes Then
        sql = "delete from G_BillSew where B_id='" & rsgrid10!B_id & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        rsgrid10.delete
        If rsgrid10.RecordCount > 0 Then
            rsgrid10.MoveFirst
        End If
    End If
End Sub

Private Sub TDBGrid10_BeforeColUpdate(ByVal colIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim a As Long
    Dim b As Double
    If colIndex = TDBGrid10.Columns("B_qtyall").colIndex Then
        If Abs(Val(TDBGrid10.Columns("B_quantity").Value)) > 0 Then
            TDBGrid10.Columns("B_qtyall").Value = Abs(Val(TDBGrid10.Columns("B_qtyall").Value))
               a = Abs(Val(TDBGrid10.Columns("B_qtyall").Value)) / Abs(Val(TDBGrid10.Columns("B_quantity").Value))
                b = Abs(Val(TDBGrid10.Columns("B_qtyall").Value)) / Abs(Val(TDBGrid10.Columns("B_quantity").Value))
                If b > a Then
                    TDBGrid10.Columns("B_boxqty").Value = a + 1
                Else
                    TDBGrid10.Columns("B_boxqty").Value = a
                End If
        End If
    End If
    If colIndex = TDBGrid10.Columns("B_quantity").colIndex Then
        TDBGrid10.Columns("B_quantity").Value = Abs(Val(TDBGrid10.Columns("B_quantity").Value))
        
        a = Abs(Val(TDBGrid10.Columns("B_qtyall").Value)) / Abs(Val(TDBGrid10.Columns("B_quantity").Value))
        b = Abs(Val(TDBGrid10.Columns("B_qtyall").Value)) / Abs(Val(TDBGrid10.Columns("B_quantity").Value))
        If b > a Then
            TDBGrid10.Columns("B_boxqty").Value = a + 1
         Else
                    TDBGrid10.Columns("B_boxqty").Value = a
        End If
    End If
End Sub
'全部保存的时候保存裁剪计划
Private Sub sewsave()
    Dim rs As New RecordSet
    Set rs = rsgrid10.Clone  '将缝制计划的记录集复制到 rs 中
    Dim rs1 As RecordSet
    Dim sql1 As String
    Dim rs2 As RecordSet
    Dim sql2 As String
    Dim rs3 As RecordSet
    Dim sql3 As String
    
    If rsgrid10.RecordCount <= 0 Then
        Exit Sub
    End If
    If sewbool = False Then
        
        Exit Sub
    End If
    
 TDBGrid10.Update
 
    rs.MoveFirst
    Do While Not rs.EOF
        Set rs1 = New RecordSet
        
        sql1 = "select * from G_BillSew where B_id='" & rs!B_id & "'"
        rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        Debug.Print rs.RecordCount
        If rs1.RecordCount > 0 Then
            Set rs2 = New RecordSet
            sql2 = "update G_BillSew set B_itemCode='" & rs!B_itemCode & "',B_label='" & rs!B_label & "',B_size='" & rs!B_size & "',B_BarCode='" & rs!B_BarCode & "' "
            sql2 = sql2 & ",B_ChiCun='" & rs!B_ChiCun & "',B_qtyall='" & rs!B_qtyall & "',B_quantity='" & rs!B_quantity & "'"
            sql2 = sql2 & ",B_boxqty='" & rs!B_BoxQty & "',B_memo='" & rs!B_memo & "',B_process='" & rs!B_process & "',B_color='" & rs!B_color & "',B_colorid='" & rs!B_colorid & "'"
            sql2 = sql2 & ",B_KuanHao='" & rs!B_KuanHao & "' where B_id='" & rs!B_id & "'"
            rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        Else
            Set rs3 = New RecordSet
             sql3 = "select * from G_BillSew where 1=1"
             rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            rs3.AddNew

            rs3!B_itemCode = rs!B_itemCode
            rs3!B_label = rs!B_label
            rs3!B_size = rs!B_size
            rs3!B_colorid = rs!B_colorid
            rs3!B_color = rs!B_color
'            rs3!B_hex = rsgrid9!B_hex
            rs3!B_BarCode = rs!B_BarCode
            rs3!B_ChiCun = rs!B_ChiCun
            rs3!B_qtyall = rs!B_qtyall
            rs3!B_quantity = rs!B_quantity
            rs3!B_BoxQty = rs!B_BoxQty
            rs3!B_memo = rs!B_memo
            rs3!B_orderitemid = rs!B_orderitemid
            rs3!B_OrderID = rs!B_OrderID
            rs3!B_process = rs!B_process
            rs3!B_KuanHao = rs!B_KuanHao
            rs3.Update
        End If
        rs.movenext
    Loop
    sewPlan
End Sub
Private Function sewbool() As Boolean
    sewbool = False

    Dim rs As New RecordSet
    Set rs = rsgrid10.Clone


    rs.MoveFirst
    Do While Not rs.EOF
        If Val(rs!B_qtyall) <= 0 Then
            MsgBox "数量小于0,不能保存", vbInformation, "提示"
            sewbool = False
            Exit Function
        End If
    
        If Len(rs!B_process) <= 0 Then
            MsgBox "工序为空,不能保存", vbInformation, "提示"
            sewbool = False
            Exit Function
        End If
        rs.movenext
    Loop

    sewbool = True
    
End Function
Private Sub TDBGrid10_ButtonClick(ByVal colIndex As Integer)
    
If TDBGrid10.Columns("B_process").colIndex = colIndex Then
     Dim frm2 As New frmPopupSew
    frm2.Show vbModal
    If frm2.bool = True Then
        setrsgrid10 (frm2.clientid)
    Else
        rsgrid10!B_process = frm2.clientid
    End If
        
    
    Unload frm2
    End If
End Sub

Private Sub setrsgrid10(ByVal a As String)
    Dim b As Long
    b = TDBGrid10.bookmark
    rsgrid10.MoveFirst
    Do While Not rsgrid10.EOF
        rsgrid10!B_process = a
        rsgrid10.movenext
    Loop
    If rsgrid10.RecordCount > 0 Then
        TDBGrid10.bookmark = b
    End If
End Sub
Private Sub Printsew()
    Dim sql As String
    Dim rs As New RecordSet
'    sql = "SELECT * FROM G_BillSew WHERE B_orderid='" & theid & "'"""
'    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
sql = "exec usp_BillSewPrint'" & theID & "'"
rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
         Dim frm1 As New frmModBLRPreviewOriColor
        Set frm1.RecordSet = rs.Clone
            
        frm1.ObjectID = "22B136"
        frm1.Show vbModal
    End If
End Sub

Private Sub prit1()
Dim sql As String
Dim rs As New RecordSet
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    
'    sql = "exec usp_OrdersContractsPrint '" & rsdetail!B_ItemID & "'"
     sql = "exec usp_OrdersContractsPrint '" & theID & "'"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount <= 0 Then
        MsgBox "此合同没有保存，请保存后再打印！", vbInformation, "提示"
        Exit Sub
    End If
    Dim frm1 As New frmModBLRPreviewOri
    Set frm1.RecordSet = rs.Clone
'    frm1.obj = "11S067"
    frm1.ObjectID = "22B134"
    frm1.Show


End Sub

'水洗标图片
Private Sub washing()
Dim frm1 As New frmOrderProduct_FZ
If rsgrid10!B_id = "" Then
    MsgBox "请先保存缝制计划！", vbInformation, "提示"
    Exit Sub
End If

frm1.m_ID = rsgrid10!B_id
frm1.m_OrderID = theID
frm1.FlatEdit1(0).Text = rsgrid10!B_KuanHao
frm1.Show vbModal


End Sub
'洗标和织标的位置
Private Sub WAW()
Dim frm1 As New frmOrderProduct_FZ_WAW
'If theid = "" Then
'    MsgBox "无效的合同计划！", vbInformation, "提示"
'    Exit Sub
'End If
'

frm1.m_OrderID = theID
frm1.Show vbModal

End Sub

'织标的位置
Private Sub Woven()
Dim frm1 As New frmOrderProduct_FZ_Woven
'If theid = "" Then
'    MsgBox "无效的合同计划！", vbInformation, "提示"
'    Exit Sub
'End If
'
frm1.m_OrderID = theID
frm1.Show vbModal

End Sub

'打开显示包装计划图片
Private Sub OpenImageBZ()
Dim rs1 As New RecordSet
                Dim sql As String
                Dim clsFile01 As New clsFile
                Dim szPic As String
                
            sql = "SELECT * FROM WVAccountImage.dbo.G_image_NEW_BZ where B_OrderID='" & theID & "'"
            rs1.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
                
                If rs1.RecordCount > 0 Then
                    szPic = App.Path & "\temp\" & "BZ" & rs1!B_id & rs1!B_OrderID & ".JPG"
                    Debug.Print szPic
                    
                    
                    PicShow2Ctl rs1!B_picture, Picture6
                    Text1.Text = IIf(IsNull(rs1!B_memo), "", rs1!B_memo)
                    'PicShow2Ctl rs!B_picture, Picture3
                Else
                    Picture6.Picture = Nothing
                End If
End Sub

