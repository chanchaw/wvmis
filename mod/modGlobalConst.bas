Attribute VB_Name = "modGlobalConst"
Option Explicit

'全局常量
Public Const B_yuanOrderid = "888888"
Public Const B_WhiteOrderid = "999999"
Public Const B_Value As Integer = "0.1"
Public Const SUPERADMIN As String = "管理员"
Public Const SUPERCOMPUTER As String = "chanchaw-lenovo"
Public Const TABLECOMPANYINFO As String = "G_CompanyInfo"  '保存客户公司信息的数据表
Public Const CONSTINIFILENAME As String = "XDFSoft.ini"

Public Const IF_DBSECTION As String = "数据库"
Public Const IF_DBSECTION_SERVERKEY As String = "服务器名"
Public Const IF_DBSECTION_DBKEY As String = "数据库名"
Public Const IF_DBSECTION_USERKEY As String = "用户名"
Public Const IF_DBSECTION_PWKEY As String = "密码"


Public Const IF_DBSECTION_SOB As String = "账套集数据库"
Public Const IF_DBSECTION_SERVERKEY_SOB As String = "服务器名"
Public Const IF_DBSECTION_DBKEY_SOB As String = "数据库名"
Public Const IF_DBSECTION_USERKEY_SOB As String = "用户名"
Public Const IF_DBSECTION_PWKEY_SOB As String = "密码"

Public Const IF_DBSECTION_Image As String = "图片数据库"
Public Const IF_DBSECTION_SERVERKEY_Image As String = "服务器名"
Public Const IF_DBSECTION_DBKEY_Image As String = "数据库名"
Public Const IF_DBSECTION_USERKEY_Image As String = "用户名"
Public Const IF_DBSECTION_PWKEY_Image As String = "密码"

'生成的白坯计划单主表中的字段
Public Const B_BID As String = "WHT"
Public Const B_ObjectID As String = "12B006"
Public Const B_BillType As String = "WHT01"

'生成的色布计划单主表中的字段
Public Const B_BID_CC As String = "CLC"
Public Const B_ObjectID_CC  As String = "12B008"
Public Const B_BillType_CC  As String = "COL01"

Public Const TheSystemKewordPrefix = "@@TSK_"  '系统关键字的前缀

Public Const CREATEFORWARD = "CFI"  '可正向生成的单据
                                    '例如：生成领用单、盘点单等等

'下面是使用ActiveBar做DOCKER效果用到的常量
'====================================================
Public Const UISPACE As Long = 45
Public Const UISPACE2X As Long = 90
Public Const UISMALLSPACE As Long = 15 ' aprox 2 pixels
Public Const UISMALLSPACE2X As Long = 30
Public Const DOCKABLEBANDPREFIXNAME As String = "IWillDockToActiveBarBand_"
Public Const DOCKABLETOOLPREFIXNAME As String = "IWillDockToActiveBarTool_"
Public Const ERR_DOCKABLETOOLNOTFOUND As Long = 12001
'====================================================



Public Const MODULE_ORDER = "订单合同"
Public Const MODULE_DATADICTIONARY = "基础资料"
Public Const MODULE_ACCESSORY = "辅料仓库"
Public Const MODULE_YARN = "原料仓库"
Public Const MODULE_WHITE = "白坯仓库"
Public Const MODULE_COLOR = "色布仓库"
Public Const MODULE_CP = "成品仓库"

Public Const MODULE_Gold = "五金仓库"
Public Const MODULE_financial = "财务系统"



'下面是单据类型
'下面是辅料仓库的单据类型
Public Const ACCPLAN As String = "ACC00"   '辅料仓库 - 计划单
Public Const ACCPIN As String = "ACC01"   '辅料仓库 - 入库单
Public Const ACCPOUT As String = "ACC02"   '辅料仓库 - 退料单
Public Const ACCPSPEND As String = "ACC03"   '辅料仓库 - 生产领用单


Public Const WHITEPLAN As String = "WHT01"   '白坯仓库 - 白坯计划单
Public Const WHITEPURCHASE As String = "WHT03"   '白坯仓库 - 采购入库单
Public Const WHITEPPROCESS As String = "WHT04"   '白坯仓库 - 外加工入库



Public Const COLORPLAN As String = "COL01"   '色布仓库 - 计划单
Public Const COLORPURCHASE As String = "COL03"   '色布仓库 - 采购入库单
Public Const COLORPROCESS As String = "COL04"   '色布仓库 - 外加工入库


'下面是单据的可生成性
Public Const CREATEFROMINVETORY As String = "CFI"   '从库存中生成的单据
Public Const CREATEREVERSE As String = "CR"   '逆向生成的单据
Public Const CANNTCREATE As String = "CTC"   '不可被生成的单据


Public Const OBJECTIDACCE As String = "12B002"  '辅料单据


Public Const COLORBC13FIRST As String = "9"  '色布流转卡13位条码的首字符





