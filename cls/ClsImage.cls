VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'本工具类用于画图，放大，缩小以及旋转
'默认在控件的正中央画图
Private A_FileName As String
Private A_stdPicture As StdPicture
Private A_ZoomedPicture As StdPicture
Private A_PictureBox As Image
Private A_Rate As Double

Private A_StartX As Double  '在PictureBox开始画的起始位置的X坐标
Private A_StartY As Double  '在PictureBox开始画的起始位置的Y坐标

Private A_ZoomRate As Double '每次缩放的比率.初始化为0.01即每次缩放的比例是1%

Private Type PictureSize
    width As Single
    height As Single
End Type

Private Type BitMap
   bmType As Long         '图像类型：0 表示是位图
   bmWidth As Long        '图像宽度（像素）
   bmHeight As Long       '图像高度（像素）
   bmWidthBytes As Long   '每一行图像的字节数
   bmPlanes As Integer    '图像的图层数
   bmBitsPixel As Integer '图像的位数
   bmBits As Long         '位图的内存指针
End Type

Private A_PictureSize As PictureSize
Private ctP180 As Double

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long


Public Property Let ZoomRate(ByVal vData As Double)
    A_ZoomRate = vData
End Property

Public Property Get ZoomRate() As Double
    ZoomRate = A_ZoomRate
End Property

Public Property Let StartX(ByVal vData As Double)
    A_StartX = vData
End Property

Public Property Get StartX() As Double
    StartX = A_StartX
End Property

Public Property Let StartY(ByVal vData As Double)
    A_StartY = vData
End Property

Public Property Get StartY() As Double
    StartY = A_StartY
End Property

'初始化着画的起始坐标，默认在控件的中间开始画
Private Sub InitStartPosition(ByVal vRate As Double)
    A_StartX = A_PictureBox.width / 2 - A_stdPicture.width * vRate / 2
    
    If A_StartX < 0 Then
        A_StartX = 0
    End If
    A_StartY = 0
End Sub

Public Sub InitCls(ByVal vFileName As String, ByRef vPictureBox As Image, _
    Optional ByVal vRate As Double = 1)
    On Error Resume Next
    ctP180 = 4 * Atn(1) '圆周率
    
    A_FileName = vFileName
    Set A_stdPicture = LoadPicture(vFileName)
    Set A_PictureBox = vPictureBox
    A_Rate = vRate
    
    
    '调用时候没有设置显示比例
    '那么自动做调整
    '自动比率=控件的宽度 / 图片的宽度
    If A_Rate = 1 Then
        A_Rate = GetZoomRate
    End If


    A_ZoomRate = 0.01

    '初始化着画的坐标点
    InitStartPosition A_Rate


    '画图
    PaintPicture A_Rate
End Sub

'根据图片的尺寸和显示图片的控件的尺寸
'获取宽度  高度的比率，以最小的比例返回
'用于显示图片的比率
Private Function GetZoomRate() As Double
    Dim rateWidth As Double
    Dim rateHeight As Double
    
    rateWidth = A_PictureBox.width / A_stdPicture.width
    rateHeight = A_PictureBox.height / A_stdPicture.height
    
    If rateWidth < rateHeight Then
        GetZoomRate = rateWidth
    Else
        GetZoomRate = rateHeight
    End If
End Function

Public Sub Dispose()
    Set A_stdPicture = Nothing
End Sub

Private Sub PaintPicture(ByVal vRate As Double)
    A_PictureBox.Cls
    A_PictureBox.PaintPicture A_stdPicture, A_StartX, A_StartY, _
        A_stdPicture.width * vRate, A_stdPicture.height * vRate

'    A_PictureBox.PaintPicture LoadPicture(A_FileName), A_StartX, A_StartY, _
'        A_stdPicture.width * vRate, A_stdPicture.height * vRate

    '记录当前的图片尺寸
    A_PictureSize.width = A_stdPicture.width * vRate
    A_PictureSize.height = A_stdPicture.height * vRate
    
    Set A_ZoomedPicture = A_PictureBox.Picture
End Sub


Private Sub PaintPicture8Size(ByVal vWidth As Single, ByVal vHeight As Single)
    A_PictureBox.Cls
    A_PictureBox.PaintPicture A_stdPicture, A_StartX, A_StartY, _
        vWidth, vHeight

    '记录当前的图片尺寸
    A_PictureSize.width = vWidth
    A_PictureSize.height = vHeight
    
    '缩放后的图像
    Set A_ZoomedPicture = A_PictureBox.Picture
End Sub

'刷新，重新画图
Public Sub Refresh()
    
End Sub

'缩放图片.在当前图片的尺寸基础上
'不是图片的原尺寸的基础
'即实现递增、递减
Public Sub ZoomPicture(ByVal vRate As Double)
    Dim dZoom As Double
    dZoom = vRate
    
    A_PictureSize.width = A_PictureSize.width * dZoom
    A_PictureSize.height = A_PictureSize.height * dZoom
    
    
    'A_PictureBox.PaintPicture A_PictureBox.Image, 0, 0, A_PictureSize.width, A_PictureSize.height
    PaintPicture8Size A_PictureSize.width, A_PictureSize.height
End Sub


Public Sub MoveUp(ByVal vScale As Double)
    A_StartY = A_StartY - vScale
    PaintPicture8Size A_PictureSize.width, A_PictureSize.height
End Sub

Public Sub MoveDown(ByVal vScale As Double)
    A_StartY = A_StartY + vScale
    PaintPicture8Size A_PictureSize.width, A_PictureSize.height
End Sub

Public Sub MoveLeft(ByVal vScale As Double)
    A_StartX = A_StartX - vScale
    PaintPicture8Size A_PictureSize.width, A_PictureSize.height
End Sub

Public Sub MoveRight(ByVal vScale As Double)
    A_StartX = A_StartX + vScale
    PaintPicture8Size A_PictureSize.width, A_PictureSize.height
End Sub

'顺时针旋转90度
Public Sub Clockwise()
    Dim W1 As Long, H1 As Long, B1() As Byte, Bs1 As Long, BytesW1 As Long, Ps1 As Long
    Dim W2 As Long, H2 As Long, B2() As Byte, Bs2 As Long, BytesW2 As Long, Ps2 As Long
    Dim S1 As Long, S2 As Long, X As Long, Y As Long, x1 As Long, y1 As Long
    Dim CenX1 As Long, CenY1 As Long, CenX2 As Long, CenY2 As Long
    Dim ToJ As Single
    
    ToJ = 90 / 180 * ctP180 '旋转角度转弧度
    
    '下面一条语句看似可有可无，实际有两个作用：恢复旋转前控件的原图像和大小
    'A_PictureBox.Picture = A_PictureBox.Picture
    'A_PictureBox.Picture = PaintPicture8Size(A_PictureSize.width, A_PictureSize.height)
    'Set A_PictureBox.Picture = A_ZoomedPicture
    
    
    '旋转前图像数据：宽度，高度，颜色数组，总字节数，每行字节数，每像素字节数
    GetBmpDat A_PictureBox, W1, H1, B1, Bs1, BytesW1, Ps1
    CenX1 = Int(W1 * 0.5): CenY1 = Int(H1 * 0.5)       '旋转前的图像中心点
    
    '计算旋转后控件的高度和宽度，要预先设置窗体和图片的 ScaleMode 为 3（像素）
    W2 = Abs(W1 * Cos(ToJ)) + Abs(H1 * Sin(ToJ))  '旋转后：图像宽度
    H2 = Abs(H1 * Cos(ToJ)) + Abs(W1 * Sin(ToJ))  '旋转后：图像高度
    X = A_PictureBox.width - A_PictureBox.ScaleWidth      '图片框边框：宽度
    Y = A_PictureBox.height - A_PictureBox.ScaleHeight    '图片框边框：高度
   
    
    'A_PictureBox.Move A_PictureBox.Left, A_PictureBox.Top, X + W2, Y + H2
    A_PictureBox.Cls
    A_PictureBox.Line (0, 0)-(W2, H2), &HFFFFFF, BF
    
    '旋转后图像数据：宽度，高度，颜色数组，总字节数，每行字节数，每像素字节数
    GetBmpDat A_PictureBox, W2, H2, B2, Bs2, BytesW2, Ps2
    CenX2 = Int(W2 * 0.5): CenY2 = Int(H2 * 0.5)       '旋转后：图像中心点

    '显示信息
    A_PictureBox.CurrentX = 5: A_PictureBox.CurrentY = 5
    A_PictureBox.Print "处理中，请稍候..."
    
    W1 = W1 - 1: H1 = H1 - 1
    For X = 0 To W2 - 1
    For Y = 0 To H2 - 1
       Zhuan -ToJ, CenX2, CenY2, X, Y, x1, y1           '用 x1,y1 获得旋转坐标
       x1 = x1 - CenX2 + CenX1: y1 = y1 - CenY2 + CenY1 '转变为旋转前的坐标
       
       S2 = XYtoIndex(X, Y, BytesW2, Ps2)    '旋转后：像素点在数组 B2 中的索引
       If x1 < 0 Or x1 > W1 Or y1 < 0 Or y1 > H1 Then
          B2(S2 + 2) = 255: B2(S2 + 1) = 255: B2(S2) = 255 '超出原图像区域，设为白色
       Else
          S1 = XYtoIndex(x1, y1, BytesW1, Ps1)  '旋转前：像素点在数组 B1 中的索引
          B2(S2 + 2) = B1(S1 + 2): B2(S2 + 1) = B1(S1 + 1): B2(S2) = B1(S1) '红绿蓝
       End If
    Next
    Next
    
    
    SetBitmapBits A_PictureBox.Image, Bs2, B2(0) '将 Picture1 的图像设置为旋转后的二进数组 B2()
End Sub

'逆时针旋转90度
Public Sub AntiClockwise()
    
End Sub


Private Sub GetBmpDat(Kj As Control, W As Long, h As Long, b() As Byte, Bs As Long, BytesW As Long, Ps As Long)
   '获取控件 Kj 的图像数据
   Dim MapInf As BitMap

   GetObject Kj.Image, Len(MapInf), MapInf '用 MapInf 得到 Kj 的图像信息
   W = MapInf.bmWidth: h = MapInf.bmHeight '图像宽度、高度（像素）
   BytesW = MapInf.bmWidthBytes            '每行占用字节数
   Ps = BytesW \ W                         '每个像素字节数(一般为4)
   Bs = W * h * Ps                         '总字节数=宽度*高度*每个像素字节
   ReDim b(0 To Bs - 1)
   GetBitmapBits Kj.Image, Bs, b(0)        '将 Kj 图像所有像素点的颜色值读入二进数组 B()
End Sub

Private Sub Zhuan(ToJ As Single, x0 As Long, y0 As Long, ByVal X As Long, ByVal Y As Long, x1 As Long, y1 As Long)
   '将点 x,y 围绕 x0,y0 顺时针旋转 ToJ 弧度，用 x1,y1 返回旋转后的位置
   '注意：要预先设置圆周率 ctP180 = 4 * Atn(1)
    Dim s As Single, j As Single
   
    X = X - x0: Y = Y - y0
    s = Sqr(X ^ 2 + Y ^ 2)               'X,Y 与 x0,y0 的距离
    If s = 0 Then j = 0 Else j = Y / s   '与水平线的夹角的正弦值

    If Abs(j) >= 1 Then
       If j > 0 Then j = ctP180 * 0.5 Else j = -ctP180 * 0.5 '90 度时的特殊情况
    Else
       j = Atn(j / Sqr(-j * j + 1)) '与水平线的夹角
    End If
   
    If X < 0 Then j = -ctP180 - j
    x1 = x0 + s * Cos(j + ToJ): y1 = y0 + s * Sin(j + ToJ) '返回旋转后的位置
End Sub

Private Function XYtoIndex(X As Long, Y As Long, BytesW As Long, Ps As Long) As Long
   '返回图像坐标 x,y 在颜色数组中的序号位置。
   'BytesW：每行图像占用字节数，Ps：每个像素点占用字节数(一般为4)
   XYtoIndex = Y * BytesW + X * Ps
End Function


