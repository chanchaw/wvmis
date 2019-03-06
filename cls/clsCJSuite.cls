VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCJSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private nTheme As AppearanceStyle 'CodejockSuiteControls的皮肤

'设置某窗体中的所有Codejock的控件的主题（皮肤）
Public Sub SetCodejockCtlTheme(ByVal vFrm As Form)
    Dim oCtl As Object
    For Each oCtl In vFrm.Controls
        If TypeOf oCtl Is XtremeSuiteControls.CheckBox Or _
           TypeOf oCtl Is XtremeSuiteControls.GroupBox Or _
           TypeOf oCtl Is XtremeSuiteControls.ColorPicker Or _
           TypeOf oCtl Is XtremeSuiteControls.ComboBox Or _
           TypeOf oCtl Is XtremeSuiteControls.CheckBox Or _
           TypeOf oCtl Is XtremeSuiteControls.RadioButton Or _
           TypeOf oCtl Is XtremeSuiteControls.PushButton Or _
           TypeOf oCtl Is XtremeSuiteControls.FlatEdit Or _
           TypeOf oCtl Is XtremeSuiteControls.ListBox Or _
           TypeOf oCtl Is XtremeSuiteControls.ListView Or _
           TypeOf oCtl Is XtremeSuiteControls.UpDown Or _
           TypeOf oCtl Is XtremeSuiteControls.HexEdit Or _
           TypeOf oCtl Is XtremeSuiteControls.TreeView Or _
           TypeOf oCtl Is XtremeSuiteControls.TabControlPage Then
           'Or TypeOf oCtl Is xtremeReportControl.ReportControl Then
                oCtl.Appearance = nTheme ' Redraw
        End If
        
    
'        If TypeOf oCtl Is XtremeSuiteControls.CheckBox Or _
'           TypeOf oCtl Is XtremeSuiteControls.GroupBox Or _
'           TypeOf oCtl Is XtremeSuiteControls.RadioButton Or _
'           TypeOf oCtl Is XtremeSuiteControls.Label Then
'            If cmbResourceFile(Index).ListIndex > 7 Then
'                oCtl.ForeColor = SuiteControlsGlobalSettings.ResourceImages.GetColor("TabManager", "TabPageTextColor")
'            Else
'                oCtl.ForeColor = RGB(0, 0, 0)
'            End If
'        End If
    Next
End Sub

Private Sub Class_Initialize()
    nTheme = xtpAppearanceResource
    SuiteControlsGlobalSettings.ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Silver.ini"
End Sub


