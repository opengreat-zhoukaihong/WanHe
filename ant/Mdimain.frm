VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H80000003&
   Caption         =   "移动通信网无线环境普查及优化分析工具 --- ANT FOR GSM"
   ClientHeight    =   1785
   ClientLeft      =   180
   ClientTop       =   4245
   ClientWidth     =   11760
   HelpContextID   =   1
   Icon            =   "Mdimain.frx":0000
   LinkTopic       =   "gsm"
   LockControls    =   -1  'True
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   741
      ButtonWidth     =   661
      ButtonHeight    =   635
      ImageList       =   "ImageList"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   34
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "选择"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "放大"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "缩小"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "移动"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "显示中心"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "信息"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "尺子"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "标注"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "打开图例"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "数据回放"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "输入文本"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "输入符号"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "半径选择"
            Object.Tag             =   ""
            ImageIndex      =   25
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "画圆"
            Object.Tag             =   ""
            ImageIndex      =   29
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "画多边形"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "转换为区域"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "地理信息"
            Object.Tag             =   ""
            ImageIndex      =   20
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "同频查找"
            Object.Tag             =   ""
            ImageIndex      =   21
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "邻频查找"
            Object.Tag             =   ""
            ImageIndex      =   22
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "NCELL 查找"
            Object.Tag             =   ""
            ImageIndex      =   30
         EndProperty
         BeginProperty Button24 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "NCELL 数据显示"
            Object.Tag             =   ""
            ImageIndex      =   26
         EndProperty
         BeginProperty Button25 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Dedicated Channel参数显示"
            Object.Tag             =   ""
            ImageIndex      =   27
         EndProperty
         BeginProperty Button26 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "双网参数比较"
            Object.Tag             =   ""
            ImageIndex      =   28
         EndProperty
         BeginProperty Button27 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "邻频载干比显示"
            Object.Tag             =   ""
            ImageIndex      =   31
         EndProperty
         BeginProperty Button28 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button29 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "创建副图"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button30 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "修改图例"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button31 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "图例变换"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button32 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "自动标注"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button33 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "图片观察"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button34 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "主邻小区显示"
            Object.Tag             =   ""
            ImageIndex      =   34
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   1800
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   1470
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   2116
            MinWidth        =   2116
            Picture         =   "Mdimain.frx":030A
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   8624
            MinWidth        =   8624
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1942
            MinWidth        =   1942
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   1185
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.txt"
      DialogTitle     =   "打开文件"
      FileName        =   "*.txt"
      Filter          =   " *.txt  *.* "
      InitDir         =   "\gsm\normal"
   End
   Begin ComctlLib.ImageList ImageList 
      Left            =   555
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   34
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":101C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":15D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":1B90
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":214A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":26F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":2CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":320C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":37C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":3D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":433A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":48E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":4E9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":5440
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":59E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":5F8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":6546
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":6B28
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":70F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":76B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":7C6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":8224
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":87DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":8D98
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":8F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":94BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":9A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":A0A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":A646
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":A750
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":AD1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":B2C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":B87E
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":BBD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdimain.frx":BF22
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MAIN_1 
      Caption         =   "&P 预处理"
      Begin VB.Menu SUB_12 
         Caption         =   "&G 数据转换"
         Begin VB.Menu ANTSurveyor 
            Caption         =   "ANT Pilot 通话测试数据转换"
            Begin VB.Menu Pilot_data 
               Caption         =   "纯数据转换"
            End
            Begin VB.Menu Pilot_data_Report 
               Caption         =   "数据转换并生成测试报告"
            End
         End
         Begin VB.Menu SUB_121 
            Caption         =   "&Tems 通话测试数据转换"
            Begin VB.Menu Tems_Data 
               Caption         =   "纯数据转换"
            End
            Begin VB.Menu Tems_Data_Report 
               Caption         =   "数据转换并生成测试报告"
            End
         End
         Begin VB.Menu Tems98_Convert 
            Caption         =   "&Tems98 通话测试数据转换"
            Begin VB.Menu Tems98_Data 
               Caption         =   "纯数据转换"
            End
            Begin VB.Menu Tems98_Report 
               Caption         =   "数据转换并生成测试报告"
            End
         End
         Begin VB.Menu SUB_Obtel 
            Caption         =   "&Grayson Surveyor 数据转换"
            Begin VB.Menu Obtel_Data 
               Caption         =   "纯数据转换"
            End
            Begin VB.Menu Obtel_Data_Report 
               Caption         =   "数据转换并生成测试报告"
            End
         End
         Begin VB.Menu SUB_123 
            Caption         =   "Tem&S 扫频测试数据转换"
         End
         Begin VB.Menu ScanPilot 
            Caption         =   "ANT Pilot 扫频测试数据转换"
         End
         Begin VB.Menu alert 
            Caption         =   "-"
         End
         Begin VB.Menu SUB_122 
            Caption         =   "&D 文档管理 "
         End
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu OPEN_ALL_MAP 
         Caption         =   "&S 打开地图"
      End
      Begin VB.Menu SUB_CENTER 
         Caption         =   "&R 显示中心"
         Enabled         =   0   'False
      End
      Begin VB.Menu SP2 
         Caption         =   "-"
      End
      Begin VB.Menu SYS_MANAGER 
         Caption         =   "&Y 系统管理"
      End
      Begin VB.Menu SUB_CONFIG 
         Caption         =   "&C 地图配置"
      End
      Begin VB.Menu SUB_61 
         Caption         =   "&A 小区设计数据"
      End
      Begin VB.Menu SUB_BASE_ADD 
         Caption         =   "&Z 站址维护"
      End
      Begin VB.Menu wert 
         Caption         =   "-"
      End
      Begin VB.Menu EXIT 
         Caption         =   "&X 退出"
      End
   End
   Begin VB.Menu MAIN_2 
      Caption         =   "&F 文件"
      Begin VB.Menu SUb_21 
         Caption         =   "&A 打开文件"
      End
      Begin VB.Menu SUB_25 
         Caption         =   "&B 保存文件"
         Enabled         =   0   'False
      End
      Begin VB.Menu SUB_23 
         Caption         =   "&C 关闭文件"
         Enabled         =   0   'False
      End
      Begin VB.Menu SUB_24 
         Caption         =   "&D 全部关闭"
         Enabled         =   0   'False
      End
      Begin VB.Menu SUB_26 
         Caption         =   "&E 另存为..."
         Enabled         =   0   'False
      End
      Begin VB.Menu SUB_28 
         Caption         =   "&F 合并文件"
         Enabled         =   0   'False
      End
      Begin VB.Menu SP21 
         Caption         =   "-"
      End
      Begin VB.Menu USERMARK 
         Caption         =   "&G 打开用户标识层"
         Enabled         =   0   'False
      End
      Begin VB.Menu CLOSEMARK 
         Caption         =   "&H 关闭用户标识层"
         Enabled         =   0   'False
      End
      Begin VB.Menu SAVEMARK 
         Caption         =   "&I 保存用户标识层"
         Enabled         =   0   'False
      End
      Begin VB.Menu plkm 
         Caption         =   "-"
      End
      Begin VB.Menu OpenWor 
         Caption         =   "&J 打开空间"
      End
      Begin VB.Menu SaveWor 
         Caption         =   "&K 保存空间"
      End
      Begin VB.Menu SaveWindows 
         Caption         =   "&W 抓取地图画面"
      End
      Begin VB.Menu SP23 
         Caption         =   "-"
      End
      Begin VB.Menu SUB_51 
         Caption         =   "&L 页面设置"
      End
      Begin VB.Menu SUB_52 
         Caption         =   "&M 页面布局"
      End
      Begin VB.Menu SUB_531 
         Caption         =   "&N 打印布局"
      End
      Begin VB.Menu pou 
         Caption         =   "-"
      End
      Begin VB.Menu SUB_TEST_REPORT 
         Caption         =   "&O 生成测试报告"
      End
   End
   Begin VB.Menu MAIN_3 
      Caption         =   "&V 观测"
      Begin VB.Menu Sub_31 
         Caption         =   "&C 当前小区参数"
         Begin VB.Menu sub_311 
            Caption         =   "&RxLevFull"
         End
         Begin VB.Menu sub_312 
            Caption         =   "Rx&QualFull"
         End
         Begin VB.Menu sub_313 
            Caption         =   "&ARFCN"
         End
         Begin VB.Menu View_Ci 
            Caption         =   "&CI"
         End
         Begin VB.Menu ll 
            Caption         =   "-"
         End
         Begin VB.Menu Sub_314 
            Caption         =   "RxLev&Sub"
         End
         Begin VB.Menu Sub_315 
            Caption         =   "RxQualSu&b"
         End
         Begin VB.Menu lllll 
            Caption         =   "-"
         End
         Begin VB.Menu Sub_316 
            Caption         =   "&Tx_Power"
         End
         Begin VB.Menu Sub_317_old 
            Caption         =   "Timing &Advance"
         End
      End
      Begin VB.Menu MnuLabelMark 
         Caption         =   "&M 采集事件标注"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu RadioLink 
         Caption         =   "&R 无线链路丢失统计"
      End
      Begin VB.Menu Hopping 
         Caption         =   "&H 跳频状态"
         Begin VB.Menu HoppingStatus 
            Caption         =   "Hopping"
         End
         Begin VB.Menu HoppingLabel 
            Caption         =   "标注跳频参数"
         End
      End
      Begin VB.Menu MnuLabel 
         Caption         =   "&L 参数标注"
      End
      Begin VB.Menu MnuDataGraph 
         Caption         =   "&G 数据统计图"
      End
      Begin VB.Menu Mnu_Replay 
         Caption         =   "&P 数据动态回放"
      End
      Begin VB.Menu ghtr466 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGsm_Dcs 
         Caption         =   "&Q 双频切换点"
      End
      Begin VB.Menu spw 
         Caption         =   "-"
      End
      Begin VB.Menu Browse 
         Caption         =   "&B 数据浏览"
      End
   End
   Begin VB.Menu MnuMAIN_4 
      Caption         =   "&A 分析"
      Begin VB.Menu MnuCover 
         Caption         =   "覆盖分析"
         Begin VB.Menu MnuDSelBcch 
            Caption         =   "选频覆盖统计"
         End
         Begin VB.Menu SUB_41 
            Caption         =   "小区覆盖区域显示"
         End
         Begin VB.Menu MnuNewCover 
            Caption         =   "网络资源覆盖盲区"
         End
         Begin VB.Menu MnuBestSCover 
            Caption         =   "最佳主小区覆盖"
         End
         Begin VB.Menu MnuBetterSCover 
            Caption         =   "次佳主小区覆盖"
         End
         Begin VB.Menu MnuGDCover 
            Caption         =   "G网与D网覆盖显示"
         End
         Begin VB.Menu MnuBYCover 
            Caption         =   "本网与异网覆盖显示"
         End
         Begin VB.Menu Sub_317 
            Caption         =   "覆盖合理性统计"
         End
      End
      Begin VB.Menu MnuDDDD 
         Caption         =   "干扰分析"
         Begin VB.Menu MnuPZ 
            Caption         =   "主邻载频碰撞"
            Begin VB.Menu SUB_441 
               Caption         =   "Ncell-->BCCH"
            End
            Begin VB.Menu SUB_442 
               Caption         =   "Ncell-->TCH"
            End
         End
         Begin VB.Menu MnuNewDisturb 
            Caption         =   "上行干扰统计"
         End
         Begin VB.Menu MnuDisturbSearch 
            Caption         =   "指定范围内同邻频查找"
         End
         Begin VB.Menu MnuC_A 
            Caption         =   "实时邻频C/A比分布"
            Begin VB.Menu MnuC_ASubUp 
               Caption         =   "C/A -1 (Sub)"
            End
            Begin VB.Menu MnuC_ASubDown 
               Caption         =   "C/A +1 (Sub)"
            End
            Begin VB.Menu MnuC_AFullUp 
               Caption         =   "C/A -1 (Full)"
            End
            Begin VB.Menu MnuC_AFullDown 
               Caption         =   "C/A +1 (Full)"
            End
         End
         Begin VB.Menu MnuBcchAdjust 
            Caption         =   "载频调整统计"
         End
      End
      Begin VB.Menu MnuNeighbor 
         Caption         =   "邻小区分析"
         Begin VB.Menu MnuNcellReason 
            Caption         =   "邻小区合理性统计"
         End
         Begin VB.Menu ViewNcell 
            Caption         =   "有效邻小区分布"
         End
         Begin VB.Menu MnuNcellDisplay 
            Caption         =   "邻小区动态显示"
            Visible         =   0   'False
         End
         Begin VB.Menu SUB_431 
            Caption         =   "岛效应查找"
         End
      End
      Begin VB.Menu MnuHO 
         Caption         =   "切换分析"
         Begin VB.Menu MnuHOPara 
            Caption         =   "切换前后参数显示"
         End
         Begin VB.Menu MnuMustHandover 
            Caption         =   "功率预算切换统计"
         End
         Begin VB.Menu MnuHOStat 
            Caption         =   "切换统计"
         End
      End
      Begin VB.Menu MnuMessage 
         Caption         =   "信令分析"
         Begin VB.Menu SUB_33 
            Caption         =   "信令地理描述"
            Begin VB.Menu SUB_330 
               Caption         =   "&M 主要信令描述"
            End
            Begin VB.Menu SUB_367 
               Caption         =   "&O 其它信令地理描述"
            End
            Begin VB.Menu MessagePlay 
               Caption         =   "&R 信令层回放"
            End
         End
         Begin VB.Menu MnuCallProcess 
            Caption         =   "通话过程统计"
         End
         Begin VB.Menu MnuCallReplay 
            Caption         =   "拨打过程重演"
         End
      End
      Begin VB.Menu MnuDoubleNet 
         Caption         =   "双网分析"
         Begin VB.Menu Test_Define 
            Caption         =   "网络定义"
         End
         Begin VB.Menu Other_Precede 
            Caption         =   "异网优于本网"
            Begin VB.Menu Rxlev_Other 
               Caption         =   "Rxlev"
            End
            Begin VB.Menu RxQual_Other 
               Caption         =   "RxQual"
            End
            Begin VB.Menu global_Other 
               Caption         =   "综合质量"
            End
         End
         Begin VB.Menu Local_Precede 
            Caption         =   "本网优于异网"
            Begin VB.Menu Rxlev_Local 
               Caption         =   "Rxlev"
            End
            Begin VB.Menu RxQual_Local 
               Caption         =   "RxQual"
            End
            Begin VB.Menu global_Local 
               Caption         =   "综合质量"
            End
         End
         Begin VB.Menu LabelTa 
            Caption         =   "双网TA分布标注"
         End
         Begin VB.Menu sp8 
            Caption         =   "-"
         End
         Begin VB.Menu SecMobileRxlev 
            Caption         =   "异网无线参数观测"
            Begin VB.Menu SecMobileRxlevf 
               Caption         =   "RxLevFull"
            End
            Begin VB.Menu SecMobileRxqualf 
               Caption         =   "RxQualFull"
            End
            Begin VB.Menu SecMobileArfcn 
               Caption         =   "ARFCN"
            End
            Begin VB.Menu sp88 
               Caption         =   "-"
            End
            Begin VB.Menu SecMobileRxlevs 
               Caption         =   "RxLevSub"
            End
            Begin VB.Menu SecMobileRxquals 
               Caption         =   "RxQualSub"
            End
            Begin VB.Menu sp888 
               Caption         =   "-"
            End
            Begin VB.Menu SecMobileTa 
               Caption         =   "Timing Advance"
            End
         End
      End
   End
   Begin VB.Menu Scan 
      Caption         =   "&S 扫频"
      Begin VB.Menu My_ScanPlay 
         Caption         =   "&P 扫频回放"
      End
      Begin VB.Menu Arfcn_Changing 
         Caption         =   "&B 生成信道变化图"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu My_Over 
         Caption         =   "&L 本地覆盖"
      End
      Begin VB.Menu SCAN_4 
         Caption         =   "&N 非本地覆盖"
      End
      Begin VB.Menu SCAN_5 
         Caption         =   "&E 干扰点标注"
      End
      Begin VB.Menu SCAN_6 
         Caption         =   "&F 盲点标注"
      End
      Begin VB.Menu Mnudistributing 
         Caption         =   "&S 选频场强分布"
      End
      Begin VB.Menu ss1 
         Caption         =   "-"
      End
      Begin VB.Menu TRAN_C_I 
         Caption         =   "&C 载干比计算"
      End
      Begin VB.Menu SCAN_7 
         Caption         =   "&O C/I1 分析"
      End
      Begin VB.Menu SCAN_8 
         Caption         =   "&T C/I2 分析"
      End
   End
   Begin VB.Menu SUB_46 
      Caption         =   "&B 查找"
      Begin VB.Menu SUB_467 
         Caption         =   "&F 按ARFCN查找小区"
      End
      Begin VB.Menu SUB_469 
         Caption         =   "&G 按CI查找小区"
      End
      Begin VB.Menu BsNo_FindCell 
         Caption         =   "&B 按BaseNo查找小区"
      End
      Begin VB.Menu SUB_468 
         Caption         =   "&H 按LAC查找小区"
      End
      Begin VB.Menu FindMyBsic 
         Caption         =   "&S 按BSIC查找小区"
      End
      Begin VB.Menu sp66 
         Caption         =   "-"
      End
      Begin VB.Menu SUB_461 
         Caption         =   "&A 同频组小区查找"
      End
      Begin VB.Menu SUB_466 
         Caption         =   "&B 邻频组小区查找"
      End
      Begin VB.Menu SUB_462 
         Caption         =   "&C 同BSIC小区查找"
      End
      Begin VB.Menu SUB_463 
         Caption         =   "&D 同LAC基站查找 "
      End
      Begin VB.Menu SUB_464 
         Caption         =   "&E 同频同BSIC检查"
      End
      Begin VB.Menu jnkl 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLacLegend 
         Caption         =   "&L Lac图例"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuBcchRetrieve 
         Caption         =   "&B 载频复用簇观测"
      End
      Begin VB.Menu MnuCellReUse 
         Caption         =   "&R 同LAC观测"
      End
      Begin VB.Menu SUB_465 
         Caption         =   "&J 相邻小区定义检查"
      End
      Begin VB.Menu FindFree 
         Caption         =   "&H 空闲ARFCN查找"
      End
      Begin VB.Menu SUB_4600 
         Caption         =   "&I 站址显示"
      End
   End
   Begin VB.Menu sts 
      Caption         =   "&T 话务"
      Begin VB.Menu STS_1 
         Caption         =   "&A 数据转换"
      End
      Begin VB.Menu sts_2 
         Caption         =   "&B 基站显示"
      End
      Begin VB.Menu STS_3 
         Caption         =   "&C 数据查询"
         Begin VB.Menu Tch_Find 
            Caption         =   "&TCH 查询"
         End
         Begin VB.Menu Cch_Find 
            Caption         =   "&CCH 查询"
         End
      End
      Begin VB.Menu sts_4 
         Caption         =   "&D 地图显示"
         Begin VB.Menu Tch_Map 
            Caption         =   "&TCH 地图"
         End
         Begin VB.Menu Cch_Map 
            Caption         =   "&CCH 地图"
         End
      End
      Begin VB.Menu CQT 
         Caption         =   "&C 呼叫测试"
         Begin VB.Menu CQT_1 
            Caption         =   "&A 测试小区选择"
         End
         Begin VB.Menu CQT_2 
            Caption         =   "&B 拨打记录表"
            Begin VB.Menu CQT_Edit 
               Caption         =   "编辑"
            End
            Begin VB.Menu CQT_View 
               Caption         =   "显示"
            End
         End
      End
   End
   Begin VB.Menu WINDOW 
      Caption         =   "&W 窗口"
      WindowList      =   -1  'True
      Begin VB.Menu OpenMap 
         Caption         =   "&O 打开地图窗口 "
      End
      Begin VB.Menu spo 
         Caption         =   "-"
      End
      Begin VB.Menu REDRAW 
         Caption         =   "&R 重画"
      End
      Begin VB.Menu CASECADE 
         Caption         =   "&C 级联 "
      End
      Begin VB.Menu TITLE 
         Caption         =   "&H 平铺"
      End
      Begin VB.Menu ICONS 
         Caption         =   "&I 排列图标"
      End
   End
   Begin VB.Menu MAIN_8 
      Caption         =   "&Y 资源"
      Begin VB.Menu MAIN_8_1 
         Caption         =   "&A 资源定义"
      End
      Begin VB.Menu MAIN_8_2 
         Caption         =   "&B 资源切换"
      End
   End
   Begin VB.Menu MAIN_7 
      Caption         =   "&H 帮助"
      Begin VB.Menu Help_content 
         Caption         =   "&C 目录"
         HelpContextID   =   4
         Shortcut        =   {F1}
      End
      Begin VB.Menu ABOUT 
         Caption         =   "&A 关于..."
      End
   End
   Begin VB.Menu MnuGraphyControl 
      Caption         =   "GraphyControl"
      Visible         =   0   'False
      Begin VB.Menu MnuFullSubSwitch 
         Caption         =   "Full/Sub 图形颜色转换"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SNDisplay_Click()
    Dim SelTbl As String
    Dim MyCival As String
    Dim MyCellName As String
    'Dim Mybsictemp As Integer
    'Dim MyBcchtemp As Integer
    Dim MyNcellBcch(5) As Integer
    Dim MyNcellBsic(5) As Integer
    Dim i As Integer
    
    On Error Resume Next
    StatusBar.Panels(2).Text = " 主邻小区显示"
    SelTbl = mapinfo.eval("selectionInfo(1)")
    If SelTbl = "" Then
       MsgBox "请选择测量数据中的一个点！", 64, "提示"
       StatusBar.Panels(2).Text = " "
       Exit Sub
    End If
    If UCase(mapinfo.eval("Columninfo( " & SelTbl & ",COL1, 1)")) <> "TIME" Then
       MsgBox "请选择测量数据中的一个点！", 64, "提示"
       StatusBar.Panels(2).Text = " "
       Exit Sub
    End If
    mapinfo.Do "x0=selection.lon"
    mapinfo.Do "y0=selection.lat"
    
    MyCival = mapinfo.eval("selection.ci_serv")
    For i = 0 To 5
        MyNcellBcch(i) = mapinfo.eval("selection.bcch_n" & Format(i + 1))
        MyNcellBsic(i) = mapinfo.eval("selection.bsic_n" & Format(i + 1))
    Next
    Call SearchCellName(0, 0, 0, 0, MyCellName, MyCival, "")
    mapinfo.Do "set map redraw off"
    mapinfo.Do "Set Map Layer 0 Editable On "
    If mapinfo.eval("x1") > 0 And mapinfo.eval("y1") > 0 Then
        mapinfo.Do "Set Style Pen MakePen(1,2,255)"
        mapinfo.Do "Set Style Brush MakeBrush(2,255,255)"
        mapinfo.Do "x1=x1 + 0.0015 * sin (x3* 0.01745329252)"
        mapinfo.Do "y1=y1 + 0.0015 * cos (x3 * 0.01745329252) "
        mapinfo.Do "create Line(x1,y1)(x0,y0)"
    End If
    mapinfo.Do "Set Style Pen MakePen(1,2,16711680)"
    mapinfo.Do "Set Style Brush MakeBrush(2,16711680,16711680)"
    For i = 0 To 5
        Call SearchCellName(MyNcellBsic(i), MyNcellBcch(i), mapinfo.eval("x0"), mapinfo.eval("y0"), MyCellName, "", "")
        If mapinfo.eval("x1") > 0 And mapinfo.eval("y1") > 0 Then
            mapinfo.Do "x1=x1 + 0.0015 * sin (x3* 0.01745329252)"
            mapinfo.Do "y1=y1 + 0.0015 * cos (x3 * 0.01745329252) "
            mapinfo.Do "create Line(x1,y1)(x0,y0)"
        End If
    Next
    
End Sub

Private Sub ABOUT_Click()
    On Error Resume Next
    Face.Show 1
    '          mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
    '          mapinfo.do "Set Next Document Parent " & MapForm.hwnd & " Style 1"
    '          mapinfo.do "Create ButtonPad ""Lee"" As ToolButton HelpMsg ""Use this tool to draw a new route"" Calling 1702  Width 3 Show"
              
    
    '          mapinfo.do "Open Window Message"
              
    '          mapinfo.do "Print ""Test"""
    '          mapinfo.do "note ""Test"""
    
End Sub

Private Sub Arfcn_Changing_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  生成信道变化图"
    Menu_Flag = 9901
    SelTable.Show 1
    StatusBar.Panels(2).Text = ""
End Sub

Private Sub MnuBcchAdjust_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 载频调整统计"
    Menu_Flag = 991107
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuBcchRetrieve_Click()
    'Dim BcchGroup() As Integer
    Dim BcchGroup() As String
    Dim BcchCount() As Integer
    Dim BcchRecPos() As Integer
    Dim MyRow As Integer, i As Integer, j As Integer
    Dim MyCounter As Integer, MyGroup As Integer
    Dim Reco As Integer, StartPos As Integer
    Dim k As Integer
    Dim LastCounter As Integer
    Dim MyColor As Long
    Dim GroupRise As Boolean
    Dim MyCellname1 As String, MyCellname2 As String, MyCellname3 As String
    Dim MyBcch1 As Integer, MyBcch2 As Integer, MyBcch3 As Integer
    Dim BcchGroupColor() As Long
    Dim MyBcchtemp As String
    Dim IsMatch As Boolean
    Dim MyCellName As String
    
    On Error Resume Next
    frmBcchRetrieve.Show 1
    If SelBcchGroup = 0 Then
        Exit Sub
    End If
    StatusBar.Panels(2).Text = "  载频复用簇观测"
    Menu_Flag = 22222
    mapinfo.Do "SELECT arfcn,count(*) " & "FROM cell where arfcn>0 GROUP BY arfcn order by 2 desc,1 desc into mytemp"
    For i = 1 To mapinfo.eval("NumTables()")
        If UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) = "BCCHRETRIEVE" Then
           mapinfo.Do "close table BcchRetrieve"
           Exit For
        End If
    Next
    mapinfo.Do "commit table mytemp as " + Chr(34) + Gsm_Path + "\BcchRetrieve.tab" + Chr(34)
    mapinfo.Do "open table " & Chr(34) + Gsm_Path + "\BcchRetrieve.tab" + Chr(34)
    MyRow = mapinfo.eval("tableinfo(BcchRetrieve,8)")
    If MyRow > 12 Then
        For i = 13 To MyRow
            mapinfo.Do "Delete From BcchRetrieve Where Rowid =" & Format(i)
        Next
    End If
    mapinfo.Do "commit table BcchRetrieve"
    'mapinfo.do "Select BcchRetrieve.arfcn,BcchRetrieve.col2,cell.lon,cell.lat,cell.bearing,cell.lon*cell.lon+cell.lat*cell.lat,cell.cell_name from cell, BcchRetrieve where cell.arfcn =BcchRetrieve.arfcn order by 2 desc ,1 desc ,6 desc into mytemp"
    mapinfo.Do "Select BcchRetrieve.arfcn,BcchRetrieve.col2,cell.lon,cell.lat,cell.bearing,cell.lon*cell.lon+cell.lat*cell.lat,cell.cell_name from cell, BcchRetrieve where cell.arfcn =BcchRetrieve.arfcn order by 6 desc,1 desc,2 desc into mytemp"
    MyRow = mapinfo.eval("tableinfo(BcchRetrieve,8)")
    'ReDim BcchGroup(Int(MyRow / 3)) As String
    'ReDim BcchGroupColor(Int(MyRow / 3)) As Long
    ReDim BcchGroup(Int(MyRow / 2)) As String
    ReDim BcchGroupColor(Int(MyRow / 2)) As Long
    mapinfo.Do "fetch first from BcchRetrieve"
    mapinfo.Do "fetch first from mytemp"
    mapinfo.Do "set map redraw off"
    mapinfo.Do "Set Map Layer 0 Editable On  "
    
    
    mapinfo.Do "set map redraw on"
    Do While UCase(mapinfo.eval("EOT(Mytemp)")) = "F"
FindAgain1:
        MyCellName = mapinfo.eval("mytemp.col7")
        MyCellname1 = GetBaseName(MyCellName)
        mapinfo.Do "x1=mytemp.col3 + 0.0015 * sin (mytemp.col5 * 0.01745329252)"
        mapinfo.Do "y1=mytemp.col4 + 0.0015 * cos (mytemp.col5 * 0.01745329252) "
        mapinfo.Do "x0=mytemp.col3"
        mapinfo.Do "y0=mytemp.col4"
        MyBcch1 = mapinfo.eval("mytemp.col1")
FindAgain2:
        mapinfo.Do "fetch next from mytemp"
        MyCellName = mapinfo.eval("mytemp.col7")
        MyCellname2 = GetBaseName(MyCellName)
        mapinfo.Do "x2=mytemp.col3 + 0.0015 * sin (mytemp.col5 * 0.01745329252)"
        mapinfo.Do "y2=mytemp.col4 + 0.0015 * cos (mytemp.col5 * 0.01745329252) "
        MyBcch2 = mapinfo.eval("mytemp.col1")
        Do While UCase(mapinfo.eval("EOT(Mytemp)")) = "F"
            If MyCellname1 = MyCellname2 Then
               Exit Do
            Else
               MyCellname1 = MyCellname2
               mapinfo.Do "x1=x2"
               mapinfo.Do "y1=y2"
               mapinfo.Do "x0=mytemp.col3"
               mapinfo.Do "y0=mytemp.col4"
               MyBcch1 = MyBcch2
               mapinfo.Do "fetch next from mytemp"
               MyCellName = mapinfo.eval("mytemp.col7")
               MyCellname2 = GetBaseName(MyCellName)
               mapinfo.Do "x2=mytemp.col3 + 0.0015 * sin (mytemp.col5 * 0.01745329252)"
               mapinfo.Do "y2=mytemp.col4 + 0.0015 * cos (mytemp.col5 * 0.01745329252) "
               MyBcch2 = mapinfo.eval("mytemp.col1")
            End If
        Loop
        If UCase(mapinfo.eval("EOT(Mytemp)")) = "T" Then
           Exit Do
        End If
        mapinfo.Do "fetch next from mytemp"
        MyCellName = mapinfo.eval("mytemp.col7")
        MyCellname3 = GetBaseName(MyCellName)
        mapinfo.Do "x3=mytemp.col3 + 0.0015 * sin (mytemp.col5 * 0.01745329252)"
        mapinfo.Do "y3=mytemp.col4 + 0.0015 * cos (mytemp.col5 * 0.01745329252) "
        MyBcch3 = mapinfo.eval("mytemp.col1")
        If MyCellname3 = MyCellname2 Then
            MyBcchtemp = Format(MyBcch1) & "," & Format(MyBcch2) & "," & Format(MyBcch3)
            IsMatch = False
            For i = 0 To UBound(BcchGroup) - 1
                If BcchGroup(i) = "" Then
                    Exit For
                ElseIf MyBcchtemp = BcchGroup(i) Then
                    IsMatch = True
                    Exit For
                End If
            Next
            If Not IsMatch Then
                BcchGroup(i) = MyBcchtemp
                If i > UBound(MyBcchColor) Then
                    BcchGroupColor(i) = MyRndColor(i)
                Else
                    BcchGroupColor(i) = MyBcchColor(i)
                End If
            End If
            mapinfo.Do "Set Style Pen MakePen(2,2," & Format(BcchGroupColor(i)) & ")"
            mapinfo.Do "Set Style Brush MakeBrush(2," & Format(BcchGroupColor(i)) & "," & Format(BcchGroupColor(i)) & ")"
               'mapinfo.do "create Line(x1,y1)(x2,y2)"
               'mapinfo.do "create Line(x2,y2)(x3,y3)"
               'mapinfo.do "create Line(x3,y3)(x1,y1)"
           If SelBcchGroup = 1 Or SelBcchGroup = 2 Then
               mapinfo.Do "Create Region 1 3(x1,y1)(x2,y2)(x3,y3)"
           End If
           GoTo FindAgain1
        Else
'********************两个频率
            MyBcchtemp = Format(MyBcch1) & "," & Format(MyBcch2)
            IsMatch = False
            For i = 0 To UBound(BcchGroup) - 1
                If BcchGroup(i) = "" Then
                    Exit For
                ElseIf MyBcchtemp = BcchGroup(i) Then
                    IsMatch = True
                    Exit For
                End If
            Next
            If Not IsMatch Then
                BcchGroup(i) = MyBcchtemp
                If i > UBound(MyBcchColor) Then
                    BcchGroupColor(i) = MyRndColor(i)
                Else
                    BcchGroupColor(i) = MyBcchColor(i)
                End If
            End If
            mapinfo.Do "Set Style Pen MakePen(2,2," & Format(BcchGroupColor(i)) & ")"
            mapinfo.Do "Set Style Brush MakeBrush(2," & Format(BcchGroupColor(i)) & "," & Format(BcchGroupColor(i)) & ")"
            If SelBcchGroup = 1 Or SelBcchGroup = 3 Then
                mapinfo.Do "Create Region 1 3(x1,y1)(x2,y2)(x0,y0)"
            End If
'********************两个频率
           MyCellname1 = MyCellname3
           mapinfo.Do "x1=x3"
           mapinfo.Do "y1=y3"
           mapinfo.Do "x0=mytemp.col3"
           mapinfo.Do "y0=mytemp.col4"
           MyBcch1 = MyBcch3
           If UCase(mapinfo.eval("EOT(Mytemp)")) = "T" Then
              Exit Do
           End If
           GoTo FindAgain2
           
        End If
    Loop
    mapinfo.Do "close table selection"
    StatusBar.Panels(2).Text = "按鼠标右键清除装饰层"
End Sub

Private Sub BlindFull_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 网络覆盖盲区"
    Menu_Flag = 6451
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub BlindSub_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 网络覆盖盲区"
    Menu_Flag = 6452
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub Browse_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 数据浏览"
    Menu_Flag = 35
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub BsNo_FindCell_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 按BaseNo查找小区"
    Menu_Flag = 4700
    CI_Cell.Show 1
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub Cam_Click()
    Dim i As Integer
    Dim mapHWnd, j As Long

    On Error Resume Next
    Dim ApiPack As APIPACKET                      'win95
    Dim portnum%                                  'win95
    Dim status%                                   'win95
    Dim majVer%, minVer%, rev%, drvrType%         'win95
    Dim adr%, datum%                              'win95
    portnum% = 4 ' CPlus-B, port 1                     'win95
    status% = RNBOcplusFormatPacket(ApiPack, 1028)     'win95
    status% = RNBOcplusInitialize(ApiPack, portnum%)   'win95
'    If status <> 0 Then GoTo VERYFY_OUT                'win95
    status% = RNBOcplusGetVersion(ApiPack, majVer%, minVer%, rev%, drvrType%)     'win95
    status% = RNBOcplusGetFullStatus(ApiPack)          'win95
    adr = 62                                           'win95
    status% = RNBOcplusRead(ApiPack, adr%, datum%)     'win95
    datum = (datum / 89) * 4 + 23                      'win95
'    If datum <> 427 Then GoTo VERYFY_OUT               'win95
    adr = 60                                           'win95
    status% = RNBOcplusRead(ApiPack, adr%, datum%)     'win95
    datum = (datum / 89) * 4 + 23                      'win95
'    If datum <> 619 Then GoTo VERYFY_OUT               'win95
    
    'dog = scread(62)     'win95
    'dog = (dog / 89) * 2 + 23     'win95
    'If dog <> 225 Then GoTo VERYFY_OUT     'win95
    
             SelTbl = mapinfo.eval("selectionInfo(1)")
             If SelTbl = "" Then
                MsgBox "请选择测量数据中的回放起始点", 64, "提示"
                StatusBar.Panels(2).Text = " "
                Exit Sub
             End If
             If UCase(mapinfo.eval("Columninfo( " & SelTbl & ",COL1, 1)")) <> "TIME" Then
                MsgBox "请选择测量数据中的回放起始点", 64, "提示"
                StatusBar.Panels(2).Text = " "
                Exit Sub
             End If

    StatusBar.Panels(2).Text = " 数据回放分析"
    If sys = 0 Then
       'Replay.Show 1
      i = Val(mapinfo.eval("selectionInfo(3)"))  ' SEL_INFO_NROWS
      If i <> 0 Then
    '      MDIMain.SUB_532.Enabled = 1
          Load MapForm
          mapHWnd = Val(mapinfo.eval("WindowInfo(" & mapid & ",12)"))
          If MapForm.WindowState = 1 Or MapForm.WindowState = 2 Then
             MapForm.WindowState = 0
          End If
          MapForm.Move 0, 10, 12000, 4050
    
          Load Graph
          'Graph.Move 0, 4050, 6950, 3150
          Graph.Move 0, 4050, 7000, 3495
    
          Load msgdis
          'msgdis.Move 6950, 4050, 5020, 3150
          msgdis.Move 6950, 4050, 5020, 3495
      End If

    Else
       SelTbl = mapinfo.eval("selectionInfo(1)")
       If UCase(mapinfo.eval("Columninfo( " & SelTbl & ",COL3, 1)")) <> "TIME" Then
          StatusBar.Panels(2).Text = " "
          Exit Sub
       End If
       MapForm.Show
       mapHWnd = Val(mapinfo.eval("WindowInfo(" & mapid & ",12)"))
       If MapForm.WindowState = 1 Or MapForm.WindowState = 2 Then
          MapForm.WindowState = 0
       End If
       MapForm.Move 0, 0, 5600, 7300
       Load Scan_Frm
       Scan_Frm.Move 5600, -35, 6000, 7300

'      ScanGra.Show
'      ScanGra.Move 1440, 4300
    End If
    StatusBar.Panels(2).Text = " "
    Exit Sub
VERYFY_OUT:
    MsgBox "加密锁错误, 请与珠海万禾公司联系！", 64, "提示"
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub CASECADE_Click()
On Error Resume Next
    On Error Resume Next
    StatusBar.Panels(2).Text = " 级联窗口"
    MDIMain.Arrange 0     'CASCADE
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub Cch_Find_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "   CCH 数据查询"
    Gsm_FileName = Gsm_Path + "\sts\cch_sts.tab"
    If UCase(Dir(Gsm_FileName, 0)) <> "CCH_STS.TAB" Then
       MsgBox " CCH_STS.tab 不存在！", 64, "提示"
       StatusBar.Panels(2).Text = "  "
       Exit Sub
    End If
    Cch_data_find.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub Cch_Map_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  打开 CCH 地图"
    Gsm_FileName = Gsm_Path + "\sts\cch_sts.tab"
    If UCase(Dir(Gsm_FileName, 0)) <> "CCH_STS.TAB" Then
       MsgBox " CCH_STS.tab 不存在！", 64, "提示"
       StatusBar.Panels(2).Text = "  "
       Exit Sub
    End If
    mapinfo.Do "open table " + Chr(34) + Gsm_FileName + Chr(34)
    If mapinfo.eval("tableinfo(cch_sts,4)") = 14 Then
       mapinfo.Do "close table cch_sts "
       Cch_emap_choice.Show 1
    Else
       mapinfo.Do "close table cch_sts"
       Cch_mmap_choice.Show 1
    End If
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub CLOSEMARK_Click()
    On Error Resume Next
    mapinfo.Do "set map redraw off"
    mapinfo.Do "Set Map Layer 0 Editable off  "
    mapinfo.Do "set map redraw on"
End Sub

Private Sub CQT_1_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  测试小区选择"
    CQT_choice.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub CQT_Edit_Click()
    Dim mypath As String, buff As String
    Dim tab_name As String, use_name As String
    
    On Error Resume Next
    StatusBar.Panels(2).Text = "  编辑拨打记录表"
open_again:
    FileDialog.DialogTitle = "打开拨打记录表"
    FileDialog.Filter = "*.tab Files|*.TAB|All Files|*.*"
    FileDialog.DefaultExt = "*.TAB"
    FileDialog.Flags = &H80000
    Gsm_FileName = Gsm_Path + "\sts"
    FileDialog.InitDir = Gsm_FileName
    FileDialog.filename = "CQT*.TAB"
    FileDialog.ShowOpen
    convert_filename(1) = Trim(FileDialog.filename)
    If convert_filename(1) = "" Or convert_filename(1) = "CQT*.TAB" Then
       FileDialog.filename = ""
       StatusBar.Panels(2).Text = " "
       Exit Sub
    End If
    FileDialog.filename = ""
    If Dir(convert_filename(1)) = "" Then
       GoTo err_exit
    End If
    If InStr(convert_filename(1), ".tab") = 0 And InStr(convert_filename(1), ".TAB") = 0 Then
       i = MsgBox("请打开拨打记录表 ", 48, "打开拨打记录表")
       GoTo open_again
    End If
    Err = 0
    On Error GoTo err_exit
    mapinfo.Do "open table " + Chr(34) + convert_filename(1) + Chr(34)
    Load CQT_Table
    CQT_Table.Move 1800, 300, 6100, 4800
    
    StatusBar.Panels(2).Text = " "
    Exit Sub
err_exit:
    i = MsgBox("无法打开文件 " + convert_filename(1), 48, "打开拨打记录表")
    GoTo open_again
End Sub

Private Sub CQT_View_Click()
    Dim mypath As String, buff As String
    Dim tab_name As String, use_name As String
    Dim open_table
    Dim pie_size As String
    Dim case_condition
    Dim center_point, center_lon, center_lat
    Dim i As Long
    Dim WinId
    
    On Error Resume Next
    StatusBar.Panels(2).Text = "  显示拨打记录表"
open_again:
    FileDialog.DialogTitle = "打开拨打记录表"
    FileDialog.Filter = "*.tab Files|*.TAB|All Files|*.*"
    FileDialog.DefaultExt = "CQT*.TAB"
    FileDialog.Flags = &H80000
    Gsm_FileName = Gsm_Path + "\sts"
    FileDialog.InitDir = Gsm_FileName
    FileDialog.filename = "CQT*.TAB"
    FileDialog.ShowOpen
    convert_filename(1) = Trim(FileDialog.filename)
    If convert_filename(1) = "" Or convert_filename(1) = "CQT*.TAB" Then
       FileDialog.filename = ""
       StatusBar.Panels(2).Text = " "
       Exit Sub
    End If
    FileDialog.filename = ""
    If Dir(convert_filename(1)) = "" Then
       GoTo err_exit
    End If
    If InStr(convert_filename(1), ".tab") = 0 And InStr(convert_filename(1), ".TAB") = 0 Then
       i = MsgBox("请打开拨打记录表 ", 48, "打开拨打记录表")
       GoTo open_again
    End If
    use_name = convert_filename(1)
    finds = InStr(use_name, "\")
    Do While finds > 0
       use_name = Right(use_name, Len(use_name) - finds)
       finds = InStr(use_name, "\")
    Loop
    use_name = Left(use_name, Len(use_name) - 4)
    If Asc(Left(use_name, 1)) > 47 And Asc(Left(use_name, 1)) < 58 Then
       use_name = "_" + use_name
    End If
    open_table = mapinfo.eval("NumTables()")
    For i = 1 To MyTableNum
        If UCase(mapinfo.eval("tableinfo(" & i & ",1)")) = UCase(use_name) Then
           GoTo opened
        End If
    Next
    Err = 0
    On Error GoTo err_exit
    mapinfo.Do "open table " + Chr(34) + convert_filename(1) + Chr(34)
opened:
    mapinfo.Do "set next document parent " & MapForm.hWnd & "style 1"
    TableNum = Val(mapinfo.eval("NumTables()"))
    If TableNum > 1 Then
       mapinfo.Do "Add Map Auto Layer" + Chr(34) + use_name + Chr(34)
       mapinfo.Do "set map zoom 6 units " + Chr(34) + "km" + Chr(34)
    Else
       mapinfo.Do "Map from " + Chr(34) + use_name + Chr(34)
       thereIsAMap = True
       mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
    If InStr(MapForm.Caption, "CQT_Table") = 0 Then
       MapForm.Caption = MapForm.Caption + ",CQT_Table"
    End If
           
    mapinfo.Do "select avg(col16) from " & use_name & " into mytemp"
    case_condition = Int(mapinfo.eval("mytemp.col1"))
    mapinfo.Do "close table mytemp"
    Select Case case_condition
           Case 0 To 10
                pie_size = 0.985
           Case 11 To 20
                pie_size = 0.885
           Case 21 To 30
                pie_size = 0.785
           Case 31 To 40
                pie_size = 0.685
           Case 41 To 50
                pie_size = 0.585
           Case 51 To 60
                pie_size = 0.485
           Case 61 To 70
                pie_size = 0.385
           Case 71 To 2000
                pie_size = 0.185
    End Select
    For i = 1 To mapinfo.eval("NumWindows()")
        If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then
           WinId = mapinfo.eval("windowid(" & i & ")")
           If WinId = mapinfo.eval("frontwindow()") Then
              Exit For
           End If
        End If
    Next
'    mapinfo.do "shade window Frontwindow() " + use_name + " with col5,col6,col7,col8,col9,col10,col11,col12 pie Angle 180 Max Size " + pie_size + " Units " + Chr(34) + "cm" + Chr(34) + " At Value 25 vary size by " + Chr(34) + "SQRT" + Chr(34) + " border Pen (1,1,0)  position center center style Brush (2,16711680,16777215) ,Brush (2,65280,16777215) ,Brush (2,255,16777215) ,Brush (2,16711935,16777215) ,Brush (2,16776960,16777215) ,Brush (2,65535,16777215) ,Brush (2,8388608,16777215) ,Brush (2,32768,16777215)  # max 25 color 0 #"
    mapinfo.Do "shade window " + WinId + use_name + " with col5,col6,col7,col8,col9,col10,col11,col12 pie Angle 180 Max Size " + pie_size + " Units " + Chr(34) + "cm" + Chr(34) + " At Value 25 vary size by " + Chr(34) + "SQRT" + Chr(34) + " border Pen (1,1,0)  position center center style Brush (2,16711680,16777215) ,Brush (2,65280,16777215) ,Brush (2,255,16777215) ,Brush (2,16711935,16777215) ,Brush (2,16776960,16777215) ,Brush (2,65535,16777215) ,Brush (2,8388608,16777215) ,Brush (2,32768,16777215)  # max 25 color 0 #"
    mapinfo.Do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + use_name + "饼状图" + Chr(34) + " Font(""宋体"",0,9,0) ascending on ranges Font(""宋体"",0,9,0) " + Chr(34) + Chr(34) + " display off ," + Chr(34) + "正常通话" + Chr(34) + " display on ," + Chr(34) + "噪音加带" + Chr(34) + " display on ," + Chr(34) + "串音加带" + Chr(34) + " display on ," + Chr(34) + "回音加带" + Chr(34) + " display on ," + Chr(34) + "无话音" + Chr(34) + " display on ," + Chr(34) + "单方通话" + Chr(34) + " display on ," + Chr(34) + "掉话" + Chr(34) + " display on ," + Chr(34) + "未接通" + Chr(34) + " display on"
    
    thereIsAMap = True
    If mapid = 0 Then
       mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
    MDIMain.SUB_23.Enabled = 1
    MDIMain.SUB_24.Enabled = 1
    MDIMain.SUB_25.Enabled = 1
    MDIMain.SUB_26.Enabled = 1
    MDIMain.SUB_28.Enabled = 1
    
    If MapForm.WindowState = 1 Or MapForm.WindowState = 2 Then
       MapForm.WindowState = 0
    End If
    MapForm.Move 0, 10, 12000, 4300
    center_point = mapinfo.eval("tableinfo(" + use_name + ",8)")
    mapinfo.Do "fetch first from " & use_name
    For i = 1 To center_point
       center_lon = mapinfo.eval(use_name & ".lon")
       center_lat = mapinfo.eval(use_name & ".lat")
       If center_lon <> 0 And center_lat <> 0 Then
          Exit For
       Else
          mapinfo.Do "fetch next from " & use_name
       End If
    Next
    mapinfo.Do "set map Center(" & center_lon & "," & center_lat & ") "
    
    mapinfo.Do "Set Next Document Parent " & MDIMain.hWnd & " Style 2"
    mapinfo.Do "set paper units ""pt"""
    mapinfo.Do "browse * from " & use_name
    mapinfo.Do "set window Frontwindow() Position(0,250) Width 600 Height 160 "
    
    StatusBar.Panels(2).Text = " "
    Exit Sub
err_exit:
    i = MsgBox("无法打开文件 " + convert_filename(1), 48, "打开拨打记录表")
    GoTo open_again
End Sub

Private Sub DATA_PRO_2_Click()
    Dim mypath As String, buff As String
    Dim finds As Integer, i As Integer
    
    On Error Resume Next
    StatusBar.Panels(2).Text = " Idle 数据转移"
    Menu_Flag = 5002
    Gsm_FileName = Gsm_Path + "\normal"
    ChDir Gsm_FileName
    For i = 1 To 50
        convert_filename(i) = ""
    Next
open_again:
    FileDialog.DialogTitle = "Idle 数据转移"
    FileDialog.Filter = "*.tab Files|*.TAB"
    FileDialog.DefaultExt = "*.TAB"
    FileDialog.Flags = &H80000 Or &H200
    Gsm_FileName = Gsm_Path + "\normal"
    FileDialog.InitDir = Gsm_FileName
    FileDialog.ShowOpen
    buff = Trim(FileDialog.filename)
    If buff = "" Then
       StatusBar.Panels(2).Text = " "
       Exit Sub
    End If
    finds = InStr(buff, Chr(0))
    If finds > 0 Then
       mypath = Left(buff, finds - 1) + "\"
       buff = Trim(Right(buff, Len(buff) - finds))
       finds = InStr(buff, Chr(0))
       i = 1
       Do While finds > 0
          convert_filename(i) = mypath + Left(buff, finds - 1)
          buff = Trim(Right(buff, Len(buff) - finds))
          finds = InStr(buff, Chr(0))
          i = i + 1
       Loop
       convert_filename(i) = mypath + buff
    Else
       convert_filename(1) = buff
    End If
    FileDialog.filename = ""
    If Dir(convert_filename(1)) = "" Then
       GoTo err_exit
    End If
    Screen.MousePointer = 11
    Idle.Show 1
    Screen.MousePointer = 0
    If Is_Done = True Then
       MsgBox "处理已完成，您可开始分析了!", 64, "提示"
    End If
    StatusBar.Panels(2).Text = " "
    Exit Sub
    
err_exit:
    i = MsgBox("无法打开文件 " + convert_filename(1), 48, "打开文件")
    GoTo open_again
End Sub

Private Sub DcsRxLevFull_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  DCS参数观测"
    Menu_Flag = 55552
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub DisturbFull_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 网络干扰区"
    Menu_Flag = 7451
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub DisturbSub_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 网络干扰区"
    Menu_Flag = 7452
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub FindFree_Click()
    Dim Layers As Variant
    Dim mySelTbl As String
    Dim QueryName As String
    Dim CellRow As Integer, SelectRow As Integer
    Dim FreeBcch() As Integer
    Dim FreeNum As Integer, i As Integer, j As Integer
    Dim OpenTableNum As Integer
    Dim NonBcchtemp As String, MyNonBcch As String
    
    On Error Resume Next
    If MaxBcch = 0 And MinBcch = 0 Then
       MsgBox "请先用半径选择选择一个范围内的天线！", 64, "提示"
       Exit Sub
    End If
    StatusBar.Panels(2).Text = "  空闲信道查找"
    mySelTbl = mapinfo.eval("selectionInfo(1)")
    QueryName = mapinfo.eval("selectionInfo(2)")
    
    If UCase(mySelTbl) <> "CELL" Or mySelTbl = "" Or QueryName = "" Then
       MsgBox "请先用半径选择选择一个范围内的天线！", 64, "提示"
    Else
       CellRow = mapinfo.eval("tableinfo(cell,8)")
       SelectRow = mapinfo.eval("tableinfo(selection,8)")
       FreeNum = MaxBcch - MinBcch
       ReDim FreeBcch(1 To FreeNum + 1) As Integer
       For i = 1 To FreeNum + 1
           FreeBcch(i) = MinBcch + i - 1
       Next
       mapinfo.Do "fetch first from selection"
       For i = 1 To SelectRow
           If Val(mapinfo.eval("selection.arfcn")) > 0 Then
              FreeBcch(Val(mapinfo.eval("selection.arfcn")) - MinBcch + 1) = 0
           End If
           NonBcchtemp = Trim(mapinfo.eval("selection.non_bcch"))
           For j = 1 To 16
              If InStr(NonBcchtemp, ",") > 0 Then
                 MyNonBcch = Left(NonBcchtemp, InStr(NonBcchtemp, ",") - 1)
                 NonBcchtemp = Trim(Right(NonBcchtemp, Len(NonBcchtemp) - InStr(NonBcchtemp, ",")))
              Else
                 MyNonBcch = NonBcchtemp
                 NonBcchtemp = ""
              End If
              If Val(MyNonBcch) > 0 Then
                 FreeBcch(Val(MyNonBcch) - MinBcch + 1) = 0
              End If
           Next
           mapinfo.Do "fetch next from selection"
       Next
       
       OpenTableNum = mapinfo.eval("NumTables()")
       For i = 1 To OpenTableNum
           If UCase(mapinfo.eval("tableinfo(" & i & ",1)")) = "UNUSEBCCH" Then
              mapinfo.Do "close table UnUseBcch"
              Exit For
           End If
       Next
       mapinfo.Do "Create Table ""UnUseBcch"" (Bcch Decimal(3,0)) file " + Chr(34) + Gsm_Path + "\UnUseBcch.tab" + Chr(34) + " TYPE NATIVE Charset ""WindowsSimpChinese"""
       mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\UnUseBcch.tab" + Chr(34)
       For i = 1 To FreeNum + 1
           If FreeBcch(i) > 0 Then
              mapinfo.Do "insert into UnUseBcch (col1) values (" & FreeBcch(i) & ")"
           End If
       Next
       mapinfo.Do "commit table UnUseBcch"
       mapinfo.Do "Set Next Document Parent " & MDIMain.hWnd & " Style 2"
       mapinfo.Do "browse * from UnUseBcch"
       mapinfo.Do "set window Frontwindow() Position(0,1) Width 2 Height 3 "
       mapinfo.Do "close table " & QueryName
       If CellLayer > 1 Then
          Mymsg = "set map order " & Format(CellLayer)
          For i = 2 To Layers
              If i <> CellLayer Then
                 Mymsg = Mymsg + "," + Format(i)
              Else
                 Mymsg = Mymsg + ",1"
              End If
          Next
          mapinfo.Do Mymsg
       End If
    
    End If
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub FindMyBsic_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 按BSIC查找小区"
    Menu_Flag = 4788
    CI_Cell.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub global_Local_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  本网优于异网综合质量分析"
    If M2_Local = True Then
       Menu_Flag = 886
    Else
       Menu_Flag = 883
    End If
    SelTable.Show 1
    StatusBar.Panels(2).Text = "   "

End Sub

Private Sub global_Other_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  异网优于本网综合质量分析"
    If M2_Local = True Then
       Menu_Flag = 883
    Else
       Menu_Flag = 886
    End If
    SelTable.Show 1
    StatusBar.Panels(2).Text = "   "

End Sub

Private Sub HoppingLabel_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  标注跳频参数"
    Menu_Flag = 1203
    SelTable.Show
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub HoppingStatus_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  跳频状态观测"
    Menu_Flag = 1202
    SelTable.Show
    StatusBar.Panels(2).Text = "  "
    
End Sub

Private Sub LabelTa_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  双网TA标注"
    Menu_Flag = 1204
    SelTable.Show
    StatusBar.Panels(2).Text = "  "
    
End Sub

Private Sub MessagePlay_Click()
    
    On Error Resume Next
    
    MDIMain.Arrange 0     'CASCADE
    StatusBar.Panels(2).Text = " 信令层回放"
    SelTbl = mapinfo.eval("selectionInfo(1)")
    If SelTbl = "" Then
       MsgBox "请选择测量数据中的一个点！", 64, "提示"
       StatusBar.Panels(2).Text = " "
       Exit Sub
    End If
    If UCase(mapinfo.eval("Columninfo( " & SelTbl & ",COL1, 1)")) <> "TIME" Then
       MsgBox "请选择测量数据中的一个点！", 64, "提示"
       StatusBar.Panels(2).Text = " "
       Exit Sub
    End If
       
    MapForm.Show
    mapHWnd = Val(mapinfo.eval("WindowInfo(" & mapid & ",12)"))
    If MapForm.WindowState = 1 Or MapForm.WindowState = 2 Then
       MapForm.WindowState = 0
    End If
    MapForm.Move 0, 0, 5800, 7230

    Load MessageReplay
    MessageReplay.Move 5800, 0, 6000, 7230
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub Mnu_Replay_Click()
    Dim i As Integer
    Dim mapHWnd, j As Long

    On Error Resume Next
    Dim ApiPack As APIPACKET                      'win95
    Dim portnum%                                  'win95
    Dim status%                                   'win95
    Dim majVer%, minVer%, rev%, drvrType%         'win95
    Dim adr%, datum%                              'win95
    portnum% = 4 ' CPlus-B, port 1                     'win95
    status% = RNBOcplusFormatPacket(ApiPack, 1028)     'win95
    status% = RNBOcplusInitialize(ApiPack, portnum%)   'win95
'    If status <> 0 Then GoTo VERYFY_OUT                'win95
    status% = RNBOcplusGetVersion(ApiPack, majVer%, minVer%, rev%, drvrType%)     'win95
    status% = RNBOcplusGetFullStatus(ApiPack)          'win95
    adr = 62                                           'win95
    status% = RNBOcplusRead(ApiPack, adr%, datum%)     'win95
    datum = (datum / 89) * 4 + 23                      'win95
'    If datum <> 427 Then GoTo VERYFY_OUT               'win95
    adr = 60                                           'win95
    status% = RNBOcplusRead(ApiPack, adr%, datum%)     'win95
    datum = (datum / 89) * 4 + 23                      'win95
'    If datum <> 619 Then GoTo VERYFY_OUT               'win95

    StatusBar.Panels(2).Text = " 数据回放分析"
    'Replay.Show 1
    
      i = Val(mapinfo.eval("selectionInfo(3)"))  ' SEL_INFO_NROWS
             SelTbl = mapinfo.eval("selectionInfo(1)")
             If SelTbl = "" Then
                MsgBox "请选择测量数据中的回放起始点", 64, "提示"
                StatusBar.Panels(2).Text = " "
                Exit Sub
             End If
             If UCase(mapinfo.eval("Columninfo( " & SelTbl & ",COL1, 1)")) <> "TIME" Then
                MsgBox "请选择测量数据中的回放起始点", 64, "提示"
                StatusBar.Panels(2).Text = " "
                Exit Sub
             End If

      If i <> 0 Then
    '      MDIMain.SUB_532.Enabled = 1
          Load MapForm
          mapHWnd = Val(mapinfo.eval("WindowInfo(" & mapid & ",12)"))
          If MapForm.WindowState = 1 Or MapForm.WindowState = 2 Then
             MapForm.WindowState = 0
          End If
          MapForm.Move 0, 10, 12000, 4050
    
          Load Graph
          'Graph.Move 0, 4050, 6950, 3150
          Graph.Move 0, 4050, 7000, 3495
    
          Load msgdis
          'msgdis.Move 6950, 4050, 5020, 3150
          msgdis.Move 6950, 4050, 5020, 3495
      End If
    StatusBar.Panels(2).Text = " "
    Exit Sub
VERYFY_OUT:
    MsgBox "加密锁错误, 请与珠海万禾公司联系！", 64, "提示"
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub EXIT_Click()
'     Set mapinfo = Nothing
'    On Error Resume Next
'    mapinfo.runmenucommand 104
'    thereIsAMap = 0
'
'    MapForm.Hide
'    Unload MapForm
'    mapinfo.do "End mapinfo "
'    Gsm_FileName = Gsm_Path + "\map\street.map"
'    Gsm_File2 = Gsm_Path + "\ncell.dbf"
'    Kill Gsm_FileName
'    FileCopy Gsm_File2, Gsm_FileName
'    End
     On Error Resume Next
     Unload Me
End Sub

Private Sub MAIN_8_1_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 系统定义"
    SysDefine.Show 1
End Sub

Private Sub MAIN_8_2_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 系统切换"
    SUB_24_Click
    SysChange.Show 1
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   MDIMain.SetFocus
End Sub

Private Sub My_Center_Click()
    On Error Resume Next
    If SUB_CENTER.Enabled = True Then
       StatusBar.Panels(2).Text = "  显示中心"
       Menu_Flag = 151
       Center.Show 1
       StatusBar.Panels(2).Text = " "
    End If
End Sub

Private Sub MnuBestSCover_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 最佳主小区覆盖"
    Menu_Flag = 991103
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuBetterSCover_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 次佳主小区覆盖"
    Menu_Flag = 991104
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuBYCover_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 本网与异网覆盖显示"
    Menu_Flag = 991106
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuC_AFullDown_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 邻频 C/A +1 (Full) 观测"
    Menu_Flag = 1117
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuC_AFullDown2_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 邻频 Full C/A +2 观测"
    Menu_Flag = 1112
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuC_AFullUp_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 邻频 C/A -1 (Full) 观测"
    Menu_Flag = 1116
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuC_AFullUp2_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 邻频 Full C/A -2 观测"
    Menu_Flag = 1113
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuC_ASubDown_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 邻频 C/A +1 (Sub) 观测"
    Menu_Flag = 1119
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuC_ASubDown2_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 邻频 Sub C/A +2 观测"
    Menu_Flag = 1115
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuC_ASubUp_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 邻频 C/A -1 (Sub) 观测"
    Menu_Flag = 1118
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuC_ASubUp2_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 邻频 Sub C/A -2 观测"
    Menu_Flag = 1114
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuCallAnalyse_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 通话分析"
    Menu_Flag = 991028
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuCallProcess_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 通话过程统计"
    Menu_Flag = 991021
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuCallReplay_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 拨打过程重演"
    Menu_Flag = 991120
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "
    
End Sub

Private Sub MnuCellReUse_Click()
    On Error Resume Next
    Dim CellLon As Variant, CellLat As Variant
    Dim NcellLon As Variant, NcellLat As Variant
    Dim MyRow As Integer, LacRow As Integer
    Dim i As Integer, j As Integer
    Dim MyLac As Variant
    Dim LineFlag As Boolean
    Dim MyColor As Long
        
    StatusBar.Panels(2).Text = " 同LAC观测"
    mapinfo.Do "Select * from base group by Lac into LacTemp"
    MyRow = mapinfo.eval("tableinfo(LacTemp,8)")
    If MyRow = 0 Then
       GoTo Exit1
    End If
    mapinfo.Do "set map redraw off"
    mapinfo.Do "Set Map Layer 0 Editable On  "
    mapinfo.Do "set map redraw on"
    mapinfo.Do "fetch first from LacTemp"
    For i = 1 To MyRow
        MyLac = mapinfo.eval("LacTemp.lac")
        If MyLac > 0 Then
           mapinfo.Do "select * from base where lac = " & MyLac & " into SelLacTemp"
           LineFlag = False
           LacRow = mapinfo.eval("tableinfo(SelLacTemp,8)")
           mapinfo.Do "fetch first from SelLacTemp"
           For j = 1 To LacRow
               CellLon = mapinfo.eval("SelLacTemp.lon")
               CellLat = mapinfo.eval("SelLacTemp.lat")
               If CellLon > 0 And CellLat > 0 Then
                  LineFlag = True
                  Exit For
               End If
               mapinfo.Do "fetch next from SelLacTemp"
           Next
           'mapinfo.do "Set Style Pen MakePen(1,4,16719904)"
           If LineFlag Then
           
          Select Case i
                 Case 1
                       MyColor = 16711680
                 Case 2
                       MyColor = 65280 + 10000
                 Case 3
                       MyColor = 255
                 Case 4
                       MyColor = 16711935
                 Case 5
                       MyColor = 16776960 + 50000
                 Case 6
                       MyColor = 65535
                 Case 7
                       MyColor = 8388608
                 Case 8
                       MyColor = 32768
                 Case 9
                       MyColor = 128
                 Case 10
                       MyColor = 8388736
                 Case 11
                       MyColor = 8421376
                 Case 12
                       MyColor = 32896
                 Case 13
                       MyColor = 16744576
                 Case 14
                       MyColor = 8454016
                 Case 15
                       MyColor = 8421631
                 Case 16
                       MyColor = 16744703
                 Case 17
                       MyColor = 16777088
                 Case 18
                       MyColor = 8454143
                 Case 19
                       MyColor = 8405056
                 Case 20
                       MyColor = 4227136
                 Case 21
                       MyColor = 4210816
                 Case 22
                       MyColor = 8405120
                 Case 23
                       MyColor = 8421440
                 Case 24
                       MyColor = 4227200
                 Case 25
                       MyColor = 16761024
                 Case 26
                       MyColor = 12648384
                 Case 27
                       MyColor = 12632319
                 Case 28
                       MyColor = 16761087
                 Case 29
                       MyColor = 16777152
                 Case 30
                       MyColor = 12648447
                 Case 31
                       MyColor = 8413280
                 Case 32
                       MyColor = 6324320
                 Case 33
                       MyColor = 6316160
                 Case 34
                       MyColor = 8413312
                 Case 35
                       MyColor = 8421472
                 Case 36
                       MyColor = 6324352

                 Case 37
                       MyColor = 16711680
                 Case 38
                       MyColor = 65280
                 Case 39
                       MyColor = 255
                 Case 40
                       MyColor = 16711935
                 Case 41
                       MyColor = 16776960
                 Case 42
                       MyColor = 65535
                 Case 43
                       MyColor = 8388608
                 Case 44
                       MyColor = 32768
                 Case 45
                       MyColor = 128
                 Case 46
                       MyColor = 8388736
                 Case 47
                       MyColor = 8421376
                 Case 48
                       MyColor = 32896
                 Case 49
                       MyColor = 16744576
                 Case 50
                       MyColor = 8454016
                 Case 51
                       MyColor = 8421631
                 Case 52
                       MyColor = 16744703
                 Case 53
                       MyColor = 16777088
                 Case 54
                       MyColor = 8454143
                 Case 55
                       MyColor = 8405056
                 Case 56
                       MyColor = 4227136
                 Case 57
                       MyColor = 4210816
                 Case 58
                       MyColor = 8405120
                 Case 59
                       MyColor = 8421440
                 Case 60
                       MyColor = 4227200
                 Case 61
                       MyColor = 16761024
                 Case 62
                       MyColor = 12648384
                 Case 63
                       MyColor = 12632319
                 Case 64
                       MyColor = 16761087
                 Case 65
                       MyColor = 16777152
                 Case 66
                       MyColor = 12648447
                 Case 67
                       MyColor = 8413280
                 Case 68
                       MyColor = 6324320
                 Case 69
                       MyColor = 6316160
                 Case 70
                       MyColor = 8413312
                 Case 71
                       MyColor = 8421472
                 Case 72
                       MyColor = 6324352
                 Case 73
                       MyColor = 16711680
                 Case 74
                       MyColor = 65280
                 Case 75
                       MyColor = 255
                 Case 76
                       MyColor = 16711935
                 Case 77
                       MyColor = 16776960
                 Case 78
                       MyColor = 65535
                 Case 79
                       MyColor = 8388608
                 Case 80
                       MyColor = 32768
                 Case 81
                       MyColor = 128
                 Case 82
                       MyColor = 8388736
                 Case 83
                       MyColor = 8421376
                 Case 84
                       MyColor = 32896
                 Case 85
                       MyColor = 16744576
                 Case 86
                       MyColor = 8454016
                 Case 87
                       MyColor = 8421631
                 Case 88
                       MyColor = 16744703
                 Case 89
                       MyColor = 16777088
                 Case 90
                       MyColor = 8454143
                 Case 91
                       MyColor = 8405056
                 Case 92
                       MyColor = 4227136
                 Case 93
                       MyColor = 4210816
                 Case 94
                       MyColor = 8405120
                 Case 95
                       MyColor = 8421440
                 Case 96
                       MyColor = 4227200
                 Case 97
                       MyColor = 16761024
                 Case 98
                       MyColor = 12648384
                 Case 99
                       MyColor = 12632319
                 Case 100
                       MyColor = 16761087
           End Select
                mapinfo.Do "Set Style Pen MakePen(1,4," & Format(MyColor) & ")"
                mapinfo.Do "fetch first from SelLacTemp"
                For j = 1 To LacRow
                     NcellLon = mapinfo.eval("SelLacTemp.lon")
                     NcellLat = mapinfo.eval("SelLacTemp.lat")
                     If NcellLon > 0 And NcellLat > 0 Then
                        mapinfo.Do "create Line(" & CellLon & "," & CellLat & ")(" & NcellLon & "," & NcellLat & ")"
                     End If
                     mapinfo.Do "fetch next from SelLacTemp"
                Next
           End If
        End If
        mapinfo.Do "fetch next from LacTemp"
    Next
    
Exit1:
    mapinfo.Do "close table LacTemp"
    mapinfo.Do "close table SelLacTemp"
    StatusBar.Panels(2).Text = "按鼠标右键清除装饰层"
End Sub

Private Sub MnuDcsBcchLabel_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  DCS参数观测"
    Menu_Flag = 55556
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub MnuDcsRxLevSub_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  DCS参数观测"
    Menu_Flag = 55551
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub MnuDcsRxQualFull_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  DCS参数观测"
    Menu_Flag = 55553
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub MnuDcsRxQualSub_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  DCS参数观测"
    Menu_Flag = 55552
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub MnuDcsTa_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  DCS参数观测"
    Menu_Flag = 55554
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub MnuDcsTchLabel_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  DCS参数观测"
    Menu_Flag = 55557
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub MnuDcsTx_Power_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  DCS参数观测"
    Menu_Flag = 55555
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub MnuCHANNELRELEASE_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 信令事件成因——CHANNEL RELEASE"
    Menu_Flag = 9008
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuDataGraph_Click()
    Dim mySelTbl As String
    On Error Resume Next
    StatusBar.Panels(2).Text = " 数据统计图"
             SelTbl = mapinfo.eval("selectionInfo(1)")
             If SelTbl = "" Then
                MsgBox "请先用半径选择选择一段测试数据！", 64, "提示"
                StatusBar.Panels(2).Text = " "
                Exit Sub
             End If
             If UCase(mapinfo.eval("Columninfo( " & SelTbl & ",COL1, 1)")) <> "TIME" Then
                MsgBox "请先用半径选择选择一段测试数据！", 64, "提示"
                StatusBar.Panels(2).Text = " "
                Exit Sub
             End If

    frmDataGraph.Show 1
    StatusBar.Panels(2).Text = " "
    
End Sub

Private Sub MnuDISCONNECT_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 信令事件成因——DISCONNECT"
    Menu_Flag = 9006
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub Mnudistributing_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 选频场强分布"
    Menu_Flag = 1999
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
    
End Sub

Private Sub MnuDisturbSearch_Click()
    Dim mySelTbl As String
    On Error Resume Next
    StatusBar.Panels(2).Text = " 指定范围内同邻频查找"
             SelTbl = mapinfo.eval("selectionInfo(1)")
             If SelTbl = "" Then
                MsgBox "请先用半径选择选择一段测试数据！", 64, "提示"
                StatusBar.Panels(2).Text = " "
                Exit Sub
             End If
             If UCase(mapinfo.eval("Columninfo( " & SelTbl & ",COL1, 1)")) <> "TIME" Then
                MsgBox "请先用半径选择选择一段测试数据！", 64, "提示"
                StatusBar.Panels(2).Text = " "
                Exit Sub
             End If

    frmDisturbSearch.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuDSelBcch_Click()
    Dim mySelTbl As String
    
    On Error Resume Next
    StatusBar.Panels(2).Text = " 选频覆盖统计"
    'SelTbl = mapinfo.eval("selectionInfo(1)")
    'If SelTbl = "" Then
    '    MsgBox "请先用半径选择选择一段测试数据！", 64, "提示"
    '    StatusBar.Panels(2).Text = " "
    '    Exit Sub
    'End If
    'If UCase(mapinfo.eval("Columninfo( " & SelTbl & ",COL1, 1)")) <> "TIME" Then
    '    MsgBox "请先用半径选择选择一段测试数据！", 64, "提示"
    '    StatusBar.Panels(2).Text = " "
    '    Exit Sub
    'End If
    'frmDurbSelBcch.Show 1
    Menu_Flag = 991112
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuFullSubSwitch_Click()

    On Error Resume Next
    If Back_Sel = 0 Then
       Back_Sel = 1
    Else
       Back_Sel = 0
    End If
End Sub

Private Sub MnuGDCover_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " G网与D网覆盖显示"
    Menu_Flag = 991105
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuGsm_Dcs_Click()
    
    On Error Resume Next
    StatusBar.Panels(2).Text = "  GSM/DCS切换点"
    Menu_Flag = 55550
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
    
End Sub

Private Sub MnuHandOverCause_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 切换发生成因"
    Menu_Flag = 9020
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "
    
End Sub

Private Sub MnuHANDOVERFAILUER_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 信令事件成因——HANDOVER FAILURE"
    Menu_Flag = 9005
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuHOPara_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  切换前后参数显示"
    Menu_Flag = 991109
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub MnuHOStat_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  切换统计"
    Menu_Flag = 991110
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub MnuLabel_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  参数标注"
    Menu_Flag = 1111
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
        
End Sub

Private Sub MnuLabelMark_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  采集事件标注"
    Menu_Flag = 1657
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub MnuLacLegend_Click()
    On Error Resume Next
    LacLegend.Show
    LacLegend.Move 9000, 100, 2040, 2520
End Sub

Private Sub MnuMessageAnalyse_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 干扰分析"
    Menu_Flag = 9002
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuMustHandover_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 功率预算切换统计"
    Menu_Flag = 9010
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuNewCallCover_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 通话覆盖分布"
    Menu_Flag = 9004
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuNcellReason_Click()
    Dim SelTbl As String, MySelName As String
    Dim MySelCellName As String, MySelCellCI As String
    
    On Error Resume Next
    StatusBar.Panels(2).Text = " 邻小区合理性统计"
    SelTbl = mapinfo.eval("selectionInfo(1)")
    MySelName = mapinfo.eval("selectioninfo(2)")
    If UCase(SelTbl) = "CELL" Then
        MyNRSelCellName = mapinfo.eval("selection.cell_name")
        MyNRSelCellCI = mapinfo.eval("selection.ci")
        MyNRSelCellLac = mapinfo.eval("selection.lac")
        MyNRSelCellBcch = mapinfo.eval("selection.arfcn")
        MyNRSelCellBsic = mapinfo.eval("selection.bsic")
        MyNRSelCellBcch_2 = 0
        MyNRSelCellBcch_3 = 0
        mapinfo.Do "x2 = selection.lon"
        mapinfo.Do "y2 = selection.lat"
        mapinfo.Do "close table " & MySelName
    ElseIf UCase(SelTbl) = "BASE" Then
        MyNRSelCellName = mapinfo.eval("selection.bs_name")
        MyNRSelCellCI = mapinfo.eval("selection.ci_1")
        MyNRSelCellBcch = mapinfo.eval("selection.bcch_1")
        MyNRSelCellBsic = mapinfo.eval("selection.bsic_1")
        If mapinfo.eval("selection.bcch_2") > 0 Then
            MyNRSelCellCI_2 = mapinfo.eval("selection.ci_2")
            MyNRSelCellBcch_2 = mapinfo.eval("selection.bcch_2")
            MyNRSelCellBsic_2 = mapinfo.eval("selection.bsic_2")
            If mapinfo.eval("selection.bcch_3") > 0 Then
                MyNRSelCellCI_3 = mapinfo.eval("selection.ci_3")
                MyNRSelCellBcch_3 = mapinfo.eval("selection.bcch_3")
                MyNRSelCellBsic_3 = mapinfo.eval("selection.bsic_3")
            End If
        End If
        MyNRSelCellLac = mapinfo.eval("selection.lac")
        mapinfo.Do "x2 = selection.lon"
        mapinfo.Do "y2 = selection.lat"
        mapinfo.Do "close table " & MySelName
    Else
        'MsgBox "请选择一个Cell", 64, "提示"
        If MySelName <> "" Then              '?
            mapinfo.Do "close table " & MySelName
        End If
        MyNRSelCellCI = ""
        MyNRSelCellBcch = 0
        MyNRSelCellBsic = 0
        MyNRSelCellLac = ""
    End If
    Menu_Flag = 991108
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuNewCover_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 网络资源覆盖盲区"
    Menu_Flag = 9001
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub MnuNewDisturb_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 干扰分析"
    Menu_Flag = 9002
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuNewDropCall_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 网络掉话区"
    Menu_Flag = 9003
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuRELEASE_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 信令事件成因——RELEASE"
    Menu_Flag = 9007
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub MnuSql_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " SQL 下行话音质量观测"
    Menu_Flag = 1120
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub My_ScanPlay_Click()
    Dim i As Integer
    Dim mapHWnd, j As Long

    On Error Resume Next
    Dim ApiPack As APIPACKET                      'win95
    Dim portnum%                                  'win95
    Dim status%                                   'win95
    Dim majVer%, minVer%, rev%, drvrType%         'win95
    Dim adr%, datum%                              'win95
    portnum% = 4 ' CPlus-B, port 1                     'win95
    status% = RNBOcplusFormatPacket(ApiPack, 1028)     'win95
    status% = RNBOcplusInitialize(ApiPack, portnum%)   'win95
'    If status <> 0 Then GoTo VERYFY_OUT                'win95
    status% = RNBOcplusGetVersion(ApiPack, majVer%, minVer%, rev%, drvrType%)     'win95
    status% = RNBOcplusGetFullStatus(ApiPack)          'win95
    adr = 62                                           'win95
    status% = RNBOcplusRead(ApiPack, adr%, datum%)     'win95
    datum = (datum / 89) * 4 + 23                      'win95
'    If datum <> 427 Then GoTo VERYFY_OUT               'win95
    adr = 60                                           'win95
    status% = RNBOcplusRead(ApiPack, adr%, datum%)     'win95
    datum = (datum / 89) * 4 + 23                      'win95
'    If datum <> 619 Then GoTo VERYFY_OUT               'win95
    
    'dog = scread(62)     'win95
    'dog = (dog / 89) * 2 + 23     'win95
    'If dog <> 225 Then GoTo VERYFY_OUT     'win95

    StatusBar.Panels(2).Text = " 数据回放分析"
'    If sys = 0 Then
'       Replay.Show 1
'    Else
       SelTbl = mapinfo.eval("selectionInfo(1)")
       If UCase(mapinfo.eval("Columninfo( " & SelTbl & ",COL3, 1)")) <> "TIME" Then
          MsgBox "请选择扫频测量数据中的一个点！", 64, "提示"
          StatusBar.Panels(2).Text = " "
          Exit Sub
       End If
       MapForm.Show
       mapHWnd = Val(mapinfo.eval("WindowInfo(" & mapid & ",12)"))
       If MapForm.WindowState = 1 Or MapForm.WindowState = 2 Then
          MapForm.WindowState = 0
       End If
       MapForm.Move 0, 0, 5600, 7300
       Load Scan_Frm
       Scan_Frm.Move 5600, -35, 6000, 7300

'      ScanGra.Show
'      ScanGra.Move 1440, 4300
'    End If
    StatusBar.Panels(2).Text = " "
    Exit Sub
VERYFY_OUT:
    MsgBox "加密锁错误, 请与珠海万禾公司联系！", 64, "提示"
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub Ncell_Map_Click()
    Dim My_Ci As String
    Dim ci(16) As String
    Dim MyRecord As Record
    Dim MySelName As String
    Dim Mymsg As String
    Dim j As Integer
    Dim CellLon As Variant, CellLat As Variant
    Dim NcellLon As Variant, NcellLat As Variant
    
    On Error Resume Next
    Gsm_FileName = Gsm_Path + "\gsm.dat"
    Open Gsm_FileName For Binary As #1
    Get #1, 1, MyRecord
    Close #1
    StatusBar.Panels(2).Text = "  NCELL观察"
    MySelName = mapinfo.eval("selectioninfo(2)")
    SelTbl = mapinfo.eval("selectionInfo(1)")
    If SelTbl <> "cell" Then
       MsgBox "请选择一个Cell ！", 64, "提示"
    Else
       CellLon = mapinfo.eval("selection.lon") + 0.0015 * Sin(mapinfo.eval("selection.bearing") * 0.01745329252)
       CellLat = mapinfo.eval("selection.lat") + 0.0015 * Cos(mapinfo.eval("selection.bearing") * 0.01745329252)
       If MyRecord.exchange = 0 Then
          My_Ci = mapinfo.eval("selection.bs_no")
          If My_Ci <> "" Then
             ci(1) = mapinfo.eval("selection.ncell1")
             If ci(1) = "" Then
                GoTo ExitSub
             End If
             Mymsg = "select * from cell where bs_no = " + Chr(34) + ci(1) + Chr(34)
             For j = 2 To 16
                 ci(j) = mapinfo.eval("selection.ncell" & Format(j))
                 If ci(j) <> "" Then
                    Mymsg = Mymsg + " or bs_no = " + Chr(34) + ci(j) + Chr(34)
                 End If
             Next j
             Mymsg = Mymsg + " into neighber_cell"
             mapinfo.Do Mymsg
             GoTo ericsson_ncell
          Else
             GoTo ExitSub
          End If
       End If
       My_Ci = Val(mapinfo.eval("selection.ci"))
       'Gsm_FileName = Gsm_Path + "\map\ncell.tab"
       'mapinfo.Do "open table " + Chr(34) + Gsm_FileName + Chr(34)
       i = 0
       row = Val(mapinfo.eval("tableinfo(cell,8)"))
       mapinfo.Do "fetch First from cell"
       Msg = mapinfo.eval("cell.ci")
       While i < row And Msg <> My_Ci
           mapinfo.Do "fetch next from cell"
           Msg = mapinfo.eval("cell.ci")
           i = i + 1
       Wend
       If i < row Then
           msg1 = "select  *  from cell  where  ci = " + Chr(34) + "ABCD" + Chr(34)
           For j = 1 To 16
               Msg = "cell.ncell" & j
               ci(j) = mapinfo.eval(Msg)
               If ci(j) <> "" And ci(j) <> "F" Then
                  msg1 = msg1 + "   or ci= " + Chr(34) + ci(j) + Chr(34)
               End If
           Next j
           msg1 = msg1 + "  into neighber_cell"
           mapinfo.Do msg1
ericsson_ncell:

           Msg = "Add Map Auto Layer " + Chr(34) + "neighber_cell" + Chr(34)
           mapinfo.Do Msg
             mapinfo.Do "set map redraw off"
             mapinfo.Do "Set Map Layer 0 Editable On  "
             mapinfo.Do "set map redraw on"
             
            row = Val(mapinfo.eval("tableinfo(neighber_cell,8)"))
            mapinfo.Do " fetch first from neighber_cell"
            'mapinfo.do "Set Style Pen MakePen(1,4,16719904)"
            mapinfo.Do "Set Style Pen MakePen(2,4,16719904)"
            For i = 1 To row
                NcellLon = mapinfo.eval("neighber_cell.lon") + 0.0015 * Sin(mapinfo.eval("neighber_cell.bearing") * 0.01745329252)
                NcellLat = mapinfo.eval("neighber_cell.lat") + 0.0015 * Cos(mapinfo.eval("neighber_cell.bearing") * 0.01745329252)
                mapinfo.Do "create Line(" & CellLon & "," & CellLat & ")(" & NcellLon & "," & NcellLat & ")"
                mapinfo.Do "fetch  next from neighber_cell"
            Next
                      
           Msg = " shade window FrontWindow()  neighber_cell with arfcn "
           Msg = Msg + "ignore 0 values  1 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0) ,"
           Msg = Msg + "2 Symbol (33,65280,8,""MapInfo Cartographic"",0,0) ,3 Symbol (33,255,8,""MapInfo Cartographic"",0,0) ,"
           Msg = Msg + "4 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0) ,5 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0) ,"
           Msg = Msg + "6 Symbol (33,65535,8,""MapInfo Cartographic"",0,0) ,7 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0) ,"
           Msg = Msg + "8 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),9 Symbol (33,128,8,""MapInfo Cartographic"",0,0),10 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),11 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "12 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),13 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),14 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),15 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "16 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0),17 Symbol (33,16777088,8,""MapInfo Cartographic"",0,0),18 Symbol (33,8454143,8,""MapInfo Cartographic"",0,0),19 Symbol (33,8405056,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "20 Symbol (33,4227136,8,""MapInfo Cartographic"",0,0),21 Symbol (33,4210816,8,""MapInfo Cartographic"",0,0),22 Symbol (33,8405120,8,""MapInfo Cartographic"",0,0),23 Symbol (33,8421440,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "24 Symbol (33,4227200,8,""MapInfo Cartographic"",0,0),25 Symbol (33,16761024,8,""MapInfo Cartographic"",0,0),26 Symbol (33,12648384,8,""MapInfo Cartographic"",0,0),27 Symbol (33,12632319,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "28 Symbol (33,16761087,8,""MapInfo Cartographic"",0,0),29 Symbol (33,16777152,8,""MapInfo Cartographic"",0,0),30 Symbol (33,12648447,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "31 Symbol (33,8413280,8,""MapInfo Cartographic"",0,0),32 Symbol (33,6324320,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "33 Symbol (33,6316160,8,""MapInfo Cartographic"",0,0),34 Symbol (33,8413312,8,""MapInfo Cartographic"",0,0),35 Symbol (33,8421472,8,""MapInfo Cartographic"",0,0),36 Symbol (33,6324352,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "37 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0),38 Symbol (33,65280,8,""MapInfo Cartographic"",0,0),39 Symbol (33,255,8,""MapInfo Cartographic"",0,0),40 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "41 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0),42 Symbol (33,65535,8,""MapInfo Cartographic"",0,0),43 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0),44 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "45 Symbol (33,128,8,""MapInfo Cartographic"",0,0),46 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),47 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),48 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "49 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),50 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),51 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),52 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "53 Symbol (33,16777088,8,""MapInfo Cartographic"",0,0),54 Symbol (33,8454143,8,""MapInfo Cartographic"",0,0),55 Symbol (33,8405056,8,""MapInfo Cartographic"",0,0),56 Symbol (33,4227136,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "57 Symbol (33,4210816,8,""MapInfo Cartographic"",0,0),58 Symbol (33,8405120,8,""MapInfo Cartographic"",0,0),59 Symbol (33,8421440,8,""MapInfo Cartographic"",0,0),60 Symbol (33,4227200,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "61 Symbol (33,16761024,8,""MapInfo Cartographic"",0,0),62 Symbol (33,12648384,8,""MapInfo Cartographic"",0,0),63 Symbol (33,12632319,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "64 Symbol (33,16761087,8,""MapInfo Cartographic"",0,0),65 Symbol (33,16777152,8,""MapInfo Cartographic"",0,0),66 Symbol (33,12648447,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "67 Symbol (33,8413280,8,""MapInfo Cartographic"",0,0),68 Symbol (33"
           Msg = Msg + ",6324320,8,""MapInfo Cartographic"",0,0),69 Symbol (33,6316160,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "70 Symbol (33,8413312,8,""MapInfo Cartographic"",0,0),71 Symbol (33,8421472,8,""MapInfo Cartographic"",0,0),72 Symbol (33,6324352,8,""MapInfo Cartographic"",0,0),73 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "74 Symbol (33,65280,8,""MapInfo Cartographic"",0,0),75 Symbol (33,255,8,""MapInfo Cartographic"",0,0),76 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0),77 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "78 Symbol (33,65535,8,""MapInfo Cartographic"",0,0),79 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0),80 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),81 Symbol (33,128,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "82 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),83 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),84 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),85 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "86 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),87 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),88 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0),89 Symbol (33,16777088,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "90 Symbol (33,8454143,8,""MapInfo Cartographic"",0,0),91 Symbol (33,8405056,8,""MapInfo Cartographic"",0,0),92 Symbol (33,4227136,8,""MapInfo Cartographic"",0,0),93 Symbol (33,4210816,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "94 Symbol (33,8405120,8,""MapInfo Cartographic"",0,0),95 Symbol (33,8421440,8,""MapInfo Cartographic"",0,0),96 Symbol (33,4227200,8,""MapInfo Cartographic"",0,0),97 Symbol (33,16761024,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "98 Symbol (33,12648384,8,""MapInfo Cartographic"",0,0),99 Symbol (33,12632319,8,""MapInfo Cartographic"",0,0),100 Symbol (33,16761087,8,""MapInfo Cartographic"",0,0),101 Symbol (33,16777152,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "102 Symbol (33,12648447,8,""MapInfo Cartographic"",0,0),103 Symbol (33,8413280,8,""MapInfo Cartographic"",0,0),104 Symbol (33,6324320,8,""MapInfo Cartographic"",0,0),105 Symbol (33,6316160,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "106 Symbol (33,8413312,8,""MapInfo Cartographic"",0,0),107 Symbol (33,8421472,8,""MapInfo Cartographic"",0,0),108 Symbol (33,6324352,8,""MapInfo Cartographic"",0,0),109 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "110 Symbol (33,65280,8,""MapInfo Cartographic"",0,0),111 Symbol (33,255,8,""MapInfo Cartographic"",0,0),112 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0),113 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "114 Symbol (33,65535,8,""MapInfo Cartographic"",0,0),115 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0),116 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),117 Symbol (33,128,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "118 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),119 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),120 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),121 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),"
           Msg = Msg + "122 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),123 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),124 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0)"
           mapinfo.Do Msg
           If legendid = 0 Then
              mapinfo.Do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
              mapinfo.Do "Create Legend From Window  Frontwindow()"
              legendid = mapinfo.eval("windowinfo(1009,12)")
           End If
           Msg = " Title " + Chr(34) + "NCell 观测  " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "图例中显示所有邻小区的载频" + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off"
           mapinfo.Do "set legend window   Frontwindow()  Layer prev " & Msg
       End If
    End If
ExitSub:
   StatusBar.Panels(2).Text = "按鼠标右键清除装饰层"
    mapinfo.Do "close table " & MySelName
End Sub

Private Sub Obtel_Data_Click()
    On Error Resume Next
    Data_Report = False
    TObtel_Click
End Sub

Private Sub Obtel_Data_Report_Click()
    On Error Resume Next
    Data_Report = True
    TObtel_Click
End Sub

Private Sub OPENWOR_Click()
    On Error Resume Next
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    mapinfo.runmenucommand 108     'M_TOOLS_TEXT
    mapinfo.Do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
    mapinfo.Do "Create Legend From Window Frontwindow()"
    SUB_26.Enabled = 1
    SUB_151.Enabled = 1
    SUB_41.Enabled = 1
    SUB_461.Enabled = 1
    SUB_462.Enabled = 1
    SUB_463.Enabled = 1
    SUB_464.Enabled = 1
    MnuBcchRetrieve.Enabled = True
    SUB_465.Enabled = 1
    MnuCellReUse.Enabled = True
    SUB_466.Enabled = 1
    SUB_467.Enabled = 1
    SUB_468.Enabled = 1
    FindMyBsic.Enabled = 1
    FindFree.Enabled = 1
    SUB_469.Enabled = 1
    BsNo_FindCell.Enabled = 1
    SUB_4600.Enabled = 1
     
    SUB_24.Enabled = 1
End Sub

Private Sub Help_Click()
     On Error Resume Next
     Gsm_FileName = "winhelp.exe  " + Gsm_Path + "\gsm.hlp    "
'     ReturnValue = Shell("winhelp.EXE  \gsm\gsm.hlp   ", 3)
     ReturnValue = Shell(Gsm_FileName, 3)
End Sub

Private Sub Label_Click()
        On Error Resume Next
        mapinfo.runmenucommand 801   'M_TOOLS_RECENTER
End Sub

Private Sub M_legend_Click()
        On Error Resume Next
        mapinfo.runmenucommand 308   'M_TOOLS_RECENTER
End Sub

Private Sub MMM_Click()
        On Error Resume Next
        mapinfo.runmenucommand 308   'M_TOOLS_RECENTER
End Sub

Private Sub MAP_DIS_Click()

    On Error Resume Next
    Select Case dis_flag
    Case 0
        mapinfo.Do "set map display position"
        dis_flag = 1
    Case 1
        mapinfo.Do "set map display scale"
        dis_flag = 2
    Case 2
        mapinfo.Do "set map display zoom"
        dis_flag = 0
    End Select
End Sub

Private Sub My_Over_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 生成本地覆盖图"
    Menu_Flag = 917
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub Pilot_data_Click()
    Dim buff As String, mypath As String
    Dim finds As Integer, i As Integer
        
    On Error Resume Next
    Menu_Flag = 2222
    tran_del = 0
    StatusBar.Panels(2).Text = " 数据转换"
    For i = 1 To 50
        convert_filename(i) = ""
    Next
open_again:
    FileDialog.DialogTitle = "ANT Pilot 通话测试数据转换"
    'FileDialog.Filter = "*.dbf Files|*.DBF"
    'FileDialog.DefaultExt = "*.DBF"
    'FileDialog.Filter = "*.log Files|*.LOG|*.ant Files|*.ANT|All Files|*.*"
    FileDialog.Filter = "*.ant Files|*.ANT|*.log Files|*.LOG|All Files|*.*"
    FileDialog.DefaultExt = "*.LOG"
    FileDialog.Flags = &H80000 Or &H200
    Gsm_FileName = Gsm_Path
    FileDialog.InitDir = Gsm_FileName
    FileDialog.ShowOpen
    buff = Trim(FileDialog.filename)
    If buff = "" Then
       StatusBar.Panels(2).Text = " "
       Exit Sub
    End If
    tran_fn = 0
    finds = InStr(buff, Chr(0))
    If finds > 0 Then
       mypath = Left(buff, finds - 1) + "\"
       buff = Trim(Right(buff, Len(buff) - finds))
       finds = InStr(buff, Chr(0))
       i = 1
       Do While finds > 0
          convert_filename(i) = mypath + Left(buff, finds - 1)
          tran_f(i) = convert_filename(i)
          buff = Trim(Right(buff, Len(buff) - finds))
          finds = InStr(buff, Chr(0))
          i = i + 1
       Loop
       convert_filename(i) = mypath + buff
       tran_f(i) = convert_filename(i)
       tran_fn = i
    Else
       convert_filename(1) = buff
       tran_f(1) = buff
       tran_fn = 1
    End If
    FileDialog.filename = ""
    If Dir(convert_filename(1)) = "" Then
       GoTo err_exit
    End If
    sinput = tran_f(1)
    FileDialog.filename = ""
    'DocManager.Show 1
    cvChoice.Show 1
    StatusBar.Panels(2).Text = " "
    Exit Sub
    
err_exit:
       i = MsgBox("无法打开文件 " + convert_filename(1), 48, "打开文件")
       GoTo open_again

End Sub

Private Sub Pilot_data_Report_Click()
    Dim buff As String, mypath As String
    Dim finds As Integer, i As Integer
        
    On Error Resume Next
    Menu_Flag = 4444
    tran_del = 0
    StatusBar.Panels(2).Text = " 数据转换"
    For i = 1 To 50
        convert_filename(i) = ""
    Next
open_again:
    FileDialog.DialogTitle = "ANT Pilot 通话测试数据转换"
    'FileDialog.Filter = "*.dbf Files|*.DBF"
    'FileDialog.DefaultExt = "*.DBF"
    FileDialog.Filter = "*.log Files|*.LOG|*.ant Files|*.ANT|All Files|*.*"
    FileDialog.DefaultExt = "*.LOG"
    FileDialog.Flags = &H80000 Or &H200
    Gsm_FileName = Gsm_Path
    FileDialog.InitDir = Gsm_FileName
    FileDialog.ShowOpen
    buff = Trim(FileDialog.filename)
    If buff = "" Then
       StatusBar.Panels(2).Text = " "
       Exit Sub
    End If
    tran_fn = 0
    finds = InStr(buff, Chr(0))
    If finds > 0 Then
       mypath = Left(buff, finds - 1) + "\"
       buff = Trim(Right(buff, Len(buff) - finds))
       finds = InStr(buff, Chr(0))
       i = 1
       Do While finds > 0
          convert_filename(i) = mypath + Left(buff, finds - 1)
          tran_f(i) = convert_filename(i)
          buff = Trim(Right(buff, Len(buff) - finds))
          finds = InStr(buff, Chr(0))
          i = i + 1
       Loop
       convert_filename(i) = mypath + buff
       tran_f(i) = convert_filename(i)
       tran_fn = i
    Else
       convert_filename(1) = buff
       tran_f(1) = buff
       tran_fn = 1
    End If
    FileDialog.filename = ""
    If Dir(convert_filename(1)) = "" Then
       GoTo err_exit
    End If
    sinput = tran_f(1)
    FileDialog.filename = ""
    'DocManager.Show 1
    cvChoice.Show 1
    StatusBar.Panels(2).Text = " "
    Exit Sub
    
err_exit:
       i = MsgBox("无法打开文件 " + convert_filename(1), 48, "打开文件")
       GoTo open_again

End Sub

Private Sub RadioLink_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  无线链路丢失状况"
    Menu_Flag = 1201
    SelTable.Show
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub Rxlev_Local_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  本网优于异网RxLev分析"
    If M2_Local = True Then
       Menu_Flag = 888
    Else
       Menu_Flag = 885
    End If
    SelTable.Show 1
    StatusBar.Panels(2).Text = "   "

End Sub

Private Sub Rxlev_Other_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  异网优于本网RxLev分析"
    If M2_Local = True Then
       Menu_Flag = 885
    Else
       Menu_Flag = 888
    End If
    SelTable.Show 1
    StatusBar.Panels(2).Text = "   "
End Sub

Private Sub RxQual_Local_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  本网优于异网RxQual分析"
    If M2_Local = True Then
       Menu_Flag = 887
    Else
       Menu_Flag = 884
    End If
    SelTable.Show 1
    StatusBar.Panels(2).Text = "   "

End Sub

Private Sub RxQual_Other_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  异网优于本网RxQual分析"
    If M2_Local = True Then
       Menu_Flag = 884
    Else
       Menu_Flag = 887
    End If
    SelTable.Show 1
    StatusBar.Panels(2).Text = "   "
End Sub

Private Sub SaveWindows_Click()
    On Error Resume Next
    If thereIsAMap Then
       mapinfo.runmenucommand 609
    End If
    
End Sub

Private Sub SaveWor_Click()
    On Error Resume Next
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    mapinfo.runmenucommand 109     'M_TOOLS_TEXT
End Sub

Private Sub SCAN_7_Click()
  On Error Resume Next
  StatusBar.Panels(2).Text = " 观测C/I1"
  Menu_Flag = 918
  SelTable.Show 1
  StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SCAN_8_Click()
  On Error Resume Next
  StatusBar.Panels(2).Text = " 观测C/I2"
  Menu_Flag = 919
  SelTable.Show 1
  StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SSCommand10_Click()
 Dim bcch_no As Integer
 Dim MySelName As String
 Dim MyRow As Long
 Dim NcellLon As Variant, NcellLat As Variant
 Dim CellLon As Variant, CellLat As Variant
 
 On Error Resume Next
  StatusBar.Panels(2).Text = "  同频观察"
  SelTbl = mapinfo.eval("selectionInfo(1)")
  MySelName = mapinfo.eval("selectioninfo(2)")
  If SelTbl <> "cell" Then
     MsgBox "请选择一个Cell", 64, "提示"
  Else
       'CellLon = mapinfo.eval("selection.lon") + 0.0015 * Sin(mapinfo.eval("selection.bearing") * 0.01745329252)
       'CellLat = mapinfo.eval("selection.lat") + 0.0015 * Cos(mapinfo.eval("selection.bearing") * 0.01745329252)
       mapinfo.Do "x1=selection.lon + 0.0015 * Sin(selection.bearing * 0.01745329252)"
       mapinfo.Do "y1=selection.lat + 0.0015 * Cos(selection.bearing * 0.01745329252)"
     CchTch_Frm.Show 1
     If SearchDistance = 19999 Then
        mapinfo.Do "close table " & MySelName
        Exit Sub
     End If
     bcch_no = Val(mapinfo.eval("selection.arfcn"))
     mapinfo.Do "close table " & MySelName
     If CELL_CCH = 1 Then
        If SearchDistance = 0 Then
           mapinfo.Do "select * from cell where ARFCN = " & bcch_no & " into same_arfcn"
        Else
           mapinfo.Do "select * from cell where ARFCN = " & bcch_no & " and ((x1-(lon+0.0015*sin(bearing*0.01745329252)))^2 +(y1-(lat+0.0015*sin(bearing*0.01745329252)))^2)<" & Format(SearchDistance * 0.0021 / 3) & " into same_arfcn"
        End If
     Else
        If SearchDistance = 0 Then
           mapinfo.Do "Select * from cell where Like(Non_bcch,""%" & Format(bcch_no) & "%"","""") = 1 into same_arfcn"
        Else
           mapinfo.Do "Select * from cell where Like(Non_bcch,""%" & Format(bcch_no) & "%"","""") = 1 and ((x1-(lon+0.0015*sin(bearing*0.01745329252)))^2 +(y1-(lat+0.0015*sin(bearing*0.01745329252)))^2)<" & Format(SearchDistance * 0.0021 / 3) & " into same_arfcn"
        End If
     End If
     If Val(mapinfo.eval("tableinfo(same_arfcn,8)")) <= 1 Then
        If CELL_CCH = 1 Then
           MsgBox "不存在BCCH同为 " & bcch_no & " 的小区", 64, "提示"
        Else
           MsgBox "不存在TCH同为 " & bcch_no & " 的小区", 64, "提示"
        End If
        mapinfo.Do "close table same_arfcn"
        Exit Sub
     End If
     mapinfo.Do "Add Map Auto Layer " + Chr(34) + "same_arfcn" + Chr(34)
             
             mapinfo.Do "set map redraw off"
             mapinfo.Do "Set Map Layer 0 Editable On  "
             mapinfo.Do "set map redraw on"
             
            MyRow = Val(mapinfo.eval("tableinfo(same_arfcn,8)"))
            mapinfo.Do " fetch first from same_arfcn"
            'mapinfo.do "Set Style Pen MakePen(1,4,16719904)"
            mapinfo.Do "Set Style Pen MakePen(2,4," & Format(Xiaoyu_Color(Linyujin)) & ")"
            Linyujin = Linyujin + 1
            Linyujin = Linyujin Mod 5
            For i = 1 To MyRow
                mapinfo.Do "x2=same_arfcn.lon + 0.0015 * Sin(same_arfcn.bearing * 0.01745329252)"
                mapinfo.Do "y2=same_arfcn.lat + 0.0015 * Cos(same_arfcn.bearing * 0.01745329252)"
                'NcellLon = mapinfo.eval("same_arfcn.lon") + 0.0015 * Sin(mapinfo.eval("same_arfcn.bearing") * 0.01745329252)
                'NcellLat = mapinfo.eval("same_arfcn.lat") + 0.0015 * Cos(mapinfo.eval("same_arfcn.bearing") * 0.01745329252)
                       'If (CellLat - NcellLat) * (CellLat - NcellLat) + (CellLon - NcellLon) * (CellLon - NcellLon) < 0.000001 Then    150米
                'If (CellLat - NcellLat) * (CellLat - NcellLat) + (CellLon - NcellLon) * (CellLon - NcellLon) < 0.0021 Then    '5公里
                    mapinfo.Do "create Line(x1,y1)(x2,y2)"
                'End If
                mapinfo.Do "fetch  next from same_arfcn"
            Next
     
     If CELL_CCH = 1 Then
        mapinfo.Do "shade window Frontwindow() same_arfcn with ARFCN values " + Chr(34) & bcch_no & Chr(34) + " Symbol (58,16711935,12)"
     Else
        mapinfo.Do "shade window Frontwindow() same_arfcn with Like(Non_bcch,""%" & Format(bcch_no) & "%"","""") values 1 Symbol (58,16711935,12)"
     End If

     If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
     End If
     If CELL_CCH = 1 Then
        If SearchDistance = 0 Then
           Msg = " Title " + Chr(34) + " 同频观测 (CCH)" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle ""全网"" Font (""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off"
        Else
           Msg = " Title " + Chr(34) + " 同频观测 (CCH)" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle ""只找出" & Format(SearchDistance) & "公里以内的小区"" Font (""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off"
        End If
     Else
       'msg = " Title " + Chr(34) + " 同频观测 " + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
       If SearchDistance = 0 Then
           Msg = " Title " + Chr(34) + " 同频观测 (TCH)" + Chr(34) + " Font (""宋体"",0,9,0) Subtitle ""全网"" Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ," & Chr(34) & Format(bcch_no) & Chr(34) & " display on"
       Else
          Msg = " Title " + Chr(34) + " 同频观测 (TCH)" + Chr(34) + " Font (""宋体"",0,9,0) Subtitle ""只找出" & Format(SearchDistance) & "公里以内的小区"" Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ," & Chr(34) & Format(bcch_no) & Chr(34) & " display on"
       End If
     End If
     mapinfo.Do "set legend window Frontwindow() Layer prev " & Msg
 End If
    StatusBar.Panels(2).Text = "按鼠标右键清除装饰层"
End Sub

Private Sub SSCommand11_Click()
 Dim bcch_no As Integer
 Dim aad, ddd As Integer
 Dim MySelName As String
 Dim MyRow As Integer
 Dim NcellLon As Variant, NcellLat As Variant
 Dim CellLon As Variant, CellLat As Variant
 
 On Error Resume Next
  StatusBar.Panels(2).Text = "  邻频观察"
  SelTbl = mapinfo.eval("selectionInfo(1)")
  MySelName = mapinfo.eval("selectioninfo(2)")
  If SelTbl <> "cell" Then
     MsgBox "请选择一个Cell ！", 64, "提示"
  Else
       mapinfo.Do "x1=selection.lon + 0.0015 * Sin(selection.bearing * 0.01745329252)"
       mapinfo.Do "y1=selection.lat + 0.0015 * Cos(selection.bearing * 0.01745329252)"
       'CellLon = mapinfo.eval("selection.lon") + 0.0015 * Sin(mapinfo.eval("selection.bearing") * 0.01745329252)
       'CellLat = mapinfo.eval("selection.lat") + 0.0015 * Cos(mapinfo.eval("selection.bearing") * 0.01745329252)
     CchTch_Frm.Show 1
     bcch_no = Val(mapinfo.eval("selection.arfcn"))
     mapinfo.Do "close table " & MySelName
     aad = bcch_no + 1
     ddd = bcch_no - 1
     If CELL_CCH = 1 Then
        If SearchDistance = 0 Then
           mapinfo.Do "select * from cell where ABS(Arfcn - " & bcch_no & ")=1 into neighber_arfcn"
        Else
           mapinfo.Do "select * from cell where (ABS(Arfcn - " & bcch_no & ")=1) and ((x1-(lon+0.0015*sin(bearing*0.01745329252)))^2 +(y1-(lat+0.0015*sin(bearing*0.01745329252)))^2)<" & Format(SearchDistance * 0.0021 / 3) & " into neighber_arfcn"
        End If
        
     Else
        'mapinfo.do "select * from cell where ABS(non_bcch_1 - " & bcch_no & ")=1 or ABS(non_bcch_2 - " & bcch_no & ")=1 or ABS(non_bcch_3 - " & bcch_no & ")=1 or ABS(non_bcch_4 - " & bcch_no & ")=1 or ABS(non_bcch_5 - " & bcch_no & ")=1 or ABS(non_bcch_6 - " & bcch_no & ")=1 into neighber_arfcn"
        If SearchDistance = 0 Then
           mapinfo.Do "Select * from cell where Like(Non_bcch,""%" & Format(bcch_no + 1) & "%"","""") = 1 or Like(Non_bcch,""%" & Format(bcch_no - 1) & "%"","""") = 1 into neighber_arfcn"
        Else
           mapinfo.Do "Select * from cell where Like(Non_bcch,""%" & Format(bcch_no + 1) & "%"","""") = 1 or Like(Non_bcch,""%" & Format(bcch_no - 1) & "%"","""") = 1 and ((x1-(lon+0.0015*sin(bearing*0.01745329252)))^2 +(y1-(lat+0.0015*sin(bearing*0.01745329252)))^2)<" & Format(SearchDistance * 0.0021 / 3) & " into neighber_arfcn"
        End If
     End If
     If Val(mapinfo.eval("tableinfo(neighber_arfcn,8)")) <= 1 Then
        If CELL_CCH = 1 Then
           MsgBox "不存在BCCH同为 " & bcch_no - 1 & " 或 " & bcch_no + 1 & " 的小区", 64, "提示"
        Else
           MsgBox "不存在TCH同为 " & bcch_no - 1 & " 或 " & bcch_no + 1 & " 的小区", 64, "提示"
        End If
        mapinfo.Do "close table neighber_arfcn"
        Exit Sub
     End If
     
     Msg = "Add Map Auto Layer " + Chr(34) + "neighber_arfcn" + Chr(34)
     mapinfo.Do Msg
             
             mapinfo.Do "set map redraw off"
             mapinfo.Do "Set Map Layer 0 Editable On  "
             mapinfo.Do "set map redraw on"
             
            MyRow = Val(mapinfo.eval("tableinfo(neighber_arfcn,8)"))
            mapinfo.Do " fetch first from neighber_arfcn"
            'mapinfo.do "Set Style Pen MakePen(1,4,16719904)"
            'mapinfo.do "Set Style Pen MakePen(1,4,255)"
            'mapinfo.do "Set Style Pen MakePen(1,4,16719904)"
            mapinfo.Do "Set Style Pen MakePen(2,4," & Format(Xiaoyu_Color(Linyujin)) & ")"
            Linyujin = Linyujin + 1
            Linyujin = Linyujin Mod 5
            
            For i = 1 To MyRow
                mapinfo.Do "x2=neighber_arfcn.lon + 0.0015 * Sin(neighber_arfcn.bearing * 0.01745329252)"
                mapinfo.Do "y2=neighber_arfcn.lat + 0.0015 * Cos(neighber_arfcn.bearing * 0.01745329252)"
                'NcellLon = mapinfo.eval("neighber_arfcn.lon") + 0.0015 * Sin(mapinfo.eval("neighber_arfcn.bearing") * 0.01745329252)
                'NcellLat = mapinfo.eval("neighber_arfcn.lat") + 0.0015 * Cos(mapinfo.eval("neighber_arfcn.bearing") * 0.01745329252)
                'If (CellLat - NcellLat) * (CellLat - NcellLat) + (CellLon - NcellLon) * (CellLon - NcellLon) < 0.0021 Then    '5公里
                   mapinfo.Do "create Line(x1,y1)(x2,y2)"
                'End If
                mapinfo.Do "fetch  next from neighber_arfcn"
            Next
     
     If CELL_CCH = 1 Then
        mapinfo.Do "shade window   Frontwindow()  neighber_arfcn with ARFCN values  " + Chr(34) & aad & Chr(34) + " Symbol (58,16711935,12), " + Chr(34) & ddd & Chr(34) + " Symbol (58,65535,12)"
     Else
        'mapinfo.do "shade window Frontwindow() neighber_arfcn with non_bcch_1,non_bcch_2,non_bcch_3,non_bcch_4,non_bcch_5,non_bcch_6 values  " + Chr(34) & aad & Chr(34) + " Symbol (58,16711935,12), " + Chr(34) & ddd & Chr(34) + " Symbol (58,65535,12)"
        mapinfo.Do "shade window Frontwindow() neighber_arfcn with Like(Non_bcch,""%" & Format(bcch_no + 1) & "%"","""") OR Like(Non_bcch,""%" & Format(bcch_no - 1) & "%"","""") values 1 Symbol (58,65535,12)"
     End If

     If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
     End If
     If CELL_CCH = 1 Then
        If SearchDistance = 0 Then
           Msg = " Title " + Chr(34) + " 邻频观测 (CCH)" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle ""全网"" Font (""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off"
        Else
           Msg = " Title " + Chr(34) + " 邻频观测 (CCH)" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle ""只找出" & Format(SearchDistance) & "公里以内的小区"" Font (""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off"
        End If
     Else
        If SearchDistance = 0 Then
           Msg = " Title " + Chr(34) + " 邻频观测 (TCH)" + Chr(34) + " Font (""宋体"",0,9,0) Subtitle ""全网"" Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ," & Chr(34) & Format(bcch_no) & Chr(34) & " display on"
        Else
           Msg = " Title " + Chr(34) + " 邻频观测 (TCH)" + Chr(34) + " Font (""宋体"",0,9,0) Subtitle ""只找出" & Format(SearchDistance) & "公里以内的小区"" Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ," & Chr(34) & Format(bcch_no) & Chr(34) & " display on"
        End If
     End If
     mapinfo.Do "set legend window   Frontwindow()  Layer prev " & Msg
 End If
    StatusBar.Panels(2).Text = "按鼠标右键清除装饰层"
End Sub

Private Sub Same_Channel()
    Dim my_BcchNo As Integer, my_BcchNo1 As Integer, my_BcchNo2 As Integer, my_BcchNo3 As Integer
    Dim my_BcchNo4 As Integer, my_BcchNo5 As Integer, my_BcchNo6 As Integer
    Dim my_msg As String
        
    On Error Resume Next
    StatusBar.Panels(2).Text = "  邻频观察"
    SelTbl = mapinfo.eval("selectionInfo(1)")
    If SelTbl <> "cell" Then
       MsgBox "请选择一个Cell ！", 64, "提示"
    Else
       my_BcchNo = Val(mapinfo.eval("selection.arfcn"))
       my_BcchNo1 = Val(mapinfo.eval("selection.non_bcch_1"))
       my_BcchNo2 = Val(mapinfo.eval("selection.non_bcch_2"))
       my_BcchNo3 = Val(mapinfo.eval("selection.non_bcch_3"))
       my_BcchNo4 = Val(mapinfo.eval("selection.non_bcch_4"))
       my_BcchNo5 = Val(mapinfo.eval("selection.non_bcch_5"))
       my_BcchNo6 = Val(mapinfo.eval("selection.non_bcch_6"))
       my_msg = "select * from cell where Arfcn = " & my_BcchNo & " or non_bcch_1 = " & my_BcchNo & " or non_bcch_2 = " & my_BcchNo & " or non_bcch_3 = " & my_BcchNo & " or non_bcch_4 = " & my_BcchNo & " or non_bcch_5 = " & my_BcchNo & " or non_bcch_6 = " & my_BcchNo
       If my_BcchNo1 > 0 Then
          my_msg = my_msg + " or Arfcn = " & my_BcchNo1 & " or non_bcch_1 = " & my_BcchNo1 & " or non_bcch_2 = " & my_BcchNo1 & " or non_bcch_3 = " & my_BcchNo1 & " or non_bcch_4 = " & my_BcchNo1 & " or non_bcch_5 = " & my_BcchNo1 & " or non_bcch_6 = " & my_BcchNo1
       End If
       If my_BcchNo2 > 0 Then
          my_msg = my_msg + " or Arfcn = " & my_BcchNo2 & " or non_bcch_1 = " & my_BcchNo2 & " or non_bcch_2 = " & my_BcchNo2 & " or non_bcch_3 = " & my_BcchNo2 & " or non_bcch_4 = " & my_BcchNo2 & " or non_bcch_5 = " & my_BcchNo2 & " or non_bcch_6 = " & my_BcchNo2
       End If
       If my_BcchNo3 > 0 Then
          my_msg = my_msg + " or Arfcn = " & my_BcchNo3 & " or non_bcch_1 = " & my_BcchNo3 & " or non_bcch_2 = " & my_BcchNo3 & " or non_bcch_3 = " & my_BcchNo3 & " or non_bcch_4 = " & my_BcchNo3 & " or non_bcch_5 = " & my_BcchNo3 & " or non_bcch_6 = " & my_BcchNo3
       End If
       'If my_BcchNo4 > 0 Then
       '   my_msg = my_msg + " or ABS(Arfcn - " & my_BcchNo4 & ")=1 or ABS(non_bcch_1 - " & my_BcchNo4 & ")=1 or ABS(non_bcch_2 - " & my_BcchNo4 & ")=1 or ABS(non_bcch_3 - " & my_BcchNo4 & ")=1 or ABS(non_bcch_4 - " & my_BcchNo4 & ")=1 or ABS(non_bcch_5 - " & my_BcchNo4 & ")=1 or ABS(non_bcch_6 - " & my_BcchNo4 & ")=1"
       'End If
       'If my_BcchNo5 > 0 Then
       '   my_msg = my_msg + " or ABS(Arfcn - " & my_BcchNo5 & ")=1 or ABS(non_bcch_1 - " & my_BcchNo5 & ")=1 or ABS(non_bcch_2 - " & my_BcchNo5 & ")=1 or ABS(non_bcch_3 - " & my_BcchNo5 & ")=1 or ABS(non_bcch_4 - " & my_BcchNo5 & ")=1 or ABS(non_bcch_5 - " & my_BcchNo5 & ")=1 or ABS(non_bcch_6 - " & my_BcchNo5 & ")=1"
       'End If
       'If my_BcchNo6 > 0 Then
       '   my_msg = my_msg + " or ABS(Arfcn - " & my_BcchNo6 & ")=1 or ABS(non_bcch_1 - " & my_BcchNo6 & ")=1 or ABS(non_bcch_2 - " & my_BcchNo6 & ")=1 or ABS(non_bcch_3 - " & my_BcchNo6 & ")=1 or ABS(non_bcch_4 - " & my_BcchNo6 & ")=1 or ABS(non_bcch_5 - " & my_BcchNo6 & ")=1 or ABS(non_bcch_6 - " & my_BcchNo6 & ")=1"
       'End If
       my_msg = my_msg + " into Same_Channel"
       mapinfo.Do my_msg
       mapinfo.Do "Set Next Document Parent " & MDIMain.hWnd & " Style 2"
       mapinfo.Do "browse * from Same_Channel"
       mapinfo.Do "set window Frontwindow() Position(0,1) Width 8 Height 1 "
 End If
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub Neighbor_Channel()
    Dim my_BcchNo As Integer, my_BcchNo1 As Integer, my_BcchNo2 As Integer, my_BcchNo3 As Integer
    Dim my_BcchNo4 As Integer, my_BcchNo5 As Integer, my_BcchNo6 As Integer
    Dim my_msg As String
        
    On Error Resume Next
    StatusBar.Panels(2).Text = "  邻频观察"
    SelTbl = mapinfo.eval("selectionInfo(1)")
    If SelTbl <> "cell" Then
       MsgBox "请选择一个Cell ！", 64, "提示"
    Else
       my_BcchNo = Val(mapinfo.eval("selection.arfcn"))
       my_BcchNo1 = Val(mapinfo.eval("selection.non_bcch_1"))
       my_BcchNo2 = Val(mapinfo.eval("selection.non_bcch_2"))
       my_BcchNo3 = Val(mapinfo.eval("selection.non_bcch_3"))
       my_BcchNo4 = Val(mapinfo.eval("selection.non_bcch_4"))
       my_BcchNo5 = Val(mapinfo.eval("selection.non_bcch_5"))
       my_BcchNo6 = Val(mapinfo.eval("selection.non_bcch_6"))
       my_msg = "select * from cell where ABS(Arfcn - " & my_BcchNo & ")=1 or ABS(non_bcch_1 - " & my_BcchNo & ")=1 or ABS(non_bcch_2 - " & my_BcchNo & ")=1 or ABS(non_bcch_3 - " & my_BcchNo & ")=1 or ABS(non_bcch_4 - " & my_BcchNo & ")=1 or ABS(non_bcch_5 - " & my_BcchNo & ")=1 or ABS(non_bcch_6 - " & my_BcchNo & ")=1"
       If my_BcchNo1 > 0 Then
          my_msg = my_msg + " or ABS(Arfcn - " & my_BcchNo1 & ")=1 or ABS(non_bcch_1 - " & my_BcchNo1 & ")=1 or ABS(non_bcch_2 - " & my_BcchNo1 & ")=1 or ABS(non_bcch_3 - " & my_BcchNo1 & ")=1 or ABS(non_bcch_4 - " & my_BcchNo1 & ")=1 or ABS(non_bcch_5 - " & my_BcchNo1 & ")=1 or ABS(non_bcch_6 - " & my_BcchNo1 & ")=1"
       End If
       If my_BcchNo2 > 0 Then
          my_msg = my_msg + " or ABS(Arfcn - " & my_BcchNo2 & ")=1 or ABS(non_bcch_1 - " & my_BcchNo2 & ")=1 or ABS(non_bcch_2 - " & my_BcchNo2 & ")=1 or ABS(non_bcch_3 - " & my_BcchNo2 & ")=1 or ABS(non_bcch_4 - " & my_BcchNo2 & ")=1 or ABS(non_bcch_5 - " & my_BcchNo2 & ")=1 or ABS(non_bcch_6 - " & my_BcchNo2 & ")=1"
       End If
       If my_BcchNo3 > 0 Then
          my_msg = my_msg + " or ABS(Arfcn - " & my_BcchNo3 & ")=1 or ABS(non_bcch_1 - " & my_BcchNo3 & ")=1 or ABS(non_bcch_2 - " & my_BcchNo3 & ")=1 or ABS(non_bcch_3 - " & my_BcchNo3 & ")=1 or ABS(non_bcch_4 - " & my_BcchNo3 & ")=1 or ABS(non_bcch_5 - " & my_BcchNo3 & ")=1 or ABS(non_bcch_6 - " & my_BcchNo3 & ")=1"
       End If
       If my_BcchNo4 > 0 Then
          my_msg = my_msg + " or ABS(Arfcn - " & my_BcchNo4 & ")=1 or ABS(non_bcch_1 - " & my_BcchNo4 & ")=1 or ABS(non_bcch_2 - " & my_BcchNo4 & ")=1 or ABS(non_bcch_3 - " & my_BcchNo4 & ")=1 or ABS(non_bcch_4 - " & my_BcchNo4 & ")=1 or ABS(non_bcch_5 - " & my_BcchNo4 & ")=1 or ABS(non_bcch_6 - " & my_BcchNo4 & ")=1"
       End If
       If my_BcchNo5 > 0 Then
          my_msg = my_msg + " or ABS(Arfcn - " & my_BcchNo5 & ")=1 or ABS(non_bcch_1 - " & my_BcchNo5 & ")=1 or ABS(non_bcch_2 - " & my_BcchNo5 & ")=1 or ABS(non_bcch_3 - " & my_BcchNo5 & ")=1 or ABS(non_bcch_4 - " & my_BcchNo5 & ")=1 or ABS(non_bcch_5 - " & my_BcchNo5 & ")=1 or ABS(non_bcch_6 - " & my_BcchNo5 & ")=1"
       End If
       If my_BcchNo6 > 0 Then
          my_msg = my_msg + " or ABS(Arfcn - " & my_BcchNo6 & ")=1 or ABS(non_bcch_1 - " & my_BcchNo6 & ")=1 or ABS(non_bcch_2 - " & my_BcchNo6 & ")=1 or ABS(non_bcch_3 - " & my_BcchNo6 & ")=1 or ABS(non_bcch_4 - " & my_BcchNo6 & ")=1 or ABS(non_bcch_5 - " & my_BcchNo6 & ")=1 or ABS(non_bcch_6 - " & my_BcchNo6 & ")=1"
       End If
       my_msg = my_msg + " into Neighbor_Channel"
       mapinfo.Do my_msg
       mapinfo.Do "Set Next Document Parent " & MDIMain.hWnd & " Style 2"
       mapinfo.Do "browse * from Neighbor_Channel"
       mapinfo.Do "set window Frontwindow() Position(0,1) Width 8 Height 1 "
 End If
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub Help_content_Click()
    
    On Error Resume Next
    Gsm_FileName = "winhelp.exe   " + Gsm_Path + "\gsm.hlp   "
    ReturnValue = Shell(Gsm_FileName, 2)
End Sub

Private Sub data_Quit_Click()
    Data_Tool.Visible = False
    On Error Resume Next
    mapinfo.Do "Set Map Layer  street Editable off selectable  off"
    StatusBar.Panels(2).Text = "  "
End Sub


Private Sub ICONS_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 排列图标"
    MDIMain.Arrange ARRANGE_ICONS
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub CELL_GRAPH_Click()
  On Error Resume Next
  StatusBar.Panels(2).Text = "  基站图象观察"
  SelTbl = mapinfo.eval("selectionInfo(1)")
  If SelTbl <> "cell" Then
     MsgBox "请选择一个Cell ！", 64, "提示"
  Else
     Bmp_Name = mapinfo.eval("selection.photo")
     CellGraph.Show 1
 End If
    StatusBar.Panels(2).Text = " "
End Sub


Private Sub SUB_320_Click()
  On Error Resume Next
  StatusBar.Panels(2).Text = "  NCELL排序转换"
  Menu_Flag = 5004
  SelTable.Show 1
  StatusBar.Panels(2).Text = ""
End Sub


Private Sub ScanPilot_Click()
    Dim buff As String, mypath As String
    Dim finds As Integer, i As Integer
        
    On Error Resume Next
    Menu_Flag = 4449
    StatusBar.Panels(2).Text = " 数据转换"
    For i = 1 To 50
        convert_filename(i) = ""
    Next
open_again:
    FileDialog.DialogTitle = "ANT Pilot 扫频测试数据转换"
    FileDialog.Filter = "*.log Files|*.LOG|All Files|*.*"
    FileDialog.DefaultExt = "*.LOG"
    FileDialog.Flags = &H80000 Or &H200
    Gsm_FileName = Gsm_Path + "\scan"
    FileDialog.InitDir = Gsm_FileName
    FileDialog.ShowOpen
    buff = Trim(FileDialog.filename)
    If buff = "" Then
       StatusBar.Panels(2).Text = " "
       Exit Sub
    End If
    tran_fn = 0
    finds = InStr(buff, Chr(0))
    If finds > 0 Then
       mypath = Left(buff, finds - 1) + "\"
       buff = Trim(Right(buff, Len(buff) - finds))
       finds = InStr(buff, Chr(0))
       i = 1
       Do While finds > 0
          convert_filename(i) = mypath + Left(buff, finds - 1)
          tran_f(i) = convert_filename(i)
          buff = Trim(Right(buff, Len(buff) - finds))
          finds = InStr(buff, Chr(0))
          i = i + 1
       Loop
       convert_filename(i) = mypath + buff
       tran_f(i) = convert_filename(i)
       tran_fn = i
    Else
       convert_filename(1) = buff
       tran_f(1) = buff
       tran_fn = 1
    End If
    FileDialog.filename = ""
    If Dir(convert_filename(1)) = "" Then
       GoTo err_exit
    End If
    sinput = tran_f(1)
    FileDialog.filename = ""
    Menu_Flag = 4449
    DocManager.Show 1
    StatusBar.Panels(2).Text = " "
    Exit Sub
    
err_exit:
       i = MsgBox("无法打开文件 " + convert_filename(1), 48, "打开文件")
       GoTo open_again

End Sub

Private Sub SecMobileArfcn_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 第二手机频率分析"
    Menu_Flag = 88388
    ARFCNSelect.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub SecMobileRxlevf_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 第二手机场强分析"
    Menu_Flag = 88311
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub SecMobileRxlevs_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "第二手机场强分析"
    Menu_Flag = 88314
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub SecMobileRxqualf_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 第二手机品质分析"
    Menu_Flag = 88312
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "

End Sub

Private Sub SecMobileRxquals_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "第二手机品质分析"
    Menu_Flag = 88315
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub SecMobileTa_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 第二手机TA分析"
    Menu_Flag = 88317
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub Sub_317_old_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " Timing Advance分析"
    'Menu_Flag = 317
    Menu_Flag = 991121
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub SUB_330_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "  主要信令描述分析"
    Menu_Flag = 330
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_431_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 岛效应点观测"
    SelTbl = mapinfo.eval("selectionInfo(1)")
    If SelTbl <> "cell" Then
       MsgBox "请选择一个Cell ！", 64, "提示"
    Else
       rmsg1 = mapinfo.eval("selection.ci")
       rmsg2 = mapinfo.eval("selection.arfcn")
       Menu_Flag = 431
       SelTable.Show 1
   End If
   StatusBar.Panels(2).Text = " "
End Sub

Private Sub SUB_432_Click()
    Dim i As Integer
    On Error Resume Next
    StatusBar.Panels(2).Text = " 岛效应分析"
    Menu_Flag = 432

    i = Val(mapinfo.eval("selectionInfo(3)"))  ' SEL_INFO_NROWS
    If i <> 0 Then
       Iland_Base.Show 1
    End If
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_468_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 按LAC查找小区"
    Find_Lac.Show 1
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub SUB_469_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "  按CI查找小区"
    Menu_Flag = 469
    CI_Cell.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_BASE_ADD_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 站址维护"
    Gsm_FileName = Gsm_Path + "\map\base_add.tab"
    If Dir(Gsm_FileName) = "" Then
       MsgBox "请先生成站址库再使用该功能！", 64, "提示"
       StatusBar.Panels(2).Text = ""
       Exit Sub
    End If
    mapinfo.Do "open table " + Chr(34) + Gsm_FileName + Chr(34)
    mapinfo.Do "Set Next Document Parent " & MDIMain.hWnd & " Style 2"
    mapinfo.Do "browse * from  base_add"
    mapinfo.Do "set window Frontwindow() Position(0,2) Width 6 Height 3 "

    SUB_23.Enabled = True
    SUB_24.Enabled = True
    SUB_25.Enabled = True
    StatusBar.Panels(2).Text = ""
End Sub

Private Sub SUB_4600_Click()
    StatusBar.Panels(2).Text = " 站址查找"
    Menu_Flag = 4600
    Center.Show 1
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub TObtel_Click()
    Dim buff As String, mypath As String
    Dim finds As Integer, i As Integer
        
    On Error Resume Next
    'Menu_Flag = 1244
    Menu_Flag = 121
    StatusBar.Panels(2).Text = " 数据转换"
    For i = 1 To 50
        convert_filename(i) = ""
    Next
open_again:
    FileDialog.DialogTitle = "Grayson Surveyor 测量数据文件选择"
    FileDialog.Filter = "*.00* Files|*.00*|All Files|*.*"
    FileDialog.DefaultExt = "*.00*"
    FileDialog.Flags = &H80000 Or &H200
    Gsm_FileName = Gsm_Path + "\normal"
    FileDialog.InitDir = Gsm_FileName
    FileDialog.ShowOpen
    buff = Trim(FileDialog.filename)
    If buff = "" Then
       StatusBar.Panels(2).Text = " "
       Exit Sub
    End If
    tran_fn = 0
    finds = InStr(buff, Chr(0))
    If finds > 0 Then
       mypath = Left(buff, finds - 1) + "\"
       buff = Trim(Right(buff, Len(buff) - finds))
       finds = InStr(buff, Chr(0))
       i = 1
       Do While finds > 0
          convert_filename(i) = mypath + Left(buff, finds - 1)
          tran_f(i) = convert_filename(i)
          buff = Trim(Right(buff, Len(buff) - finds))
          finds = InStr(buff, Chr(0))
          i = i + 1
       Loop
       convert_filename(i) = mypath + buff
       tran_f(i) = convert_filename(i)
       tran_fn = i
    Else
       convert_filename(1) = buff
       tran_f(1) = buff
       tran_fn = 1
    End If
    FileDialog.filename = ""
    If Dir(convert_filename(1)) = "" Then
       GoTo err_exit
    End If
    'sinput = tran_f(1)
    'DocManager.Show 1
    'StatusBar.Panels(2).Text = " "
    sinput = tran_f(1)
    cvChoice.Show 1
    StatusBar.Panels(2).Text = " "
    Exit Sub
    
err_exit:
       i = MsgBox("无法打开文件 " + convert_filename(1), 48, "打开文件")
       GoTo open_again
End Sub


Private Sub Sub_ORBITEL_Click()
    On Error Resume Next
    MsgBox Chr(10) + "请与珠海万禾技术集成公司联系！" + Chr(10) + Chr(10) + "Tel:0756-3367710", 64, "提示"
End Sub

Private Sub Sub_RS9951_Click()
    On Error Resume Next
    MsgBox Chr(10) + "请与珠海万禾技术集成公司联系！" + Chr(10) + Chr(10) + "Tel:0756-3367710", 64, "提示"
End Sub

Private Sub SUB_TEST_REPORT_Click()
    On Error Resume Next
    Dim mypath As String, buff As String
    Dim MyDbName As String, MyTableName As String
    Dim lpParameters As String
    Dim code As Integer
    Dim i As Integer, finds As Integer, dd As Integer
    
    On Error Resume Next
    'IsQuickConvert = False
    For i = 1 To 50
        convert_filename(i) = ""
    Next
open_again:
    MDIMain.FileDialog.DialogTitle = "统计文件选择"
    MDIMain.FileDialog.InitDir = Gsm_Path + "\normal"
    MDIMain.FileDialog.Filter = "Ant Files|*.ant|DBF Files|*.dbf|All Files|*.*"
    MDIMain.FileDialog.DefaultExt = "*.ant"
    MDIMain.FileDialog.Flags = &H80000 Or &H200
    MDIMain.FileDialog.ShowOpen
    buff = Trim(MDIMain.FileDialog.filename)
    If buff = "" Then
       Exit Sub
    End If
    finds = InStr(buff, Chr(0))
    If finds > 0 Then
       mypath = Left(buff, finds - 1) + "\"
       buff = Trim(Right(buff, Len(buff) - finds))
       finds = InStr(buff, Chr(0))
       i = 1
       Do While finds > 0
          convert_filename(i) = mypath + Left(buff, finds - 1)
          buff = Trim(Right(buff, Len(buff) - finds))
          finds = InStr(buff, Chr(0))
          i = i + 1
       Loop
       convert_filename(i) = mypath + buff
    '   If UCase(Mid(convert_filename(i), Len(convert_filename(i)) - 4, 1)) = "F" Then
    '      IsQuickConvert = True
    '   End If
    Else
       convert_filename(1) = buff
    '   If UCase(Mid(convert_filename(1), Len(convert_filename(1)) - 4, 1)) = "F" Then
    '      IsQuickConvert = True
    '   End If
    End If
    MDIMain.FileDialog.filename = ""
    If Dir(convert_filename(1)) = "" Then
       GoTo err_exit
    End If
    For i = 1 To 50
        If convert_filename(i) = "" Then
           stre_num = i - 1
           Exit For
        End If
        Err = 0
        If Err Then
           dd = MsgBox("无法打开文件   " + convert_filename(i), 48, "打开文件")
           Exit Sub
        End If
        stre_tab(i) = convert_filename(i)
        finds = InStr(stre_tab(i), ".")
        If finds > 0 Then
           stre_tab(i) = Left(stre_tab(i), finds - 1)
        End If
        finds = InStr(stre_tab(i), "\")
        Do While finds > 0
           stre_tab(i) = Right(stre_tab(i), Len(stre_tab(i)) - finds)
           finds = InStr(stre_tab(i), "\")
        Loop
        If Asc(Left(stre_tab(i), 1)) > 47 And Asc(Left(stre_tab(i), 1)) < 58 Then
           stre_tab(i) = "_" + stre_tab(i)
        End If
    Next
    TEST_REPORT
    'Stre_Sel.Show 1
    'If CancelFlag Then
    '    Exit Sub
    'End If
    'MyTableName = stcname
    'MyDbName = ""
    'Do While InStr(MyTableName, "\") > 0
    '   MyDbName = MyDbName & Left(MyTableName, InStr(MyTableName, "\"))
    '   MyTableName = Right(MyTableName, Len(MyTableName) - InStr(MyTableName, "\"))
    'Loop
    'MyDbName = Left(MyDbName, Len(MyDbName) - 1)

    
    'code = ShellExecute(mdimain.hwnd, "open", MyTableName, lpParameters, MyDbName, 4)
    
    'RetVal = Shell(stcname, 1)    ' 完成Calculator。
    'AppActivate RetVal
    Exit Sub
err_exit:
    i = MsgBox("无法打开文件 " + convert_filename(1), 48, "打开文件")
    GoTo open_again
End Sub

Private Sub Tch_Find_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "   TCH 数据查询"
    Gsm_FileName = Gsm_Path + "\sts\tch_sts.tab"
    If UCase(Dir(Gsm_FileName, 0)) <> "TCH_STS.TAB" Then
       MsgBox " TCH_STS.tab 不存在！", 64, "提示"
       StatusBar.Panels(2).Text = "  "
       Exit Sub
    End If
    Tch_data_find.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub Tch_Map_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  打开 TCH 地图"
    Gsm_FileName = Gsm_Path + "\sts\tch_sts.tab"
    If UCase(Dir(Gsm_FileName, 0)) <> "TCH_STS.TAB" Then
       MsgBox " TCH_STS.tab 不存在！", 64, "提示"
       StatusBar.Panels(2).Text = "  "
       Exit Sub
    End If
    mapinfo.Do "open table " + Chr(34) + Gsm_FileName + Chr(34)
    If mapinfo.eval("tableinfo(tch_sts,4)") = 18 Then
       mapinfo.Do "close table tch_sts "
       Tch_emap_choice.Show 1
    Else
       mapinfo.Do "close table tch_sts"
       Tch_mmap_choice.Show 1
    End If
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub Tems_Data_Click()
    On Error Resume Next
    Data_Report = False
    Menu_Flag = 121
    T121_Click
End Sub

Private Sub Tems_Data_Report_Click()
    On Error Resume Next
    Data_Report = True
    Menu_Flag = 121
    T121_Click
End Sub

Private Sub Tems98_Data_Click()
    On Error Resume Next
    Data_Report = False
    Menu_Flag = 128
    T121_Click

End Sub

Private Sub Tems98_Report_Click()
    On Error Resume Next
    Data_Report = True
    Menu_Flag = 128
    T121_Click

End Sub

Private Sub Test_Define_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "   对抗测试定义"
'    Load Cope_Define
'    Cope_Define.Move 3500, 2200, 3620, 2700
    Cope_Define.Show 1
    StatusBar.Panels(2).Text = "   "
End Sub

Private Sub TOG_LEG_Click()
  On Error Resume Next
    If Legend_Tog = 0 Then
       StatusBar.Panels(2).Text = "  多段图例"
       Legend_Tog = 1
    Else
      If Legend_Tog = 1 Then
       Legend_Tog = 0
       StatusBar.Panels(2).Text = "  3段图例"
      End If
    End If
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    Dim AllRows As Variant
        
    On Error Resume Next
    Select Case Button.Index
        Case 1
             SSCommand1_Click
        Case 2
             SSCommand2_Click
        Case 3
             SSCommand3_Click
        Case 4
             SSCommand4_Click
        Case 5
             My_Center_Click
        Case 7
             SSCommand5_Click
        Case 9
             SSCommand6_Click
        Case 8
             SSCommand7_Click
        Case 13
             SSCommand8_Click
        Case 14
             SSCommand9_Click
        Case 10
             SSClengend_Click
        Case 29
             VMap_Click
        Case 11
             Cam_Click
        Case 16
             StatusBar.Panels(2).Text = " 当前工具：画圆"
             mapinfo.Do "Set Style brush Makebrush(1,0,0)"
             mapinfo.runmenucommand 1715
        Case 17
             Pline_Click      '''曲线
        Case 18
             Region_Click     '''转换为区域
        Case 33
             CELL_GRAPH_Click
        Case 30
             M_legend_Click
        Case 31
             TOG_LEG_Click
        Case 32
             Label_Click
        Case 19
             MAP_DIS_Click
        Case 21
             SSCommand10_Click
        Case 22
             SSCommand11_Click
        Case 23
             Ncell_Map_Click
        Case 34
             'Help_Click
             SNDisplay_Click
        Case 15
             Radius_Select
        Case 24, 25, 26, 27
             SelTbl = mapinfo.eval("selectionInfo(1)")
             If SelTbl = "" Then
                MsgBox "请选择测量数据中的一个点！", 64, "提示"
                StatusBar.Panels(2).Text = " "
                Exit Sub
             End If
             If UCase(mapinfo.eval("Columninfo( " & SelTbl & ",COL1, 1)")) <> "TIME" Then
                MsgBox "请选择测量数据中的一个点！", 64, "提示"
                StatusBar.Panels(2).Text = " "
                Exit Sub
             End If
             Select Case Button.Index
                 Case 24
                    StatusBar.Panels(2).Text = " NCELL 数据显示"
                    If NcellWinFlag Then
                       Ncell_graph.Form_Load
                    Else
                       Ncell_graph.Show
                    End If
                    SetWindowPos Ncell_graph.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                    StatusBar.Panels(2).Text = "  "
                 Case 25
                    StatusBar.Panels(2).Text = " Dedicated Channel参数显示"
                    Dedicated.Show 1
                    StatusBar.Panels(2).Text = "  "
                 Case 26
                    StatusBar.Panels(2).Text = " 双网参数比较"
                    TowMobile.Show 1
                    StatusBar.Panels(2).Text = "  "
                 Case 27
                    AllRows = Val(mapinfo.eval("tableinfo(" & SelTbl & ",4)"))
                    If AllRows <> 88 And AllRows <> 150 Then
                       MsgBox "该文件不能进行邻频载干比显示。", 64, "提示"
                       Exit Sub
                    End If
                    StatusBar.Panels(2).Text = " 邻频载干比显示"
                    FrmCICA.Show
                    SetWindowPos FrmCICA.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                    StatusBar.Panels(2).Text = "  "
             End Select
    End Select
End Sub

Private Sub TRAN_C_I_Click()
  On Error Resume Next
  StatusBar.Panels(2).Text = " C/I计算"
  Gsm_FileName = Gsm_Path + "\scan"
  ChDir Gsm_FileName
  CMDialog1.filename = "*.tab"
  CMDialog1.InitDir = Gsm_FileName
  CMDialog1.DefaultExt = "*.tab"
  CMDialog1.Filter = "Text Files (*.tab*)|*.tab|*.*"
  CMDialog1.FilterIndex = 1
'  CMDialog1.Move 2000, 2000

  On Error GoTo Go_OUT
  CMDialog1.Action = 1

  sinput = CMDialog1.filename

  pot = Len(sinput) - 4
  stab = Left$(sinput, pot)
  stab = stab + ".tab"

  ttab = stab
  pot = 1
  While (pot <> 0)
        ttab = Mid$(ttab, pot + 1, Len(ttab) - pot + 1)
        pot = InStr(1, ttab, "\")
  Wend

  If Mid$(ttab, 1, 1) < "9" Then
       my_tab = "_" + Mid$(ttab, 1, Len(ttab) - 4)
  Else
       my_tab = Mid$(ttab, 1, Len(ttab) - 4)
  End If

  tblname = my_tab
  If sinput = "*.scn" Then
       Exit Sub
  Else
       mapinfo.Do "open  table " + Chr(34) + sinput + Chr(34)
       C_I.Show 1
       mapinfo.Do "close  table " & my_tab
  End If
Go_OUT:
  StatusBar.Panels(2).Text = "  "
End Sub

Private Sub MDIForm_Load()
    Dim dog As Integer, i As Integer
    
    Dim ApiPack As APIPACKET                      'win95
    Dim portnum%                                  'win95
    Dim status%                                   'win95
    Dim majVer%, minVer%, rev%, drvrType%         'win95
    Dim adr%, datum%                              'win95
    
    On Error GoTo cant_createobject

    Gsm_Path = CurDir$
    Gsm_Path = App.path
    'Gsm_Path = "G:\Ant2000"

    portnum% = 4 ' CPlus-B, port 1                     'win95
    status% = RNBOcplusFormatPacket(ApiPack, 1028)     'win95
    status% = RNBOcplusInitialize(ApiPack, portnum%)   'win95
    If status <> 0 Then GoTo VERYFY_OUT                'win95
    status% = RNBOcplusGetVersion(ApiPack, majVer%, minVer%, rev%, drvr6Type%)     'win95
    status% = RNBOcplusGetFullStatus(ApiPack)          'win95
    adr = 62                                           'win95
    status% = RNBOcplusRead(ApiPack, adr%, datum%)     'win95
    datum = (datum / 89) * 4 + 23                      'win95
    If datum <> 427 Then GoTo VERYFY_OUT               'win95
    adr = 60                                           'win95
    status% = RNBOcplusRead(ApiPack, adr%, datum%)     'win95
    datum = (datum / 89) * 4 + 23                      'win95
    If datum <> 619 Then GoTo VERYFY_OUT               'win95
    
    On Error Resume Next
    Randomize 2
    For i = 0 To 375
        MyRndColor(i) = Int(Rnd * 16777000 + 1)
    Next
    For i = 0 To 124
        MyCellRndColor(i) = Int(Rnd * 16777000 + 1)
    Next
    For i = 0 To 100
        MyLacColor(i) = Int(Rnd * 16777000 + 1)
    Next
    
    MyBcchColor(0) = 16711680
    MyBcchColor(1) = 65280
    MyBcchColor(2) = 255
    MyBcchColor(3) = 16711935
    MyBcchColor(4) = 16776960
    MyBcchColor(5) = 65535
    MyBcchColor(6) = 16756952
    MyBcchColor(7) = 15257855
    MyBcchColor(8) = 13689087
    MyBcchColor(9) = 13697023
    MyBcchColor(10) = 13893520
    MyBcchColor(11) = 16777168
    MyBcchColor(12) = 16756912
    MyBcchColor(13) = 12615935
    MyBcchColor(14) = 10535167
    MyBcchColor(15) = 16765088
    MyBcchColor(16) = 15745088
    
    Xiaoyu_Color(0) = 16719904
    Xiaoyu_Color(1) = 255
    Xiaoyu_Color(2) = 65280
    Xiaoyu_Color(3) = 65535
    Xiaoyu_Color(4) = 16711935
    Linyujin = 0
    
    Set mapinfo = CreateObject("mapinfo.Application")
    Gsm_FileName = Gsm_Path + "\gsm.hlp"
    App.HelpFile = Gsm_FileName
    'indicate that we don't have a legend window
    legendid = 0
    NcellWinFlag = False

    'Initialize the application
    mapinfo.Do "Set Application Window " & MDIMain.hWnd

    'Disable the help subsystem - we don't support help in this application
'    mapinfo.do "Set Window Help Off"

    Set myCallback = New Micallback  'micallback *.cls  'win95
    mapinfo.setcallback myCallback

    'Disable the <Add...> control in the Layer Control dialog box
'    mapinfo.do "Alter mapinfoDialog 1800 Control 12 Disable"

    'Make the Info tool parented to our form for when
    ' the user uses that tool
    mapinfo.Do "Set Window Info Parent " & MDIMain.hWnd
'    mapinfo.do "Set Window Info ReadOnly"

    mapinfo.Do "Set Window ruler Parent " & MDIMain.hWnd
    
    'Reprogram the map's shortcut menu (right-click on map)
    ' note well that we reprogram the Layer Control & Create Thematic
    ' to call us back via DDE so that we can do maintenence after these run
    ' mapinfo.do "Create Menu ""MapperShortcut"" ID 17 as ""&Layer Control..."" ID 801 calling DDE ""FindZip"",""MainForm"",""Create &Thematic Map..."" ID 307 calling DDE ""FindZip"",""MainForm"",""View &Entire Layer..."" calling 807"
     
     Msg = " Create Menu ""MapperShortcut"" ID 17 as""图层控制[&L]...""  calling 801 ,""(-"" ,""清除装饰图层[&Y]""  calling 810  ,""改变视图[&V]..."" calling 805 ,""前一视图[&P]"" calling 806 ,""查看整个图层[&E]..."" calling 807 ,""(-"" ,""编辑对象"" as ""对象[&O]"" ,""(-"" ,""获取信息[&I]...\tF7/W%118/Mi/XF7"" calling 207"

     On Error Resume Next
     mapinfo.Do Msg
     My_Ver = Get_date

    'Create our custom "Query" tool
    'mapinfo.do "Alter ButtonPad ID 3 Add ToolButton Calling DDE ""FindZip"",""MainForm"" Cursor 128 DrawMode 34 ID 101"

    'Select the grabber tool as the active tool,
 '   mapinfo.runmenucommand 1702

    Face_show = 0
    Legend_Tog = 0

    west = 109
    south = 20
    xx = 117 - 109
    yy = 25.3 - 20
    thereIsAMap = False
    Map_No = 0
    Gsm_FileName = Gsm_Path + "\map"
    If Dir(Gsm_FileName, 16) <> "" Then
       Gsm_File2 = Gsm_Path + "\map\gsm.tag"
       If Dir(Gsm_File2, 0) <> "" Then
          i = restore_street
       End If
    End If
    
    BackColor = RGB(64, 128, 128)
    mapinfo.Do "dim jj as integer"
    mapinfo.Do "dim x1 as float"
    mapinfo.Do "dim y1 as float"
    mapinfo.Do "dim x2 as float"
    mapinfo.Do "dim y2 as float"
    mapinfo.Do "dim x3 as float"
    mapinfo.Do "dim y3 as float"
    mapinfo.Do "dim x0 as float"
    mapinfo.Do "dim y0 as float"
    mapinfo.Do "Dim region As Object"
    mapinfo.Do "dim sts_mypoint as object"
    
    mapinfo.Do "Set Style pen Makepen(1,46,16711680)"
    
    On Error Resume Next
    ChDir Gsm_Path
    If Dir(Gsm_Path & "\user", 16) = "" Then
        MkDir Gsm_Path & "\user"
    End If

    Dim MyRecord As Record
    Gsm_FileName = Gsm_Path + "\gsm.dat"
    Open Gsm_FileName For Binary As #1
    Get #1, 1, MyRecord
    Close #1
    USERNAME = Trim(MyRecord.Name)
    If InStr(USERNAME, Chr(0)) > 0 Then
       USERNAME = Trim(Left(USERNAME, InStr(USERNAME, Chr(0)) - 1))
    End If

    On Error Resume Next
    Gsm_FileName = Gsm_Path + "\mysymb"
    mapinfo.Do "reload custom symbols from " + Chr(34) + Gsm_FileName + Chr(34)
    Exit Sub

cant_createobject:
    MsgBox "不能与地图系统联结，请重新运行！", 64, "提示"
    End
VERYFY_OUT:
    MsgBox "加密锁错误, 请与珠海万禾公司联系！", 64, "提示"
    End
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
'   AutherName.Caption = ""
'   MDIMain.SetFocus
 
 If sys = 0 Then
    MnuCover.Enabled = True
    MnuDDDD.Enabled = True
    MnuNeighbor.Enabled = True
    MnuHO.Enabled = True
    MnuMessage.Enabled = True
    MnuDoubleNet.Enabled = True
    MnuDataGraph.Enabled = True
    MnuGsm_Dcs.Enabled = True
    
    SUB_123.Enabled = 0
    ScanPilot.Enabled = 0
'    SCAN_2.Enabled = 0
'    SCAN_3.Enabled = 0
    My_ScanPlay.Enabled = False
    Arfcn_Changing.Enabled = False
    SCAN_4.Enabled = 0
    SCAN_5.Enabled = 0
    SCAN_6.Enabled = 0
    SCAN_7.Enabled = 0
    SCAN_8.Enabled = 0
    My_Over.Enabled = 0
    Mnudistributing.Enabled = 0
    TRAN_C_I.Enabled = 0

    SUB_121.Enabled = 1
    Tems98_Convert.Enabled = True
    ANTSurveyor.Enabled = 1
    SUB_Obtel.Enabled = 1
    Sub_31.Enabled = 1
    ViewNcell.Enabled = 1
    MnuC_A.Enabled = 1
    'MnuSql.Enabled = 1
    MnuLabel.Enabled = True
    MnuLabelMark.Enabled = 1
    RadioLink.Enabled = 1
    Hopping.Enabled = 1
    SUB_431.Enabled = 1
'    NetworkBlind.Enabled = 1
'    NetworkDisturb.Enabled = 1
'    View_Cope.Enabled = 1
    'SUB_32.Enabled = 1
    SUB_33.Enabled = 1
    SUB_41.Enabled = 1
'    SUB_42.Enabled = 1
    SUB_441.Enabled = 1
    SUB_442.Enabled = 1
   ' SUB_443.Enabled = 1
    'SUB_43.Enabled = 1
'    SUB_45.Enabled = 1
    Mnu_Replay.Enabled = 1
    StatusBar.Panels(4).Text = "通话分析"
 Else
    MnuCover.Enabled = False
    MnuDDDD.Enabled = False
    MnuNeighbor.Enabled = False
    MnuHO.Enabled = False
    MnuMessage.Enabled = False
    MnuDoubleNet.Enabled = False
    MnuDataGraph.Enabled = False
    MnuGsm_Dcs.Enabled = False
    
    SUB_121.Enabled = 0
    Tems98_Convert.Enabled = False
    ANTSurveyor.Enabled = 0
    SUB_Obtel.Enabled = 0
    Sub_31.Enabled = 0
    ViewNcell.Enabled = 0
    MnuC_A.Enabled = 0
    'MnuSql.Enabled = 0
    MnuLabel.Enabled = False
    MnuLabelMark.Enabled = 0
    RadioLink.Enabled = 0
    Hopping.Enabled = 0
    SUB_431.Enabled = 0
'    NetworkBlind.Enabled = 0
  '  NetworkDisturb.Enabled = 0
 '   View_Cope.Enabled = 0
    
    'SUB_32.Enabled = 0
    SUB_33.Enabled = 0
    SUB_41.Enabled = 0
'    SUB_42.Enabled = 0
'    SUB_43.Enabled = 0
    SUB_441.Enabled = 0
    SUB_442.Enabled = 0
 '   SUB_443.Enabled = 0
    Mnu_Replay.Enabled = 0
'    SUB_45.Enabled = 0
    SUB_123.Enabled = 1
    ScanPilot.Enabled = 1
'    SCAN_2.Enabled = 1
'    SCAN_3.Enabled = 1
    My_ScanPlay.Enabled = True
    Arfcn_Changing.Enabled = True
    SCAN_4.Enabled = 1
    SCAN_5.Enabled = 1
    SCAN_6.Enabled = 1
    SCAN_7.Enabled = 1
    SCAN_8.Enabled = 1
    TRAN_C_I.Enabled = 1
    My_Over.Enabled = 1
    Mnudistributing.Enabled = 1
'    OPen_Str_Data.Enabled = 0
'    Static_Pad.Enabled = 0
'    report.Enabled = 0
'    STREET_AN.Enabled = 0
    StatusBar.Panels(4).Text = "扫频分析"
    
  End If
  If Map_No = 1 Then
     SUB_24.Enabled = 1
  Else
    SUB_23.Enabled = 0
'    SUB_24.Enabled = 0
    SUB_25.Enabled = 0
    SUB_26.Enabled = 0
    SUB_28.Enabled = 0
    SUB_26.Enabled = 0
    
    USERMARK.Enabled = 0
    CLOSEMARK.Enabled = 0
    SAVEMARK.Enabled = 0
    
    SUB_461.Enabled = 0
    SUB_462.Enabled = 0
    SUB_463.Enabled = 0
    SUB_464.Enabled = 0
    MnuBcchRetrieve.Enabled = False
    SUB_465.Enabled = 0
    MnuCellReUse.Enabled = False
    SUB_466.Enabled = 0
    SUB_467.Enabled = 0
    SUB_468.Enabled = 0
    FindMyBsic.Enabled = 0
    FindFree.Enabled = 0
    SUB_469.Enabled = 0
    BsNo_FindCell.Enabled = 0
    SUB_4600.Enabled = 0
     
'    OPen_Str_Data.Enabled = 0
'    Static_Pad.Enabled = 0
'    report.Enabled = 0
    SUB_CENTER.Enabled = 0
    Toolbar.Buttons(33).Enabled = False
  End If

 If thereIsAMap = 0 Then
        SUB_23.Enabled = 0
'        SUB_24.Enabled = 0
        SUB_25.Enabled = 0
        SUB_26.Enabled = 0
 End If
End Sub

Private Sub MDIForm_Resize()
    Dim CellHeadData As ScanHead
    
    On Error Resume Next
    If Face_show = 0 Then
       Face.Show 1
       Face_show = 1
       If Dir(Gsm_Path + "\map\cell.dbf", 0) <> "" Then
          hDbfFile = FreeFile
          Open Gsm_Path + "\map\cell.dbf" For Binary As #hDbfFile
          Get #hDbfFile, , CellHeadData
          Close #hDbfFile
          If CellHeadData.RecordLen <> (35 + 1) * 32 + 1 And CellHeadData.RecordLen <> 336 + 5 Then
             If (MsgBox("系统检测到你的基站库结构是旧的，" + Chr(10) + "如果不更新有些功能将不能正常使用。" + Chr(10) + Chr(10) + "想现在就更新吗？", 36, "提示")) = 6 Then
                UpdateFileName = Gsm_Path + "\map\cell.dbf"
                Menu_Flag = 9999
                Data_Convert.Show 1
             End If
          End If
       End If
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error Resume Next
'******************************************
    Gsm_FileName = Gsm_Path + "\NARF_Jam.*"
    If Dir(Gsm_FileName) <> "" Then
       Kill Gsm_FileName
    End If
    Gsm_FileName = Gsm_Path + "\sARF_Jam.*"
    If Dir(Gsm_FileName) <> "" Then
       Kill Gsm_FileName
    End If
    Gsm_FileName = Gsm_Path + "\UnUseBcch.*"
    If Dir(Gsm_FileName) <> "" Then
       Kill Gsm_FileName
    End If
    Gsm_FileName = Gsm_Path + "\gsm_temp.*"
    If Dir(Gsm_FileName) <> "" Then
       Kill Gsm_FileName
    End If
    Gsm_FileName = Gsm_Path + "\local.*"
    If Dir(Gsm_FileName) <> "" Then
       Kill Gsm_FileName
    End If
    Gsm_FileName = Gsm_Path + "\outer.*"
    If Dir(Gsm_FileName) <> "" Then
       Kill Gsm_FileName
    End If
    Gsm_FileName = Gsm_Path + "\jam.*"
    If Dir(Gsm_FileName) <> "" Then
       Kill Gsm_FileName
    End If
    Gsm_FileName = Gsm_Path + "\blind.*"
    If Dir(Gsm_FileName) <> "" Then
       Kill Gsm_FileName
    End If
    Gsm_FileName = Gsm_Path + "\NeighborLay1.*"
    If Dir(Gsm_FileName) <> "" Then
       Kill Gsm_FileName
    End If
    Gsm_FileName = Gsm_Path + "\NeighborLay2.*"
    If Dir(Gsm_FileName) <> "" Then
       Kill Gsm_FileName
    End If
    Gsm_FileName = Gsm_Path + "\NeighborLay3.*"
    If Dir(Gsm_FileName) <> "" Then
       Kill Gsm_FileName
    End If
    
'******************************************
    
    mapinfo.runmenucommand 104
    thereIsAMap = 0
    MapForm.Hide
    Unload MapForm
    mapinfo.setcallback Nothing
    Set myCallback = Nothing
    mapinfo.Do "End mapinfo "
    Gsm_FileName = Gsm_Path + "\map\street.map"
    Gsm_File2 = Gsm_Path + "\map\street.tab"
    Kill Gsm_FileName
    FileCopy Gsm_File2, Gsm_FileName
    End
End Sub

Public Sub OPen_All_Map_Click()
    Dim i As Integer
    Dim MapWinFlag As Boolean
    Dim DCSFlag As Boolean, GSMFlag As Boolean

 'Dim conf_val As String * 8
 Dim conf_val As String * 25
 On Error Resume Next
 
 Gsm_FileName = Gsm_Path + "\ant.cfg"
 If Dir(Gsm_FileName, 0) = "" Then
    Config_frm.Show 1
    Exit Sub
 End If
 StatusBar.Panels(2).Text = " 区域图"
 Open Gsm_FileName For Binary As #1
 Get #1, , conf_val
 Close
 
 If Mid(conf_val, 1, 1) = "1" Then
    Map_No = 1
    Gsm_FileName = Gsm_Path + "\map"
    ChDir Gsm_FileName
    Msg = "Open Table " + Chr(34) + Gsm_Path + "\map\Area" + Chr(34)
    mapinfo.Do Msg

    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    MapForm.Caption = MapForm.Caption + "Area"
    TableNum = Val(mapinfo.eval("NumTables()"))
    MapWinFlag = False
    
                  For i = 1 To mapinfo.eval("NumWindows()")     'win95
                      If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then     'win95
                         MapWinFlag = True
                         Exit For
                      End If     'win95
                  Next     'win95
    
    
    If MapWinFlag Then
      Msg = "Add Map Auto Layer" + Chr(34) + "AREA" + Chr(34)
      mapinfo.Do Msg
     Else
      Msg = "Map from " + Chr(34) + "Area" + Chr(34)
      mapinfo.Do Msg

      Map_No = 1
      Msg = Chr(34) + "km" + Chr(34)
      mapinfo.Do "set map zoom 30 units " & Msg
      thereIsAMap = True
      mapid = Val(mapinfo.eval("FrontWindow()"))
      
      If Mid(conf_val, 22, 1) = "1" Then
        mapinfo.Do "set map redraw off"
        mapinfo.Do "Set Map Layer ""Area"" Label Font (""宋体"",256,9,14680288,16777215) Auto On Visibility Zoom (0, 100) Units ""km"""
        mapinfo.Do "set map redraw on"
      End If
    End If

    mapinfo.Do "set map redraw off"
    mapinfo.Do "set Map Layer Area Selectable Off"
    mapinfo.Do "set map redraw on"
End If
If Mid(conf_val, 4, 1) = "1" Then
    Map_No = 1
    Gsm_FileName = Gsm_Path + "\map"
    ChDir Gsm_FileName
    StatusBar.Panels(2).Text = " 打开街区图"
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    Msg = "Open Table " + Chr(34) + Gsm_Path + "\map\block" + Chr(34)
    mapinfo.Do Msg
    MapForm.Caption = MapForm.Caption + ",block"
    TableNum = Val(mapinfo.eval("NumTables()"))
    If TableNum > 1 Then
      Msg = "Add Map Auto Layer" + Chr(34) + "block" + Chr(34)
      mapinfo.Do Msg
     Else
      Msg = "Map from " + Chr(34) + "block" + Chr(34)
      mapinfo.Do Msg

      Map_No = 1
      Msg = Chr(34) + "km" + Chr(34)
      mapinfo.Do "set map zoom 30 units " & Msg
      thereIsAMap = True
      mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
End If

If Mid(conf_val, 5, 1) = "1" Then
    StatusBar.Panels(2).Text = " 打开市镇"
    Map_No = 1
    Gsm_FileName = Gsm_Path + "\map\town.tab"
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    If Dir(Gsm_FileName, 0) <> "" Then
       Msg = "Open Table " + Chr(34) + Gsm_Path + "\map\town" + Chr(34)
       mapinfo.Do Msg
       MapForm.Caption = MapForm.Caption + ",town"
       TableNum = Val(mapinfo.eval("NumTables()"))
       If TableNum > 1 Then
         Msg = "Add Map Auto Layer" + Chr(34) + "town" + Chr(34)
         mapinfo.Do Msg
       Else
          Msg = "Map from " + Chr(34) + "town" + Chr(34)
          mapinfo.Do Msg
         Msg = Chr(34) + "km" + Chr(34)
         mapinfo.Do "set map zoom 30 units " & Msg
         thereIsAMap = True
         mapid = Val(mapinfo.eval("FrontWindow()"))
       End If
      If Mid(conf_val, 23, 1) = "1" Then
        mapinfo.Do "set map redraw off"
        mapinfo.Do "Set Map Layer ""town"" Label Font (""宋体"",256,9,27552,16777215) Auto On Visibility Zoom (0, 100) Units ""km"""
        mapinfo.Do "set map redraw on"
      End If
   
    Else
       Msg = "Open Table " + Chr(34) + Gsm_Path + "\map\vip" + Chr(34)
       mapinfo.Do Msg
       MapForm.Caption = MapForm.Caption + ",vip"
       TableNum = Val(mapinfo.eval("NumTables()"))
       If TableNum > 1 Then
         Msg = "Add Map Auto Layer" + Chr(34) + "vip" + Chr(34)
         mapinfo.Do Msg
       Else
         Msg = "Map from " + Chr(34) + "vip" + Chr(34)
         mapinfo.Do Msg
         Msg = Chr(34) + "km" + Chr(34)
         mapinfo.Do "set map zoom 30 units " & Msg
         thereIsAMap = True
         mapid = Val(mapinfo.eval("FrontWindow()"))
       End If
      If Mid(conf_val, 23, 1) = "1" Then
        mapinfo.Do "set map redraw off"
        mapinfo.Do "Set Map Layer ""vip"" Label Font (""宋体"",256,9,27552,16777215) Auto On Visibility Zoom (0, 100) Units ""km"""
        mapinfo.Do "set map redraw on"
      End If
     
     End If
End If

If Mid(conf_val, 2, 1) = "1" Then
    Map_No = 1
    StatusBar.Panels(2).Text = " 打开绿化图"
    Gsm_FileName = Gsm_Path + "\map"
    ChDir Gsm_FileName
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    Msg = "Open Table " + Chr(34) + Gsm_Path + "\map\Landmark" + Chr(34)
    mapinfo.Do Msg

    TableNum = Val(mapinfo.eval("NumTables()"))
    If TableNum > 1 Then
      Msg = "Add Map Auto Layer" + Chr(34) + "landmark" + Chr(34)
      mapinfo.Do Msg
     Else
      Msg = "Map from " + Chr(34) + "landmark" + Chr(34)
      mapinfo.Do Msg

      Msg = Chr(34) + "km" + Chr(34)
      mapinfo.Do "set map zoom 30 units " & Msg
      thereIsAMap = True
      mapid = Val(mapinfo.eval("FrontWindow()"))
    End If

    MapForm.Caption = MapForm.Caption + "," + "Landmark"
    mapinfo.Do "Set Map Layer LandMark Editable off selectable  off"
    StatusBar.Panels(2).Text = " "

    StatusBar.Panels(2).Text = " 打开水域图"
    Gsm_FileName = Gsm_Path + "\map"
    ChDir Gsm_FileName
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    Msg = "Open Table " + Chr(34) + Gsm_Path + "\map\water" + Chr(34)
    mapinfo.Do Msg
    MapForm.Caption = MapForm.Caption + "," + "Water"
    Msg = "Add Map window   Frontwindow()  Layer" + Chr(34) + "Water" + Chr(34)
    mapinfo.Do Msg
    mapinfo.Do "Set Map Layer  water Editable off selectable  off"

    StatusBar.Panels(2).Text = " 打开山峰图"
    Gsm_FileName = Gsm_Path + "\map"
    ChDir Gsm_FileName
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    Msg = "Open Table " + Chr(34) + Gsm_Path + "\map\mountain" + Chr(34)
    mapinfo.Do Msg
    MapForm.Caption = MapForm.Caption + "," + "mountain"
    Msg = "Add Map window   Frontwindow()  Layer" + Chr(34) + "mountain" + Chr(34)
    mapinfo.Do Msg
    mapinfo.Do "Set Map Layer  mountain Editable off selectable  off"
End If

If Mid(conf_val, 3, 1) = "1" Then
    Map_No = 1
    StatusBar.Panels(2).Text = " 打开基站"
    Gsm_FileName = Gsm_Path + "\map"
    ChDir Gsm_FileName
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    Msg = "Open Table " + Chr(34) + Gsm_Path + "\map\base" + Chr(34) + " Interactive"
    mapinfo.Do Msg
    MapForm.Caption = MapForm.Caption + "," + "base"

    TableNum = Val(mapinfo.eval("NumTables()"))
    If TableNum > 1 Then
      Msg = "Add Map Auto Layer" + Chr(34) + "base" + Chr(34)
      mapinfo.Do Msg
     Else
      Msg = "Map from " + Chr(34) + "base" + Chr(34)
      mapinfo.Do Msg

      Msg = Chr(34) + "km" + Chr(34)
      mapinfo.Do "set map zoom 30 units " & Msg
      thereIsAMap = True
      mapid = Val(mapinfo.eval("FrontWindow()"))
    End If

    'If Mid(conf_val, 17, 1) = "1" And Mid(conf_val, 18, 1) = "1" And Mid(conf_val, 19, 1) = "1" Then
    DCSFlag = False
    GSMFlag = False
    If Mid(conf_val, 17, 1) = "1" Or Mid(conf_val, 18, 1) = "1" Or Mid(conf_val, 19, 1) = "1" Or Mid(conf_val, 20, 1) = "1" Then
        mapinfo.Do "Open Table " + Chr(34) + Gsm_Path + "\map\cell" + Chr(34)
    End If
    MapForm.Caption = MapForm.Caption + "," + "cell"
    If Mid(conf_val, 20, 1) = "1" Then
        mapinfo.Do "select * from cell where basetype =""3"" into DCSCell"
        mapinfo.Do "Add Map Auto Layer" + Chr(34) + "DCSCell" + Chr(34)
        DCSFlag = True
    End If
    If Mid(conf_val, 17, 1) = "1" Or Mid(conf_val, 18, 1) = "1" Or Mid(conf_val, 19, 1) = "1" Then
       Msg = "select * from cell where"
       If Mid(conf_val, 17, 1) = "1" Then
          Msg = Msg & " basetype =""0"" or basetype="""" or"
       End If
       If Mid(conf_val, 18, 1) = "1" Then
          Msg = Msg & " basetype =""1"" or"
       End If
       If Mid(conf_val, 19, 1) = "1" Then
          Msg = Msg & " basetype =""2"" or"
       End If
       
       Msg = Left(Msg, Len(Msg) - 2)
       Msg = Msg & " into GSMCell"
       mapinfo.Do Msg
       mapinfo.Do "Add Map Auto Layer" + Chr(34) + "GSMCell" + Chr(34)
       GSMFlag = True
    End If
    mapinfo.Do "close table selection"
    mapinfo.runmenucommand 610
    'msg = "Open Table " + Chr(34) + Gsm_Path + "\map\cell" + Chr(34)
    'mapinfo.Do msg
    'MapForm.Caption = MapForm.Caption + "," + "cell"
    'msg = "Add Map Auto Layer" + Chr(34) + "cell" + Chr(34)
    'mapinfo.Do msg

    mapinfo.Do "Set Map  Order  2, 1"
    
    SUB_CENTER.Enabled = 1
'    NCELL.Enabled = 1
    Toolbar.Buttons(33).Enabled = True
    SUB_26.Enabled = 1
    SUB_41.Enabled = 1
    SUB_461.Enabled = 1
    SUB_462.Enabled = 1
    SUB_463.Enabled = 1
    SUB_464.Enabled = 1
    MnuBcchRetrieve.Enabled = True
    SUB_465.Enabled = 1
    MnuCellReUse.Enabled = True
    SUB_466.Enabled = 1
    SUB_467.Enabled = 1
    SUB_468.Enabled = 1
    FindMyBsic.Enabled = 1
    FindFree.Enabled = 1
    SUB_469.Enabled = 1
    BsNo_FindCell.Enabled = 1
    SUB_4600.Enabled = 1
     
    If Mid(conf_val, 7, 1) = "1" Then
       If Mid(conf_val, 17, 1) = "1" Or Mid(conf_val, 18, 1) = "1" Or Mid(conf_val, 19, 1) = "1" Or Mid(conf_val, 20, 1) = "1" Then
          mapinfo.Do "fetch first from cell"
          mapinfo.Do "set map redraw off"
          If Val(mapinfo.eval("cell.bearing")) = 0 Then
             mapinfo.Do "Set Map Layer " + Chr(34) + "Base" + Chr(34) + " Label  Position Left Visibility  Font (""宋体"",257,9,0,16777136) Auto On offset 11 Overlap On Visibility Zoom (0, 20) Units ""km"""
          Else
             mapinfo.Do "Set Map Layer " + Chr(34) + "Base" + Chr(34) + " Label  Position above Visibility  Font (""宋体"",257,9,0,16777136) Auto On offset 11 Overlap On Visibility Zoom (0, 20) Units ""km"""
          End If
            'mapinfo.do "Set Map Layer " + Chr(34) + "cell" + Chr(34) + " Label Visibility  Font (""Arial"",257,8,255,16777215)  With Arfcn Auto On Overlap On Duplicates On Position Center "
            'Position Above Auto On Offset 0
          If GSMFlag Then
             mapinfo.Do "Set Map Layer " + Chr(34) + "GSMCell" + Chr(34) + " Label Visibility  Font (""Arial"",257,8,255,16777215)  With Arfcn Auto On Overlap On Duplicates On Position  Above  Offset 0 Visibility Zoom (0, 20) Units ""km"""
          End If
          If DCSFlag Then
             mapinfo.Do "Set Map Layer " + Chr(34) + "DCScell" + Chr(34) + " Label Visibility  Font (""Arial"",257,8,255,16777215)  With Arfcn Auto On Overlap On Duplicates On Position  Above  Offset 0 Visibility Zoom (0, 20) Units ""km"""
          End If
          mapinfo.Do "set map redraw on"
       End If
   End If
   If Mid(conf_val, 6, 1) = "1" Then
       mapinfo.Do "fetch first from cell"
       mapinfo.Do "set map redraw off"
       If Val(mapinfo.eval("cell.bearing")) = 0 Then
          mapinfo.Do "Set Map Layer " + Chr(34) + "Base" + Chr(34) + " Label  Position Left Visibility  Font (""宋体"",257,9,0,16777136) Auto On offset 11 Overlap On Visibility Zoom (0, 20) Units ""km"""
       Else
          mapinfo.Do "Set Map Layer " + Chr(34) + "Base" + Chr(34) + " Label  Position above Visibility  Font (""宋体"",257,9,0,16777136) Auto On offset 11 Overlap On Visibility Zoom (0, 20) Units ""km"""
       End If
       If DCSFlag Then
          mapinfo.Do "Set Map Layer " + Chr(34) + "DCSCell" + Chr(34) + " Label Font (""Arial"",256,8,255,16777215) With Non_bcch Auto On Overlap On Duplicates On Position Center Visibility Zoom (0, 20) Units ""km"""    'Zoom (0, 5) Units ""mi"""
       End If
       If GSMFlag Then
          mapinfo.Do "Set Map Layer " + Chr(34) + "GSMcell" + Chr(34) + " Label Font (""Arial"",256,8,255,16777215) With Non_bcch Auto On Overlap On Duplicates On Position Center Visibility Zoom (0, 20) Units ""km"""    'Zoom (0, 5) Units ""mi"""
       End If
       'mapinfo.do "Set Map Layer " + Chr(34) + "cell" + Chr(34) + " Label Visibility  Font (""Arial"",257,8,255,16777215)  With non_bcch Auto On Overlap On Duplicates On Position Center "
       mapinfo.Do "set map redraw on"
   End If
   If Mid(conf_val, 21, 1) = "1" Then
       mapinfo.Do "fetch first from cell"
       mapinfo.Do "set map redraw off"
       If Val(mapinfo.eval("cell.bearing")) = 0 Then
          mapinfo.Do "Set Map Layer " + Chr(34) + "Base" + Chr(34) + " Label  Position Left Visibility  Font (""宋体"",257,9,0,16777136) Auto On offset 11 Overlap On Visibility Zoom (0, 20) Units ""km"""
       Else
          mapinfo.Do "Set Map Layer " + Chr(34) + "Base" + Chr(34) + " Label  Position above Visibility  Font (""宋体"",257,9,0,16777136) Auto On offset 11 Overlap On Visibility Zoom (0, 20) Units ""km"""
       End If
       If DCSFlag Then
          mapinfo.Do "Set Map Layer " + Chr(34) + "DCSCell" + Chr(34) + " Label Font (""Arial"",256,8,255,16777215) With ci Auto On Overlap On Duplicates On Position Center Visibility Zoom (0, 20) Units ""km"""    'Zoom (0, 5) Units ""mi"""
       End If
       If GSMFlag Then
          mapinfo.Do "Set Map Layer " + Chr(34) + "GSMcell" + Chr(34) + " Label Font (""Arial"",256,8,255,16777215) With ci Auto On Overlap On Duplicates On Position Center Visibility Zoom (0, 20) Units ""km"""    'Zoom (0, 5) Units ""mi"""
       End If
       'mapinfo.do "Set Map Layer " + Chr(34) + "cell" + Chr(34) + " Label Visibility  Font (""Arial"",257,8,255,16777215)  With non_bcch Auto On Overlap On Duplicates On Position Center "
       mapinfo.Do "set map redraw on"
   End If
   
End If

'If Mid(conf_val, 6, 1) = "1" Then
'    Map_No = 1
'    StatusBar.Panels(2).Text = "打开竞争对手基站"
'    Gsm_FileName = Gsm_Path + "\map"
'    ChDir Gsm_FileName
'    mapinfo.do "Set Next Document Parent " & MapForm.hwnd & " Style 1"
'    msg = "Open Table " + Chr(34) + Gsm_Path + "\map\compbase" + Chr(34) + " Interactive"
'    mapinfo.do msg

'    MapForm.Caption = MapForm.Caption + "," + "compbase"
'    TableNum = Val(mapinfo.eval("NumTables()"))
'    If TableNum > 1 Then
'      msg = "Add Map Auto Layer" + Chr(34) + "compbase" + Chr(34)
'      mapinfo.do msg
'     Else
'      msg = "Map from " + Chr(34) + "compbase" + Chr(34)
'      mapinfo.do msg

'      msg = Chr(34) + "km" + Chr(34)
'      mapinfo.do "set map zoom 30 units " & msg
'      thereIsAMap = True
'      mapid = Val(mapinfo.eval("FrontWindow()"))
'    End If

'    mapinfo.do "set map redraw off"
'    mapinfo.do "Set Map Layer " + Chr(34) + "Base" + Chr(34) + " Position above Label Font (""宋体"",257,9,0,16777136) Auto On Overlap On "
'    mapinfo.do "set map redraw on"


'    msg = "Open Table " + Chr(34) + Gsm_Path + "\map\compcell" + Chr(34)
'    mapinfo.do msg

'    mapinfo.do "fetch Rec 10 from compcell"
'    CM = "compcell.obj"
'    ver_x = Val(mapinfo.eval("centroidx(" & CM & ")")) - west
'    ver_y = Val(mapinfo.eval("centroidy(" & CM & ")")) - south
'    If ver_y < 0 Or ver_y > yy Or ver_x < 0 Or ver_x > xx Then
'        MsgBox "本软件仅供授权的合法用户使用，欢迎您申请后使用！", 64, "提示"
'        GoTo VER_OUT
'    End If

'    MapForm.Caption = MapForm.Caption + "," + "compcell"
'    msg = "Add Map Auto Layer" + Chr(34) + "compcell" + Chr(34)
'    mapinfo.do msg
'End If

'Gsm_FileName = Gsm_Path + "\ant.cfg"
'Open Gsm_FileName For Binary As #1
' Get #1, 9, conf_val
' Close
 
 If Mid(conf_val, 1 + 8, 1) = "1" Then
    StatusBar.Panels(2).Text = " 打开街道图"
    Map_No = 1
    Gsm_FileName = Gsm_Path + "\map"
    ChDir Gsm_FileName
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    Msg = "Open Table " + Chr(34) + Gsm_Path + "\map\street" + Chr(34)
    mapinfo.Do Msg

'    mapinfo.do "fetch first  from street"
'    mapinfo.do "fetch next  from street"
'    mapinfo.do "fetch next  from street"
'    CM = "street.obj"
'    ver_x = Val(mapinfo.eval("centroidx(" & CM & ")")) - west
'    ver_y = Val(mapinfo.eval("centroidy(" & CM & ")")) - south
'    If ver_y < 0 Or ver_y > yy Or ver_x < 0 Or ver_x > xx Then
'        MsgBox "本软件仅供授权的合法用户使用，欢迎您申请后使用！", 64, "提示"
'        GoTo VER_OUT
'    End If

    MapForm.Caption = MapForm.Caption + ",street"
    TableNum = Val(mapinfo.eval("NumTables()"))
    If TableNum > 1 Then
      Msg = "Add Map Auto Layer" + Chr(34) + "street" + Chr(34)
      mapinfo.Do Msg
    Else
      Msg = "Map from " + Chr(34) + "street" + Chr(34)
      mapinfo.Do Msg

      Msg = Chr(34) + "km" + Chr(34)
      mapinfo.Do "set map zoom 30 units " & Msg
      thereIsAMap = True
      mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
    SUB_23.Enabled = 1
    SUB_24.Enabled = 1
    SUB_25.Enabled = 1
    SUB_26.Enabled = 1
'    SUB_13.Enabled = 0
    USERMARK.Enabled = 1
    CLOSEMARK.Enabled = 1
    SAVEMARK.Enabled = 1
'    STREET_AN.Enabled = 1
    Over = 0
    mapinfo.Do "Set Map Layer  street Editable off selectable  off"
 End If

 If Mid(conf_val, 2 + 8, 1) = "1" Then
    StatusBar.Panels(2).Text = " 打开PUBLIC"
    Map_No = 1
    Gsm_FileName = Gsm_Path + "\map"
    ChDir Gsm_FileName
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    Msg = "Open Table " + Chr(34) + Gsm_Path + "\map\public" + Chr(34)
    mapinfo.Do Msg

    MapForm.Caption = MapForm.Caption + ",public"
    TableNum = Val(mapinfo.eval("NumTables()"))
    If TableNum > 1 Then
      Msg = "Add Map Auto Layer" + Chr(34) + "public" + Chr(34)
      mapinfo.Do Msg
    Else
      Msg = "Map from " + Chr(34) + "public" + Chr(34)
      mapinfo.Do Msg

      Msg = Chr(34) + "km" + Chr(34)
      mapinfo.Do "set map zoom 30 units " & Msg
      thereIsAMap = True
      mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
    SUB_23.Enabled = 1
    SUB_24.Enabled = 1
    SUB_25.Enabled = 1
    SUB_26.Enabled = 1
'    SUB_13.Enabled = 0
    USERMARK.Enabled = 1
    CLOSEMARK.Enabled = 1
    SAVEMARK.Enabled = 1
 End If

 If Mid(conf_val, 3 + 8, 1) = "1" Then
 End If

 If Mid(conf_val, 4 + 8, 1) = "1" Then
    StatusBar.Panels(2).Text = " 打开POST"
    Map_No = 1
    Gsm_FileName = Gsm_Path + "\map"
    ChDir Gsm_FileName
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    Msg = "Open Table " + Chr(34) + Gsm_Path + "\map\post" + Chr(34)
    mapinfo.Do Msg

    MapForm.Caption = MapForm.Caption + ",post"
    TableNum = Val(mapinfo.eval("NumTables()"))
    If TableNum > 1 Then
      Msg = "Add Map Auto Layer" + Chr(34) + "post" + Chr(34)
      mapinfo.Do Msg
    Else
      Msg = "Map from " + Chr(34) + "post" + Chr(34)
      mapinfo.Do Msg

      Msg = Chr(34) + "km" + Chr(34)
      mapinfo.Do "set map zoom 30 units " & Msg
      thereIsAMap = True
      mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
    SUB_23.Enabled = 1
    SUB_24.Enabled = 1
    SUB_25.Enabled = 1
    SUB_26.Enabled = 1
    SUB_13.Enabled = 0
    USERMARK.Enabled = 1
    CLOSEMARK.Enabled = 1
    SAVEMARK.Enabled = 1
 End If

 If Mid(conf_val, 5 + 8, 1) = "1" Then
    StatusBar.Panels(2).Text = " 打开USER_1"
    Map_No = 1
    Gsm_FileName = Gsm_Path + "\map"
    ChDir Gsm_FileName
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    Msg = "Open Table " + Chr(34) + Gsm_Path + "\user\user_1" + Chr(34)
    mapinfo.Do Msg

    MapForm.Caption = MapForm.Caption + ",user_1"
    TableNum = Val(mapinfo.eval("NumTables()"))
    If TableNum > 1 Then
      Msg = "Add Map Auto Layer" + Chr(34) + "user_1" + Chr(34)
      mapinfo.Do Msg
    Else
      Msg = "Map from " + Chr(34) + "user_1" + Chr(34)
      mapinfo.Do Msg

      Msg = Chr(34) + "km" + Chr(34)
      mapinfo.Do "set map zoom 30 units " & Msg
      thereIsAMap = True
      mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
 End If

 If Mid(conf_val, 6 + 8, 1) = "1" Then
    StatusBar.Panels(2).Text = " 打开USER_1"
    Map_No = 1
    Gsm_FileName = Gsm_Path + "\map"
    ChDir Gsm_FileName
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    Msg = "Open Table " + Chr(34) + Gsm_Path + "\user\user_2" + Chr(34)
    mapinfo.Do Msg

    MapForm.Caption = MapForm.Caption + ",user_2"
    TableNum = Val(mapinfo.eval("NumTables()"))
    If TableNum > 1 Then
      Msg = "Add Map Auto Layer" + Chr(34) + "user_2" + Chr(34)
      mapinfo.Do Msg
    Else
      Msg = "Map from " + Chr(34) + "user_2" + Chr(34)
      mapinfo.Do Msg

      Msg = Chr(34) + "km" + Chr(34)
      mapinfo.Do "set map zoom 30 units " & Msg
      thereIsAMap = True
      mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
 End If

 If Mid(conf_val, 7 + 8, 1) = "1" Then
    StatusBar.Panels(2).Text = " 打开USER_3"
    Map_No = 1
    Gsm_FileName = Gsm_Path + "\map"
    ChDir Gsm_FileName
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    Msg = "Open Table " + Chr(34) + Gsm_Path + "\user\user_3" + Chr(34)
    mapinfo.Do Msg

    MapForm.Caption = MapForm.Caption + ",user_3"
    TableNum = Val(mapinfo.eval("NumTables()"))
    If TableNum > 1 Then
      Msg = "Add Map Auto Layer" + Chr(34) + "user_3" + Chr(34)
      mapinfo.Do Msg
    Else
      Msg = "Map from " + Chr(34) + "user_3" + Chr(34)
      mapinfo.Do Msg

      Msg = Chr(34) + "km" + Chr(34)
      mapinfo.Do "set map zoom 30 units " & Msg
      thereIsAMap = True
      mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
 End If

 If Map_No = 1 Then
    SUb_21.Enabled = 1
    SUB_23.Enabled = 1
    SUB_24.Enabled = 1
    SUB_25.Enabled = 1
    SUB_26.Enabled = 1
    USERMARK.Enabled = 1
    CLOSEMARK.Enabled = 1
    SAVEMARK.Enabled = 1
 End If
Go_OUT:
VER_OUT:

    mapinfo.Do "set map display position"
    dis_flag = 1
    StatusBar.Panels(2).Text = " "
    Exit Sub
No_Map:
    MsgBox "无地图窗口", 64, "提示"
    Unload MapForm
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub OpenMap_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 打开地图窗口"
    Menu_Flag = 81
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub Pline_Click()
        StatusBar.Panels(2).Text = " 当前工具：曲线"
        On Error Resume Next
        mapinfo.runmenucommand 1713
End Sub

Private Sub REDRAW_Click()
    StatusBar.Panels(2).Text = " 重画窗口"
    On Error Resume Next
    mapinfo.runmenucommand 610
'    MDIMain.Arrange REDRAW
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub Region_Click()
    StatusBar.Panels(2).Text = "工具：转换为区域"
    On Error Resume Next
    mapinfo.Do "Set Style brush Makebrush(2,16777215,0)"
    mapinfo.runmenucommand 1607
End Sub

Private Sub SAVEMARK_Click()
    StatusBar.Panels(2).Text = " 保存用户标识层"
    On Error Resume Next
    mapinfo.runmenucommand 809
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SCAN_2_Click()
  On Error Resume Next
  StatusBar.Panels(2).Text = " 生成场强优化图"
  Menu_Flag = 912
  SelTable.Show 1
  StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SCAN_3_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 生成场强弱区图"
    Menu_Flag = 913
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SCAN_4_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 生成非本地覆盖图"
    Menu_Flag = 914
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SCAN_5_Click()
    On Error Resume Next
    Menu_Flag = 915
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SCAN_6_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 生成干扰点图"
    Menu_Flag = 916
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SSClengend_Click()
    Dim i As Integer, windows_num As Integer
    Dim WinId As Variant
    Dim Legend_Win As Variant
        
    On Error Resume Next
    StatusBar.Panels(2).Text = " 打开图例"
    For i = 1 To mapinfo.eval("NumWindows()")
        If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then
           WinId = mapinfo.eval("windowid(" & i & ")")
           If WinId = mapinfo.eval("frontwindow()") Then
              Exit For
           End If
        End If
    Next
    If thereIsAMap Then
       mapinfo.Do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
       'If legendid = 0 Then
          mapinfo.Do "Create Legend From Window " & WinId
          mapinfo.runmenucommand 606
          legendid = mapinfo.eval("windowinfo(1009,12)")
       'Else
       '   windows_num = mapinfo.eval("numallwindows()")
       '   For i = -1 To -windows_num Step -1
       '       If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1009 Then
       '          Legend_Win = mapinfo.eval("WindowID(" & i & ")")
       '          mapinfo.do "Close Window " & Legend_Win
       '       End If
       '   Next
       '   legendid = 0
       'End If
    End If
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SSCommand1_Click()
'        mapinfo.RunMenuCommand M_TOOLS_SELECT
        StatusBar.Panels(2).Text = " 当前工具：选择"
        On Error Resume Next
        mapinfo.runmenucommand 1701
End Sub

Private Sub SSCommand2_Click()
        StatusBar.Panels(2).Text = " 当前工具：放大"
        On Error Resume Next
        mapinfo.runmenucommand 1705  'M_TOOLS_EXPAND
End Sub

Private Sub SSCommand3_Click()
        StatusBar.Panels(2).Text = " 当前工具：缩小"
        On Error Resume Next
        mapinfo.runmenucommand 1706  'M_TOOLS_SHRINK
End Sub

Private Sub SSCommand4_Click()
        StatusBar.Panels(2).Text = " 当前工具：移动"
        On Error Resume Next
        mapinfo.runmenucommand 1702   'M_TOOLS_RECENTER
End Sub

Private Sub SSCommand5_Click()
        StatusBar.Panels(2).Text = " 当前工具：信息"
        On Error Resume Next
        mapinfo.runmenucommand 1707   'M_TOOLS_PNT_QUERY
End Sub

Private Sub SSCommand6_Click()
        StatusBar.Panels(2).Text = " 当前工具：标注"
        On Error Resume Next
        mapinfo.runmenucommand 1708   'M_TOOLS_LABELER
End Sub

Private Sub SSCommand7_Click()
        StatusBar.Panels(2).Text = " 当前工具：尺子"
        On Error Resume Next
        mapinfo.runmenucommand 1710     'M_TOOLS_RULER
End Sub

Private Sub SSCommand8_Click()
        StatusBar.Panels(2).Text = " 当前工具：文本"
        On Error Resume Next
        mapinfo.runmenucommand 1709     'M_TOOLS_TEXT
End Sub

Private Sub SSCommand9_Click()
        StatusBar.Panels(2).Text = " 当前工具：符号"
        On Error Resume Next
        mapinfo.runmenucommand 1711
End Sub

Private Sub STS_1_Click()
    Dim MyRecord As Record
    Dim mypath As String, buff As String
    Dim finds As Integer, i As Integer
    Dim row, ci
    Dim Nametemp As String
    
    On Error Resume Next
    
    StatusBar.Panels(2).Text = " 数据转换"
    Menu_Flag = 2301
    Gsm_FileName = Gsm_Path + "\gsm.dat"
    Open Gsm_FileName For Binary As #1
    Get #1, 1, MyRecord  ' Read third record.
    Close #1
    If Val(MyRecord.exchange) = 0 Or Val(MyRecord.exchange) = 1 Then
       For i = 1 To 50
           convert_filename(i) = ""
       Next
open_again:
       FileDialog.DialogTitle = "数据转换文件选择"
       If Val(MyRecord.exchange) = 0 Then
          FileDialog.Filter = "*.xls Files|*.XLS|All Files|*.*"
          FileDialog.DefaultExt = "*.XLS"
          FileDialog.Flags = &H80000
       Else
          FileDialog.Filter = "*.txt Files|*.TXT|All Files|*.*"
          FileDialog.DefaultExt = "*.TXT"
          FileDialog.Flags = &H200 Or &H80000
       End If
       Gsm_FileName = Gsm_Path + "\sts"
       FileDialog.InitDir = Gsm_FileName
       FileDialog.ShowOpen
       buff = Trim(FileDialog.filename)
       If buff = "" Then
          StatusBar.Panels(2).Text = " "
          Exit Sub
       End If
       finds = InStr(buff, Chr(0))
       If finds > 0 Then
          mypath = Left(buff, finds - 1) + "\"
          buff = Trim(Right(buff, Len(buff) - finds))
          finds = InStr(buff, Chr(0))
          i = 1
          Do While finds > 0
             convert_filename(i) = mypath + Left(buff, finds - 1)
             buff = Trim(Right(buff, Len(buff) - finds))
             finds = InStr(buff, Chr(0))
             i = i + 1
          Loop
          convert_filename(i) = mypath + buff
       Else
          convert_filename(1) = buff
       End If
       FileDialog.filename = ""
       If Dir(convert_filename(1)) = "" Then
          GoTo err_exit
       End If
       If Val(MyRecord.exchange) = 1 Then
          Screen.MousePointer = 11
          Mot_Sts1.Show 1
          Screen.MousePointer = 0
          On Error GoTo Sts_Out
          
          mapinfo.Do "Register Table  " + " " + Chr(34) + Gsm_Path + "\sts\tch_sts.dbf" + Chr(34) + "Type " + " " + Chr(34) + "DBF" + Chr(34) + "Into  " + Chr(34) + Gsm_Path + "\sts\tch_sts.tab" + Chr(34)
          mapinfo.Do "Open Table " + Chr(34) + Gsm_Path + "\sts\tch_sts.tab" + Chr(34)
          mapinfo.Do "Register Table  " + " " + Chr(34) + Gsm_Path + "\sts\cch_sts.dbf" + Chr(34) + "Type " + " " + Chr(34) + "DBF" + Chr(34) + "Into  " + Chr(34) + Gsm_Path + "\sts\cch_sts.tab" + Chr(34)
          mapinfo.Do "Open Table " + Chr(34) + Gsm_Path + "\sts\cch_sts.tab" + Chr(34)
'********************************************************
          mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\map\cell" + Chr(34)
          mapinfo.Do "fetch first from tch_sts"
          row = Val(mapinfo.eval("tableinfo(tch_sts,8)"))
          mapinfo.Do "Create Map For tch_sts CoordSys Earth Projection 1, 0 "
          mapinfo.Do "fetch first from cch_sts"
          mapinfo.Do "Create Map For cch_sts CoordSys Earth Projection 1, 0 "
          mapinfo.Do "Set Style Pen MakePen(1,60,0)"
          mapinfo.Do "set style brush  makebrush(2,0,0) "
          mapinfo.Do "Set Style Symbol MakeSymbol(33,0,2)"
          For j = 1 To row
              ci = mapinfo.eval("tch_sts.col2")
              mapinfo.Do "select * from cell where col3 = " + Chr(34) + ci + Chr(34) + " into temp"
              temp_row = Val(mapinfo.eval("tableinfo(temp,8)"))
              If temp_row > 0 Then
                 Nametemp = mapinfo.eval("temp.col1")
                 If InStr(Nametemp, Chr(0)) > 0 Then
                    Nametemp = Trim(Left(Nametemp, InStr(Nametemp, Chr(0)) - 1))
                 End If
                 lon = mapinfo.eval("temp.lon")
                 lat = mapinfo.eval("temp.lat")
                 bearing = mapinfo.eval("temp.bearing")
                 lon = lon + 0.0015 * Sin(bearing * 0.01745329252)
                 lat = lat + 0.0015 * Cos(bearing * 0.01745329252)
                 mapinfo.Do " update tch_sts set col1 = " + Chr(34) + Nametemp + Chr(34) + ",lon = " + str(lon) + ",lat = " + str(lat) + ",bearing = " + str(bearing) + " where rowid = " & j
                 mapinfo.Do "create point into variable sts_mypoint (" & lon & "," & lat & ") symbol(34,7585792,2)"
                 mapinfo.Do "update tch_sts set Obj=sts_mypoint  where rowid=" & j
                 mapinfo.Do " update cch_sts set col1 = " + Chr(34) + Nametemp + Chr(34) + ", lon = " + str(lon) + ",lat = " + str(lat) + ",bearing = " + str(bearing) + " where rowid = " & j
                 mapinfo.Do "create point into variable sts_mypoint (" & lon & "," & lat & ") symbol(34,7585792,2)"
                 mapinfo.Do "update cch_sts set Obj=sts_mypoint  where rowid=" & j
              End If
              mapinfo.Do "fetch next from tch_sts"
              mapinfo.Do "fetch next from cch_sts"
              mapinfo.Do "fetch first from cell"
          Next
          mapinfo.Do "commit table tch_sts"
          mapinfo.Do "commit table cch_sts"
          mapinfo.Do "close table cell"
'**************************************************************************************
'******************************************************
          Screen.MousePointer = 0
          
          mapinfo.Do "close table tch_sts"
          mapinfo.Do "close table cch_sts"
          StatusBar.Panels(2).Text = " "
          Exit Sub
Sts_Out:
          Screen.MousePointer = 0
          MsgBox "有错误！没有生成STS数据! ", 64, "提示"
          StatusBar.Panels(2).Text = " "
          Exit Sub
       Else
          Screen.MousePointer = 11
          Data_Convert.Show 1
          Screen.MousePointer = 0
       End If
       StatusBar.Panels(2).Text = " "
       Exit Sub
err_exit:
       i = MsgBox("无法打开文件 " + convert_filename(1), 48, "打开文件")
       GoTo open_again
    Else
       MsgBox "该交换机类型的数据转换暂未挂接!", 64, "提示"
    End If
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub sts_2_Click()
    On Error Resume Next
    Gsm_FileName = Gsm_Path + "\sts\tch_sts.tab"
    Msg = UCase(Dir(Gsm_FileName, 0))
    If Msg <> "TCH_STS.TAB" Then
       MsgBox " TCH_STS.tab 不存在！", 64, "提示"
    Else
     Gsm_FileName = Gsm_Path + "\sts\cch_sts.tab"
     Msg = UCase(Dir(Gsm_FileName, 0))
     If Msg <> "CCH_STS.TAB" Then
          MsgBox " CCH_STS.tab 不存在！", 64, "提示"
     Else
          TCH_CCH_SEL.Show 1
     End If
   End If
End Sub

Private Sub SUB_11_Click()
        Dim ReturnValue As Integer
        On Error Resume Next
        ChDir "\tems"
        ReturnValue = Shell("\tems\tems.EXE", 3)
        ChDir Gsm_Path
End Sub

Private Sub T121_Click()
    Dim buff As String, mypath As String
    Dim finds As Integer, i As Integer
        
    On Error Resume Next
    StatusBar.Panels(2).Text = " 数据转换"
    For i = 1 To 50
        convert_filename(i) = ""
    Next
open_again:
    FileDialog.DialogTitle = "Tems 通话测试数据转换"
    FileDialog.Filter = "*.txt Files|*.TXT|All Files|*.*"
    FileDialog.DefaultExt = "*.TXT"
    FileDialog.Flags = &H80000 Or &H200
    Gsm_FileName = Gsm_Path + "\normal"
    FileDialog.InitDir = Gsm_FileName
    FileDialog.ShowOpen
    buff = Trim(FileDialog.filename)
    If buff = "" Then
       StatusBar.Panels(2).Text = " "
       Exit Sub
    End If
    tran_fn = 0
    finds = InStr(buff, Chr(0))
    If finds > 0 Then
       mypath = Left(buff, finds - 1) + "\"
       buff = Trim(Right(buff, Len(buff) - finds))
       finds = InStr(buff, Chr(0))
       i = 1
       Do While finds > 0
          convert_filename(i) = mypath + Left(buff, finds - 1)
          tran_f(i) = convert_filename(i)
          buff = Trim(Right(buff, Len(buff) - finds))
          finds = InStr(buff, Chr(0))
          i = i + 1
       Loop
       convert_filename(i) = mypath + buff
       tran_f(i) = convert_filename(i)
       tran_fn = i
    Else
       convert_filename(1) = buff
       tran_f(1) = buff
       tran_fn = 1
    End If
    FileDialog.filename = ""
    If Dir(convert_filename(1)) = "" Then
       GoTo err_exit
    End If
    sinput = tran_f(1)
    cvChoice.Show 1
    StatusBar.Panels(2).Text = " "
    Exit Sub
    
err_exit:
       i = MsgBox("无法打开文件 " + convert_filename(1), 48, "打开文件")
       GoTo open_again
End Sub

Private Sub SUB_122_Click()
     On Error Resume Next
     StatusBar.Panels(2).Text = " 文档管理"
     Menu_Flag = 122
     DocManager.Show 1
     StatusBar.Panels(2).Text = " "
End Sub

Private Sub SUB_123_Click()
    Dim buff As String, mypath As String
    Dim finds As Integer, i As Integer
        
    On Error Resume Next
    Menu_Flag = 123
    StatusBar.Panels(2).Text = " 数据转换"
    For i = 1 To 50
        convert_filename(i) = ""
    Next
open_again:
    FileDialog.DialogTitle = "Tems 扫频测试数据转换"
    FileDialog.Filter = "*.scn Files|*.SCN|All Files|*.*"
    FileDialog.DefaultExt = "*.SCN"
    FileDialog.Flags = &H80000 Or &H200
    Gsm_FileName = Gsm_Path + "\scan"
    FileDialog.InitDir = Gsm_FileName
    FileDialog.ShowOpen
    buff = Trim(FileDialog.filename)
    If buff = "" Then
       StatusBar.Panels(2).Text = " "
       Exit Sub
    End If
    tran_fn = 0
    finds = InStr(buff, Chr(0))
    If finds > 0 Then
       mypath = Left(buff, finds - 1) + "\"
       buff = Trim(Right(buff, Len(buff) - finds))
       finds = InStr(buff, Chr(0))
       i = 1
       Do While finds > 0
          convert_filename(i) = mypath + Left(buff, finds - 1)
          tran_f(i) = convert_filename(i)
          buff = Trim(Right(buff, Len(buff) - finds))
          finds = InStr(buff, Chr(0))
          i = i + 1
       Loop
       convert_filename(i) = mypath + buff
       tran_f(i) = convert_filename(i)
       tran_fn = i
    Else
       convert_filename(1) = buff
       tran_f(1) = buff
       tran_fn = 1
    End If
    FileDialog.filename = ""
    If Dir(convert_filename(1)) = "" Then
       GoTo err_exit
    End If
    sinput = tran_f(1)
    FileDialog.filename = ""
    Menu_Flag = 123
    DocManager.Show 1
    StatusBar.Panels(2).Text = " "
    Exit Sub
    
err_exit:
       i = MsgBox("无法打开文件 " + convert_filename(1), 48, "打开文件")
       GoTo open_again
End Sub

Private Sub SUB_21_Click()
    Dim CM, cm1, TITLE As String
    Dim num As Integer
    
    StatusBar.Panels(2).Text = " 打开文件"
    On Error Resume Next
        
    On Error Resume Next
    If sys = 0 Then
       Gsm_FileName = Gsm_Path + "\normal"
       If Dir(Gsm_FileName, 16) = "" Then
          MkDir Gsm_FileName
       End If
       ChDir Gsm_FileName
    Else
       Gsm_FileName = Gsm_Path + "\scan"
       If Dir(Gsm_FileName, 16) = "" Then
          MkDir Gsm_FileName
       End If
       ChDir Gsm_FileName
    End If
    
    num = Val(mapinfo.eval("NumTables()"))
    
    On Error Resume Next
    Gsm_FileName = Gsm_Path + "\normal"
    ChDir Gsm_FileName
    mapinfo.Do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
    mapinfo.runmenucommand 102

    On Error Resume Next
    TableNum = Val(mapinfo.eval("NumTables()"))
'    mapinfo.do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
'    If TableNum > Num Then
       On Error Resume Next
       CM = mapinfo.eval("tableinfo(0,1)")
       TITLE = MapForm.Caption + "," + CM
       MapForm.Caption = TITLE
    
'       mapinfo.do "fetch Rec 20 from " & CM
'       CM = CM + ".obj"
'       ver_x = Val(mapinfo.eval("centroidx(" & cm1 & ")")) - west
'       ver_y = Val(mapinfo.eval("centroidy(" & cm1 & ")")) - south
'       If ver_y < 0 Or ver_y > yy Or ver_x < 0 Or ver_x > xx Then
'           MsgBox "本软件仅供授权的合法用户使用，欢迎您申请后使用！", 64, "提示"
'           GoTo VER_OUT
'       End If

       On Error Resume Next

       thereIsAMap = True
'Lin       If mapid = 0 Then
          mapid = Val(mapinfo.eval("FrontWindow()"))
'Lin       End If
       SUB_23.Enabled = 1
       SUB_24.Enabled = 1
       SUB_25.Enabled = 1
       SUB_26.Enabled = 1
       SUB_28.Enabled = 1

    StatusBar.Panels(2).Text = "  "
    ChDir Gsm_Path
    Exit Sub
VER_OUT:
    ChDir Gsm_Path
    Exit Sub
Go_OUT:
    ChDir Gsm_Path
End Sub

Private Sub SUB_22_Click()
    On Error Resume Next
    mapinfo.runmenucommand 108
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_23_Click()
    StatusBar.Panels(2).Text = " 关闭文件"
    On Error Resume Next
    mapinfo.runmenucommand 103
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub SUB_24_Click()
    Dim i, m As Integer

    On Error Resume Next
    StatusBar.Panels(2).Text = " 关闭所有文件及窗口"
    If MapGraphflag Then
       Unload frmMapGraph
    End If
    i = Map_No
    While i > 1
       i = i - 1
       Unload ViceMap(i)
    Wend

    mapinfo.runmenucommand 104

    MapForm.Hide
    Unload MapForm

    thereIsAMap = 0
'    SUB_13.Enabled = 1
    SUB_23.Enabled = 0
'    SUB_24.Enabled = 0
    SUB_25.Enabled = 0
    SUB_26.Enabled = 0
    SUB_28.Enabled = 0
    SUB_26.Enabled = 0
    
    USERMARK.Enabled = 0
    CLOSEMARK.Enabled = 0
    SAVEMARK.Enabled = 0
    
    Map_No = 0
'    SUB_151.Enabled = 0
    SUB_41.Enabled = 0
    SUB_461.Enabled = 0
    SUB_462.Enabled = 0
    SUB_463.Enabled = 0
    SUB_464.Enabled = 0
    MnuBcchRetrieve.Enabled = False
    SUB_465.Enabled = 0
    MnuCellReUse.Enabled = False
    SUB_466.Enabled = 0
    SUB_467.Enabled = 0
    SUB_468.Enabled = 0
    FindMyBsic.Enabled = 0
    FindFree.Enabled = 0
    SUB_469.Enabled = 0
    BsNo_FindCell.Enabled = 0
    SUB_4600.Enabled = 0
     
'    OPen_Str_Data.Enabled = 0
'    Static_Pad.Enabled = 0
'    report.Enabled = 0
    SUB_CENTER.Enabled = 0
    Toolbar.Buttons(33).Enabled = False

    StatusBar.Panels(2).Text = " "
    MDIMain.StatusBar.Panels(3).Text = " "
End Sub

Private Sub SUB_25_Click()
    StatusBar.Panels(2).Text = " 保存文件"
    On Error Resume Next
    mapinfo.runmenucommand 105
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub SUB_26_Click()
    StatusBar.Panels(2).Text = " 另存文件"
    Menu_Flag = 26
    On Error Resume Next
    SelTable.Show 1
'    mapinfo.runmenucommand 106
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub SUB_28_Click()
     On Error Resume Next
     StatusBar.Panels(2).Text = " 合并文件"
     Menu_Flag = 1998
     UniteTable.Show 1
     StatusBar.Panels(2).Text = " "
     'mapinfo.runmenucommand 411

End Sub

Private Sub sub_311_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 当前小区场强分析"
    Menu_Flag = 311
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub sub_312_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 当前小区品质分析"
    Menu_Flag = 312
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub sub_313_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 当前小区频率分析"
     Menu_Flag = 300
    ARFCNSelect.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub Sub_314_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "当前小区RxlevSub分析"
    Menu_Flag = 314
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub Sub_315_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "当前小区RxQualSub分析"
    Menu_Flag = 315
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub Sub_316_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " Tx_Power分析"
    Menu_Flag = 316
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub Sub_317_Click()
  On Error Resume Next
    'StatusBar.Panels(2).Text = " Timing Advance分析"
    StatusBar.Panels(2).Text = " 覆盖合理性统计"
    Menu_Flag = 317
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUb_3211_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 相邻小区A场强分析"
    Menu_Flag = 3211
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub
Private Sub sub_3212_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 相邻小区A频率分析"
    Menu_Flag = 3212
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub sub_3221_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 相邻小区B场强分析"
    Menu_Flag = 3221
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub SUB_3222_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 相邻小区B频率分析"
    Menu_Flag = 3222
    SelTable.Show 1

End Sub

Private Sub SUB_3231_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 相邻小区C场强分析"
    Menu_Flag = 3231
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub SUB_3232_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 相邻小区C频率分析"
    Menu_Flag = 3232
    SelTable.Show 1
End Sub
Private Sub SUB_3241_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 相邻小区D场强分析"
    Menu_Flag = 3241
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub SUB_3242_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 相邻小区D频率分析"
    Menu_Flag = 3242
    SelTable.Show 1
End Sub
Private Sub SUB_3251_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 相邻小区E场强分析"
    Menu_Flag = 3251
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub SUB_3252_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 相邻小区E频率分析"
    Menu_Flag = 3252
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub
Private Sub SUB_3261_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 相邻小区F场强分析"
    Menu_Flag = 3261
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "

End Sub

Private Sub SUB_3262_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 相邻小区F频率分析"
    Menu_Flag = 3262
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub
Private Sub SUB_331_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "  SETUP分析"
    Menu_Flag = 331
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_332_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "  CONNECT分析"
    Menu_Flag = 332
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_333_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "  RELEASE分析"
    Menu_Flag = 333
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_334_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "  HANDOVER分析"
    Menu_Flag = 334
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_335_Click()
  On Error Resume Next
 StatusBar.Panels(2).Text = "  LOCATION UPDATE分析"
    Menu_Flag = 335
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_336_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "  IDLE 分析"
    Menu_Flag = 336
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_367_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "  其它信令分析"
    Menu_Flag = 337
    OTHER3.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_41_Click()
  On Error Resume Next
  'dog = scread(62)     'win95
  'dog = (dog / 89) * 3 + 23     'win95
  'If dog <> 326 Then GoTo VERYFY_OUT     'win95
    Dim ApiPack As APIPACKET                      'win95
    Dim portnum%                                  'win95
    Dim status%                                   'win95
    Dim majVer%, minVer%, rev%, drvrType%         'win95
    Dim adr%, datum%                              'win95
    portnum% = 4 ' CPlus-B, port 1                     'win95
    status% = RNBOcplusFormatPacket(ApiPack, 1028)     'win95
    status% = RNBOcplusInitialize(ApiPack, portnum%)   'win95
'    If status <> 0 Then GoTo VERYFY_OUT                'win95
    status% = RNBOcplusGetVersion(ApiPack, majVer%, minVer%, rev%, drvrType%)     'win95
    status% = RNBOcplusGetFullStatus(ApiPack)          'win95
    adr = 62                                           'win95
    status% = RNBOcplusRead(ApiPack, adr%, datum%)     'win95
    datum = (datum / 89) * 4 + 23                      'win95
'    If datum <> 427 Then GoTo VERYFY_OUT               'win95
    adr = 60                                           'win95
    status% = RNBOcplusRead(ApiPack, adr%, datum%)     'win95
    datum = (datum / 89) * 4 + 23                      'win95
'    If datum <> 619 Then GoTo VERYFY_OUT               'win95

  StatusBar.Panels(2).Text = " 小区覆盖区域显示"
  Menu_Flag = 41
  SelTable.Show 1
  StatusBar.Panels(2).Text = ""
  Exit Sub

VERYFY_OUT:
    MsgBox "加密锁错误, 请与珠海万禾公司联系！", 64, "提示"
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub SUB_42_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 确定小区覆盖盲点"
    Menu_Flag = 42
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_441_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " Ncell 对 BCCH 频率碰撞提取"
    Menu_Flag = 441
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_442_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " Ncell 对 TCH 频率碰撞提取"
    Menu_Flag = 442
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_443_Click()
  Dim i As Integer

  On Error Resume Next
  MDIMain.Arrange 0     'CASCADE
  StatusBar.Panels(2).Text = " 多径衰落与干扰趋势"
  i = Val(mapinfo.eval("selectionInfo(3)"))  ' SEL_INFO_NROWS
  If i <> 0 Then
      MapForm.Show
      mapHWnd = Val(mapinfo.eval("WindowInfo(" & mapid & ",12)"))
      If MapForm.WindowState = 1 Or MapForm.WindowState = 2 Then
         MapForm.WindowState = 0
      End If
      MapForm.Move 0, 10, 11920, 3350

      Graphjam.Show
      Graphjam.Move 0, 3350, 11920, 3950
  End If
  StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_45_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 生成质差图层"
    Menu_Flag = 45
    SelTable.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_461_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = "  同频组小区查找"
    Menu_Flag = 461
    Base.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_462_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "  同BSIC小区查找"
    Menu_Flag = 462
    Base.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_463_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "  同LAC小区查找"
    Menu_Flag = 463
    Base.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_464_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "  同频同BSIC小区查找"
    Menu_Flag = 464
    Base.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_465_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "  相邻小区设计检查"
    Menu_Flag = 465
    Base.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_466_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "  邻频查找"
    Menu_Flag = 466
    Base.Show 1
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_467_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 按频率查找小区"
    ARFCN.Show 1
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub SUB_47_Click()
  Dim num As Integer

    StatusBar.Panels(2).Text = " 自定义分析"
    On Error Resume Next
    TableNum = Val(mapinfo.eval("NumTables()"))
    mapinfo.runmenucommand 302
    num = Val(mapinfo.eval("NumTables()"))
    If TableNum <> num Then
       tblname = mapinfo.eval("tableinfo(0,1)")
       mapinfo.Do "Add Map Auto Layer  " & tblname
    End If
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_51_Click()
    StatusBar.Panels(2).Text = " 页面设置"
    On Error Resume Next
    mapinfo.runmenucommand 111
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_52_Click()
    StatusBar.Panels(2).Text = " 页面布局"
    On Error Resume Next
    mapinfo.Do "set style pen  makepen(1,2,0)"

    mapinfo.Do "Set Next Document Parent " & MDIMain.hWnd & " Style 2"
    mapinfo.runmenucommand 604
    StatusBar.Panels(2).Text = "  "
End Sub

Private Sub SUB_531_Click()
    StatusBar.Panels(2).Text = " 打印布局"
    On Error Resume Next
    mapinfo.runmenucommand 112
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub SUB_532_Click()
    StatusBar.Panels(2).Text = " 打印放像结果"
  On Error Resume Next
    Graph.PrintForm
    MapForm.PrintForm
    MsgView.PrintForm
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub SUB_533_Click()
    StatusBar.Panels(2).Text = " 打印放像结果"
  On Error Resume Next
    Graphjam.PrintForm
    MapForm.PrintForm
    MsgView.PrintForm
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub SUB_61_Click()
    StatusBar.Panels(2).Text = "  小区设计数据"
    On Error Resume Next
    MaxBcch = 0
    MinBcch = 0
    Call SUB_24_Click
    Menu_Flag = 61
    PassWord.Show 1
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub SUB_CENTER_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = "  显示中心"
    Menu_Flag = 151
    Center.Show 1
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub SUB_CONFIG_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 地图配置"
    Call SUB_24_Click
    Config_frm.Show 1
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub SYS_MANAGER_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 系统管理"
    Menu_Flag = 62
    PassWord.Show 1
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub Title_Click()
  On Error Resume Next
    StatusBar.Panels(2).Text = " 平铺窗口"
'    MDIMain.Arrange TITLE_HORIZONTAL
    MDIMain.Arrange 2   'TITLE
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub USERMARK_Click()
    On Error Resume Next
    mapinfo.Do "set map redraw off"
    mapinfo.Do "Set Map Layer 0 Editable On  "
    mapinfo.Do "set map redraw on"
End Sub

Private Sub View_Ci_Click()
    On Error Resume Next
    StatusBar.Panels(2).Text = " 服务小区分布观察"
    Menu_Flag = 318
    SelTable.Show 1
    StatusBar.Panels(2).Text = " "
End Sub

Private Sub ViewNcell_Click()
  On Error Resume Next
  StatusBar.Panels(2).Text = "  有效邻小区分布"
  Menu_Flag = 5003
  SelTable.Show 1
  StatusBar.Panels(2).Text = ""
End Sub

Private Sub VMap_Click()
  On Error Resume Next
 StatusBar.Panels(2).Text = " 创建副图"
 If Map_No > 0 And Map_No < 4 Then
'    SUB_24.Enabled = 0
    ReDim ViceMap(Map_No)
    ViceMap(Map_No).Caption = "副本视图：" + MapForm.Caption
    mapinfo.Do "Set Next Document Parent " & ViceMap(Map_No).hWnd & " Style 1"

    On Error Resume Next
    mapinfo.Do "Run Command WindowInfo(" & mapid & ",15)"
    Map_No = Map_No + 1
 End If
 StatusBar.Panels(2).Text = "  "
End Sub

Private Sub Radius_Select()
    Dim WinId As Variant, Layers As Variant
    Dim i As Integer
    Dim Mymsg As String
    Dim MyTableNum As Integer
    
    On Error Resume Next

'       For i = 1 To mapinfo.eval("NumWindows()")
'           If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then
'              WinId = mapinfo.eval("windowid(" & i & ")")
'              If WinId = mapinfo.eval("frontwindow()") Then
'                 Exit For
'              End If
'           End If
'       Next
'       If Val(WinId) = 0 Then
'          StatusBar.Panels(2).Text = " "
'          Exit Sub
'       End If
'       Layers = mapinfo.eval("mapperinfo(" & WinId & ",9)")
'       If Val(Layers) = 0 Then
'          StatusBar.Panels(2).Text = " "
'          Exit Sub
'       End If
       StatusBar.Panels(2).Text = "  半径选择"
       mapinfo.runmenucommand 1703
       GoTo NotChangeLayer
       
       CellLayer = 0
       For i = 1 To Layers
           If UCase(mapinfo.eval("layerinfo(" & WinId & "," & i & ",1)")) = "CELL" Then
              CellLayer = i
              Exit For
           End If
       Next
       If CellLayer = 0 Then
          StatusBar.Panels(2).Text = " "
          Exit Sub
       End If
       Mymsg = "set map order " & Format(CellLayer)
       For i = 2 To Layers
           If i <> CellLayer Then
              Mymsg = Mymsg + "," + Format(i)
           Else
              Mymsg = Mymsg + ",1"
           End If
       Next
       mapinfo.Do Mymsg
NotChangeLayer:
        
        MyTableNum = mapinfo.eval("NumTables()")
        For i = 1 To MyTableNum
            If UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) = "CELL" Then
               GetBcchMaxMin
               Exit For
            End If
        Next
    'End If
End Sub
