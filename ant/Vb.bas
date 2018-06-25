Attribute VB_Name = "GSM_VB"
'****************************************************************************
'
' Declare sub procedures
'
'****************************************************************************

'Declare Function Convert Lib "whdll.dll" (ByVal file1 As String, ByVal file2 As String) As Integer
'Declare Function conv_scan Lib "scandll.dll" Alias "Conv_scan" (ByVal file1 As String, ByVal file2 As String) As Integer
'Declare Function bcapp Lib "bc30rtl.dll" (ByVal file1 As String, ByVal file2 As String) As Integer
'****************************************************************************
'Declare Function MoveWindow Lib "User" (ByVal hwnd As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Integer) As Integer   'win32
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long    'win95
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Declare Sub MessageBeep Lib "User" (ByVal n As Integer)
Declare Sub CloseWindow Lib "User" (ByVal hWnd As Integer)
Public mapinfo As Object, word  As Object
Public ViceMap() As New ViceMapForm
Public msg, tblname, USERNAME As String, Auther, Msg_3_Layer As String
Public thereIsAMap, TableNum, Menu_Flag, XY_flag, Over, DisFlag As Integer
Public mapid As Long
Public legendid As Long
Public MessageId1 As Long
Public MessageId2 As Long
Public ver_y, west, south, yy, xx As Double
Public sinput, soutput As String
Public Face_show, Map_No, sys, GPS_NO, Data_Tran_Flag As Integer
Public Gsm_Path As String, Gsm_FileName As String, Gsm_File2 As String
Public UpdateFileName As String
Public M2_Local As Boolean
Public SearchDistance As Integer    'ͬƵ��Ƶ

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public myCallback As Object
Public dis_flag As Integer
Public NcellWinFlag As Boolean

Option Explicit

Public CellLayer As Integer
Public DATA_NO, Beg_Rec, End_Rec, Full_Flag  As Integer
Public SelTbl, Sel_Street, mscode, str_len, str_name, Bmp_Name As String

Public Type Record ' Define user-defined type.
    Name As String * 30
    Pass As String * 6
    exchange As String * 1
    Antenna As String * 5
End Type

Public Type doc
           DOCNAME                  As String * 30
           GPS                      As String * 10
           DATE                     As String * 10
           Partner                  As String * 8
           IMG                      As String * 12
           TESTOBJECT               As String * 1
           TESTDIST                 As String * 1
           TESTBACK                 As String * 1
           WEATHER                  As String * 1
'** Total **                                74
End Type

Type field
     londbf As String * 10
     latdbf As String * 10
     timedbf As String * 11
     coldbf(1 To 30) As String * 4
End Type

Type ScanHead
     ver As Byte
     year As Byte
     month As Byte
     day As Byte
     recordno As Long
     HeaderLen As Integer
     RecordLen As Integer
     Zero As String * 20
End Type
Type WriteField
     Name As String * 11
     Type As String * 1
     Pos As Long
     length As Byte
     Dec As Byte
     Zero As String * 14
End Type

Type MessageType
    RecordTime As String
    RecordFrame As String
    RecordMessage As String
End Type

'Public strNcellChinese(2) As String
Public Const LB_SETANCHORINDEX = &H19C

Public EditFrmFlag As Byte
Public varBookmark(2) As Variant
Public EditFlag(2) As Boolean
Public SortType As Byte
Public CellFileName As String
Public StatString As String
Public MyRndColor(375) As Long
Public MyCellRndColor(124) As Long
Public MyLacColor(100) As Long
Public MyBcchColor(16) As Long
Public ShowValueFlag As Boolean
Public CheckValue(2) As Boolean
Public GSMDCSBCCH As Byte
Public MapGraphflag As Boolean
Public SelBcchGroup As Integer
Public CMIsCDD As Boolean
Public HOParaFlag As Byte
Public MyNRSelCellName As String, MyNRSelCellCI As String, MyNRSelCellLac As String
Public MyNRSelCellBcch As Integer, MyNRSelCellBsic As Integer
Public MyNRSelCellName_2 As String, MyNRSelCellCI_2 As String
Public MyNRSelCellBcch_2 As Integer, MyNRSelCellBsic_2 As Integer
Public MyNRSelCellName_3 As String, MyNRSelCellCI_3 As String
Public MyNRSelCellBcch_3 As Integer, MyNRSelCellBsic_3 As Integer
Public Linyujin As Integer
Public Xiaoyu_Color(4) As Long
Public frmNcellRS_2 As New frmNcellRS
Public frmNcellRS_3 As New frmNcellRS
Public CurrentNcellRS As Byte

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0

'Global Const BLACK = 0
'Global Const WHITE = 16777215
'Global Const RED = 16711680
'Global Const GREEN = 65280
'Global Const BLUE = 255
'Global Const CYAN = 65535
'Global Const MAGENTA = 16711935
'Global Const YELLOW = 16776960
