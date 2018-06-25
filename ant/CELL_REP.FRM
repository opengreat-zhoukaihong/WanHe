VERSION 4.00
Begin VB.Form Cell_Report_Frm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "小区选择"
   ClientHeight    =   2625
   ClientLeft      =   2025
   ClientTop       =   2955
   ClientWidth     =   3900
   BeginProperty Font 
      name            =   "System"
      charset         =   1
      weight          =   700
      size            =   12
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   3030
   Left            =   1965
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2625
   ScaleWidth      =   3900
   Top             =   2610
   Width           =   4020
   Begin VB.ListBox List1 
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   255
      TabIndex        =   3
      Top             =   495
      Width           =   1995
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取  消"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2505
      TabIndex        =   2
      Top             =   1125
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确  定"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2505
      TabIndex        =   1
      Top             =   525
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "小区名："
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   0
      Top             =   165
      Width           =   960
   End
End
Attribute VB_Name = "Cell_Report_Frm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Integer, j As Integer
    
    On Error Resume Next
    If List1.ListIndex = -1 Then
       j = 1
    Else
       j = List1.ListIndex + 1
    End If
    select_name = List1.List(j - 1)
    For i = 1 To stre_num
        mapinfo.do "select * from " + stre_tab(i) + " where ci_serv = " + Chr(34) + Trim(Rep_Ci(j)) + Chr(34) + " into my_temp" & Format(i)
    Next
    Unload Me
    My_Report
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim i As Integer, j As Integer, p As Integer
    Dim cellci As String * 5, oldci As String * 5
    Dim cellall As Integer
    Dim dtx() As String
    Dim lon() As String, bearing() As String, down() As String
    Dim lat() As String
    Dim str_all As Integer, dtlog As Integer
    Dim mapci As Variant
    
    On Error Resume Next
    
    For i = 1 To stre_num
        mapinfo.do "fetch first from " & stre_tab(i)
    Next
    mapinfo.do "open table " + Chr(34) + Gsm_Path + "\map\cell" + Chr(34)
    
       cellci = space$(5)
       oldci = space$(5)
       cellall = mapinfo.eval("tableinfo(cell,8)")
       ReDim Rcellname(1 To cellall) As String
       ReDim dtx(1 To cellall) As String
       ReDim Rep_Ci(1 To cellall) As String * 5
       ReDim lon(1 To cellall) As String
       ReDim lat(1 To cellall) As String
       ReDim down(1 To cellall) As String
       ReDim bearing(1 To cellall) As String
       RCellNo = 1
       For j = 1 To stre_num
           str_all = mapinfo.eval("tableinfo(" + stre_tab(j) + ",8)")
           mapinfo.do "fetch first from " & stre_tab(j)
           For i = 1 To str_all
               mapci = mapinfo.eval(stre_tab(j) + ".col16")
               mapci = Val(mapci)
               cellci = mapci
              If cellci <> oldci Then
                 If RCellNo > 1 Then
                    For p = 1 To RCellNo - 1
                        If Rep_Ci(p) = cellci Then
                           GoTo ddt_mov_cell
                        End If
                    Next
                 End If
                 Rep_Ci(RCellNo) = cellci
                 Rcellname(RCellNo) = Findcell(cellci)
                 dtx(RCellNo) = mapinfo.eval(stre_tab(j) + ".col40")
                 lon(RCellNo) = mapinfo.eval("cell.lon")
                 lat(RCellNo) = mapinfo.eval("cell.lat")
                 down(RCellNo) = mapinfo.eval("cell.col16")
                 bearing(RCellNo) = mapinfo.eval("cell.col6")
                 dtlog = lencell(Rcellname(RCellNo))
                 dtlog = 14 - dtlog
                 If dtlog < 0 Then dtlog = 0
                 List1.AddItem Rcellname(RCellNo)
                 RCellNo = RCellNo + 1
                 oldci = cellci
              End If
ddt_mov_cell:
              mapinfo.do "fetch next from " & stre_tab(j)
          Next
       Next
       
'    For i = 1 To RCellNo
'        List1.AddItem Rcellname(i)
'    Next

End Sub

