VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmHOStat 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "ÇÐ»»Í³¼Æ"
   ClientHeight    =   4965
   ClientLeft      =   3285
   ClientTop       =   1845
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "ËÎÌå"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHOStat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7155
   Begin VB.CommandButton Command2 
      Caption         =   "¹Ø±Õ"
      DragIcon        =   "frmHOStat.frx":000C
      Height          =   320
      Left            =   3075
      TabIndex        =   0
      Top             =   4500
      Width           =   1080
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4080
      Left            =   285
      TabIndex        =   1
      Top             =   225
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   7197
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Ê±¼ä"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "×´Ì¬"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ÇÐ»»ÀàÐÍ"
         Object.Width           =   1164
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Ô´BCCH"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "TN"
         Object.Width           =   212
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Ä¿±êBCCH"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "TN"
         Object.Width           =   212
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "¾àÀë"
         Object.Width           =   1058
      EndProperty
   End
End
Attribute VB_Name = "frmHOStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer, j As Integer
    Dim MyRows As Integer
    Dim HCommand As Boolean, HComplete As Boolean, HCFail As Boolean
    Dim MyMark1 As String
    Dim strHC As String, strHSCHFC As String
    Dim itmX As ListItem
    Dim OldCi As String, NewCi As String
    Dim MyCellName As String
    Dim Oldcelllon As Single, Oldcelllat As Single
    Dim Newcelllon As Single, Newcelllat As Single
    Dim MyTableNum As Integer
    Dim CellIsOpen As Boolean
    Dim MyDistance As Single
    
    On Error Resume Next
    MyTableNum = mapinfo.eval("NumTables()")
    For i = 1 To MyTableNum
        If UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) = "CELL" Then
           CellIsOpen = True
           Exit For
        End If
    Next
    
    MyRows = mapinfo.eval("tableinfo(HOStat,8)")
    mapinfo.do "fetch first from HOStat"
    For i = 1 To MyRows
        If Left(mapinfo.eval("HOStat.mark1"), 3) = "HOA" Then
            MyMark1 = mapinfo.eval("HOStat.mark1")
            HCommand = True
            HComplete = False
            HCFail = False
        ElseIf HCommand Then
            If Left(mapinfo.eval("HOStat.mark1"), 3) = "HOS" Or Left(mapinfo.eval("HOStat.mark1"), 3) = "HOF" Then
                Set itmX = ListView1.ListItems.ADD
                itmX.Text = mapinfo.eval("HOStat.time")
                Select Case Right(Left(MyMark1, InStr(MyMark1, ",") - 1), 1)
                    Case "1"
                        itmX.SubItems(2) = "Ê±Ï¶ÇÐ»»"
                    Case "2"
                        itmX.SubItems(2) = "Ð¡ÇøÇÐ»»"
                    Case "3"
                        itmX.SubItems(2) = "ÏµÍ³ÇÐ»»"
                End Select
                If Left(mapinfo.eval("HOStat.mark1"), 3) = "HOS" Then
                    itmX.SubItems(1) = "³É¹¦"
                Else
                    itmX.SubItems(1) = "Ê§°Ü"
                End If
                
                For j = 3 To 6
                    MyMark1 = Right(MyMark1, Len(MyMark1) - InStr(MyMark1, ","))
                    If j = 6 Then
                        If itmX.SubItems(1) = "Ê§°Ü" Then
                            itmX.SubItems(j) = Trim(MyMark1) & "(³¢ÊÔ)"
                        Else
                            itmX.SubItems(j) = Trim(MyMark1)
                        End If
                    Else
                        If j = 5 And itmX.SubItems(1) = "Ê§°Ü" Then
                            itmX.SubItems(j) = Trim(Left(MyMark1, InStr(MyMark1, ",") - 1)) & "(³¢ÊÔ)"
                        Else
                            itmX.SubItems(j) = Trim(Left(MyMark1, InStr(MyMark1, ",") - 1))
                        End If
                    End If
                Next
        
            Else
                MyMark1 = mapinfo.eval("HOStat.mark1")
                HCommand = False
                For j = 1 To 11
                    MyMark1 = Right(MyMark1, Len(MyMark1) - InStr(MyMark1, ","))
                    If j = 3 Then
                        OldCi = Trim(Left(MyMark1, InStr(MyMark1, ",") - 1))
                    ElseIf j = 11 Then
                        If itmX.SubItems(1) = "Ê§°Ü" Then
                            NewCi = Trim(Left(MyMark1, InStr(MyMark1, ",") - 1))
                        End If
                    End If
                Next
                If CellIsOpen Then
                    Call SearchCellName(0, 0, 0, 0, MyCellName, OldCi, "")
                    Oldcelllon = mapinfo.eval("x1")
                    Oldcelllat = mapinfo.eval("y1")
                    If Oldcelllon <> 0 And Oldcelllat <> 0 Then
                        If itmX.SubItems(1) = "Ê§°Ü" Then
                            itmX.SubItems(7) = "N/A"
                        Else
                            Call SearchCellName(0, 0, 0, 0, MyCellName, NewCi, "")
                            Newcelllon = mapinfo.eval("x1")
                            Newcelllat = mapinfo.eval("y1")
                            If Newcelllon <> 0 And Newcelllat <> 0 Then
                                MyDistance = Sqr((60 * 60 * (Newcelllon - Oldcelllon)) ^ 2 + (60 * 60 * (Newcelllat - Oldcelllat)) ^ 2) * 30
                                itmX.SubItems(7) = Format(MyDistance, "0") & "Ã×"
                            Else
                                itmX.SubItems(7) = "N/A"
                            End If
                        End If
                    Else
                        itmX.SubItems(7) = "N/A"
                    End If
                Else
                    itmX.SubItems(7) = "N/A"
                End If
            End If
        End If
        mapinfo.do "fetch next from HOStat"
    Next
    
End Sub
