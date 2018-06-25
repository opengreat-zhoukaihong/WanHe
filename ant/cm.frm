VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim ApiPack As APIPACKET
    Dim portnum%
    Dim status%
    Dim majVer%, minVer%, rev%, drvrType%
    Dim adr%, datum%
    
    portnum% = 4 ' CPlus-B, port 1
    status% = RNBOcplusFormatPacket(ApiPack, 1028)
    status% = RNBOcplusInitialize(ApiPack, portnum%)
    If status <> 0 Then GoTo VERYFY_OUT
    
    status% = RNBOcplusGetVersion(ApiPack, majVer%, minVer%, rev%, drvrType%)
    status% = RNBOcplusGetFullStatus(ApiPack)
    adr = 62
    status% = RNBOcplusRead(ApiPack, adr%, datum%)
    
    datum = (datum / 89) * 4 + 23
    If datum <> 427 Then GoTo VERYFY_OUT
    
    adr = 60
    status% = RNBOcplusRead(ApiPack, adr%, datum%)
    datum = (datum / 89) * 4 + 23
    If datum <> 619 Then GoTo VERYFY_OUT

End Sub
