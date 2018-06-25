Attribute VB_Name = "Module1"
Type APIPACKET
  Data(1028) As Byte
End Type

Declare Function RNBOcplusFormatPacket% Lib "sc32w.dll" (ap As APIPACKET, ByVal aplen As Integer)
Declare Function RNBOcplusInitialize% Lib "sc32w.dll" (ap As APIPACKET, ByVal portnum%)
Declare Function RNBOcplusGetFullStatus% Lib "sc32w.dll" (ap As APIPACKET)
Declare Function RNBOcplusGetVersion% Lib "sc32w.dll" (ap As APIPACKET, majv As Integer, minv As Integer, rev As Integer, osdt As Integer)
Declare Function RNBOcplusRead% Lib "sc32w.dll" (ap As APIPACKET, ByVal adr As Integer, datum As Integer)


