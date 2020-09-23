VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEthernet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarDestinationMAC As String
Private mvarSourceMAC As String
Private mvarTypeCode As Integer
Private mvarstypecode As String
Private mvarStart As Integer
Private mvarLastByte As Integer
Private mvarLLCDSAP As Byte
Private mvarLLCSNAP As Byte
Private mvarLLCCTRL As Byte

Private Const SOURCE_IP = 26
Private Const DEST_IP = 30
Private Const destmac = 0
Private Const srcmac = 5



Public Property Get LastByte() As Integer
    LastByte = mvarLastByte
End Property

Public Property Let start(ByVal vData As Integer)
    mvarStart = vData
End Property

Public Sub Decode(data() As Byte)

    mvarDestinationMAC = Byte2Hex(data, mvarStart, 5)
    mvarSourceMAC = Byte2Hex(data, mvarStart + 6, 5)
    
  ' CopyMemory mvarTypeCode, data(12), 2
  ' mvarTypeCode = ((data(12) And &HFF) * (2 ^ 8)) Or (data(13) And &HFF)
   mvarTypeCode = MergeShort(data(12), data(13))
   
    mvarLastByte = 13
   Select Case mvarTypeCode
    
    Case Is <= &H5DC 'cmp ,jbe
    mvarstypecode = "[Ethernet IEEE 802.3]"
    mvarLLCDSAP = data(14)
    mvarLLCSNAP = data(15)
    mvarLLCCTRL = data(16)
    
    mvarLastByte = 21 ' 16 + snap OUI + snap Prot ID
    
  If mvarLLCDSAP = mvarLLCSNAP Then
    
    Select Case mvarLLCDSAP
    Case &HF0
     mvarstypecode = mvarstypecode & " [NETBIOS LLC Encapsulation]"
    Case &HFF
     mvarstypecode = mvarstypecode & " [IPX RAW 802.3 Encapsulation]"
    End Select
    
    End If
    Case &H800
    mvarstypecode = "[IPv4]"
    Case &H806
    mvarstypecode = "[Arp]"
     Case &H8035
    mvarstypecode = "[Rarp]"
       Case &H8137
    mvarstypecode = "[Novell Netware IPX]"

    
    Case &H86DD
    mvarstypecode = "[IPv6]"
    Case &H8037
    mvarstypecode = "[IPX]"
    Case &H809B
    mvarstypecode = "[Ethertalk]"
    
    Case &H814C
    mvarstypecode = "[SNMP over Ethernet]"
    Case Else
       mvarstypecode = " Unknown ->TypeCode: " & Hex(mvarTypeCode)
    End Select
    

End Sub

Public Property Get TypeCode() As String
    TypeCode = mvarTypeCode
End Property

Public Property Get SourceMAC() As String
    SourceMAC = mvarSourceMAC
End Property

Public Property Get DestinationMAC() As String
    DestinationMAC = mvarDestinationMAC
End Property

Public Property Get TypeCodeString() As String
    TypeCodeString = mvarstypecode
End Property


