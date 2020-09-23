VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private mvarVer As Byte 'Copia locale.
Private mvarIHL As Byte 'Copia locale.
Private mvarTotalLength As Integer 'Copia locale.

Private mvarIdentification As Integer 'Copia locale.
Private mvarFragmentOffset As Integer 'Copia locale.
Private mvarTTL As Integer 'Copia locale.
Private mvarProtocol As Integer 'Copia locale.
Private mvarsProtocol As String
Private mvarHD_Cheksum As String 'Copia locale.
Private mvarSrcIP As String 'Copia locale.
Private mvarDestIP As String 'Copia locale.
'Private mvarOptions As Long 'Copia locale.

Private Enum IPHDR
eVer = 14
eIHL = 14
etos = 15
etlen1 = 16
etlen2 = 17
eIdentification1 = 18
eIdentification2 = 19
eflags = 20
eFragmentOffset1 = 20
eFragmentOffset2 = 21
eTTL = 22
eproto = 23
ecrc1 = 24
ecrc2 = 25
esaddr = 26
edaddr = 30
eop = 34
epad = 34
End Enum

Public Property Get Options() As Long
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Options
    Options = mvarOptions
End Property





Public Property Get DestIP() As String
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.DestIP
    DestIP = mvarDestIP
End Property




Public Property Get SrcIP() As String
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.SrcIP
    SrcIP = mvarSrcIP
End Property





Public Property Get HD_Cheksum() As String
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.HD_Cheksum
    HD_Cheksum = mvarHD_Cheksum
End Property





Public Property Get Protocol() As Long
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Protocol
    Protocol = mvarProtocol
End Property





Public Property Get TTL() As Long
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.TTL
    TTL = mvarTTL
End Property




Public Property Get FragmentOffset() As Long
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.FragmentOffset
    FragmentOffset = mvarFragmentOffset
End Property





Public Property Get Identification() As Long
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Identification
    Identification = mvarIdentification
End Property





Public Property Get TotalLength() As Long
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.TotalLength
    TotalLength = mvarTotalLength
End Property





Public Property Get IHL() As Byte
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.IHL
    IHL = mvarIHL
End Property




Public Property Get Ver() As Byte
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Ver
    Ver = mvarVer
End Property

Public Function Decode(bff() As Byte)
mvarVer = (bff(IPHDR.eVer) And &HF0) / 16 ' 2^4
mvarIHL = (bff(IPHDR.eIHL) And &HF)
mvarTotalLength = MergeShort(bff(IPHDR.etlen1), bff(IPHDR.etlen2))
mvarIdentification = MergeShort(IPHDR.eIdentification1, IPHDR.eIdentification2)
mvarFLAGS = (bff(IPHDR.eflags) And &HE0) / 32
mvarFragmentOffset = MergeShort(bff(IPHDR.eFragmentOffset1), bff(IPHDR.eFragmentOffset2)) And &H1FFF
mvarTTL = bff(IPHDR.eTTL)
mvarProtocol = bff(IPHDR.eproto)

GetProtocolString

mvarHD_Cheksum = Hex(MergeShort(bff(IPHDR.ecrc1), bff(IPHDR.ecrc2)))
'sostituire sotto con IPfromByte
mvarSrcIP = bff(IPHDR.esaddr) & "." & bff(IPHDR.esaddr + 1) & "." & bff(IPHDR.esaddr + 2) & "." & bff(IPHDR.esaddr + 3)
mvarDestIP = bff(IPHDR.edaddr) & "." & bff(IPHDR.edaddr + 1) & "." & bff(IPHDR.edaddr + 2) & "." & bff(IPHDR.edaddr + 3)

End Function

Private Sub GetProtocolString()
Select Case mvarProtocol
Case 1
mvarsProtocol = "ICMP"
Case 6
mvarsProtocol = "TCP"
Case 17
mvarsProtocol = "UDP"
Case Else
mvarsProtocol = "UNKNOWN"
End Select

End Sub

Public Property Get sProtocol() As String
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Protocol
    sProtocol = mvarsProtocol
End Property
